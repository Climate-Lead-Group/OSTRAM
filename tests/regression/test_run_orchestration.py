from __future__ import annotations

import argparse
import ast
import importlib.util
import io
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import contextmanager, redirect_stderr, redirect_stdout
from itertools import product
from pathlib import Path
from unittest import mock

from ostram.pipeline.compilation import orchestrator as b1_orchestrator


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
RUN_PATH = REPO_ROOT / "ostram" / "pipeline" / "orchestration.py"


def _load_launcher(label: str):
    module_name = f"_ostram_run_orchestration_{label}"
    spec = importlib.util.spec_from_file_location(module_name, RUN_PATH)
    if spec is None or spec.loader is None:
        raise AssertionError(f"could not load module spec for {RUN_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    finally:
        sys.modules.pop(module_name, None)
    return module


@contextmanager
def _working_directory(path: Path):
    previous = Path.cwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(previous)


class LauncherHarness:
    def __init__(
        self,
        launcher,
        *,
        snapshot_exists: bool = True,
        active_scenarios: tuple[str, ...] = ("A", "B", "C"),
        dvc_remote: bool = False,
        guessed_env: str | None = "yaml-env",
        pipeline_failure: tuple[str, int] | None = None,
    ) -> None:
        self.launcher = launcher
        self.snapshot_exists = snapshot_exists
        self.active_scenarios = list(active_scenarios)
        self.dvc_remote = dvc_remote
        self.guessed_env = guessed_env
        self.pipeline_failure = pipeline_failure
        self.events: list[tuple[object, ...]] = []
        self.patches: list[mock._patch] = []

    def _record(self, name: str, *args: object) -> None:
        self.events.append((name, *args, Path.cwd()))

    def _guess_env(self, env_file: str) -> str | None:
        self._record("guess_env", env_file)
        return self.guessed_env

    def _check_tool(self, tool: str) -> None:
        self._record("check_tool", tool)

    def _create_env(self, env_name: str, env_file: str) -> None:
        self._record("create_env", env_name, env_file)

    def _ensure_deps(self, env_name: str) -> None:
        self._record("ensure_deps", env_name)

    def _ensure_dvc(self, env_name: str) -> None:
        self._record("ensure_dvc", env_name)

    def _has_remote(self, env_name: str) -> bool:
        self._record("has_dvc_remote", env_name)
        return self.dvc_remote

    def _dvc_command(self, env_name: str, args: str) -> None:
        self._record("dvc_command", env_name, args)

    def _snapshot(self, path: Path, roots: tuple[str, ...]) -> bool:
        self._record("snapshot_exists", path, roots)
        return self.snapshot_exists

    def _load_registry(self):
        active = tuple(self.active_scenarios)

        class Registry:
            def select(self, requested):
                if requested is None or requested == "":
                    return active
                names = [
                    name.strip()
                    for name in requested.split(",")
                    if name.strip()
                ]
                duplicates = sorted(
                    {name for name in names if names.count(name) > 1}
                )
                if duplicates:
                    raise ValueError(
                        f"duplicate scenario selection: {duplicates}"
                    )
                unknown = [name for name in names if name not in active]
                if unknown:
                    raise ValueError(
                        f"unknown scenario selection: {unknown}"
                    )
                selected = set(names)
                return tuple(name for name in active if name in selected)

            def required_roots(self, selected):
                return tuple(selected)

        return Registry()

    def _pipeline(self, env_name: str, script: Path, extra_args: str = "") -> None:
        self._record("pipeline", env_name, script, extra_args)
        if self.pipeline_failure is not None and self.pipeline_failure[0] == script.name:
            raise subprocess.CalledProcessError(
                self.pipeline_failure[1], f"recorded {script.name}"
            )

    def _a3(
        self,
        env_name: str,
        script: Path,
        scenarios: list[str],
        seed: str | None,
    ) -> None:
        for scenario in scenarios:
            self._record("a3", env_name, script, scenario, seed)

    class _Reporter:
        @contextmanager
        def capture_output(self):
            yield

        def note(self, *_args, **_kwargs):
            return None

        def stage_start(self, *_args, **_kwargs):
            return None

        def stage_complete(self, *_args, **_kwargs):
            return None

        def stage_skip(self, *_args, **_kwargs):
            return None

        def stage_fail(self, *_args, **_kwargs):
            return None

        def finish(self, *_args, **_kwargs):
            return None

    def _reporter(self, *_args, **_kwargs):
        return self._Reporter()

    def __enter__(self):
        replacements = {
            "guess_env_name_from_yaml": self._guess_env,
            "check_tool_available": self._check_tool,
            "create_env_if_missing": self._create_env,
            "ensure_deps": self._ensure_deps,
            "ensure_dvc_repo": self._ensure_dvc,
            "has_dvc_remote": self._has_remote,
            "dvc_command": self._dvc_command,
            "root_snapshots_exist": self._snapshot,
            "load_registry": self._load_registry,
            "ensure_root_output_directories": lambda *_: None,
            "run_pipeline_script": self._pipeline,
            "run_a3_for_scenarios": self._a3,
            "_create_run_reporter": self._reporter,
        }
        self.patches = [
            mock.patch.object(self.launcher, name, replacement)
            for name, replacement in replacements.items()
        ]
        for patch in self.patches:
            patch.start()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        for patch in reversed(self.patches):
            patch.stop()


def _event_names(events: list[tuple[object, ...]]) -> list[str]:
    return [str(event[0]) for event in events]


def _tree_state(root: Path) -> tuple[tuple[str, str, bytes | None], ...]:
    entries: list[tuple[str, str, bytes | None]] = []
    for path in sorted(root.rglob("*"), key=lambda item: item.relative_to(root).as_posix()):
        relative = path.relative_to(root).as_posix()
        if path.is_dir():
            entries.append((relative, "directory", None))
        elif path.is_file():
            entries.append((relative, "file", path.read_bytes()))
    return tuple(entries)


class ImportAndCliCharacterizationTests(unittest.TestCase):
    def test_import_has_no_process_or_pipeline_side_effects(self) -> None:
        with (
            mock.patch.object(subprocess, "check_call") as check_call,
            mock.patch.object(subprocess, "check_output") as check_output,
            mock.patch.object(subprocess, "run") as process_run,
            redirect_stdout(io.StringIO()) as stdout,
            redirect_stderr(io.StringIO()) as stderr,
        ):
            launcher = _load_launcher("import_safety")

        check_call.assert_not_called()
        check_output.assert_not_called()
        process_run.assert_not_called()
        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")
        self.assertTrue(callable(launcher.main))

    def test_parse_args_accepts_explicit_argv_and_preserves_cli_contract(self) -> None:
        launcher = _load_launcher("parse_args")
        with mock.patch.object(sys, "argv", ["python -m ostram run", "--not-an-option"]):
            defaults = launcher.parse_args([])
            explicit = launcher.parse_args(
                [
                    "--env-name",
                    "custom env",
                    "--env-file",
                    "config/custom environment.yaml",
                    "--dvc-file",
                    "config/custom dvc.yaml",
                    "--skip-pull",
                    "--skip-a3",
                    "--skip-b1",
                    "--skip-b2",
                    "--scenarios",
                    " C , A, C ",
                ]
            )

        self.assertEqual(
            vars(defaults),
            {
                "env_name": None,
                "env_file": launcher.ENV_FILE_DEFAULT,
                "dvc_file": launcher.DVC_FILE_DEFAULT,
                "skip_pull": False,
                "skip_a3": False,
                "skip_b1": False,
                "skip_b2": False,
                "scenarios": None,
                "a_result_seed": None,
                "compile_only": False,
                "verbose": False,
            },
        )
        self.assertEqual(
            vars(explicit),
            {
                "env_name": "custom env",
                "env_file": "config/custom environment.yaml",
                "dvc_file": "config/custom dvc.yaml",
                "skip_pull": True,
                "skip_a3": True,
                "skip_b1": True,
                "skip_b2": True,
                "scenarios": " C , A, C ",
                "a_result_seed": None,
                "compile_only": False,
                "verbose": False,
            },
        )

    def test_defaults_and_free_form_cli_values_are_forwarded_exactly(self) -> None:
        launcher = _load_launcher("cli_values")
        argv = [
            "python -m ostram run",
            "--env-name",
            "custom env",
            "--env-file",
            "config/custom environment.yaml",
            "--dvc-file",
            "config/custom dvc.yaml",
            "--skip-pull",
            "--skip-a3",
            "--skip-b1",
            "--skip-b2",
        ]
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            cwd = Path(temp).resolve()
            with (
                _working_directory(cwd),
                LauncherHarness(launcher) as harness,
                mock.patch.object(sys, "argv", argv),
                redirect_stdout(io.StringIO()) as stdout,
            ):
                result = launcher.main()

        self.assertIsNone(result)
        self.assertEqual(
            _event_names(harness.events),
            [
                "check_tool",
                "create_env",
                "ensure_deps",
                "snapshot_exists",
            ],
        )
        self.assertEqual(harness.events[0], ("check_tool", "conda", cwd))
        self.assertEqual(
            harness.events[1],
            (
                "create_env",
                "custom env",
                REPO_ROOT / "config" / "custom environment.yaml",
                cwd,
            ),
        )
        self.assertEqual(
            harness.events[-1],
            (
                "snapshot_exists",
                REPO_ROOT / "workspace" / "preparation" / "A1_Outputs",
                ("A", "B", "C"),
                cwd,
            ),
        )
        output = stdout.getvalue()
        self.assertIn("Using environment: custom env", output)
        self.assertIn(
            f"DVC config: {REPO_ROOT / 'config' / 'custom dvc.yaml'}",
            output,
        )
        self.assertIn(
            "Skipping DVC repository setup and `dvc pull` by request.",
            output,
        )

    def test_default_environment_resolution_order(self) -> None:
        launcher = _load_launcher("env_defaults")
        cases = (
            ("from-yaml", "from-yaml"),
            (None, launcher.ENV_NAME_DEFAULT),
        )
        for guessed, expected in cases:
            with self.subTest(guessed=guessed):
                with (
                    LauncherHarness(launcher, guessed_env=guessed) as harness,
                    mock.patch.object(sys, "argv", ["python -m ostram run", "--skip-pull", "--skip-a3", "--skip-b1", "--skip-b2"]),
                    redirect_stdout(io.StringIO()),
                ):
                    launcher.main()

                self.assertEqual(
                    harness.events[0][0:2],
                    ("guess_env", REPO_ROOT / "environment.yaml"),
                )
                self.assertEqual(
                    harness.events[2][0:3],
                    ("create_env", expected, REPO_ROOT / "environment.yaml"),
                )

    def test_unknown_option_and_help_use_argparse_exit_codes_before_setup(self) -> None:
        launcher = _load_launcher("argparse_exits")
        cases = ((["python -m ostram run", "--not-an-option"], 2), (["python -m ostram run", "--help"], 0))
        for argv, expected_code in cases:
            with self.subTest(argv=argv):
                with (
                    mock.patch.object(launcher, "check_tool_available") as check_tool,
                    mock.patch.object(sys, "argv", argv),
                    redirect_stdout(io.StringIO()),
                    redirect_stderr(io.StringIO()),
                ):
                    with self.assertRaises(SystemExit) as raised:
                        launcher.main()
                self.assertEqual(raised.exception.code, expected_code)
                check_tool.assert_not_called()

    def test_main_guard_maps_process_and_other_failures_to_current_exit_codes(self) -> None:
        tree = ast.parse(RUN_PATH.read_text(encoding="utf-8-sig"), filename="python -m ostram run")
        guard = next(
            node
            for node in tree.body
            if isinstance(node, ast.If)
            and isinstance(node.test, ast.Compare)
            and isinstance(node.test.left, ast.Name)
            and node.test.left.id == "__name__"
        )
        guard_source = ast.get_source_segment(
            RUN_PATH.read_text(encoding="utf-8-sig"), guard
        )
        self.assertIn("main()", guard_source)
        self.assertIn("except subprocess.CalledProcessError as e:", guard_source)
        self.assertIn("sys.exit(e.returncode)", guard_source)
        self.assertIn("except Exception as e:", guard_source)
        self.assertIn("sys.exit(1)", guard_source)


class StageAndScenarioCharacterizationTests(unittest.TestCase):
    def test_every_skip_flag_combination_preserves_selected_stage_order(self) -> None:
        launcher = _load_launcher("skip_matrix")
        for skip_a3, skip_b1, skip_b2 in product((False, True), repeat=3):
            with self.subTest(skip_a3=skip_a3, skip_b1=skip_b1, skip_b2=skip_b2):
                argv = ["python -m ostram run", "--skip-pull"]
                if skip_a3:
                    argv.append("--skip-a3")
                if skip_b1:
                    argv.append("--skip-b1")
                if skip_b2:
                    argv.append("--skip-b2")
                with (
                    LauncherHarness(launcher, active_scenarios=("A", "B")) as harness,
                    mock.patch.object(sys, "argv", argv),
                    redirect_stdout(io.StringIO()),
                ):
                    launcher.main()

                stage_events = [
                    event[0]
                    for event in harness.events
                    if event[0] in {"enumerate_active", "a3", "pipeline"}
                ]
                expected: list[str] = []
                if not skip_a3:
                    expected.extend(["a3", "a3"])
                if not skip_b1:
                    expected.append("pipeline")
                if not skip_b2:
                    expected.append("pipeline")
                self.assertEqual(stage_events, expected)

                pipeline_names = [
                    event[2].name for event in harness.events if event[0] == "pipeline"
                ]
                expected_names: list[str] = []
                if not skip_b1:
                    expected_names.append("runner.py")
                if not skip_b2:
                    expected_names.append("runner.py")
                self.assertEqual(pipeline_names, expected_names)

    def test_a1_a2_snapshot_gate_is_independent_of_all_skip_flags(self) -> None:
        launcher = _load_launcher("snapshot_gate")
        argv = [
            "python -m ostram run",
            "--skip-pull",
            "--skip-a3",
            "--skip-b1",
            "--skip-b2",
        ]
        for snapshot_exists, expected_scripts in (
            (True, []),
            (False, ["base_inputs.py", "transmission.py"]),
        ):
            with self.subTest(snapshot_exists=snapshot_exists):
                with (
                    LauncherHarness(launcher, snapshot_exists=snapshot_exists) as harness,
                    mock.patch.object(sys, "argv", argv),
                    redirect_stdout(io.StringIO()),
                ):
                    launcher.main()
                scripts = [
                    event[2].name for event in harness.events if event[0] == "pipeline"
                ]
                self.assertEqual(scripts, expected_scripts)

    def test_scenario_filter_uses_one_canonical_order_for_a3_b1_and_b2(self) -> None:
        launcher = _load_launcher("scenario_order")
        with (
            LauncherHarness(launcher, active_scenarios=("A", "B", "C")) as harness,
            mock.patch.object(
                sys, "argv", ["python -m ostram run", "--skip-pull", "--scenarios", " C , A "]
            ),
            redirect_stdout(io.StringIO()),
        ):
            launcher.main()

        self.assertEqual(
            [event[3] for event in harness.events if event[0] == "a3"], ["A", "C"]
        )
        self.assertEqual(
            [event[3] for event in harness.events if event[0] == "pipeline"],
            [["--scenarios", "A,C"], ["--scenarios", "A,C"]],
        )

    def test_a3_filter_does_not_auto_add_the_bau_prerequisite(self) -> None:
        launcher = _load_launcher("no_prerequisite_closure")
        active = (
            "BAU",
            "A_Calibrated_BAU",
            "B_Optimised_VRE",
            "C_Target_VRE",
        )
        with (
            LauncherHarness(launcher, active_scenarios=active) as harness,
            mock.patch.object(
                sys,
                "argv",
                [
                    "python -m ostram run",
                    "--skip-pull",
                    "--scenarios",
                    "A_Calibrated_BAU",
                ],
            ),
            redirect_stdout(io.StringIO()),
        ):
            launcher.main()

        self.assertEqual(
            [event[3] for event in harness.events if event[0] == "a3"],
            ["A_Calibrated_BAU"],
        )

    def test_duplicate_scenario_selection_fails_before_any_stage(self) -> None:
        launcher = _load_launcher("unknown_duplicates")
        with (
            LauncherHarness(launcher, active_scenarios=("BAU", "A")) as harness,
            mock.patch.object(
                sys,
                "argv",
                [
                    "python -m ostram run",
                    "--skip-pull",
                    "--scenarios",
                    "Missing,A,Missing,Other",
                ],
            ),
            redirect_stdout(io.StringIO()),
        ):
            with self.assertRaisesRegex(
                RuntimeError,
                r"duplicate scenario selection: \['Missing'\]",
            ):
                launcher.main()

        self.assertEqual([event for event in harness.events if event[0] == "a3"], [])
        self.assertEqual(
            [event for event in harness.events if event[0] == "pipeline"],
            [],
        )

    def test_empty_scenario_filter_fails_before_all_pipeline_stages(self) -> None:
        launcher = _load_launcher("empty_scenarios")
        with (
            LauncherHarness(launcher, active_scenarios=("A", "B")) as harness,
            mock.patch.object(
                sys, "argv", ["python -m ostram run", "--skip-pull", "--scenarios", ", , "]
            ),
            redirect_stdout(io.StringIO()),
        ):
            with self.assertRaisesRegex(RuntimeError, "selected no"):
                launcher.main()

        self.assertEqual([event for event in harness.events if event[0] == "a3"], [])
        self.assertEqual(
            [event for event in harness.events if event[0] == "pipeline"], []
        )

    def test_explicit_empty_string_uses_default_but_whitespace_fails(self) -> None:
        launcher = _load_launcher("empty_string_distinction")
        for raw, expected_a3, should_fail in (
            ("", ["A", "B"], False),
            ("   ", [], True),
        ):
            with self.subTest(raw=raw):
                with (
                    LauncherHarness(
                        launcher,
                        active_scenarios=("A", "B"),
                    ) as harness,
                    mock.patch.object(
                        sys,
                        "argv",
                        ["python -m ostram run", "--skip-pull", "--scenarios", raw],
                    ),
                    redirect_stdout(io.StringIO()),
                ):
                    if should_fail:
                        with self.assertRaisesRegex(RuntimeError, "selected no"):
                            launcher.main()
                    else:
                        launcher.main()

                self.assertEqual(
                    [event[3] for event in harness.events if event[0] == "a3"],
                    expected_a3,
                )
                self.assertEqual(
                    [event[3] for event in harness.events if event[0] == "pipeline"],
                    []
                    if should_fail
                    else [["--scenarios", "A,B"], ["--scenarios", "A,B"]],
                )

    def test_unknown_scenario_is_rejected_before_a1_a2_a3_b1_b2(self) -> None:
        launcher = _load_launcher("unknown_scenario")
        with (
            LauncherHarness(
                launcher, snapshot_exists=False, active_scenarios=("A", "B")
            ) as harness,
            mock.patch.object(
                sys, "argv", ["python -m ostram run", "--skip-pull", "--scenarios", "Missing"]
            ),
            redirect_stdout(io.StringIO()),
        ):
            with self.assertRaisesRegex(RuntimeError, "unknown scenario selection"):
                launcher.main()

        pipeline_names = [
            event[2].name for event in harness.events if event[0] == "pipeline"
        ]
        self.assertEqual(pipeline_names, [])
        self.assertEqual([event for event in harness.events if event[0] == "a3"], [])

    def test_skip_a3_still_validates_the_shared_scenario_contract(self) -> None:
        launcher = _load_launcher("skip_a3_validation")
        with (
            LauncherHarness(launcher, active_scenarios=("A", "B")) as harness,
            mock.patch.object(
                sys,
                "argv",
                ["python -m ostram run", "--skip-pull", "--skip-a3", "--scenarios", "Missing"],
            ),
            redirect_stdout(io.StringIO()),
        ):
            with self.assertRaisesRegex(RuntimeError, "unknown scenario selection"):
                launcher.main()

        self.assertNotIn("enumerate_active", _event_names(harness.events))
        self.assertEqual(
            [event for event in harness.events if event[0] == "pipeline"],
            [],
        )

    def test_skip_pull_actual_route_keeps_fresh_explicit_project_dvc_free(self) -> None:
        import ostram.__main__ as canonical_cli

        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            fixture_root = Path(temp).resolve()
            project_root = fixture_root / "fresh project"
            external_workspace = fixture_root / "external workspace"
            caller_cwd = fixture_root / "caller cwd"
            invalid_project_root = fixture_root / "invalid environment project"
            invalid_workspace = fixture_root / "invalid environment workspace"

            for directory in (
                project_root / ".git",
                project_root / "ostram",
                project_root / "inputs",
                project_root / "config",
                project_root / "model",
                caller_cwd,
            ):
                directory.mkdir(parents=True, exist_ok=True)
            (project_root / "ostram" / "__init__.py").write_text(
                '"""Fresh project fixture."""\n',
                encoding="utf-8",
            )
            (project_root / "environment.yaml").write_text(
                "name: fixture-env\n",
                encoding="utf-8",
            )
            (project_root / "dvc.yaml").write_text("stages: {}\n", encoding="utf-8")

            self.assertFalse((project_root / ".dvc").exists())
            self.assertFalse(external_workspace.exists())
            project_before = _tree_state(project_root)

            with mock.patch.dict(
                os.environ,
                {
                    "OSTRAM_PROJECT_ROOT": str(project_root),
                    "OSTRAM_WORKSPACE": str(external_workspace),
                },
                clear=False,
            ):
                launcher = _load_launcher("fresh_explicit_skip_pull")

            with (
                _working_directory(caller_cwd),
                LauncherHarness(
                    launcher,
                    active_scenarios=("BAU",),
                ) as harness,
                mock.patch.dict(
                    os.environ,
                    {
                        "OSTRAM_PROJECT_ROOT": str(invalid_project_root),
                        "OSTRAM_WORKSPACE": str(invalid_workspace),
                    },
                    clear=False,
                ),
                mock.patch.object(
                    canonical_cli,
                    "_load_route_module",
                    return_value=launcher,
                ),
                mock.patch.object(launcher.subprocess, "check_call") as check_call,
                mock.patch.object(launcher.subprocess, "check_output") as check_output,
                redirect_stdout(io.StringIO()) as stdout,
            ):
                result = canonical_cli.main(
                    [
                        "--project-root",
                        str(project_root),
                        "--workspace",
                        str(external_workspace),
                        "run",
                        "--skip-pull",
                        "--skip-b2",
                        "--scenarios",
                        "BAU",
                    ]
                )

            self.assertEqual(result, 0)
            self.assertNotIn("ensure_dvc", _event_names(harness.events))
            self.assertNotIn("has_dvc_remote", _event_names(harness.events))
            self.assertNotIn("dvc_command", _event_names(harness.events))
            check_call.assert_not_called()
            check_output.assert_not_called()

            create_env = next(event for event in harness.events if event[0] == "create_env")
            self.assertEqual(create_env[2], project_root / "environment.yaml")
            snapshot = next(event for event in harness.events if event[0] == "snapshot_exists")
            self.assertEqual(
                snapshot[1],
                external_workspace / "preparation" / "A1_Outputs",
            )
            self.assertEqual(
                [event[0] for event in harness.events if event[0] in {"a3", "pipeline"}],
                ["a3", "pipeline"],
            )
            self.assertEqual(
                next(event for event in harness.events if event[0] == "a3")[2],
                project_root / "ostram" / "pipeline" / "scenarios" / "materializer.py",
            )
            self.assertEqual(
                next(event for event in harness.events if event[0] == "pipeline")[2],
                project_root / "ostram" / "pipeline" / "compilation" / "runner.py",
            )
            self.assertIn(
                "Skipping DVC repository setup and `dvc pull` by request.",
                stdout.getvalue(),
            )
            self.assertEqual(_tree_state(project_root), project_before)
            self.assertFalse((project_root / ".dvc").exists())
            self.assertFalse(external_workspace.exists())

    def test_dvc_remote_check_and_pull_selection(self) -> None:
        launcher = _load_launcher("dvc_pull")
        for remote, expected_tail in (
            (False, ["ensure_dvc", "has_dvc_remote"]),
            (True, ["ensure_dvc", "has_dvc_remote", "dvc_command"]),
        ):
            with self.subTest(remote=remote):
                argv = ["python -m ostram run", "--skip-a3", "--skip-b1", "--skip-b2"]
                with (
                    LauncherHarness(launcher, dvc_remote=remote) as harness,
                    mock.patch.object(sys, "argv", argv),
                    redirect_stdout(io.StringIO()),
                ):
                    launcher.main()
                dvc_events = [
                    name
                    for name in _event_names(harness.events)
                    if name in {"ensure_dvc", "has_dvc_remote", "dvc_command"}
                ]
                self.assertEqual(dvc_events, expected_tail)
                if remote:
                    event = next(e for e in harness.events if e[0] == "dvc_command")
                    self.assertEqual(event[1:3], ("yaml-env", "pull"))

    def test_pipeline_failure_propagates_exactly_and_stops_later_stages(self) -> None:
        launcher = _load_launcher("failure_stop")
        failure = ("transmission.py", 7)
        with (
            LauncherHarness(
                launcher, snapshot_exists=False, pipeline_failure=failure
            ) as harness,
            mock.patch.object(sys, "argv", ["python -m ostram run", "--skip-pull"]),
            redirect_stdout(io.StringIO()),
        ):
            with self.assertRaises(subprocess.CalledProcessError) as raised:
                launcher.main()

        self.assertEqual(raised.exception.returncode, 7)
        self.assertEqual(
            [event[2].name for event in harness.events if event[0] == "pipeline"],
            ["base_inputs.py", "transmission.py"],
        )
        self.assertNotIn("enumerate_active", _event_names(harness.events))

    def test_canonical_outer_route_propagates_b1_child_failure_and_skips_b2(
        self,
    ) -> None:
        import ostram.__main__ as canonical_cli

        launcher = _load_launcher("b1_child_failure")
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            fixture_root = Path(temp).resolve()
            workspace = fixture_root / "workspace"
            script_dir = fixture_root / "compilation"
            scenario_root = fixture_root / "A1_Outputs"
            scenario_root.mkdir(parents=True)
            (scenario_root / "A1_Outputs_A").mkdir()
            script_dir.mkdir()
            config = script_dir / "Config_MOMF_T1_A.yaml"
            original_config = b"xtra_scen:\r\n  Main_Scenario: ORIGINAL\r\n"
            config.write_bytes(original_config)
            compiler = script_dir / "compiler.py"
            compiler.write_text("# controlled compiler fixture\n", encoding="utf-8")
            b1_paths = b1_orchestrator.B1Paths(
                script_dir=script_dir,
                config_path=config,
                compiler_path=compiler,
                scenarios_root=scenario_root,
            )
            stage_calls: list[Path] = []

            def run_stage(_env_name, script, _extra_args=()):
                stage_calls.append(script)
                if script == launcher.B1_SCRIPT_DEFAULT:
                    return b1_orchestrator.orchestrate(
                        argparse.Namespace(scenarios="A"),
                        b1_paths,
                        compiler_runner=lambda _path: 1,
                    )
                raise AssertionError(f"B2 was called after B1 failed: {script}")

            with (
                LauncherHarness(
                    launcher,
                    active_scenarios=("A",),
                ),
                mock.patch.object(
                    launcher,
                    "run_pipeline_script",
                    side_effect=run_stage,
                ),
                mock.patch.object(
                    canonical_cli,
                    "_load_route_module",
                    return_value=launcher,
                ),
                redirect_stdout(io.StringIO()) as stdout,
                redirect_stderr(io.StringIO()) as stderr,
            ):
                result = canonical_cli.main(
                    [
                        "--project-root",
                        str(REPO_ROOT),
                        "--workspace",
                        str(workspace),
                        "run",
                        "--skip-pull",
                        "--skip-a3",
                        "--scenarios",
                        "A",
                    ]
                )

            self.assertEqual(result, 1)
            self.assertEqual(stage_calls, [launcher.B1_SCRIPT_DEFAULT])
            self.assertNotIn(launcher.B2_SCRIPT_DEFAULT, stage_calls)
            self.assertIn(
                "B1_Compiler.py exited with code 1 for scenario 'A'",
                stdout.getvalue(),
            )
            self.assertNotIn("Pipeline completed", stdout.getvalue())
            self.assertIn("Command failed (exit 1)", stderr.getvalue())
            self.assertIn(str(compiler), stderr.getvalue())
            self.assertEqual(config.read_bytes(), original_config)


class CommandBoundaryCharacterizationTests(unittest.TestCase):
    def test_run_pins_project_pythonpath_and_inherits_cwd(self) -> None:
        launcher = _load_launcher("run_env")
        cwd = Path.cwd()
        with (
            mock.patch.dict(
                launcher.os.environ,
                {"EXISTING": "kept", "PYTHONHASHSEED": "random"},
                clear=True,
            ),
            mock.patch.object(launcher.subprocess, "check_call") as check_call,
        ):
            result = launcher.run("tool --flag value")

        self.assertIsNone(result)
        check_call.assert_called_once_with(
            ["tool", "--flag", "value"],
            cwd=None,
            env={
                "EXISTING": "kept",
                "PYTHONHASHSEED": "0",
                "PYTHONDONTWRITEBYTECODE": "1",
                "PYTHONPATH": str(REPO_ROOT),
            },
        )
        self.assertEqual(Path.cwd(), cwd)

    def test_run_pipeline_and_a3_construct_module_commands(self) -> None:
        launcher = _load_launcher("command_strings")
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            cwd = Path(temp).resolve()
            pipeline_script = cwd / "ostram" / "pipeline" / "compilation" / "runner.py"
            pipeline_script.parent.mkdir(parents=True)
            pipeline_script.write_text("# fixture only\n", encoding="utf-8")
            a3_script = cwd / "ostram" / "pipeline" / "scenarios" / "materializer.py"
            a3_script.parent.mkdir(parents=True)
            a3_script.write_text("# fixture only\n", encoding="utf-8")
            commands: list[tuple[list[str], Path, Path | None]] = []

            def record(command, *, cwd=None) -> None:
                commands.append((list(command), Path.cwd(), cwd))

            path_fixture = mock.Mock(project_root=cwd)
            path_fixture.stage_workspace.side_effect = (
                lambda stage, create=False: cwd / "workspace" / stage
            )

            with (
                _working_directory(cwd),
                mock.patch.object(launcher, "run", side_effect=record),
                mock.patch.object(
                    launcher, "resolve_paths", return_value=path_fixture
                ),
                redirect_stdout(io.StringIO()),
            ):
                launcher.run_pipeline_script(
                    "Env Name", pipeline_script, '--scenarios "C,A"'
                )
                launcher.run_a3_for_scenario("Env Name", a3_script, "Scenario A")

        self.assertEqual(
            commands,
            [
                (
                    [
                        sys.executable,
                        "-B",
                        "-m",
                        "ostram.pipeline.compilation.runner",
                        "--scenarios",
                        "C,A",
                    ],
                    cwd,
                    cwd / "workspace" / "compilation",
                ),
                (
                    [
                        sys.executable,
                        "-B",
                        "-m",
                        "ostram.pipeline.scenarios.materializer",
                        "--scenarios",
                        "Scenario A",
                    ],
                    cwd,
                    cwd / "workspace" / "scenarios",
                ),
            ],
        )

    def test_pipeline_resolution_does_not_depend_on_caller_directory(self) -> None:
        launcher = _load_launcher("relative_display")
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as script_temp:
            script = Path(script_temp).resolve() / "stage.py"
            script.write_text("# fixture only\n", encoding="utf-8")
            with tempfile.TemporaryDirectory(dir=TEST_ROOT) as cwd_temp:
                with (
                    _working_directory(Path(cwd_temp).resolve()),
                    mock.patch.object(launcher, "run") as process_boundary,
                    redirect_stdout(io.StringIO()),
                ):
                    launcher.run_pipeline_script("env", script)
            process_boundary.assert_called_once()

    def test_missing_stage_files_fail_before_command_execution(self) -> None:
        launcher = _load_launcher("missing_scripts")
        missing = TEST_ROOT / "does-not-exist.py"
        with mock.patch.object(launcher, "run") as process_boundary:
            with self.assertRaises(FileNotFoundError):
                launcher.run_pipeline_script("env", missing)
            with self.assertRaises(FileNotFoundError):
                launcher.run_a3_for_scenario("env", missing, "A")
        process_boundary.assert_not_called()


if __name__ == "__main__":
    unittest.main()
