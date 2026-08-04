from __future__ import annotations

import ast
import importlib
import importlib.util
import io
import os
import subprocess
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
CLI_PATH = REPO_ROOT / "ostram" / "__main__.py"


def _load_cli(label: str):
    module_name = f"_ostram_canonical_cli_{label}"
    spec = importlib.util.spec_from_file_location(module_name, CLI_PATH)
    if spec is None or spec.loader is None:
        raise AssertionError(f"could not load module spec for {CLI_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    finally:
        sys.modules.pop(module_name, None)
    return module


def _capture_exit(call):
    stdout = io.StringIO()
    stderr = io.StringIO()
    with redirect_stdout(stdout), redirect_stderr(stderr):
        try:
            result = call()
        except SystemExit as error:
            code = error.code
        else:
            code = result
    return code, stdout.getvalue(), stderr.getvalue()


def _exact_main_guard(path: Path) -> ast.If:
    tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
    guards = []
    for node in tree.body:
        if not isinstance(node, ast.If):
            continue
        test = node.test
        if (
            isinstance(test, ast.Compare)
            and isinstance(test.left, ast.Name)
            and test.left.id == "__name__"
            and len(test.ops) == 1
            and isinstance(test.ops[0], ast.Eq)
            and len(test.comparators) == 1
            and isinstance(test.comparators[0], ast.Constant)
            and test.comparators[0].value == "__main__"
        ):
            guards.append(node)
    if len(guards) != 1:
        raise AssertionError(f"expected one exact __main__ guard in {path}")
    if tree.body[-1] is not guards[0]:
        raise AssertionError(f"expected __main__ guard to be final in {path}")
    return guards[0]


def _route_effect_boundary(command: str, module):
    if command == "run":
        owner, attribute = module, "check_tool_available"
    elif command == "transform":
        owner, attribute = module._orchestrator, "orchestrate_a3"
    elif command == "compile-inputs":
        owner, attribute = module._impl, "orchestrate"
    else:  # pragma: no cover - route registry test freezes the complete set
        raise AssertionError(f"missing effect sentinel for {command}")
    boundary = mock.Mock(
        side_effect=AssertionError(f"{command} effect boundary must be unreachable")
    )
    return mock.patch.object(owner, attribute, boundary), boundary


def _run_smoke(arguments: list[str], cwd: Path, *, expose_repo: bool = False):
    environment = os.environ.copy()
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    if expose_repo:
        existing = environment.get("PYTHONPATH")
        environment["PYTHONPATH"] = (
            str(REPO_ROOT)
            if not existing
            else str(REPO_ROOT) + os.pathsep + existing
        )
    return subprocess.run(
        [sys.executable, "-B", *arguments],
        cwd=str(cwd),
        env=environment,
        capture_output=True,
        text=True,
        timeout=30,
        check=False,
    )


class CanonicalCliImportAndHelpTests(unittest.TestCase):
    def test_import_ostram_package_is_silent_and_imports_no_entrypoint(self) -> None:
        package_names = ("ostram", "ostram.__main__")
        historical_names = (
            "run",
            "ostram.pipeline.scenarios.transform",
            "ostram.pipeline.compilation.runner",
            "ostram.pipeline.compilation.compiler",
            "ostram.pipeline.execution.runner",
        )
        saved_packages = {
            name: sys.modules.pop(name)
            for name in package_names
            if name in sys.modules
        }
        historical_before = {
            name: sys.modules.get(name) for name in historical_names
        }
        try:
            with (
                mock.patch.object(sys, "path", [str(REPO_ROOT), *sys.path]),
                mock.patch.object(os, "chdir") as change_directory,
                mock.patch.object(subprocess, "run") as process_run,
                mock.patch.object(subprocess, "check_call") as check_call,
                mock.patch.object(subprocess, "check_output") as check_output,
                mock.patch.object(subprocess, "Popen") as process_open,
                redirect_stdout(io.StringIO()) as stdout,
                redirect_stderr(io.StringIO()) as stderr,
            ):
                package = importlib.import_module("ostram")

            self.assertEqual(package.__name__, "ostram")
            self.assertEqual(stdout.getvalue(), "")
            self.assertEqual(stderr.getvalue(), "")
            self.assertNotIn("ostram.__main__", sys.modules)
            for name, previous in historical_before.items():
                self.assertIs(sys.modules.get(name), previous)
            change_directory.assert_not_called()
            process_run.assert_not_called()
            check_call.assert_not_called()
            check_output.assert_not_called()
            process_open.assert_not_called()
        finally:
            for name in package_names:
                sys.modules.pop(name, None)
            sys.modules.update(saved_packages)

    def test_import_is_silent_lazy_and_crosses_no_effect_boundary(self) -> None:
        with (
            mock.patch.object(importlib, "import_module") as import_module,
            mock.patch.object(os, "chdir") as change_directory,
            redirect_stdout(io.StringIO()) as stdout,
            redirect_stderr(io.StringIO()) as stderr,
        ):
            module = _load_cli("import_safety")

        import_module.assert_not_called()
        change_directory.assert_not_called()
        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")
        self.assertTrue(callable(module.main))

    def test_no_arguments_and_top_level_help_describe_only_truthful_routes(self) -> None:
        cli = _load_cli("top_help")
        with mock.patch.object(cli, "_load_route_module") as load_route:
            no_args = _capture_exit(lambda: cli.main([]))
            explicit_help = _capture_exit(lambda: cli.main(["--help"]))

        load_route.assert_not_called()
        self.assertEqual(no_args[0], 0)
        self.assertEqual(explicit_help[0], 0)
        self.assertEqual(no_args[1], explicit_help[1])
        self.assertEqual(no_args[2], "")
        help_text = no_args[1]
        self.assertIn("python -m ostram", help_text)
        self.assertIn("run", help_text)
        self.assertIn("transform", help_text)
        self.assertIn("compile-inputs", help_text)
        self.assertIn("inspect-resources", help_text)
        self.assertNotIn("prepare-model", help_text)
        self.assertNotIn("solve", help_text)
        self.assertIn("project-root", help_text)

    def test_unknown_deferred_and_malformed_top_level_commands_exit_two(self) -> None:
        cli = _load_cli("top_errors")
        for argv in (
            ["unknown"],
            ["prepare-model"],
            ["solve"],
            ["--unknown"],
        ):
            with self.subTest(argv=argv):
                with mock.patch.object(cli, "_load_route_module") as load_route:
                    code, stdout, stderr = _capture_exit(lambda: cli.main(argv))
                self.assertEqual(code, 2)
                self.assertEqual(stdout, "")
                self.assertIn("error:", stderr)
                load_route.assert_not_called()

    def test_subcommand_help_is_byte_for_byte_historical_parser_help(self) -> None:
        cli = _load_cli("subcommand_help")
        for command, route in cli.ROUTES.items():
            with self.subTest(command=command):
                historical = importlib.import_module(route.module_name)
                boundary_patch, boundary = _route_effect_boundary(
                    command, historical
                )

                def direct_help():
                    with mock.patch.object(sys, "argv", [route.program, "--help"]):
                        return historical.main()

                with boundary_patch:
                    direct = _capture_exit(direct_help)
                    canonical = _capture_exit(
                        lambda: cli.main([command, "--help"])
                    )
                self.assertEqual(canonical, direct)
                self.assertEqual(canonical[0], 0)
                boundary.assert_not_called()

    def test_malformed_subcommand_arguments_keep_historical_argparse_exit_two(self) -> None:
        cli = _load_cli("subcommand_errors")
        cases = (
            ("run", ["--env-name"]),
            ("run", ["--unknown"]),
            ("transform", ["--scenario"]),
            ("transform", ["--unknown"]),
            ("compile-inputs", ["--scenarios"]),
            ("compile-inputs", ["--unknown"]),
        )
        for command, forwarded in cases:
            with self.subTest(command=command, forwarded=forwarded):
                route = cli.ROUTES[command]
                historical = importlib.import_module(route.module_name)
                boundary_patch, boundary = _route_effect_boundary(
                    command, historical
                )

                def direct_malformed():
                    with mock.patch.object(
                        sys, "argv", [route.program, *forwarded]
                    ):
                        return historical.main()

                with boundary_patch:
                    direct = _capture_exit(direct_malformed)
                    canonical = _capture_exit(
                        lambda: cli.main([command, *forwarded])
                    )
                self.assertEqual(canonical, direct)
                self.assertEqual(canonical[0], 2)
                self.assertEqual(canonical[1], "")
                self.assertIn("error:", canonical[2])
                boundary.assert_not_called()

    def test_cli_help_cannot_cross_any_real_process_boundary(self) -> None:
        cli = _load_cli("help_fail_closed")
        with (
            mock.patch.object(subprocess, "run") as process_run,
            mock.patch.object(subprocess, "check_call") as check_call,
            mock.patch.object(subprocess, "check_output") as check_output,
            mock.patch.object(subprocess, "Popen") as process_open,
        ):
            for command in cli.ROUTES:
                with self.subTest(command=command):
                    code, _stdout, _stderr = _capture_exit(
                        lambda: cli.main([command, "--help"])
                    )
                    self.assertEqual(code, 0)

        process_run.assert_not_called()
        check_call.assert_not_called()
        check_output.assert_not_called()
        process_open.assert_not_called()

    def test_user_documentation_records_mappings_and_deliberate_deferrals(self) -> None:
        for relative in (
            "README.md",
            "docs/quickstart.md",
            "docs/lineage.md",
        ):
            with self.subTest(relative=relative):
                text = (REPO_ROOT / relative).read_text(encoding="utf-8")
                self.assertIn("python -m ostram run", text)
                self.assertIn("python -m ostram transform", text)
                self.assertIn("python -m ostram compile-inputs", text)
                self.assertIn("inspect-resources", text)
                self.assertIn("solve", text)
                self.assertNotIn("run.py", text)
                self.assertNotIn("t1_confection", text)


class CanonicalCliSubprocessSmokeTests(unittest.TestCase):
    def test_real_top_level_help_and_unknown_from_root_and_spaced_cwd(self) -> None:
        with tempfile.TemporaryDirectory(
            prefix="ostram cli subprocess cwd with spaces "
        ) as temporary:
            spaced_cwd = Path(temporary).resolve()
            self.assertIn(" ", str(spaced_cwd))
            locations = (
                ("repository-root", REPO_ROOT, False),
                ("spaced-cwd", spaced_cwd, True),
            )

            canonical_results = {}
            for label, cwd, expose_repo in locations:
                with self.subTest(location=label, route="canonical-help"):
                    help_result = _run_smoke(
                        ["-m", "ostram", "--help"],
                        cwd,
                        expose_repo=expose_repo,
                    )
                    self.assertEqual(help_result.returncode, 0)
                    self.assertIn("python -m ostram", help_result.stdout)
                    self.assertEqual(help_result.stderr, "")

                with self.subTest(location=label, route="canonical-unknown"):
                    unknown_result = _run_smoke(
                        ["-m", "ostram", "unknown"],
                        cwd,
                        expose_repo=expose_repo,
                    )
                    self.assertEqual(unknown_result.returncode, 2)
                    self.assertEqual(unknown_result.stdout, "")
                    self.assertIn("error:", unknown_result.stderr)
                    self.assertIn("unknown", unknown_result.stderr)

                canonical_results[label] = (
                    help_result.stdout,
                    help_result.stderr,
                    unknown_result.stdout,
                    unknown_result.stderr,
                )

            self.assertEqual(
                canonical_results["repository-root"],
                canonical_results["spaced-cwd"],
            )


class CanonicalCliDispatchTests(unittest.TestCase):
    def test_route_registry_maps_canonical_package_modules_and_defers_b2(self) -> None:
        cli = _load_cli("registry")
        self.assertEqual(
            {
                name: (route.module_name, route.program, route.exit_policy)
                for name, route in cli.ROUTES.items()
            },
            {
                "run": (
                    "ostram.pipeline.orchestration",
                    "python -m ostram run",
                    "run-guard",
                ),
                "transform": (
                    "ostram.pipeline.scenarios.transform",
                    "python -m ostram transform",
                    "main-result",
                ),
                "compile-inputs": (
                    "ostram.pipeline.compilation.runner",
                    "python -m ostram compile-inputs",
                    "natural-zero",
                ),
            },
        )
        source = CLI_PATH.read_text(encoding="utf-8-sig")
        self.assertNotIn("B1_Compiler", source)
        self.assertNotIn("B2_Executing_OG_Model", source)
        self.assertNotIn("shell=True", source)
        self.assertNotIn("subprocess.run", source)
        self.assertNotIn("subprocess.Popen", source)
        self.assertNotIn("os.chdir", source)
        guard = _exact_main_guard(CLI_PATH)
        self.assertEqual(len(guard.body), 1)
        self.assertEqual(guard.orelse, [])
        statement = guard.body[0]
        self.assertIsInstance(statement, ast.Raise)
        self.assertIsNone(statement.cause)
        self.assertIsInstance(statement.exc, ast.Call)
        self.assertIsInstance(statement.exc.func, ast.Name)
        self.assertEqual(statement.exc.func.id, "SystemExit")
        self.assertEqual(statement.exc.keywords, [])
        self.assertEqual(len(statement.exc.args), 1)
        main_call = statement.exc.args[0]
        self.assertIsInstance(main_call, ast.Call)
        self.assertIsInstance(main_call.func, ast.Name)
        self.assertEqual(main_call.func.id, "main")
        self.assertEqual(main_call.args, [])
        self.assertEqual(main_call.keywords, [])

    def test_each_canonical_route_calls_historical_main_once_with_exact_argv(self) -> None:
        cli = _load_cli("one_call")
        raw = ["--scenarios", " C , A, C ", "--literal=two words"]
        for command, route in cli.ROUTES.items():
            with self.subTest(command=command):
                events: list[tuple[object, ...]] = []

                def historical_main():
                    events.append(
                        (
                            "main",
                            tuple(sys.argv),
                            Path.cwd(),
                            dict(os.environ),
                        )
                    )
                    return 23

                fake = types.SimpleNamespace(main=mock.Mock(side_effect=historical_main))
                original_argv = sys.argv
                with (
                    mock.patch.object(cli, "_load_route_module", return_value=fake) as load,
                    mock.patch.dict(os.environ, {"OSTRAM_SENTINEL": "kept"}, clear=True),
                ):
                    result = cli.main([command, *raw])

                load.assert_called_once_with(route)
                fake.main.assert_called_once_with()
                self.assertEqual(len(events), 1)
                self.assertEqual(events[0][1], (route.program, *raw))
                self.assertEqual(events[0][2], Path.cwd())
                self.assertEqual(events[0][3]["OSTRAM_SENTINEL"], "kept")
                self.assertEqual(
                    Path(events[0][3]["OSTRAM_PROJECT_ROOT"]), REPO_ROOT
                )
                self.assertEqual(
                    Path(events[0][3]["OSTRAM_WORKSPACE"]),
                    REPO_ROOT / "workspace",
                )
                self.assertIs(sys.argv, original_argv)
                expected_result = 23 if route.exit_policy == "main-result" else 0
                self.assertEqual(result, expected_result)

    def test_mocked_canonical_and_historical_dispatch_traces_are_identical(self) -> None:
        cli = _load_cli("trace_parity")
        raw = ["--scenarios", " C , A, C "]
        for command, route in cli.ROUTES.items():
            with self.subTest(command=command):
                traces: list[tuple[object, ...]] = []

                def shared_main():
                    traces.append(
                        (
                            tuple(sys.argv),
                            Path.cwd(),
                            os.environ.get("TRACE_SENTINEL"),
                        )
                    )
                    print("shared stdout")
                    print("shared stderr", file=sys.stderr)
                    return 0

                fake = types.SimpleNamespace(main=mock.Mock(side_effect=shared_main))
                with mock.patch.dict(os.environ, {"TRACE_SENTINEL": "same"}, clear=False):
                    with mock.patch.object(sys, "argv", [route.program, *raw]):
                        direct = _capture_exit(shared_main)
                    with mock.patch.object(cli, "_load_route_module", return_value=fake):
                        canonical = _capture_exit(lambda: cli.main([command, *raw]))

                self.assertEqual(canonical, direct)
                self.assertEqual(traces[0], traces[1])
                fake.main.assert_called_once_with()

    def test_actual_modules_reach_identical_first_boundaries_without_effects(
        self,
    ) -> None:
        cli = _load_cli("actual_boundary_trace")
        boundary_stop_message = "fixture stopped at first run effect boundary"

        def historical_run_guard(module):
            try:
                module.main()
            except subprocess.CalledProcessError as error:
                print(
                    f"\nCommand failed (exit {error.returncode}): {error.cmd}",
                    file=sys.stderr,
                )
                return error.returncode
            except Exception as error:
                print(f"\nError: {error}", file=sys.stderr)
                return 1
            return 0

        for command, forwarded in (
            (
                "run",
                [
                    "--env-name",
                    "fixture-env",
                    "--skip-pull",
                    "--skip-a3",
                    "--skip-b1",
                    "--skip-b2",
                ],
            ),
            ("transform", ["--scenario", "fixture scenario", "--keep-workdir"]),
            ("compile-inputs", ["--scenarios", " C , A, C "]),
        ):
            with self.subTest(command=command):
                route = cli.ROUTES[command]
                module = importlib.import_module(route.module_name)
                events: list[tuple[object, ...]] = []

                if command == "run":
                    def boundary(tool):
                        events.append(
                            (
                                "check_tool_available",
                                tool,
                                tuple(sys.argv),
                                Path.cwd(),
                                os.environ.get("ACTUAL_TRACE_SENTINEL"),
                            )
                        )
                        raise RuntimeError(boundary_stop_message)

                    boundary_patch = mock.patch.object(
                        module, "check_tool_available", side_effect=boundary
                    )
                elif command == "transform":
                    def boundary(cli_args, paths, dependencies, input_files):
                        events.append(
                            (
                                "orchestrate_a3",
                                cli_args.scenario,
                                cli_args.keep_workdir,
                                tuple(input_files),
                                tuple(sys.argv),
                                Path.cwd(),
                                os.environ.get("ACTUAL_TRACE_SENTINEL"),
                            )
                        )
                        return 13

                    boundary_patch = mock.patch.object(
                        module._orchestrator,
                        "orchestrate_a3",
                        side_effect=boundary,
                    )
                else:
                    def boundary(cli_args, paths, **kwargs):
                        events.append(
                            (
                                "orchestrate",
                                cli_args.scenarios,
                                Path(paths.script_dir),
                                tuple(sorted(kwargs)),
                                tuple(sys.argv),
                                Path.cwd(),
                                os.environ.get("ACTUAL_TRACE_SENTINEL"),
                            )
                        )
                        return None

                    boundary_patch = mock.patch.object(
                        module._impl, "orchestrate", side_effect=boundary
                    )

                def direct():
                    with mock.patch.object(
                        sys, "argv", [route.program, *forwarded]
                    ):
                        if command == "run":
                            return historical_run_guard(module)
                        return module.main()

                with (
                    boundary_patch as patched_boundary,
                    mock.patch.dict(
                        os.environ,
                        {"ACTUAL_TRACE_SENTINEL": "preserved"},
                        clear=False,
                    ),
                ):
                    direct_result = _capture_exit(direct)
                    canonical_result = _capture_exit(
                        lambda: cli.main([command, *forwarded])
                    )

                self.assertEqual(len(events), 2)
                self.assertEqual(events[0], events[1])
                self.assertEqual(patched_boundary.call_count, 2)
                self.assertEqual(
                    canonical_result[1:], direct_result[1:]
                )
                if command == "compile-inputs":
                    self.assertIsNone(direct_result[0])
                    self.assertEqual(canonical_result[0], 0)
                else:
                    self.assertEqual(canonical_result[0], direct_result[0])

    def test_no_downstream_arguments_keep_historical_parser_defaults(self) -> None:
        cli = _load_cli("no_downstream_arguments")
        original_argv = sys.argv

        run_module = importlib.import_module("ostram.pipeline.orchestration")
        with mock.patch.object(
            run_module,
            "check_tool_available",
            side_effect=RuntimeError("fixture stopped at first run effect"),
        ) as run_boundary:
            run_result = _capture_exit(lambda: cli.main(["run"]))
        self.assertEqual(run_result[0], 1)
        self.assertIn("Using environment:", run_result[1])
        self.assertEqual(
            run_result[2], "\nError: fixture stopped at first run effect\n"
        )
        run_boundary.assert_called_once_with("conda")

        transform_module = importlib.import_module(
            "ostram.pipeline.scenarios.transform"
        )
        with mock.patch.object(
            transform_module._orchestrator,
            "orchestrate_a3",
            return_value=31,
        ) as transform_boundary:
            transform_result = cli.main(["transform"])
        self.assertEqual(transform_result, 31)
        transform_boundary.assert_called_once()
        transform_args = transform_boundary.call_args.args[0]
        self.assertEqual(transform_args.scenario, "BAU")
        self.assertIsNone(transform_args.soasia)
        self.assertFalse(transform_args.keep_workdir)

        compile_module = importlib.import_module(
            "ostram.pipeline.compilation.runner"
        )
        with mock.patch.object(
            compile_module._impl,
            "orchestrate",
            return_value=None,
        ) as compile_boundary:
            compile_result = cli.main(["compile-inputs"])
        self.assertEqual(compile_result, 0)
        compile_boundary.assert_called_once()
        compile_args = compile_boundary.call_args.args[0]
        self.assertIsNone(compile_args.scenarios)
        self.assertIs(sys.argv, original_argv)

    def test_caller_cwd_with_spaces_and_environment_are_preserved(self) -> None:
        cli = _load_cli("cwd_env")
        starting_cwd = Path.cwd()
        with tempfile.TemporaryDirectory(
            prefix="ostram cli caller with spaces "
        ) as temporary:
            caller = Path(temporary).resolve()
            self.assertIn(" ", str(caller))
            try:
                os.chdir(caller)
                for command, route in cli.ROUTES.items():
                    with self.subTest(command=command):
                        observed: list[
                            tuple[Path, tuple[str, ...], dict[str, str]]
                        ] = []

                        def historical_main():
                            observed.append(
                                (Path.cwd(), tuple(sys.argv), dict(os.environ))
                            )
                            return 0

                        fake = types.SimpleNamespace(main=historical_main)
                        original_argv = sys.argv
                        environment = {
                            "PATH_WITH_SPACES": str(caller / "tool bin")
                        }
                        with (
                            mock.patch.dict(os.environ, environment, clear=True),
                            mock.patch.object(
                                cli, "_load_route_module", return_value=fake
                            ),
                        ):
                            result = cli.main(
                                [command, "--value", "argument with spaces"]
                            )
                            self.assertEqual(Path.cwd(), caller)
                            self.assertEqual(os.environ, environment)

                        self.assertIn(result, (0, None))
                        self.assertEqual(observed[0][0], caller)
                        self.assertEqual(
                            observed[0][1],
                            (route.program, "--value", "argument with spaces"),
                        )
                        self.assertEqual(
                            observed[0][2]["PATH_WITH_SPACES"],
                            environment["PATH_WITH_SPACES"],
                        )
                        self.assertEqual(
                            Path(observed[0][2]["OSTRAM_PROJECT_ROOT"]),
                            REPO_ROOT,
                        )
                        self.assertEqual(
                            Path(observed[0][2]["OSTRAM_WORKSPACE"]),
                            REPO_ROOT / "workspace",
                        )
                        self.assertIs(sys.argv, original_argv)
            finally:
                os.chdir(starting_cwd)

    def test_sys_argv_is_restored_for_all_exit_and_interruption_routes(self) -> None:
        cli = _load_cli("argv_restore")
        failures = (
            SystemExit(17),
            KeyboardInterrupt(),
        )
        for command in cli.ROUTES:
            for failure in failures:
                with self.subTest(command=command, failure=type(failure).__name__):
                    fake = types.SimpleNamespace(main=mock.Mock(side_effect=failure))
                    original_argv = sys.argv
                    with mock.patch.object(cli, "_load_route_module", return_value=fake):
                        with self.assertRaises(type(failure)):
                            cli.main([command, "--sentinel"])
                    self.assertIs(sys.argv, original_argv)

        for command in ("transform", "compile-inputs"):
            for failure in (
                RuntimeError("ordinary"),
                subprocess.CalledProcessError(9, ["child", "with spaces"]),
            ):
                with self.subTest(
                    command=command, failure=type(failure).__name__
                ):
                    fake = types.SimpleNamespace(
                        main=mock.Mock(side_effect=failure)
                    )
                    original_argv = sys.argv
                    stdout = io.StringIO()
                    stderr = io.StringIO()
                    with (
                        mock.patch.object(
                            cli, "_load_route_module", return_value=fake
                        ),
                        redirect_stdout(stdout),
                        redirect_stderr(stderr),
                    ):
                        with self.assertRaises(type(failure)) as raised:
                            cli.main([command])
                    self.assertIs(raised.exception, failure)
                    self.assertEqual(stdout.getvalue(), "")
                    self.assertEqual(stderr.getvalue(), "")
                    self.assertIs(sys.argv, original_argv)

    def test_run_import_failure_propagates_before_guard_translation(self) -> None:
        cli = _load_cli("run_import_failure")
        failure = ModuleNotFoundError("fixture route import failed")
        original_argv = sys.argv
        stdout = io.StringIO()
        stderr = io.StringIO()
        with (
            mock.patch.object(
                cli, "_load_route_module", side_effect=failure
            ) as load_route,
            redirect_stdout(stdout),
            redirect_stderr(stderr),
        ):
            with self.assertRaises(ModuleNotFoundError) as raised:
                cli.main(["run", "--sentinel"])

        self.assertIs(raised.exception, failure)
        load_route.assert_called_once_with(cli.ROUTES["run"])
        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")
        self.assertIs(sys.argv, original_argv)

    def test_run_route_preserves_child_and_ordinary_exception_translation(self) -> None:
        cli = _load_cli("run_failures")
        cases = (
            (
                subprocess.CalledProcessError(7, ["child", "argument with spaces"]),
                7,
                "\nCommand failed (exit 7): ['child', 'argument with spaces']\n",
            ),
            (RuntimeError("ordinary failure"), 1, "\nError: ordinary failure\n"),
        )
        for failure, expected_code, expected_stderr in cases:
            with self.subTest(failure=type(failure).__name__):
                fake = types.SimpleNamespace(main=mock.Mock(side_effect=failure))
                original_argv = sys.argv
                with mock.patch.object(cli, "_load_route_module", return_value=fake):
                    code, stdout, stderr = _capture_exit(
                        lambda: cli.main(["run", "--sentinel"])
                    )
                self.assertEqual(code, expected_code)
                self.assertEqual(stdout, "")
                self.assertEqual(stderr, expected_stderr)
                fake.main.assert_called_once_with()
                self.assertIs(sys.argv, original_argv)


if __name__ == "__main__":
    unittest.main()
