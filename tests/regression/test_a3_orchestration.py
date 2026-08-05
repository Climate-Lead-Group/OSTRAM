from __future__ import annotations

import argparse
import ast
import builtins
import importlib.util
import inspect
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import types
import unittest
from contextlib import contextmanager, redirect_stderr, redirect_stdout
from pathlib import Path
from types import SimpleNamespace
from unittest import mock


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
SCENARIO_PACKAGE = REPO_ROOT / "ostram" / "pipeline" / "scenarios"
A3_ENTRYPOINT = SCENARIO_PACKAGE / "transform.py"
A3_ORCHESTRATOR = SCENARIO_PACKAGE / "orchestrator.py"
SCENARIO_HELPER = SCENARIO_PACKAGE / "transformations" / "scenario_workbooks.py"
SCENARIO_REGISTRY = REPO_ROOT / "config" / "scenarios" / "registry.json"

ACTIVE_SCENARIOS = (
    "BAU",
    "A_Calibrated_BAU",
    "B_Optimised_VRE",
    "C_Target_VRE",
)

INPUT_FILES = (
    "A-O_AR_Model_Base_Year.xlsx",
    "A-O_AR_Projections.xlsx",
    "A-O_Demand.xlsx",
    "A-O_Parametrization.xlsx",
)


def _load_module(path: Path, label: str):
    parent = (
        "ostram.pipeline.scenarios.transformations"
        if path == SCENARIO_HELPER
        else "ostram.pipeline.scenarios"
    )
    module_name = f"{parent}._characterization_{label}"
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"could not load module spec for {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    finally:
        sys.modules.pop(module_name, None)
    return module


def _load_a3(label: str):
    return _load_module(A3_ENTRYPOINT, label)


def _load_orchestrator(label: str):
    return _load_module(A3_ORCHESTRATOR, f"orchestrator_{label}")


def _load_scenarios(label: str):
    return _load_module(SCENARIO_HELPER, f"scenarios_{label}")


@contextmanager
def _working_directory(path: Path):
    previous = Path.cwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(previous)


class _FixedDateTime:
    @classmethod
    def now(cls):
        return cls()

    def strftime(self, _format: str) -> str:
        return "20260716_120000"


class A3TraceHarness:
    """Run A3 orchestration against disposable files and inert stage doubles."""

    def __init__(self, module) -> None:
        self.module = module
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name).resolve()
        self.caller = self.root / "caller"
        self.caller.mkdir()
        self.preparation_workspace = self.root / "preparation_workspace"
        self.process_dir = self.preparation_workspace / "scenarios"
        self.process_dir.mkdir(parents=True)
        self.rules_dir = self.process_dir / "rules_scripts"
        self.rules_dir.mkdir()
        self.soasia = self.caller / "scenario-template.xlsx"
        self.soasia.write_bytes(b"fixture template marker")
        payload = json.loads(SCENARIO_REGISTRY.read_text(encoding="utf-8"))
        scenario_names = (
            payload["support_scenarios"] + payload["decision_scenarios"]
        )
        self.snapshots = {}
        for scenario_name in scenario_names:
            snapshot = (
                self.preparation_workspace
                / "A1_Outputs"
                / f"_post_a2_snapshot_{scenario_name}"
            )
            snapshot.mkdir(parents=True)
            for name in INPUT_FILES:
                (snapshot / name).write_bytes(
                    f"snapshot:{name}".encode("utf-8")
                )
            self.snapshots[scenario_name] = snapshot
        self.snapshot = self.snapshots["A_Calibrated_BAU"]
        self.input_dir = self.preparation_workspace / "relative-input"
        self.input_dir.mkdir()
        (self.input_dir / "stale.txt").write_text("stale", encoding="utf-8")
        self.output_dir = self.preparation_workspace / "relative-output"
        self.workdir = self.process_dir / "_run_20260716_120000"
        self.paths = {
            "wd": self.workdir,
            "s1": self.workdir / "stage1",
            "s1b": self.workdir / "stage1b",
            "s2": self.workdir / "stage2",
            "s3": self.workdir / "stage3",
            "s5": self.workdir / "stage5",
        }
        self.events: list[tuple[object, ...]] = []
        self.stdout = ""
        self.stderr = ""

    def close(self) -> None:
        self._temp.cleanup()

    def _args(
        self,
        *,
        keep_workdir: bool = False,
        scenario: str = "A_Calibrated_BAU",
    ) -> argparse.Namespace:
        return argparse.Namespace(
            scenario=scenario,
            soasia=self.soasia,
            rules_script=None,
            inherit_from=None,
            input_dir=Path("relative-input"),
            output_dir=Path("relative-output"),
            keep_workdir=keep_workdir,
        )

    def run(
        self,
        *,
        fail_stage: str | None = None,
        failure: BaseException | None = None,
        keep_workdir: bool = False,
        scenario: str = "A_Calibrated_BAU",
    ) -> int:
        module = self.module
        events = self.events
        real_copy = shutil.copy
        real_copytree = shutil.copytree
        real_rmtree = shutil.rmtree

        def resolve_config(args, soasia):
            events.append(("resolve_config", args.scenario, soasia, Path.cwd()))
            return (
                args.scenario,
                ["set_retirement_schedule.py", "set_min_capacity_floors.py"],
                ["BAU"],
            )

        def remove_tree(path, *args, **kwargs):
            path = Path(path)
            events.append(("rmtree", path, bool(kwargs.get("ignore_errors", False))))
            return real_rmtree(path, *args, **kwargs)

        def copy_tree(source, destination, *args, **kwargs):
            source = Path(source)
            destination = Path(destination)
            events.append(("copytree", source, destination))
            return real_copytree(source, destination, *args, **kwargs)

        def copy_file(source, destination, *args, **kwargs):
            source = Path(source)
            destination = Path(destination)
            events.append(("copy", source, destination))
            return real_copy(source, destination, *args, **kwargs)

        def build_workdir(parent, timestamp, rules_scripts, scenario):
            events.append(
                (
                    "build_workdir",
                    Path(parent),
                    timestamp,
                    tuple(rules_scripts),
                    scenario,
                )
            )
            for path in self.paths.values():
                path.mkdir(parents=True, exist_ok=True)
            return dict(self.paths)

        def materialize(soasia, scenario, destination):
            events.append(
                (
                    "materialize",
                    Path(soasia),
                    scenario,
                    Path(destination),
                    os.environ.get("OSTRAM_TEMPLATE_PATH"),
                    Path.cwd(),
                )
            )
            Path(destination).write_bytes(b"materialized fixture")

        def inert_stage(name, return_value=None):
            def call(*args):
                events.append(
                    (
                        name,
                        *tuple(Path(arg) if isinstance(arg, Path) else arg for arg in args),
                        os.environ.get("OSTRAM_TEMPLATE_PATH"),
                        Path.cwd(),
                    )
                )
                if name == fail_stage:
                    raise failure if failure is not None else RuntimeError("stage failed")
                return return_value

            return call

        def deliver(s5, output_dir):
            events.append(
                (
                    "deliver_outputs",
                    Path(s5),
                    Path(output_dir),
                    os.environ.get("OSTRAM_TEMPLATE_PATH"),
                    Path.cwd(),
                )
            )
            if fail_stage == "deliver_outputs":
                raise failure if failure is not None else RuntimeError("delivery failed")

        fake_scenarios = types.ModuleType("_scenarios")
        fake_scenarios.materialize_scenario_template = materialize
        stage3_result = self.paths["s3"] / "post-trn-cap.xlsx"
        stage_replacements = {
            "stage_1_scripts_1_to_5": inert_stage("stage_1_scripts_1_to_5"),
            "stage_1b": inert_stage("stage_1b"),
            "stage_2_and_2_5": inert_stage("stage_2_and_2_5"),
            "stage_3_fix_2": inert_stage("stage_3_fix_2", stage3_result),
            "stage_4_consolidate": inert_stage("stage_4_consolidate"),
            "stage_4_5_apply_inherited_restrictions": inert_stage(
                "stage_4_5_apply_inherited_restrictions"
            ),
            "stage_5_rules_scripts": inert_stage("stage_5_rules_scripts"),
            "stage_ws3_interconnector_costs": inert_stage(
                "stage_ws3_interconnector_costs"
            ),
            "stage_ws3_internal_transmission": inert_stage(
                "stage_ws3_internal_transmission"
            ),
            "stage_ws3_internal_tx_losses": inert_stage(
                "stage_ws3_internal_tx_losses"
            ),
            "stage_ws4_pwr_min_pin": inert_stage(
                "stage_ws4_pwr_min_pin"
            ),
            "stage_6_sync_og_to_ts20": inert_stage("stage_6_sync_og_to_ts20"),
            "stage_6_persist_restrictions": inert_stage(
                "stage_6_persist_restrictions"
            ),
        }

        stdout = io.StringIO()
        stderr = io.StringIO()
        previous_template = os.environ.pop("OSTRAM_TEMPLATE_PATH", None)
        patches = [
            mock.patch.object(
                module,
                "PREPARATION_WORKSPACE",
                self.preparation_workspace,
            ),
            mock.patch.object(module, "RULES_SCRIPTS_DIR", self.rules_dir),
            mock.patch.object(module, "A3_WORKSPACE", self.process_dir),
            mock.patch.object(module, "SOASIA_V18", self.soasia),
            mock.patch.object(
                module,
                "parse_cli_args",
                return_value=self._args(
                    keep_workdir=keep_workdir,
                    scenario=scenario,
                ),
            ),
            mock.patch.object(module, "_resolve_scenario_config", side_effect=resolve_config),
            mock.patch.object(module, "build_workdir", side_effect=build_workdir),
            mock.patch.object(module, "deliver_outputs", side_effect=deliver),
            mock.patch.object(module, "banner", side_effect=lambda message: events.append(("banner", message))),
            mock.patch.object(module, "datetime", _FixedDateTime),
            mock.patch.object(module.time, "time", side_effect=(100.0, 105.25)),
            mock.patch.object(module.shutil, "rmtree", side_effect=remove_tree),
            mock.patch.object(module.shutil, "copytree", side_effect=copy_tree),
            mock.patch.object(module.shutil, "copy", side_effect=copy_file),
            mock.patch.dict(sys.modules, {"_scenarios": fake_scenarios}),
            *(mock.patch.object(module, name, replacement) for name, replacement in stage_replacements.items()),
        ]

        try:
            with (
                _working_directory(self.caller),
                redirect_stdout(stdout),
                redirect_stderr(stderr),
                mock.patch.object(subprocess, "run") as process_run,
            ):
                for patcher in patches:
                    patcher.start()
                try:
                    result = module.main()
                finally:
                    for patcher in reversed(patches):
                        patcher.stop()
            process_run.assert_not_called()
        finally:
            self.stdout = stdout.getvalue()
            self.stderr = stderr.getvalue()
            if previous_template is not None:
                os.environ["OSTRAM_TEMPLATE_PATH"] = previous_template
        return result


class A3ImportAndCliCharacterizationTests(unittest.TestCase):
    def test_import_is_silent_and_crosses_no_process_or_artifact_boundary(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            with (
                _working_directory(root),
                mock.patch.object(subprocess, "run") as process_run,
                mock.patch.object(shutil, "copy") as copy_file,
                mock.patch.object(shutil, "copytree") as copy_tree,
                mock.patch.object(shutil, "rmtree") as remove_tree,
                redirect_stdout(io.StringIO()) as stdout,
                redirect_stderr(io.StringIO()) as stderr,
            ):
                module = _load_a3("import_safety")

        process_run.assert_not_called()
        copy_file.assert_not_called()
        copy_tree.assert_not_called()
        remove_tree.assert_not_called()
        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")
        self.assertTrue(callable(module.main))

    def test_package_relative_orchestrator_import_is_silent_and_import_safe(self) -> None:
        with (
            mock.patch.object(subprocess, "run") as process_run,
            redirect_stdout(io.StringIO()) as stdout,
            redirect_stderr(io.StringIO()) as stderr,
        ):
            module = _load_a3("package_relative_import")

        process_run.assert_not_called()
        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")
        self.assertEqual(
            module._orchestrator.__name__,
            "ostram.pipeline.scenarios.orchestrator",
        )

    def test_existing_callable_helper_surface_remains_available(self) -> None:
        module = _load_a3("callable_surface")
        expected = (
            "_resolve",
            "parse_cli_args",
            "banner",
            "step",
            "run_subproc",
            "_read_script_yaml_name",
            "_resolve_script_yaml",
            "build_workdir",
            "stage_1_scripts_1_to_5",
            "stage_1b",
            "stage_2_and_2_5",
            "stage_3_fix_2",
            "stage_4_consolidate",
            "stage_4_5_apply_inherited_restrictions",
            "stage_5_rules_scripts",
            "stage_ws3_interconnector_costs",
            "stage_ws3_internal_transmission",
            "stage_ws3_internal_tx_losses",
            "stage_ws4_pwr_min_pin",
            "stage_6_sync_og_to_ts20",
            "stage_6_persist_restrictions",
            "deliver_outputs",
            "_resolve_scenario_config",
            "main",
        )
        self.assertEqual(
            [name for name in expected if not callable(getattr(module, name, None))],
            [],
        )
        self.assertFalse(hasattr(module, "stage_0_5_rnwbio"))

    def test_runtime_assets_use_authorities_without_staging_retired_machinery(
        self,
    ) -> None:
        module = _load_a3("runtime_authority_assets")
        build_source = inspect.getsource(module.build_workdir)
        stage1_source = inspect.getsource(module.stage_1_scripts_1_to_5)
        self.assertIn("OSTRAM_Timeslice_Inputs.xlsx", build_source)
        self.assertIn('script == "relax_interconnectors.py"', build_source)
        self.assertNotIn("stage1_sources", build_source)
        self.assertNotIn("OSTRAM_AO_Extensions_FILLED.xlsx", build_source)
        self.assertNotIn(
            "A-O_Parametrization_REFERENCE_with_RNWBIO.xlsx",
            build_source,
        )
        self.assertNotIn("fix_rnwbio_restore.py", build_source)
        self.assertIn(
            "ostram.pipeline.scenarios.transformations.ao_extension_decisions",
            stage1_source,
        )
        self.assertNotIn("OSTRAM_AO_Extensions_FILLED.xlsx", stage1_source)

    def test_cli_defaults_and_explicit_values_match_predecessor(self) -> None:
        module = _load_a3("cli_values")
        with mock.patch.object(sys, "argv", ["python -m ostram transform"]):
            defaults = module.parse_cli_args()
        self.assertEqual(
            vars(defaults),
            {
                "scenario": "BAU",
                "soasia": None,
                "rules_script": None,
                "inherit_from": None,
                "input_dir": None,
                "output_dir": None,
                "keep_workdir": False,
                "run_state_out": None,
                "restrictions_source": None,
            },
        )

        argv = [
            "python -m ostram transform",
            "--scenario",
            "Scenario X",
            "--soasia",
            "relative template.xlsx",
            "--rules-script",
            "first.py, second.py",
            "--inherit-from",
            "BAU, Prior",
            "--input-dir",
            "relative input",
            "--output-dir",
            "relative output",
            "--keep-workdir",
        ]
        with mock.patch.object(sys, "argv", argv):
            explicit = module.parse_cli_args()
        self.assertEqual(explicit.scenario, "Scenario X")
        self.assertEqual(explicit.soasia, Path("relative template.xlsx"))
        self.assertEqual(explicit.rules_script, "first.py, second.py")
        self.assertEqual(explicit.inherit_from, "BAU, Prior")
        self.assertEqual(explicit.input_dir, Path("relative input"))
        self.assertEqual(explicit.output_dir, Path("relative output"))
        self.assertTrue(explicit.keep_workdir)

    def test_help_unknown_options_and_missing_values_keep_argparse_exit_codes(self) -> None:
        module = _load_a3("argparse_exits")
        cases = (
            (["python -m ostram transform", "--help"], 0),
            (["python -m ostram transform", "--unknown"], 2),
            (["python -m ostram transform", "--scenario"], 2),
            (["python -m ostram transform", "--soasia"], 2),
            (["python -m ostram transform", "--rules-script"], 2),
            (["python -m ostram transform", "--inherit-from"], 2),
            (["python -m ostram transform", "--input-dir"], 2),
            (["python -m ostram transform", "--output-dir"], 2),
        )
        for argv, code in cases:
            with self.subTest(argv=argv):
                stdout = io.StringIO()
                with (
                    mock.patch.object(sys, "argv", argv),
                    redirect_stdout(stdout),
                    redirect_stderr(io.StringIO()),
                    mock.patch.object(module, "build_workdir") as build,
                    mock.patch.object(module, "_orchestration_paths") as paths,
                    mock.patch.object(
                        module,
                        "_orchestration_dependencies",
                    ) as dependencies,
                    mock.patch.object(
                        module._orchestrator,
                        "orchestrate_a3",
                    ) as orchestrate,
                ):
                    with self.assertRaises(SystemExit) as raised:
                        module.main()
                self.assertEqual(raised.exception.code, code)
                build.assert_not_called()
                paths.assert_not_called()
                dependencies.assert_not_called()
                orchestrate.assert_not_called()
                if code == 0:
                    help_text = stdout.getvalue()
                    self.assertIn("python -m ostram transform", help_text)
                    for option in (
                        "--scenario",
                        "--soasia",
                        "--rules-script",
                        "--inherit-from",
                        "--input-dir",
                        "--output-dir",
                        "--keep-workdir",
                    ):
                        self.assertIn(option, help_text)

    def test_direct_script_guard_returns_main_result_through_sys_exit(self) -> None:
        tree = ast.parse(A3_ENTRYPOINT.read_text(encoding="utf-8-sig"), filename=str(A3_ENTRYPOINT))
        guard = next(
            node
            for node in tree.body
            if isinstance(node, ast.If)
            and isinstance(node.test, ast.Compare)
            and isinstance(node.test.left, ast.Name)
            and node.test.left.id == "__name__"
        )
        source = ast.get_source_segment(A3_ENTRYPOINT.read_text(encoding="utf-8-sig"), guard)
        self.assertIn("sys.exit(main())", source)


class A3ScenarioPlanningCharacterizationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.scenarios = _load_scenarios("planning")
        cls.a3 = _load_a3("scenario_resolution")

    def test_canonical_control_has_exact_active_order_and_only_bau_to_a_dependency(self) -> None:
        configs = self.scenarios.read_control_sheet(self.scenarios.DEFAULT_SOASIA)
        self.assertEqual(tuple(config.scenario for config in configs), ACTIVE_SCENARIOS)
        self.assertTrue(all(config.active for config in configs))
        ordered = self.scenarios.topological_order(configs)
        self.assertEqual(tuple(config.scenario for config in ordered), ACTIVE_SCENARIOS)
        dependency_edges = {
            (source, config.scenario)
            for config in configs
            for source in config.inherit_restrictions_from
            if source != config.scenario
        }
        self.assertEqual(dependency_edges, {("BAU", "A_Calibrated_BAU")})

    def test_topological_order_drops_inactive_dependencies_and_detects_cycles(self) -> None:
        config = self.scenarios.ScenarioConfig
        ordered = self.scenarios.topological_order(
            [
                config("B", True, inherit_restrictions_from=["Inactive"]),
                config("Inactive", False),
                config("A", True),
            ]
        )
        self.assertEqual([item.scenario for item in ordered], ["B", "A"])
        with self.assertRaisesRegex(ValueError, "Inheritance cycle"):
            self.scenarios.topological_order(
                [
                    config("A", True, inherit_restrictions_from=["B"]),
                    config("B", True, inherit_restrictions_from=["A"]),
                ]
            )

    def test_explicit_filter_semantics_do_not_add_prerequisites(self) -> None:
        ordered = list(ACTIVE_SCENARIOS)
        cases = (
            (None, ordered, []),
            ("A_Calibrated_BAU", ["A_Calibrated_BAU"], []),
            (
                " C_Target_VRE,BAU,C_Target_VRE ",
                ["BAU", "C_Target_VRE"],
                [],
            ),
            ("Missing,Missing,BAU", ["BAU"], ["Missing", "Missing"]),
            (", ,", [], []),
            ("   ", [], []),
            ("", ordered, []),
        )
        for raw, expected, unknown in cases:
            with self.subTest(raw=raw):
                if raw:
                    requested = [part.strip() for part in raw.split(",") if part.strip()]
                    actual_unknown = [name for name in requested if name not in ordered]
                    selected = [name for name in ordered if name in requested]
                else:
                    actual_unknown = []
                    selected = ordered
                self.assertEqual(selected, expected)
                self.assertEqual(actual_unknown, unknown)
        self.assertNotIn("BAU", ["A_Calibrated_BAU"])

    def test_scenario_config_preserves_control_chain_and_cli_override_order(self) -> None:
        config = SimpleNamespace(
            scenario="A_Calibrated_BAU",
            rules_scripts=("control-first.py", "control-second.py"),
            inherit_restrictions_from=("BAU",),
        )
        fake_scenarios = types.ModuleType("_scenarios")
        fake_scenarios.read_control_sheet = lambda _path: [config]
        soasia = Path("exists.xlsx")
        args = argparse.Namespace(
            scenario="A_Calibrated_BAU",
            rules_script=None,
            inherit_from=None,
        )
        with (
            mock.patch.object(Path, "is_file", return_value=True),
            mock.patch.dict(sys.modules, {"_scenarios": fake_scenarios}),
        ):
            resolved = self.a3._resolve_scenario_config(args, soasia)
        self.assertEqual(
            resolved,
            (
                "A_Calibrated_BAU",
                ["control-first.py", "control-second.py"],
                ["BAU"],
            ),
        )

        args.rules_script = " second.py, first.py, second.py\nthird.py "
        args.inherit_from = " Prior, BAU, Prior "
        with (
            mock.patch.object(Path, "is_file", return_value=True),
            mock.patch.dict(sys.modules, {"_scenarios": fake_scenarios}),
        ):
            resolved = self.a3._resolve_scenario_config(args, soasia)
        self.assertEqual(
            resolved,
            (
                "A_Calibrated_BAU",
                ["second.py", "first.py", "second.py", "third.py"],
                ["Prior", "BAU", "Prior"],
            ),
        )

    def test_missing_v18_and_unknown_scenario_failures_keep_system_exit_messages(self) -> None:
        args = argparse.Namespace(
            scenario="Not_BAU",
            rules_script=None,
            inherit_from=None,
        )
        with mock.patch.object(Path, "is_file", return_value=False):
            with self.assertRaisesRegex(
                SystemExit,
                "required scenario-input workbook not found",
            ):
                self.a3._resolve_scenario_config(args, Path("missing.xlsx"))

        args.scenario = "BAU"
        args.rules_script = ""
        with mock.patch.object(Path, "is_file", return_value=False):
            with self.assertRaisesRegex(
                SystemExit,
                "required scenario-input workbook not found",
            ):
                self.a3._resolve_scenario_config(args, Path("missing.xlsx"))


class A3IsolatedBoundaryTests(unittest.TestCase):
    def test_isolated_module_has_no_b1_b2_compiler_matrix_or_solver_boundary(self) -> None:
        source = A3_ORCHESTRATOR.read_text(encoding="utf-8-sig")
        forbidden = (
            "B1_Run_Compiler",
            "B1_Compiler",
            "B2_Executing_OG_Model",
            "main_executer",
            "multiprocessing",
            "SolverAdapter",
            "subprocess",
            "create_matrix",
            "glpsol",
            "cbc",
            "cplex",
            "gurobi",
        )
        self.assertEqual(
            [marker for marker in forbidden if marker in source],
            [],
        )

    def test_public_main_forwards_explicit_plan_and_dependency_seams(self) -> None:
        module = _load_a3("wrapper_forwarding")
        cli_args = object()
        paths = object()
        dependencies = object()
        with (
            mock.patch.object(module, "parse_cli_args", return_value=cli_args),
            mock.patch.object(module, "_orchestration_paths", return_value=paths),
            mock.patch.object(
                module,
                "_orchestration_dependencies",
                return_value=dependencies,
            ),
            mock.patch.object(
                module._orchestrator,
                "orchestrate_a3",
                return_value=23,
            ) as orchestrate,
        ):
            result = module.main()

        self.assertEqual(result, 23)
        orchestrate.assert_called_once_with(
            cli_args,
            paths,
            dependencies,
            module.INPUT_FILES,
        )

    def test_resolved_path_run_planning_is_pure(self) -> None:
        orchestrator = _load_orchestrator("pure_plan")
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp).resolve()
            preparation_workspace = root / "preparation_workspace"
            process_dir = preparation_workspace / "scenarios"
            process_dir.mkdir(parents=True)
            paths = orchestrator.A3Paths(
                preparation_workspace=preparation_workspace,
                process_dir=process_dir,
                default_soasia=process_dir / "OSTRAM_Scenario_Inputs.xlsx",
            )

            events: list[tuple[object, ...]] = []

            def resolve_config(args, soasia):
                events.append(("resolve_config", args.scenario, soasia))
                return "Scenario_A", ["one.py", "two.py"], ["BAU"]

            def resolve_path(value):
                events.append(("resolve_path", value))
                return preparation_workspace / Path(value)

            dependencies = SimpleNamespace(
                resolve_scenario_config=resolve_config,
                resolve_path=resolve_path,
            )
            args = argparse.Namespace(
                scenario="Scenario_A",
                soasia=Path("caller-relative.xlsx"),
                input_dir=Path("input override"),
                output_dir=Path("output override"),
                keep_workdir=True,
            )
            plan = orchestrator.resolve_plan(args, paths, dependencies)

        self.assertEqual(
            events,
            [
                ("resolve_config", "Scenario_A", Path("caller-relative.xlsx")),
                ("resolve_path", Path("input override")),
                ("resolve_path", Path("output override")),
            ],
        )
        self.assertEqual(plan.scenario, "Scenario_A")
        self.assertEqual(plan.rules_scripts, ("one.py", "two.py"))
        self.assertEqual(plan.inherit_from, ("BAU",))
        self.assertEqual(plan.soasia, Path("caller-relative.xlsx"))
        self.assertEqual(plan.input_dir, preparation_workspace / "input override")
        self.assertEqual(plan.output_dir, preparation_workspace / "output override")
        self.assertEqual(
            plan.snapshot_dir,
            preparation_workspace / "A1_Outputs" / "_post_a2_snapshot_Scenario_A",
        )
        self.assertEqual(plan.workdir_base, process_dir)
        self.assertTrue(plan.keep_workdir)

    def test_default_plan_keeps_script_anchored_paths_and_in_place_delivery(self) -> None:
        orchestrator = _load_orchestrator("default_plan")
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp).resolve()
            process_dir = root / "A3_process"
            process_dir.mkdir()
            default_soasia = process_dir / "template.xlsx"
            paths = orchestrator.A3Paths(
                preparation_workspace=root,
                process_dir=process_dir,
                default_soasia=default_soasia,
            )
            args = argparse.Namespace(
                scenario="BAU",
                soasia=None,
                input_dir=None,
                output_dir=None,
                keep_workdir=False,
            )
            dependencies = SimpleNamespace(
                resolve_scenario_config=lambda _args, _soasia: ("BAU", [], []),
                resolve_path=mock.Mock(side_effect=AssertionError("not expected")),
            )
            plan = orchestrator.resolve_plan(args, paths, dependencies)

        expected_input = root / "A1_Outputs" / "A1_Outputs_BAU"
        self.assertEqual(plan.soasia, default_soasia)
        self.assertEqual(plan.input_dir, expected_input)
        self.assertEqual(plan.output_dir, expected_input)
        self.assertFalse(plan.keep_workdir)
        dependencies.resolve_path.assert_not_called()


class A3EffectAndFailureCharacterizationTests(unittest.TestCase):
    def test_mocked_trace_matches_frozen_predecessor_observables(self) -> None:
        module = _load_a3("predecessor_trace")
        harness = A3TraceHarness(module)
        original_cwd = Path.cwd()
        try:
            result = harness.run()
            materialized = harness.workdir / "_materialized_A_Calibrated_BAU.xlsx"
            expected = [
                ("resolve_config", "A_Calibrated_BAU", harness.soasia, harness.caller),
                ("banner", "A3 workflow run @ 20260716_120000"),
                ("rmtree", harness.input_dir, False),
                ("copytree", harness.snapshot, harness.input_dir),
                (
                    "build_workdir",
                    harness.process_dir,
                    "20260716_120000",
                    ("set_retirement_schedule.py", "set_min_capacity_floors.py"),
                    "A_Calibrated_BAU",
                ),
                ("banner", "Stage 0 — materialize scenario template for 'A_Calibrated_BAU'"),
                (
                    "materialize",
                    harness.soasia,
                    "A_Calibrated_BAU",
                    materialized,
                    None,
                    harness.caller,
                ),
            ]
            expected.extend(
                ("copy", harness.input_dir / name, harness.paths["s1"] / name)
                for name in INPUT_FILES
            )
            template_value = str(materialized)
            expected.extend(
                [
                    ("stage_1_scripts_1_to_5", harness.paths["s1"], template_value, harness.caller),
                    ("stage_1b", harness.workdir, harness.paths["s1"], harness.paths["s1b"], template_value, harness.caller),
                    ("stage_2_and_2_5", harness.workdir, harness.paths["s1b"], harness.paths["s2"], template_value, harness.caller),
                    ("stage_3_fix_2", harness.paths["s2"], harness.paths["s3"], template_value, harness.caller),
                    (
                        "stage_4_consolidate",
                        harness.paths["s1"],
                        harness.paths["s3"],
                        harness.paths["s5"],
                        harness.paths["s3"] / "post-trn-cap.xlsx",
                        template_value,
                        harness.caller,
                    ),
                    (
                        "stage_4_5_apply_inherited_restrictions",
                        harness.paths["s5"],
                        harness.soasia,
                        ["BAU"],
                        template_value,
                        harness.caller,
                    ),
                    (
                        "stage_5_rules_scripts",
                        harness.workdir,
                        harness.paths["s5"],
                        ["set_retirement_schedule.py", "set_min_capacity_floors.py"],
                        template_value,
                        harness.caller,
                    ),
                    (
                        "stage_ws3_interconnector_costs",
                        harness.paths["s5"],
                        harness.soasia,
                        materialized,
                        template_value,
                        harness.caller,
                    ),
                    ("stage_ws3_internal_transmission", harness.paths["s5"], template_value, harness.caller),
                    ("stage_ws3_internal_tx_losses", harness.paths["s5"], template_value, harness.caller),
                    (
                        "stage_ws4_pwr_min_pin",
                        harness.paths["s5"],
                        "A_Calibrated_BAU",
                        template_value,
                        harness.caller,
                    ),
                    ("stage_6_sync_og_to_ts20", harness.workdir, harness.paths["s1"], template_value, harness.caller),
                    (
                        "copy",
                        harness.soasia,
                        harness.workdir
                        / "_scenario_run_state_A_Calibrated_BAU.xlsx",
                    ),
                    (
                        "stage_6_persist_restrictions",
                        harness.paths["s5"],
                        harness.workdir
                        / "_scenario_run_state_A_Calibrated_BAU.xlsx",
                        "A_Calibrated_BAU",
                        ["set_retirement_schedule.py", "set_min_capacity_floors.py"],
                        template_value,
                        harness.caller,
                    ),
                    ("deliver_outputs", harness.paths["s5"], harness.output_dir, template_value, harness.caller),
                    ("rmtree", harness.workdir, True),
                    ("banner", "DONE in 5.2s"),
                ]
            )
            self.assertEqual(result, 0)
            self.assertEqual(harness.events, expected)
            self.assertEqual(Path.cwd(), original_cwd)
            self.assertNotIn("OSTRAM_TEMPLATE_PATH", os.environ)
            self.assertFalse(harness.workdir.exists())
            self.assertTrue((harness.input_dir / INPUT_FILES[0]).is_file())
            self.assertFalse((harness.input_dir / "stale.txt").exists())
            self.assertIn(f"input-dir         : {harness.input_dir}", harness.stdout)
            self.assertIn(f"output-dir        : {harness.output_dir}", harness.stdout)
            self.assertIn(f"snapshot (source) : {harness.snapshot}", harness.stdout)
            self.assertEqual(harness.stderr, "")
        finally:
            harness.close()

    def test_keep_workdir_preserves_runtime_directory_but_clears_environment(self) -> None:
        module = _load_a3("keep_workdir")
        harness = A3TraceHarness(module)
        try:
            self.assertEqual(harness.run(keep_workdir=True), 0)
            self.assertTrue(harness.workdir.is_dir())
            self.assertNotIn(("rmtree", harness.workdir, True), harness.events)
            self.assertNotIn("OSTRAM_TEMPLATE_PATH", os.environ)
            self.assertIn(f"Workdir preserved: {harness.workdir}", harness.stdout)
        finally:
            harness.close()

    def test_static_pin_dispatches_only_for_exact_canonical_roots(self) -> None:
        payload = json.loads(SCENARIO_REGISTRY.read_text(encoding="utf-8"))
        scenarios = (
            payload["support_scenarios"] + payload["decision_scenarios"]
        )
        expected_roots = {
            "A_Calibrated_BAU",
            "B_Optimised_VRE",
            "C_Target_VRE",
        }
        for index, scenario in enumerate(scenarios):
            with self.subTest(scenario=scenario):
                module = _load_a3(f"pin_dispatch_{index}")
                harness = A3TraceHarness(module)
                try:
                    self.assertEqual(harness.run(scenario=scenario), 0)
                    pin_events = [
                        event
                        for event in harness.events
                        if event[0] == "stage_ws4_pwr_min_pin"
                    ]
                    self.assertEqual(len(pin_events), int(scenario in expected_roots))
                    if pin_events:
                        self.assertEqual(pin_events[0][1:3], (
                            harness.paths["s5"],
                            scenario,
                        ))
                finally:
                    harness.close()

    def test_static_pin_wrapper_uses_exact_assets_and_cli(self) -> None:
        module = _load_a3("pin_wrapper")
        with tempfile.TemporaryDirectory() as temp:
            rules_dir = Path(temp)
            script = rules_dir / "apply_base_year_pin.py"
            rules = rules_dir / "pwr_min_2023_2026_pin.csv"
            script.write_text("# fixture\n", encoding="utf-8")
            rules.write_text("fixture\n", encoding="utf-8")
            stage5 = rules_dir / "stage5"
            stage5.mkdir()
            for scenario in sorted(module.PIN_ROOT_SCENARIOS):
                with self.subTest(scenario=scenario):
                    with (
                        mock.patch.object(module, "RULES_SCRIPTS_DIR", rules_dir),
                        mock.patch.object(module, "SCENARIO_RULE_DATA", rules_dir),
                        mock.patch.object(module, "banner"),
                        mock.patch.object(module, "run_subproc") as run_subproc,
                    ):
                        module.stage_ws4_pwr_min_pin(stage5, scenario)
                    run_subproc.assert_called_once_with(
                        "ostram.pipeline.scenarios.rules.apply_base_year_pin",
                        [
                            "--input-dir",
                            stage5,
                            "--scenario",
                            scenario,
                            "--rules-csv",
                            rules,
                            "--skip-backup",
                        ],
                        label="apply base-year pin",
                    )

    def test_static_pin_wrapper_honors_profile_policy_gate(self) -> None:
        module = _load_a3("pin_policy_gate")
        with tempfile.TemporaryDirectory() as temp:
            rules_dir = Path(temp)
            (rules_dir / "apply_base_year_pin.py").write_text(
                "# fixture\n", encoding="utf-8"
            )
            (rules_dir / "pwr_min_2023_2026_pin.csv").write_text(
                "fixture\n", encoding="utf-8"
            )
            stage5 = rules_dir / "stage5"
            stage5.mkdir()
            cases = (
                ('{"apply_pwr_min_pin": false}', 0),
                ('{"apply_pwr_min_pin": true}', 1),
                (None, 1),
            )
            for policies, expected_calls in cases:
                with self.subTest(policies=policies):
                    environ = {"OSTRAM_PROFILE_POLICIES": policies} if policies else {}
                    with (
                        mock.patch.dict(os.environ, environ, clear=False),
                        mock.patch.object(module, "RULES_SCRIPTS_DIR", rules_dir),
                        mock.patch.object(module, "SCENARIO_RULE_DATA", rules_dir),
                        mock.patch.object(module, "banner"),
                        mock.patch.object(module, "run_subproc") as run_subproc,
                    ):
                        if policies is None:
                            os.environ.pop("OSTRAM_PROFILE_POLICIES", None)
                        module.stage_ws4_pwr_min_pin(stage5, "B_Optimised_VRE")
                    self.assertEqual(run_subproc.call_count, expected_calls)

    def test_static_pin_wrapper_fails_closed_on_scenario_or_missing_asset(self):
        module = _load_a3("pin_wrapper_failures")
        with tempfile.TemporaryDirectory() as temp:
            rules_dir = Path(temp)
            with (
                mock.patch.object(module, "RULES_SCRIPTS_DIR", rules_dir),
                mock.patch.object(module, "run_subproc") as run_subproc,
            ):
                with self.assertRaisesRegex(ValueError, "unsupported"):
                    module.stage_ws4_pwr_min_pin(rules_dir, "BAU")
                with self.assertRaisesRegex(FileNotFoundError, "asset missing"):
                    module.stage_ws4_pwr_min_pin(
                        rules_dir,
                        "A_Calibrated_BAU",
                    )
            run_subproc.assert_not_called()

    def test_static_pin_failure_prevents_stage6_and_delivery(self) -> None:
        module = _load_a3("pin_failure")
        harness = A3TraceHarness(module)
        failure = RuntimeError("pin fixture failure")
        try:
            with self.assertRaisesRegex(RuntimeError, "pin fixture failure"):
                harness.run(
                    fail_stage="stage_ws4_pwr_min_pin",
                    failure=failure,
                )
            names = [event[0] for event in harness.events]
            self.assertIn("stage_ws4_pwr_min_pin", names)
            self.assertNotIn("stage_6_sync_og_to_ts20", names)
            self.assertNotIn("stage_6_persist_restrictions", names)
            self.assertNotIn("deliver_outputs", names)
        finally:
            os.environ.pop("OSTRAM_TEMPLATE_PATH", None)
            harness.close()

    def test_expected_subprocess_failure_is_fail_fast_system_exit(self) -> None:
        module = _load_a3("subprocess_failure")
        completed = SimpleNamespace(
            returncode=9,
            stdout="before\nfixture stdout\n",
            stderr="fixture stderr\n",
        )
        with (
            mock.patch.object(module.subprocess, "run", return_value=completed) as process_run,
            redirect_stdout(io.StringIO()) as stdout,
        ):
            with self.assertRaisesRegex(SystemExit, "FAILED: fixture stage"):
                module.run_subproc(
                    "ostram.pipeline.fixture",
                    ["--flag", "value"],
                    cwd=Path("fixture-cwd"),
                    label="fixture stage",
                )
        process_run.assert_called_once_with(
            [
                sys.executable,
                "-B",
                "-m",
                "ostram.pipeline.fixture",
                "--flag",
                "value",
            ],
            cwd=str(Path("fixture-cwd").resolve()),
            env=mock.ANY,
            capture_output=True,
            text=True,
        )
        self.assertIn("--- stdout ---", stdout.getvalue())
        self.assertIn("--- stderr ---", stdout.getvalue())

    def test_stage_exit_exception_and_interrupt_leave_predecessor_cleanup_state(self) -> None:
        cases: tuple[BaseException, ...] = (
            SystemExit("FAILED: fixture stage"),
            RuntimeError("unexpected fixture failure"),
            KeyboardInterrupt(),
        )
        for index, failure in enumerate(cases):
            with self.subTest(failure=type(failure).__name__):
                module = _load_a3(f"abrupt_failure_{index}")
                harness = A3TraceHarness(module)
                try:
                    with self.assertRaises(type(failure)):
                        harness.run(
                            fail_stage="stage_2_and_2_5",
                            failure=failure,
                        )
                    self.assertTrue(harness.workdir.is_dir())
                    self.assertNotIn(("rmtree", harness.workdir, True), harness.events)
                    self.assertEqual(
                        os.environ.get("OSTRAM_TEMPLATE_PATH"),
                        str(harness.workdir / "_materialized_A_Calibrated_BAU.xlsx"),
                    )
                    self.assertFalse(
                        any(event[0] == "stage_3_fix_2" for event in harness.events)
                    )
                    self.assertFalse(
                        any(event[0] == "deliver_outputs" for event in harness.events)
                    )
                finally:
                    os.environ.pop("OSTRAM_TEMPLATE_PATH", None)
                    harness.close()

    def test_missing_snapshot_exits_before_workdir_or_transformation_dispatch(self) -> None:
        module = _load_a3("missing_snapshot")
        harness = A3TraceHarness(module)
        shutil.rmtree(harness.snapshot)
        try:
            with self.assertRaisesRegex(SystemExit, "snapshot post-A2 not found"):
                harness.run()
            self.assertEqual(
                harness.events,
                [
                    (
                        "resolve_config",
                        "A_Calibrated_BAU",
                        harness.soasia,
                        harness.caller,
                    )
                ],
            )
            self.assertFalse(harness.workdir.exists())
            self.assertNotIn("OSTRAM_TEMPLATE_PATH", os.environ)
        finally:
            harness.close()


if __name__ == "__main__":
    unittest.main()
