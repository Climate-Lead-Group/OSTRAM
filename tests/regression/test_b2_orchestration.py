from __future__ import annotations

import ast
import importlib.util
import io
import itertools
import json
import multiprocessing
import os
import shutil
import subprocess
import sys
import tempfile
import types
import unittest
from contextlib import ExitStack, redirect_stderr, redirect_stdout
from pathlib import Path
from types import SimpleNamespace
from unittest import mock


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
B2_ENTRYPOINT = REPO_ROOT / "ostram" / "pipeline" / "execution" / "runner.py"
B2_ORCHESTRATOR = REPO_ROOT / "ostram" / "pipeline" / "execution" / "orchestrator.py"


def _source() -> str:
    return B2_ENTRYPOINT.read_text(encoding="utf-8-sig")


def _orchestrator_source() -> str:
    return B2_ORCHESTRATOR.read_text(encoding="utf-8-sig")


def _main_guard(tree: ast.Module) -> ast.If:
    for node in tree.body:
        if not isinstance(node, ast.If) or not isinstance(node.test, ast.Compare):
            continue
        if isinstance(node.test.left, ast.Name) and node.test.left.id == "__name__":
            return node
    raise AssertionError("B2 entrypoint has no __main__ guard")


def _call_name(node: ast.Call) -> str:
    def dotted(expr: ast.expr) -> str:
        if isinstance(expr, ast.Name):
            return expr.id
        if isinstance(expr, ast.Attribute):
            prefix = dotted(expr.value)
            return f"{prefix}.{expr.attr}" if prefix else expr.attr
        return ""

    return dotted(node.func)


def _selected_calls(node: ast.AST, selected: set[str]) -> list[str]:
    calls = [item for item in ast.walk(node) if isinstance(item, ast.Call)]
    calls.sort(key=lambda item: (item.lineno, item.col_offset))
    return [_call_name(item) for item in calls if _call_name(item) in selected]


def _function(tree: ast.Module, name: str) -> ast.FunctionDef:
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return node
    raise AssertionError(f"function {name!r} not found")


def _load_b2(label: str):
    module_name = f"ostram.pipeline.execution._characterization_{label}"
    spec = importlib.util.spec_from_file_location(module_name, B2_ENTRYPOINT)
    if spec is None or spec.loader is None:
        raise AssertionError(f"could not load module spec for {B2_ENTRYPOINT}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    finally:
        sys.modules.pop(module_name, None)
    return module


def _load_b2_guard_as_callable(label: str):
    """Expose the guarded entrypoint body without executing it at import time."""
    source = _source()
    tree = ast.parse(source, filename=str(B2_ENTRYPOINT))
    guard = _main_guard(tree)
    guard_index = tree.body.index(guard)
    callable_guard = ast.FunctionDef(
        name="_characterized_main_guard",
        args=ast.arguments(
            posonlyargs=[],
            args=[],
            kwonlyargs=[],
            kw_defaults=[],
            defaults=[],
        ),
        body=guard.body,
        decorator_list=[],
        returns=None,
        type_comment=None,
    )
    ast.copy_location(callable_guard, guard)
    tree.body[guard_index] = callable_guard
    ast.fix_missing_locations(tree)

    module_name = f"ostram.pipeline.execution._guard_characterization_{label}"
    module = types.ModuleType(module_name)
    module.__file__ = str(B2_ENTRYPOINT)
    module.__package__ = "ostram.pipeline.execution"
    sys.modules[module_name] = module
    try:
        exec(compile(tree, str(B2_ENTRYPOINT), "exec"), module.__dict__)
    finally:
        sys.modules.pop(module_name, None)
    return module


def _base_params(**overrides: object) -> dict[str, object]:
    params: dict[str, object] = {
        "A2_output": "A2_Output_Params",
        "A2_output_otoole": "A2_Outputs_Params_otoole",
        "Miscellaneous": "Miscellaneous",
        "templates": "templates",
        "executables": "Executables",
        "outputs": "Outputs",
        "concatenate_folder": "concatenate_files",
        "otoole_config": "conversion_format.yaml",
        "preprocess_data": "preprocess_data.py",
        "osemosys_model": "osemosys_fast_preprocessed.txt",
        "conv_format": "conversion_format.yaml",
        "concat_csvs": "concatenate_ostram.py",
        "inputs_file": "Inputs.csv",
        "outputs_file": "Outputs.csv",
        "prefix_final_files": "OSTRAM_",
        "preprocess_data_name": "Pre_processed_",
        "output_files": "_output",
        "solver": "cplex",
        "iteration_time": 20000,
        "cbc_random_seed": 12345,
        "cplex_threads": 4,
        "cplex_random_seed": 12345,
        "gurobi_threads": 3,
        "gurobi_seed": 12345,
        "glpk_option": "new",
        "del_files": False,
        "only_main_scenario": False,
        "parallel": False,
        "max_x_per_iter": 2,
        "A2_otoole_outputs": True,
        "write_txt_model": True,
        "create_matrix": False,
        "execute_model": False,
        "reuse_existing_sol": False,
        "concat_otoole_csv": False,
        "concat_scenarios_csv": False,
        "annualize_capital": False,
        "storage_delay_active": False,
        "strip_storage_active": False,
        "open_pwrbck_active": False,
        "reserve_margin_repair_active": False,
        "reserve_margin_xlsx_active": False,
    }
    params.update(overrides)
    return params


class B2Fixture:
    def __init__(
        self,
        *entries: str,
        main_scenario: str = "A",
        file_entries: tuple[str, ...] = (),
        **overrides: object,
    ) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name).resolve()
        self.entrypoint = self.root / "python -m ostram run"
        self.entrypoint.write_text("# fixture anchor only\n", encoding="utf-8")
        self.params = _base_params(**overrides)
        self.main_scenario = main_scenario

        scenario_root = self.root / str(self.params["A2_output"])
        scenario_root.mkdir()
        file_names = set(file_entries)
        for entry in entries:
            path = scenario_root / entry
            if entry in file_names:
                path.write_text("fixture entry\n", encoding="utf-8")
            else:
                path.mkdir()

        (self.root / "Config_MOMF_T1_AB.yaml").write_text(
            json.dumps(self.params), encoding="utf-8"
        )
        (self.root / "Config_MOMF_T1_A.yaml").write_text(
            json.dumps({"xtra_scen": {"Main_Scenario": main_scenario}}),
            encoding="utf-8",
        )

    def close(self) -> None:
        self._temp.cleanup()

    def __enter__(self) -> "B2Fixture":
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()


class GuardHarness:
    def __init__(
        self,
        module,
        fixture: B2Fixture,
        *,
        conversions: dict[str, bool] | None = None,
        solver_boundary=None,
    ) -> None:
        self.module = module
        self.fixture = fixture
        self.conversions = {} if conversions is None else dict(conversions)
        self.solver_boundary = solver_boundary
        self.events: list[tuple[object, ...]] = []
        self.params_seen: list[dict[str, object]] = []
        self.stdout = ""
        self.stderr = ""
        self.cwd_after: Path | None = None

    def _remember(self, params: dict[str, object]) -> None:
        self.params_seen.append(params)

    def _process(self, **kwargs) -> None:
        self._remember(kwargs.pop("params", {}) if "params" in kwargs else {})
        self.events.append(
            (
                "process",
                kwargs["base_input_path"],
                kwargs["template_path"],
                kwargs["base_output_path"],
                kwargs["scenario_name"],
            )
        )

    def _convert(self, **kwargs) -> bool:
        params = kwargs["params"]
        self._remember(params)
        scenario = kwargs["scenario_name"]
        self.events.append(("convert", kwargs["base_output_path"], scenario))
        return self.conversions.get(scenario, True)

    def _stage(self, name: str):
        def run(params, scenario_name) -> None:
            self._remember(params)
            self.events.append((name, scenario_name))

        return run

    def _combined(self, input_folder, output_folder, scenario_name):
        self.events.append(
            ("combined", input_folder, output_folder, scenario_name)
        )
        return str(Path(output_folder) / f"{scenario_name}_Input.csv"), None

    def _export(self, here, params, scenario_name, export_name=None):
        self._remember(params)
        self.events.append(("export", here, scenario_name, export_name))
        return Path(here).parent / (export_name or "OSTRAM_data.txt")

    def _solver(self, params, scenario_name, here) -> None:
        self._remember(params)
        self.events.append(("solver_boundary", scenario_name, here))
        if self.solver_boundary is not None:
            self.solver_boundary(params, scenario_name, here)

    def _concat_scenarios(self, here, params):
        self._remember(params)
        self.events.append(("concat_scenarios", here))
        return "inputs.csv", "outputs.csv", "combined.csv"

    def run(self, argv: list[str]) -> None:
        previous_cwd = Path.cwd()
        stdout = io.StringIO()
        stderr = io.StringIO()
        patches = {
            "process_scenario_folder": self._process,
            "run_otoole_conversion": self._convert,
            "run_preprocessing_script": self._stage("preprocess"),
            "run_days_in_day_type_patcher": self._stage("days"),
            "run_storage_delay_patcher": self._stage("storage_delay"),
            "run_strip_storage_patcher": self._stage("strip_storage"),
            "run_open_pwrbck_patcher": self._stage("open_pwrbck"),
            "run_reserve_margin_repair_patcher": self._stage("reserve_margin"),
            "run_reserve_margin_xlsx_patcher": self._stage("reserve_margin_xlsx"),
            "generate_combined_input_file": self._combined,
            "export_root_datafile": self._export,
            "main_executer": self._solver,
            "concatenate_all_scenarios": self._concat_scenarios,
        }
        try:
            with ExitStack() as stack:
                for name, replacement in patches.items():
                    stack.enter_context(
                        mock.patch.object(self.module, name, replacement)
                    )
                stack.enter_context(
                    mock.patch.object(
                        self.module.b2_orchestrator,
                        "resolve_here",
                        return_value=self.fixture.root,
                    )
                )
                stack.enter_context(
                    mock.patch.object(
                        self.module.time,
                        "time",
                        side_effect=(10.0, 14.0, 20.0, 23.0),
                    )
                )
                stack.enter_context(mock.patch.object(sys, "argv", argv))
                stack.enter_context(redirect_stdout(stdout))
                stack.enter_context(redirect_stderr(stderr))
                self.module._characterized_main_guard()
        finally:
            self.stdout = stdout.getvalue()
            self.stderr = stderr.getvalue()
            self.cwd_after = Path.cwd()
            os.chdir(previous_cwd)


class B2ImportAndCliCharacterizationTests(unittest.TestCase):
    def test_import_is_silent_and_crosses_no_effect_boundary(self) -> None:
        with (
            mock.patch.object(subprocess, "run") as process_runner,
            mock.patch.object(multiprocessing, "Process") as process_factory,
            mock.patch.object(os, "chdir") as change_directory,
            mock.patch.object(os, "makedirs") as make_directories,
            mock.patch.object(shutil, "copy2") as copy_file,
            mock.patch.object(shutil, "rmtree") as remove_tree,
            redirect_stdout(io.StringIO()) as stdout,
            redirect_stderr(io.StringIO()) as stderr,
        ):
            module = _load_b2("import_safety")

        process_runner.assert_not_called()
        process_factory.assert_not_called()
        change_directory.assert_not_called()
        make_directories.assert_not_called()
        copy_file.assert_not_called()
        remove_tree.assert_not_called()
        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")
        self.assertTrue(callable(module.main_executer))

    def test_help_unknown_option_and_missing_value_keep_argparse_contract(self) -> None:
        cases = (
            (["python -m ostram run", "--help"], 0),
            (["python -m ostram run", "--unknown"], 2),
            (["python -m ostram run", "--scenarios"], 2),
        )
        for index, (argv, expected_code) in enumerate(cases):
            with self.subTest(argv=argv):
                module = _load_b2_guard_as_callable(f"cli_{index}")
                with (
                    mock.patch.object(sys, "argv", argv),
                    mock.patch.object(module.subprocess, "run") as process_runner,
                    redirect_stdout(io.StringIO()) as stdout,
                    redirect_stderr(io.StringIO()),
                    self.assertRaises(SystemExit) as raised,
                ):
                    module._characterized_main_guard()

                self.assertEqual(raised.exception.code, expected_code)
                process_runner.assert_not_called()
                if expected_code == 0:
                    help_text = stdout.getvalue()
                    self.assertIn("Execute OSeMOSYS model across scenarios", help_text)
                    self.assertIn("--scenarios SCENARIOS", help_text)
                    self.assertIn("When omitted, runs", help_text)
                    self.assertIn("all scenarios found in the A2 output directory.", help_text)


class B2AstBoundaryCharacterizationTests(unittest.TestCase):
    def test_top_level_stage_order_includes_both_solver_routes_and_postprocessing(self) -> None:
        entry_tree = ast.parse(_source(), filename=str(B2_ENTRYPOINT))
        orchestrator_tree = ast.parse(
            _orchestrator_source(), filename=str(B2_ORCHESTRATOR)
        )
        self.assertEqual(
            _selected_calls(_main_guard(entry_tree), {"main"}),
            ["main"],
        )
        self.assertEqual(
            _selected_calls(
                _function(entry_tree, "main"),
                {"b2_orchestrator.orchestrate_b2"},
            ),
            ["b2_orchestrator.orchestrate_b2"],
        )

        top_level = _function(orchestrator_tree, "orchestrate_b2")
        self.assertEqual(
            _selected_calls(
                top_level,
                {
                    "build_run_plan",
                    "run_compiled_input_stage",
                    "run_execution_stage",
                    "run_cleanup_stage",
                    "run_final_postprocessing_stage",
                },
            ),
            [
                "build_run_plan",
                "run_compiled_input_stage",
                "run_execution_stage",
                "run_cleanup_stage",
                "run_final_postprocessing_stage",
            ],
        )

        compiled_input = _function(orchestrator_tree, "run_compiled_input_stage")
        compiled_calls = {
            "dependencies.process_scenario_folder",
            "dependencies.run_otoole_conversion",
            "dependencies.run_preprocessing_script",
            "dependencies.run_days_in_day_type_patcher",
            "dependencies.run_storage_delay_patcher",
            "dependencies.run_strip_storage_patcher",
            "dependencies.run_open_pwrbck_patcher",
            "dependencies.run_reserve_margin_repair_patcher",
            "dependencies.run_reserve_margin_xlsx_patcher",
            "dependencies.generate_combined_input_file",
            "dependencies.export_root_datafile",
        }
        self.assertEqual(
            _selected_calls(compiled_input, compiled_calls),
            [
                "dependencies.process_scenario_folder",
                "dependencies.run_otoole_conversion",
                "dependencies.run_preprocessing_script",
                "dependencies.run_days_in_day_type_patcher",
                "dependencies.run_storage_delay_patcher",
                "dependencies.run_strip_storage_patcher",
                "dependencies.run_open_pwrbck_patcher",
                "dependencies.run_reserve_margin_repair_patcher",
                "dependencies.run_reserve_margin_xlsx_patcher",
                "dependencies.generate_combined_input_file",
                "dependencies.export_root_datafile",
            ],
        )

        execution = _function(orchestrator_tree, "run_execution_stage")
        self.assertEqual(
            _selected_calls(
                execution,
                {
                    "dependencies.chunk_scenarios",
                    "dependencies.mp_module.Process",
                    "dependencies.main_executer",
                },
            ),
            [
                "dependencies.chunk_scenarios",
                "dependencies.mp_module.Process",
                "dependencies.main_executer",
            ],
        )

        cleanup = _function(orchestrator_tree, "run_cleanup_stage")
        self.assertEqual(
            _selected_calls(cleanup, {"dependencies.delete_files"}),
            ["dependencies.delete_files"],
        )
        postprocessing = _function(
            orchestrator_tree, "run_final_postprocessing_stage"
        )
        self.assertEqual(
            _selected_calls(
                postprocessing,
                {
                    "dependencies.concatenate_all_scenarios",
                    "dependencies.load_annualizer",
                    "annualize_capital_investment",
                    "shutil.copy2",
                },
            ),
            [
                "dependencies.concatenate_all_scenarios",
                "dependencies.load_annualizer",
                "annualize_capital_investment",
                "shutil.copy2",
                "shutil.copy2",
            ],
        )

    def test_exact_outer_and_parallel_guards_dominate_the_two_executor_routes(self) -> None:
        source = _orchestrator_source()
        tree = ast.parse(source, filename=str(B2_ORCHESTRATOR))
        execution = _function(tree, "run_execution_stage")
        execution_source = ast.get_source_segment(source, execution)
        assert execution_source is not None
        self.assertIn(
            'if params["execute_model"] or params["create_matrix"]:',
            execution_source,
        )
        self.assertIn('if params["parallel"]:', execution_source)

        process_calls = [
            call
            for call in ast.walk(execution)
            if isinstance(call, ast.Call)
            and _call_name(call) == "dependencies.mp_module.Process"
        ]
        direct_calls = [
            call
            for call in ast.walk(execution)
            if isinstance(call, ast.Call)
            and _call_name(call) == "dependencies.main_executer"
        ]
        self.assertEqual(len(process_calls), 1)
        self.assertEqual(len(direct_calls), 1)
        target = next(
            keyword.value
            for keyword in process_calls[0].keywords
            if keyword.arg == "target"
        )
        self.assertIsInstance(target, ast.Attribute)
        self.assertEqual(target.attr, "main_executer")
        self.assertIsInstance(target.value, ast.Name)
        self.assertEqual(target.value.id, "dependencies")

    def test_named_matrix_and_solver_boundaries_are_injectable(self) -> None:
        entry_tree = ast.parse(_source(), filename=str(B2_ENTRYPOINT))
        entry_executor = _function(entry_tree, "main_executer")
        self.assertEqual(
            _selected_calls(
                entry_executor,
                {
                    "b2_orchestrator.ScenarioExecutionDependencies",
                    "b2_orchestrator.execute_scenario",
                },
            ),
            [
                "b2_orchestrator.ScenarioExecutionDependencies",
                "b2_orchestrator.execute_scenario",
            ],
        )

        tree = ast.parse(_orchestrator_source(), filename=str(B2_ORCHESTRATOR))
        for name in ("run_matrix_preparation", "invoke_solver_command"):
            boundary = _function(tree, name)
            self.assertEqual(
                _selected_calls(boundary, {"process_runner"}),
                ["process_runner"],
            )
            process_call = next(
                item for item in ast.walk(boundary) if isinstance(item, ast.Call)
            )
            keyword_names = [keyword.arg for keyword in process_call.keywords]
            self.assertEqual(keyword_names, ["cwd", "check"])
            self.assertIsInstance(process_call.keywords[0].value, ast.Call)
            self.assertIs(process_call.keywords[1].value.value, True)


class B2ScenarioAndTraceCharacterizationTests(unittest.TestCase):
    def test_predecessor_trace_freezes_scenario_order_stage_args_and_status_order(self) -> None:
        module = _load_b2_guard_as_callable("predecessor_trace")
        with B2Fixture(
            "C",
            "Default",
            "A",
            main_scenario="A",
            execute_model=True,
            concat_scenarios_csv=True,
        ) as fixture:
            harness = GuardHarness(module, fixture)
            caller_cwd = Path.cwd()
            harness.run(
                ["python -m ostram run", "--scenarios", " C, A, C "]
            )

            base_input = str(fixture.root / "A2_Output_Params")
            template = str(fixture.root / "Miscellaneous" / "templates")
            base_output = str(fixture.root / "A2_Outputs_Params_otoole")
            expected: list[tuple[object, ...]] = []
            for scenario in ("C", "A"):
                expected.extend(
                    [
                        ("process", base_input, template, base_output, scenario),
                        ("convert", base_output, scenario),
                        ("preprocess", scenario),
                        ("days", scenario),
                        ("storage_delay", scenario),
                        ("strip_storage", scenario),
                        ("open_pwrbck", scenario),
                        ("reserve_margin", scenario),
                        ("reserve_margin_xlsx", scenario),
                        (
                            "combined",
                            str(Path(base_output) / scenario),
                            str(fixture.root / "Executables" / f"{scenario}_0"),
                            f"{scenario}_0",
                        ),
                    ]
                )
            expected.extend(
                [
                    ("export", fixture.root, "A", None),
                    ("solver_boundary", "C", fixture.root),
                    ("solver_boundary", "A", fixture.root),
                    ("concat_scenarios", fixture.root),
                ]
            )
            self.assertEqual(harness.events, expected)
            self.assertTrue(harness.params_seen)
            first_params = next(item for item in harness.params_seen if item)
            self.assertTrue(all(item is first_params for item in harness.params_seen if item))
            self.assertEqual(first_params["execute_model"], True)
            self.assertEqual(first_params["create_matrix"], False)
            self.assertEqual(harness.cwd_after, caller_cwd)

            markers = (
                "[INFO] Scenario filter active: ['C', 'A']",
                "Started linear executions",
                "Inputs and outputs concatenated for all scenarios successfully",
                "For all effects, we have finished the work of this script",
            )
            positions = [harness.stdout.index(marker) for marker in markers]
            self.assertEqual(positions, sorted(positions))

    def test_compile_only_stops_before_every_solver_and_result_boundary(self) -> None:
        module = _load_b2_guard_as_callable("compile_only_gate")
        with B2Fixture(
            "A",
            execute_model=True,
            create_matrix=True,
            concat_otoole_csv=True,
            concat_scenarios_csv=True,
            annualize_capital=True,
            del_files=True,
        ) as fixture:
            harness = GuardHarness(module, fixture)
            harness.run(
                [
                    "python -m ostram run",
                    "--scenarios",
                    "A",
                    "--compile-only",
                ]
            )

            self.assertEqual(
                [event[0] for event in harness.events],
                [
                    "process",
                    "convert",
                    "preprocess",
                    "days",
                    "storage_delay",
                    "strip_storage",
                    "open_pwrbck",
                    "reserve_margin",
                    "reserve_margin_xlsx",
                    "combined",
                    "export",
                ],
            )
            self.assertNotIn(
                "solver_boundary", [event[0] for event in harness.events]
            )
            self.assertNotIn(
                "concat_scenarios", [event[0] for event in harness.events]
            )
            params = next(item for item in harness.params_seen if item)
            for key in (
                "execute_model",
                "create_matrix",
                "reuse_existing_sol",
                "concat_otoole_csv",
                "concat_scenarios_csv",
                "annualize_capital",
                "del_files",
            ):
                self.assertFalse(params[key], key)
            self.assertIn("Compile-only gate complete", harness.stdout)
            self.assertNotIn(
                "For all effects, we have finished the work of this script",
                harness.stdout,
            )

    def test_unknowns_preserve_duplicates_and_abort_before_any_stage(self) -> None:
        module = _load_b2_guard_as_callable("unknown_scenarios")
        with B2Fixture("B", "Default", "A") as fixture:
            harness = GuardHarness(module, fixture)
            with self.assertRaises(SystemExit) as raised:
                harness.run(
                    [
                        "python -m ostram run",
                        "--scenarios",
                        "Missing,A,Missing,Other",
                    ]
                )

            self.assertEqual(raised.exception.code, 1)
            self.assertEqual(harness.events, [])
            self.assertIn("['Missing', 'Missing', 'Other']", harness.stdout)
            self.assertIn("Discovered: ['A', 'B']", harness.stdout)

    def test_truthy_empty_filter_selects_nothing_but_finishes(self) -> None:
        module = _load_b2_guard_as_callable("truthy_empty")
        with B2Fixture("B", "Default", "A") as fixture:
            harness = GuardHarness(module, fixture)
            harness.run(["python -m ostram run", "--scenarios", " , , "])

            self.assertEqual(harness.events, [])
            self.assertIn("[INFO] Scenario filter active: []", harness.stdout)
            self.assertIn(
                "For all effects, we have finished the work of this script",
                harness.stdout,
            )

    def test_only_main_scenario_replaces_discovery_before_cli_validation(self) -> None:
        module = _load_b2_guard_as_callable("only_main")
        with B2Fixture(
            "A", "B", "Default", main_scenario="B", only_main_scenario=True
        ) as fixture:
            selected = GuardHarness(module, fixture)
            selected.run(["python -m ostram run"])
            self.assertEqual(
                [event[-1] for event in selected.events if event[0] == "process"],
                ["B"],
            )

        module = _load_b2_guard_as_callable("only_main_filter")
        with B2Fixture(
            "A", "B", "Default", main_scenario="B", only_main_scenario=True
        ) as fixture:
            rejected = GuardHarness(module, fixture)
            with self.assertRaises(SystemExit) as raised:
                rejected.run(
                    ["python -m ostram run", "--scenarios", "A"]
                )
            self.assertEqual(raised.exception.code, 1)
            self.assertIn("Discovered: ['B']", rejected.stdout)
            self.assertEqual(rejected.events, [])

    def test_discovery_keeps_non_directory_entries_and_only_removes_exact_default(self) -> None:
        module = _load_b2_guard_as_callable("raw_discovery")
        with B2Fixture(
            "Default", "A", "Z_file", file_entries=("Z_file",)
        ) as fixture:
            harness = GuardHarness(module, fixture)
            harness.run(["python -m ostram run"])

            self.assertEqual(
                [event[-1] for event in harness.events if event[0] == "process"],
                ["A", "Z_file"],
            )

    def test_failed_conversion_skips_remaining_scenario_work_but_continues(self) -> None:
        module = _load_b2_guard_as_callable("conversion_continue")
        with B2Fixture("A", "B", main_scenario="A") as fixture:
            harness = GuardHarness(module, fixture, conversions={"A": False})
            harness.run(["python -m ostram run"])

            names_by_scenario = [
                (event[0], event[-1])
                for event in harness.events
                if event[0]
                in {
                    "process",
                    "convert",
                    "preprocess",
                    "days",
                    "storage_delay",
                    "strip_storage",
                    "open_pwrbck",
                    "reserve_margin",
                    "reserve_margin_xlsx",
                }
            ]
            self.assertEqual(
                names_by_scenario,
                [
                    ("process", "A"),
                    ("convert", "A"),
                    ("process", "B"),
                    ("convert", "B"),
                    ("preprocess", "B"),
                    ("days", "B"),
                    ("storage_delay", "B"),
                    ("strip_storage", "B"),
                    ("open_pwrbck", "B"),
                    ("reserve_margin", "B"),
                    ("reserve_margin_xlsx", "B"),
                ],
            )
            self.assertNotIn(
                "A_0", [event[-1] for event in harness.events if event[0] == "combined"]
            )
            self.assertIn(
                ("export", fixture.root, "A", None), harness.events
            )
            self.assertIn("Skipping preprocessing for 'A'", harness.stdout)


class B2ConfigurationMatrixCharacterizationTests(unittest.TestCase):
    def test_all_sixteen_execution_and_concatenation_combinations(self) -> None:
        for execute_model, create_matrix, concat_otoole, concat_scenarios in itertools.product(
            (False, True), repeat=4
        ):
            with self.subTest(
                execute_model=execute_model,
                create_matrix=create_matrix,
                concat_otoole=concat_otoole,
                concat_scenarios=concat_scenarios,
            ):
                module = _load_b2_guard_as_callable(
                    f"matrix_{int(execute_model)}{int(create_matrix)}"
                    f"{int(concat_otoole)}{int(concat_scenarios)}"
                )
                dispatched: list[str] = []

                if not execute_model and not create_matrix:
                    def dispatch(params, scenario_name, here) -> None:
                        raise AssertionError(
                            "solver boundary reached while execute_model and "
                            "create_matrix were both disabled"
                        )
                else:
                    def dispatch(params, scenario_name, here) -> None:
                        dispatched.append("boundary")
                        if params["create_matrix"]:
                            dispatched.append("matrix")
                        if params["execute_model"]:
                            dispatched.extend(("solve", "results"))
                        if params["concat_otoole_csv"]:
                            dispatched.append("concat_otoole")

                with B2Fixture(
                    "A",
                    execute_model=execute_model,
                    create_matrix=create_matrix,
                    concat_otoole_csv=concat_otoole,
                    concat_scenarios_csv=concat_scenarios,
                    A2_otoole_outputs=False,
                    write_txt_model=False,
                ) as fixture:
                    harness = GuardHarness(
                        module, fixture, solver_boundary=dispatch
                    )
                    harness.run(["python -m ostram run"])

                expected_dispatch: list[str] = []
                if execute_model or create_matrix:
                    expected_dispatch.append("boundary")
                    if create_matrix:
                        expected_dispatch.append("matrix")
                    if execute_model:
                        expected_dispatch.extend(("solve", "results"))
                    if concat_otoole:
                        expected_dispatch.append("concat_otoole")
                self.assertEqual(dispatched, expected_dispatch)
                self.assertEqual(
                    sum(event[0] == "solver_boundary" for event in harness.events),
                    int(execute_model or create_matrix),
                )
                self.assertEqual(
                    sum(event[0] == "concat_scenarios" for event in harness.events),
                    int(concat_scenarios),
                )

    def test_explicit_solver_adapter_is_fail_closed_for_compile_only_and_matrix_only(
        self,
    ) -> None:
        module = _load_b2("solver_adapter_fail_closed")
        orchestrator = module.b2_orchestrator

        class SentinelSolverAdapter:
            def prepare_command(self, *args, **kwargs):
                raise AssertionError("solver command preparation was reachable")

            def invoke(self, *args, **kwargs):
                raise AssertionError("solver invocation was reachable")

        def reject_process(*args, **kwargs):
            raise AssertionError("external process boundary was reachable")

        dependencies = orchestrator.ScenarioExecutionDependencies(
            run_process=reject_process,
            check_environment=lambda solver: None,
            get_executable=lambda executable: f"fixture-{executable}",
            path_exists=lambda path: False,
            remove_file=lambda path: None,
            python_executable=sys.executable,
        )
        root = Path("C:/fixture/execution_workspace")
        sentinel = SentinelSolverAdapter()

        with redirect_stdout(io.StringIO()):
            orchestrator.execute_scenario(
                _base_params(
                    solver="cplex",
                    execute_model=False,
                    create_matrix=False,
                    concat_otoole_csv=False,
                ),
                "A",
                root,
                dependencies,
                solver_adapter=sentinel,
                matrix_runner=reject_process,
            )

        matrix_commands: list[str] = []

        def record_matrix(command, process_runner) -> None:
            self.assertIs(process_runner, reject_process)
            matrix_commands.append(command)

        with redirect_stdout(io.StringIO()):
            orchestrator.execute_scenario(
                _base_params(
                    solver="cplex",
                    execute_model=False,
                    create_matrix=True,
                    concat_otoole_csv=False,
                ),
                "A",
                root,
                dependencies,
                solver_adapter=sentinel,
                matrix_runner=record_matrix,
            )

        self.assertEqual(len(matrix_commands), 1)
        self.assertIn("glpsol", matrix_commands[0])
        self.assertIn("--check", matrix_commands[0])

    def test_parallel_route_batches_process_targets_and_ignores_child_exitcodes(self) -> None:
        module = _load_b2_guard_as_callable("parallel_route")
        process_events: list[tuple[str, str]] = []

        class FakeProcess:
            def __init__(self, *, target, args) -> None:
                self.target = target
                self.args = args
                self.scenario = args[1]
                self.exitcode = 9
                process_events.append(("construct", self.scenario))

            def start(self) -> None:
                process_events.append(("start", self.scenario))
                self.target(*self.args)

            def join(self) -> None:
                process_events.append(("join", self.scenario))

        with B2Fixture(
            "C",
            "A",
            "B",
            execute_model=True,
            parallel=True,
            max_x_per_iter=2,
            concat_scenarios_csv=True,
            A2_otoole_outputs=False,
            write_txt_model=False,
        ) as fixture:
            harness = GuardHarness(module, fixture)
            with mock.patch.object(module.mp, "Process", FakeProcess):
                harness.run(["python -m ostram run"])

            self.assertEqual(
                process_events,
                [
                    ("construct", "A"),
                    ("start", "A"),
                    ("construct", "B"),
                    ("start", "B"),
                    ("join", "A"),
                    ("join", "B"),
                    ("construct", "C"),
                    ("start", "C"),
                    ("join", "C"),
                ],
            )
            self.assertEqual(
                [event[1] for event in harness.events if event[0] == "solver_boundary"],
                ["A", "B", "C"],
            )
            self.assertEqual(harness.events[-1], ("concat_scenarios", fixture.root))
            self.assertIn("Started parallelization of model execution", harness.stdout)

    def test_linear_solver_failure_propagates_and_stops_cleanup_and_postprocessing(self) -> None:
        module = _load_b2_guard_as_callable("linear_failure")
        solver_calls: list[str] = []

        def fail_first(params, scenario_name, here) -> None:
            solver_calls.append(scenario_name)
            raise RuntimeError("fixture solver boundary failure")

        with B2Fixture(
            "A",
            "B",
            execute_model=True,
            concat_scenarios_csv=True,
            A2_otoole_outputs=False,
            write_txt_model=False,
        ) as fixture:
            harness = GuardHarness(module, fixture, solver_boundary=fail_first)
            with self.assertRaisesRegex(RuntimeError, "solver boundary failure"):
                harness.run(["python -m ostram run"])

            self.assertEqual(solver_calls, ["A"])
            self.assertNotIn(
                "concat_scenarios", [event[0] for event in harness.events]
            )
            self.assertNotIn(
                "For all effects, we have finished the work of this script",
                harness.stdout,
            )


class B2MainExecutorCommandCharacterizationTests(unittest.TestCase):
    def _run_executor(
        self,
        module,
        root: Path,
        params: dict[str, object],
        *,
        solution_exists: bool = True,
        process_side_effect=None,
    ):
        removed: list[str] = []

        def exists(path) -> bool:
            return solution_exists and str(path).endswith(".sol")

        def remove(path) -> None:
            removed.append(str(path))

        with (
            mock.patch.object(module.os.path, "exists", side_effect=exists),
            mock.patch.object(module.os, "remove", side_effect=remove),
            mock.patch.object(module, "check_enviro_variables") as check_environment,
            mock.patch.object(
                module,
                "get_env_executable",
                side_effect=lambda name: f"ENV-{name}",
            ),
            mock.patch.object(
                module.subprocess,
                "run",
                side_effect=process_side_effect,
            ) as process_runner,
            redirect_stdout(io.StringIO()) as stdout,
        ):
            module.main_executer(params, "A", root)

        return SimpleNamespace(
            runner=process_runner,
            check_environment=check_environment,
            removed=removed,
            stdout=stdout.getvalue(),
        )

    def test_exact_solver_command_for_every_supported_solver(self) -> None:
        module = _load_b2("solver_commands")
        with tempfile.TemporaryDirectory() as temp:
            here = Path(temp).resolve() / "execution_workspace"
            folder = os.path.join(str(here), "Executables", "A_0")
            data_file = os.path.join(folder, "Pre_processed_A_0")
            output_file = os.path.join(folder, "Pre_processed_A_0_output")
            expected = {
                "glpk": [
                    "glpsol", "-m", "osemosys_fast_preprocessed.txt",
                    "-d", f"{data_file}.txt", "--wglp", f"{output_file}.glp",
                    "--write", f"{output_file}.sol",
                ],
                "cbc": [
                    "cbc", f"{output_file}.lp", "randomSeed", "12345",
                    "randomCbcSeed", "12345", "-seconds", "20000",
                    "solve", "-solu", f"{output_file}.sol",
                ],
                "cplex": [
                    "cplex", "-c", f"set logfile {output_file}.cplex.log",
                    f"read {output_file}.lp", "set threads 4",
                    "set randomseed 12345", "set parallel 1", "optimize",
                    f"write {output_file}.sol",
                ],
                "gurobi": [
                    "gurobi_cl", "Threads=3", "Seed=12345",
                    f"ResultFile={output_file}.sol", f"{output_file}.lp",
                ],
            }
            environment_name = {
                "glpk": "glpsol",
                "cbc": "cbc",
                "cplex": "cplex",
                "gurobi": "gurobi_cl",
            }

            for solver, expected_command in expected.items():
                with self.subTest(solver=solver):
                    params = _base_params(
                        solver=solver,
                        execute_model=True,
                        create_matrix=False,
                        concat_otoole_csv=False,
                    )
                    result = self._run_executor(module, here, params)
                    commands = [call.args[0] for call in result.runner.call_args_list]
                    self.assertEqual(commands[0], expected_command)
                    self.assertEqual(len(commands), 2)
                    for recorded in result.runner.call_args_list:
                        self.assertTrue(recorded.kwargs["check"])
                        self.assertNotIn("shell", recorded.kwargs)
                        self.assertTrue(Path(recorded.kwargs["cwd"]).is_absolute())
                    result.check_environment.assert_called_once_with(
                        environment_name[solver]
                    )

    def test_cplex_matrix_solve_results_and_concat_commands_keep_exact_order(self) -> None:
        module = _load_b2("cplex_full_chain")
        with tempfile.TemporaryDirectory() as temp:
            here = Path(temp).resolve() / "execution_workspace"
            params = _base_params(
                solver="cplex",
                execute_model=True,
                create_matrix=True,
                concat_otoole_csv=True,
            )
            result = self._run_executor(module, here, params)

            folder = os.path.join(str(here), "Executables", "A_0")
            data_file = os.path.join(folder, "Pre_processed_A_0")
            output_file = os.path.join(folder, "Pre_processed_A_0_output")
            outputs = os.path.join(folder, "Outputs")
            template = os.path.join(str(here), "A2_Outputs_Params_otoole", "A")
            conversion = os.path.join(
                str(here), "Miscellaneous", "conversion_format.yaml"
            )
            expected_commands = [
                [
                    "glpsol", "-m", "osemosys_fast_preprocessed.txt",
                    "-d", f"{data_file}.txt", "--wlp", f"{output_file}.lp",
                    "--check",
                ],
                [
                    "cplex", "-c", f"set logfile {output_file}.cplex.log",
                    f"read {output_file}.lp", "set threads 4",
                    "set randomseed 12345", "set parallel 1", "optimize",
                    f"write {output_file}.sol",
                ],
                [
                    "ENV-otoole", "results", "cplex", "csv",
                    f"{output_file}.sol", outputs, "csv", template, conversion,
                ],
                [
                    sys.executable,
                    "-B",
                    "-m",
                    "ostram.pipeline.execution.concatenate",
                    outputs,
                    output_file,
                ],
            ]
            self.assertEqual(
                [call.args[0] for call in result.runner.call_args_list],
                expected_commands,
            )
            for recorded in result.runner.call_args_list:
                self.assertTrue(recorded.kwargs["check"])
                self.assertNotIn("shell", recorded.kwargs)
                self.assertTrue(Path(recorded.kwargs["cwd"]).is_absolute())
                self.assertNotIn("env", recorded.kwargs)
            self.assertEqual(
                result.removed,
                [output_file + ".sol", output_file + ".feasopt.sol"],
            )
            result.check_environment.assert_called_once_with("cplex")
            self.assertIn("Scenario A_0 solved successfully", result.stdout)
            self.assertIn("Outputs concatenated to A_0_Output.csv", result.stdout)

    def test_command_failure_propagates_before_success_and_postprocessing(self) -> None:
        module = _load_b2("command_failure")
        with tempfile.TemporaryDirectory() as temp:
            here = Path(temp).resolve() / "execution_workspace"
            params = _base_params(
                solver="cplex",
                execute_model=False,
                create_matrix=True,
                concat_otoole_csv=True,
            )
            error = subprocess.CalledProcessError(7, "fixture matrix command")
            with self.assertRaises(subprocess.CalledProcessError) as raised:
                self._run_executor(
                    module,
                    here,
                    params,
                    process_side_effect=error,
                )
            self.assertEqual(raised.exception.returncode, 7)

    def test_missing_expected_solution_raises_before_results_or_concat(self) -> None:
        module = _load_b2("missing_solution")
        with tempfile.TemporaryDirectory() as temp:
            here = Path(temp).resolve() / "execution_workspace"
            params = _base_params(
                solver="cplex",
                execute_model=True,
                create_matrix=False,
                concat_otoole_csv=True,
            )
            with self.assertRaisesRegex(
                FileNotFoundError, "did not create the expected solution file"
            ):
                self._run_executor(
                    module,
                    here,
                    params,
                    solution_exists=False,
                )


if __name__ == "__main__":
    unittest.main()
