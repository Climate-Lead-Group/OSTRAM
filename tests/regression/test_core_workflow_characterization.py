from __future__ import annotations

import argparse
import ast
import importlib.util
import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
CHARACTERIZATION_DOC = REPO_ROOT / "docs" / "core-workflow-characterization.md"
SCENARIO_INVENTORY = TEST_ROOT / "scenarios.yaml"
COMPILED_REPORT = TEST_ROOT / "reports" / "final_compiled_input_equivalence_15.json"

PRIMARY_CORE_ENTRYPOINTS = (
    "run.py",
    "t1_confection/A0_generate_tech_country_matrix.py",
    "t1_confection/A1_Pre_processing_OG_csvs.py",
    "t1_confection/A2_AddTx.py",
    "t1_confection/A3_process.py",
    "t1_confection/B1_Run_Compiler.py",
    "t1_confection/B1_Compiler.py",
    "t1_confection/B2_Executing_OG_Model.py",
)

CANONICAL_CLI_MODULES = (
    "ostram/__init__.py",
    "ostram/__main__.py",
)

CORE_IMPLEMENTATION_MODULES = (
    "t1_confection/a3_orchestrator.py",
    "t1_confection/b1_runner.py",
    "t1_confection/b2_orchestrator.py",
)

OPTIONAL_MODEL_WRITING_ENTRYPOINTS = (
    "t1_confection/D1_generate_editor_template.py",
    "t1_confection/D2_update_secondary_techs.py",
    "t1_confection/Z_AUX_D1b_set_trn_limits_from_flows.py",
)

ANALYSIS_UTILITIES = (
    "tools/analysis/check_combined.py",
    "tools/analysis/ostram_scenario_analysis.py",
    "tools/analysis/ostram_trn_plotter.py",
    "tools/analysis/slice_by_country.py",
    "tools/analysis/analyse_sensitivity.py",
    "tools/analysis/concat_all_scenarios.py",
    "tools/analysis/reproduce_A1_A6.py",
    "tools/analysis/visualization/Z_AUX_generate_interactive_dashboards_aggregated.py",
    "tools/analysis/visualization/Z_AUX_generate_RES_diagram.py",
    "tools/analysis/visualization/Z_AUX_generate_transmission_maps.py",
    "tools/analysis/visualization/Z_AUX_interconnections_dashboard.py",
    "t1_confection/check_combined.py",
    "t1_confection/ostram_scenario_analysis.py",
    "t1_confection/ostram_trn_plotter.py",
    "t1_confection/slice_by_country.py",
    "t1_confection/analyse_sensitivity.py",
    "t1_confection/concat_all_scenarios_2.py",
    "t1_confection/reproduce_A1_A6.py",
    "t1_confection/sensitivity_expansion/analyse_ws4_vs_phaseB.py",
    "t1_confection/Z_AUX_generate_interactive_dashboards_aggregated.py",
    "t1_confection/Z_AUX_generate_RES_diagram.py",
    "t1_confection/Z_AUX_generate_transmission_maps.py",
    "t1_confection/Z_AUX_interconnections_dashboard.py",
)

PRESERVED_SCENARIOS = (
    "BAU",
    "A_Calibrated_BAU",
    "A_Calibrated_BAU_Clipped",
    "B_Optimised_VRE",
    "B_Opt_Clipped",
    "B_Opt_DirBidir",
    "B_Opt_DirContractual",
    "B_Opt_IndiaCosts",
    "B_Opt_IndiaCostsFuel",
    "B_Opt_LinkFreeze",
    "B_Opt_SolarCapex130",
    "B_Opt_SolarCapexHi",
    "B_Opt_SolarCapexSpike",
    "B_Opt_SolarHi10",
    "B_Opt_TradeCap15",
    "B_Opt_TradeCap30",
    "B_Opt_TradeCap50",
    "B_Opt_TxCap150",
    "C_Target_VRE",
    "C_Target_VRE_Clipped",
)

SUPERSEDED_PROTECTED_SCENARIOS = {
    "B_Opt_LinkFreeze",
    "B_Opt_SolarHi10",
    "B_Opt_TradeCap30",
    "B_Opt_TradeCap50",
}

DECISION_RELEVANT_SCENARIOS = {
    "A_Calibrated_BAU",
    "A_Calibrated_BAU_Clipped",
    "B_Optimised_VRE",
    "B_Opt_Clipped",
    "B_Opt_DirBidir",
    "B_Opt_DirContractual",
    "B_Opt_IndiaCosts",
    "B_Opt_IndiaCostsFuel",
    "B_Opt_SolarCapex130",
    "B_Opt_SolarCapexHi",
    "B_Opt_SolarCapexSpike",
    "B_Opt_TradeCap15",
    "B_Opt_TxCap150",
    "C_Target_VRE",
    "C_Target_VRE_Clipped",
}


def _source(relative: str) -> str:
    return (REPO_ROOT / relative).read_text(encoding="utf-8-sig")


def _tree(relative: str) -> ast.Module:
    return ast.parse(_source(relative), filename=relative)


def _load_module(relative: str, label: str):
    path = REPO_ROOT / relative
    module_name = f"_ostram_characterization_{label}"
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


def _call_name(node: ast.Call) -> str:
    def dotted(expr: ast.expr) -> str:
        if isinstance(expr, ast.Name):
            return expr.id
        if isinstance(expr, ast.Attribute):
            prefix = dotted(expr.value)
            return f"{prefix}.{expr.attr}" if prefix else expr.attr
        return ""

    return dotted(node.func)


def _calls(node: ast.AST, selected: set[str] | None = None) -> list[str]:
    calls = [item for item in ast.walk(node) if isinstance(item, ast.Call)]
    calls.sort(key=lambda item: (item.lineno, item.col_offset))
    names = [_call_name(item) for item in calls]
    return [name for name in names if selected is None or name in selected]


def _function(tree: ast.Module, name: str) -> ast.FunctionDef:
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return node
    raise AssertionError(f"function {name!r} not found")


def _main_guard(tree: ast.Module) -> ast.If:
    for node in tree.body:
        if not isinstance(node, ast.If) or not isinstance(node.test, ast.Compare):
            continue
        if (
            isinstance(node.test.left, ast.Name)
            and node.test.left.id == "__name__"
            and any(
                isinstance(comparator, ast.Constant)
                and comparator.value == "__main__"
                for comparator in node.test.comparators
            )
        ):
            return node
    raise AssertionError("module has no __main__ guard")


class EntrypointClassificationTests(unittest.TestCase):
    def test_core_entrypoints_and_analysis_utilities_are_disjoint_and_parse(self) -> None:
        core = (
            set(PRIMARY_CORE_ENTRYPOINTS)
            | set(CANONICAL_CLI_MODULES)
            | set(CORE_IMPLEMENTATION_MODULES)
            | set(OPTIONAL_MODEL_WRITING_ENTRYPOINTS)
        )
        utilities = set(ANALYSIS_UTILITIES)
        self.assertTrue(core.isdisjoint(utilities))

        for relative in sorted(core | utilities):
            path = REPO_ROOT / relative
            self.assertTrue(path.is_file(), path)
            ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))

    def test_primary_core_entrypoints_do_not_delegate_to_analysis_utilities(self) -> None:
        utility_names = {Path(relative).name for relative in ANALYSIS_UTILITIES}
        for relative in PRIMARY_CORE_ENTRYPOINTS:
            source = _source(relative)
            referenced = {name for name in utility_names if name in source}
            self.assertEqual(referenced, set(), f"{relative} references {referenced}")

        for relative in CORE_IMPLEMENTATION_MODULES:
            source = _source(relative)
            referenced = {name for name in utility_names if name in source}
            self.assertEqual(referenced, set(), f"{relative} references {referenced}")

    def test_classification_and_boundaries_are_documented(self) -> None:
        text = CHARACTERIZATION_DOC.read_text(encoding="utf-8")
        for relative in (
            *PRIMARY_CORE_ENTRYPOINTS,
            *CANONICAL_CLI_MODULES,
            *CORE_IMPLEMENTATION_MODULES,
            *OPTIONAL_MODEL_WRITING_ENTRYPOINTS,
            *ANALYSIS_UTILITIES,
        ):
            self.assertIn(f"`{relative}`", text)

        boundary_markers = (
            "A1_Pre_processing_OG_csvs.py` -> `A2_AddTx.py` -> per-scenario `A3_process.py`",
            "`list_scenario_suffixes` once, then `update_main_scenario` ->",
            "run_otoole_conversion` -> `run_preprocessing_script` -> "
            "`run_days_in_day_type_patcher`",
            "stage_0_5_rnwbio` -> `stage_1_scripts_1_to_5` -> `stage_1b`",
        )
        for marker in boundary_markers:
            self.assertIn(marker, text)


class ScenarioPolicyTests(unittest.TestCase):
    def test_inventory_preserves_the_exact_twenty_scenario_definitions(self) -> None:
        payload = json.loads(SCENARIO_INVENTORY.read_text(encoding="utf-8"))
        names = tuple(item["name"] for item in payload["scenarios"])
        self.assertEqual(names, PRESERVED_SCENARIOS)
        self.assertEqual(len(names), 20)
        self.assertEqual(len(set(names)), 20)

    def test_static_cleanup_scope_is_exactly_bau_plus_the_decision_scope(self) -> None:
        payload = json.loads(SCENARIO_INVENTORY.read_text(encoding="utf-8"))
        accepted = {
            item["name"]
            for item in payload["scenarios"]
            if item["cleanup_acceptance"]
        }
        self.assertEqual(accepted, DECISION_RELEVANT_SCENARIOS | {"BAU"})
        self.assertEqual(len(accepted), 16)
        self.assertEqual(set(PRESERVED_SCENARIOS) - accepted, SUPERSEDED_PROTECTED_SCENARIOS)

    def test_compiled_input_report_is_the_exact_decision_relevant_fifteen(self) -> None:
        report = json.loads(COMPILED_REPORT.read_text(encoding="utf-8"))
        compiled = {
            item["scenario"] for item in report["final_solver_consumed_files"]
        }
        policy = report["policy"]
        result = report["result"]

        self.assertEqual(compiled, DECISION_RELEVANT_SCENARIOS)
        self.assertEqual(len(compiled), 15)
        self.assertNotIn("BAU", compiled)
        self.assertEqual(policy["preservation_scenario_count"], 20)
        self.assertEqual(policy["static_cleanup_acceptance_scenario_count"], 16)
        self.assertEqual(policy["final_compiled_equivalence_scenario_count"], 15)
        self.assertEqual(
            set(policy["superseded_protected_scenarios"]),
            SUPERSEDED_PROTECTED_SCENARIOS,
        )
        self.assertEqual(result["expected_final_compiled_files"], 15)
        self.assertEqual(result["byte_exact_files"], 15)
        self.assertEqual(result["status"], "pass")


class DiscoveryCharacterizationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.b1 = _load_module("t1_confection/B1_Run_Compiler.py", "b1_discovery")
        cls.launcher = _load_module("run.py", "launcher_discovery")
        cls.a3 = _load_module("t1_confection/A3_process.py", "a3_discovery")
        cls.config = _load_module(
            "t1_confection/Z_AUX_config_loader.py", "config_discovery"
        )

    def test_b1_discovers_sorted_scenario_directories_and_excludes_backups(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp)
            included = ("A", "B", "Default")
            excluded = (
                "",
                "A_backup",
                "B_SNAPSHOT",
                "C_pre_experiment",
                "D_20260513",
            )
            for suffix in (*included, *excluded):
                (root / f"A1_Outputs_{suffix}").mkdir()
            (root / "A1_Outputs_file").write_text("not a directory", encoding="utf-8")
            (root / "unrelated").mkdir()

            self.assertEqual(self.b1.list_scenario_suffixes(root), list(included))

    def test_b1_filter_preserves_discovery_order_and_restores_config_on_failure(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            script_dir = Path(temp)
            (script_dir / "B1_Run_Compiler.py").write_text("", encoding="utf-8")
            (script_dir / "B1_Compiler.py").write_text("", encoding="utf-8")
            yaml_path = script_dir / "Config_MOMF_T1_A.yaml"
            original = b"xtra_scen:\n  Main_Scenario: ORIGINAL\n"
            yaml_path.write_bytes(original)
            outputs = script_dir / "A1_Outputs"
            outputs.mkdir()
            for scenario in ("A", "B", "C"):
                (outputs / f"A1_Outputs_{scenario}").mkdir()

            events: list[tuple[str, str]] = []

            def update(path: Path, scenario: str) -> None:
                events.append(("update", scenario))
                path.write_text(f"scenario: {scenario}\n", encoding="utf-8")

            def compiler(_script_dir: Path) -> int:
                scenario = events[-1][1]
                events.append(("compile", scenario))
                if scenario == "C":
                    raise RuntimeError("fixture compiler failure")
                return 0

            with (
                mock.patch.object(self.b1, "__file__", str(script_dir / "B1_Run_Compiler.py")),
                mock.patch.object(
                    self.b1,
                    "parse_cli_args",
                    return_value=argparse.Namespace(scenarios="C,A"),
                ),
                mock.patch.object(self.b1, "update_main_scenario", side_effect=update),
                mock.patch.object(self.b1, "run_compiler", side_effect=compiler),
                redirect_stdout(io.StringIO()),
                self.assertRaisesRegex(RuntimeError, "fixture compiler failure"),
            ):
                self.b1.main()

            self.assertEqual(
                events,
                [("update", "A"), ("compile", "A"), ("update", "C"), ("compile", "C")],
            )
            self.assertEqual(yaml_path.read_bytes(), original)
            self.assertFalse(yaml_path.with_suffix(".yaml.bak").exists())

    def test_run_scenario_enumeration_uses_helper_output_without_launching_it(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            t1_dir = Path(temp)
            helper = t1_dir / "A3_process" / "_scenarios.py"
            helper.parent.mkdir()
            helper.write_text("# fixture only\n", encoding="utf-8")
            with (
                mock.patch.object(self.launcher, "T1_DIR", t1_dir),
                mock.patch.object(
                    self.launcher.subprocess,
                    "check_output",
                    return_value="Loading environment\nBAU\nB_Optimised_VRE\n",
                ) as check_output,
            ):
                result = self.launcher.enumerate_active_scenarios("OSTRAM-env")

            self.assertEqual(result, ["BAU", "B_Optimised_VRE"])
            command = check_output.call_args.args[0]
            self.assertIn(str(helper.resolve()), command)
            self.assertIn("list-active", command)

    def test_run_scenario_enumeration_falls_back_to_bau_when_helper_is_absent(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            with mock.patch.object(self.launcher, "T1_DIR", Path(temp)):
                self.assertEqual(
                    self.launcher.enumerate_active_scenarios("OSTRAM-env"), ["BAU"]
                )

    def test_a3_rule_yaml_resolution_prefers_scenario_override_then_default(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            rules = Path(temp)
            script = rules / "rule.py"
            script.write_text('YAML_FILE_NAME = "rule.yaml"\n', encoding="utf-8")
            default = rules / "rule.yaml"
            default.write_text("source: default\n", encoding="utf-8")
            override = rules / "configs" / "Scenario_A" / "rule.yaml"
            override.parent.mkdir(parents=True)
            override.write_text("source: scenario\n", encoding="utf-8")

            with mock.patch.object(self.a3, "RULES_SCRIPTS_DIR", rules):
                self.assertEqual(
                    self.a3._resolve_script_yaml("rule.py", "Scenario_A"), override
                )
                self.assertEqual(
                    self.a3._resolve_script_yaml("rule.py", "Scenario_B"), default
                )

    def test_config_loader_is_script_anchored_sorted_and_cached(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            config_path = Path(temp) / "Config_country_codes.yaml"
            config_path.write_text(
                "country_data:\n"
                "  ZZZ: {english_name: Zed, ostram_name: Zeta}\n"
                "  AAA: {english_name: Aye, ostram_name: Alfa}\n"
                "countries: [ZZZ, AAA]\n",
                encoding="utf-8",
            )
            with mock.patch.object(self.config, "CONFIG_PATH", config_path):
                self.config._cached_config = None
                self.assertEqual(self.config.get_countries(), ["AAA", "ZZZ"])
                self.assertEqual(self.config.get_model_countries_list(), ["ZZZ", "AAA"])

                config_path.write_text("country_data: {BBB: {}}\n", encoding="utf-8")
                self.assertEqual(self.config.get_countries(), ["AAA", "ZZZ"])
                self.config._cached_config = None
                self.assertEqual(self.config.get_countries(), ["BBB"])

        self.config._cached_config = None
        source = _source("t1_confection/Z_AUX_config_loader.py")
        self.assertIn('CONFIG_PATH = SCRIPT_DIR / "Config_country_codes.yaml"', source)


class CallPathBoundaryTests(unittest.TestCase):
    def test_run_main_keeps_stage_order_and_scenario_propagation(self) -> None:
        launcher = _load_module("run.py", "launcher_main")
        events: list[tuple[str, str]] = []

        def pipeline(_env: str, script: Path, extra: str = "") -> None:
            events.append((script.name, extra))

        def a3(_env: str, _script: Path, scenario: str) -> None:
            events.append(("A3", scenario))

        argv = ["run.py", "--skip-pull", "--scenarios", "C,A"]
        with (
            mock.patch.object(sys, "argv", argv),
            mock.patch.object(launcher, "check_tool_available"),
            mock.patch.object(launcher, "create_env_if_missing"),
            mock.patch.object(launcher, "ensure_deps"),
            mock.patch.object(launcher, "ensure_dvc_repo"),
            mock.patch.object(launcher, "post_a2_snapshot_exists", return_value=False),
            mock.patch.object(
                launcher, "enumerate_active_scenarios", return_value=["A", "B", "C"]
            ),
            mock.patch.object(launcher, "run_pipeline_script", side_effect=pipeline),
            mock.patch.object(launcher, "run_a3_for_scenario", side_effect=a3),
            redirect_stdout(io.StringIO()),
        ):
            launcher.main()

        self.assertEqual(
            events,
            [
                ("A1_Pre_processing_OG_csvs.py", ""),
                ("A2_AddTx.py", ""),
                ("A3", "A"),
                ("A3", "C"),
                ("B1_Run_Compiler.py", '--scenarios "C,A"'),
                ("B2_Executing_OG_Model.py", '--scenarios "C,A"'),
            ],
        )

    def test_b1_wrapper_and_isolated_boundaries_are_explicit(self) -> None:
        wrapper_tree = _tree("t1_confection/B1_Run_Compiler.py")
        wrapper_main = _function(wrapper_tree, "main")
        self.assertEqual(
            _calls(
                wrapper_main,
                {
                    "parse_cli_args",
                    "_impl.B1Paths.from_entrypoint",
                    "_impl.orchestrate",
                },
            ),
            [
                "_impl.orchestrate",
                "parse_cli_args",
                "_impl.B1Paths.from_entrypoint",
            ],
        )
        wrapper_source = _source("t1_confection/B1_Run_Compiler.py")
        self.assertNotIn("subprocess", wrapper_source)
        self.assertNotIn("shutil", wrapper_source)

        tree = _tree("t1_confection/b1_runner.py")
        for name in (
            "resolve_scenarios",
            "build_compiler_command",
            "execute_command",
            "preserved_configuration",
            "orchestrate",
        ):
            _function(tree, name)

        runner = _function(tree, "run_compiler")
        self.assertEqual(
            _calls(runner, {"build_compiler_command", "execute_command"}),
            ["build_compiler_command", "execute_command"],
        )
        executor = _function(tree, "execute_command")
        executor_source = ast.get_source_segment(
            _source("t1_confection/b1_runner.py"), executor
        )
        self.assertIn("list(command.argv)", executor_source)
        self.assertIn("cwd=str(command.cwd)", executor_source)
        self.assertNotIn("env=", executor_source)

    def test_a3_main_keeps_snapshot_and_stage_order(self) -> None:
        entry_tree = _tree("t1_confection/A3_process.py")
        main = _function(entry_tree, "main")
        self.assertEqual(
            _calls(
                main,
                {
                    "_orchestrator.orchestrate_a3",
                    "parse_cli_args",
                    "_orchestration_paths",
                    "_orchestration_dependencies",
                },
            ),
            [
                "_orchestrator.orchestrate_a3",
                "parse_cli_args",
                "_orchestration_paths",
                "_orchestration_dependencies",
            ],
        )

        tree = _tree("t1_confection/a3_orchestrator.py")
        execute = _function(tree, "execute_plan")
        stages = {
            "dependencies.copy_tree",
            "dependencies.build_workdir",
            "dependencies.materialize_scenario_template",
            "dependencies.stage_0_5_rnwbio",
            "dependencies.stage_1_scripts_1_to_5",
            "dependencies.stage_1b",
            "dependencies.stage_2_and_2_5",
            "dependencies.stage_3_fix_2",
            "dependencies.stage_4_consolidate",
            "dependencies.stage_4_5_apply_inherited_restrictions",
            "dependencies.stage_5_rules_scripts",
            "dependencies.stage_ws3_interconnector_costs",
            "dependencies.stage_ws3_internal_transmission",
            "dependencies.stage_ws3_internal_tx_losses",
            "dependencies.stage_6_sync_og_to_ts20",
            "dependencies.stage_6_persist_restrictions",
            "dependencies.deliver_outputs",
        }
        self.assertEqual(
            _calls(execute, stages),
            [
                "dependencies.copy_tree",
                "dependencies.build_workdir",
                "dependencies.materialize_scenario_template",
                "dependencies.stage_0_5_rnwbio",
                "dependencies.stage_1_scripts_1_to_5",
                "dependencies.stage_1b",
                "dependencies.stage_2_and_2_5",
                "dependencies.stage_3_fix_2",
                "dependencies.stage_4_consolidate",
                "dependencies.stage_4_5_apply_inherited_restrictions",
                "dependencies.stage_5_rules_scripts",
                "dependencies.stage_ws3_interconnector_costs",
                "dependencies.stage_ws3_internal_transmission",
                "dependencies.stage_ws3_internal_tx_losses",
                "dependencies.stage_6_sync_og_to_ts20",
                "dependencies.stage_6_persist_restrictions",
                "dependencies.deliver_outputs",
            ],
        )
        plan = _function(tree, "resolve_plan")
        plan_source = ast.get_source_segment(
            _source("t1_confection/a3_orchestrator.py"), plan
        )
        execute_source = ast.get_source_segment(
            _source("t1_confection/a3_orchestrator.py"), execute
        )
        self.assertIn('"_post_a2_snapshot_BAU"', plan_source)
        self.assertLess(
            execute_source.index("dependencies.copy_tree("),
            execute_source.index("dependencies.build_workdir("),
        )

    def test_b2_discovery_generation_patch_and_solver_boundaries_are_explicit(
        self,
    ) -> None:
        entry_tree = _tree("t1_confection/B2_Executing_OG_Model.py")
        orchestrator_tree = _tree("t1_confection/b2_orchestrator.py")
        guard = _main_guard(entry_tree)
        self.assertEqual(_calls(guard, {"main"}), ["main"])

        resolution = _function(orchestrator_tree, "resolve_scenarios")
        self.assertEqual(
            _calls(resolution, {"sorted", "os.listdir", "scenarios.remove"}),
            ["sorted", "os.listdir", "scenarios.remove"],
        )
        resolution_source = ast.get_source_segment(
            _source("t1_confection/b2_orchestrator.py"), resolution
        )
        self.assertIn(
            "scenarios = [scenario for scenario in scenarios if scenario in requested]",
            resolution_source,
        )
        self.assertIn(
            'params_a2["xtra_scen"]["Main_Scenario"]',
            resolution_source,
        )

        compiled_input = _function(orchestrator_tree, "run_compiled_input_stage")
        selected_compiled_calls = {
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
            _calls(compiled_input, selected_compiled_calls),
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
            _calls(
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
        processes = [
            call
            for call in ast.walk(execution)
            if isinstance(call, ast.Call)
            and _call_name(call) == "dependencies.mp_module.Process"
        ]
        self.assertEqual(len(processes), 1)
        target = next(
            keyword.value
            for keyword in processes[0].keywords
            if keyword.arg == "target"
        )
        self.assertIsInstance(target, ast.Attribute)
        self.assertEqual(target.attr, "main_executer")
        self.assertIsInstance(target.value, ast.Name)
        self.assertEqual(target.value.id, "dependencies")

        executor = _function(entry_tree, "main_executer")
        self.assertEqual(
            _calls(
                executor,
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

        solver_boundary = _function(orchestrator_tree, "invoke_solver_command")
        self.assertEqual(
            _calls(solver_boundary, {"process_runner"}),
            ["process_runner"],
        )
        solver_boundary_source = ast.get_source_segment(
            _source("t1_confection/b2_orchestrator.py"), solver_boundary
        )
        self.assertIn(
            "process_runner(command, shell=True, check=True)",
            solver_boundary_source,
        )

        solver_adapter = next(
            node
            for node in orchestrator_tree.body
            if isinstance(node, ast.ClassDef) and node.name == "SolverAdapter"
        )
        prepare_command = next(
            node
            for node in solver_adapter.body
            if isinstance(node, ast.FunctionDef) and node.name == "prepare_command"
        )
        adapter_source = ast.get_source_segment(
            _source("t1_confection/b2_orchestrator.py"), prepare_command
        )
        for solver in ("glpsol", "cbc", "cplex", "gurobi_cl"):
            self.assertIn(solver, adapter_source)


class NoSolverSafetyTests(unittest.TestCase):
    PROCESS_CALLS = {
        "subprocess.run",
        "subprocess.call",
        "subprocess.check_call",
        "subprocess.check_output",
        "subprocess.Popen",
        "os.system",
        "os.popen",
        "os.spawnl",
        "os.spawnle",
        "os.spawnlp",
        "os.spawnlpe",
        "os.spawnv",
        "os.spawnve",
        "os.spawnvp",
        "os.spawnvpe",
        "multiprocessing.Process",
        "mp.Process",
        "asyncio.create_subprocess_exec",
        "asyncio.create_subprocess_shell",
    }

    def test_regression_suite_allows_only_git_metadata_and_cli_help_smoke(self) -> None:
        found: list[tuple[Path, str, str, ast.Call]] = []

        class Visitor(ast.NodeVisitor):
            def __init__(self, path: Path) -> None:
                self.path = path
                self.functions: list[str] = []

            def visit_FunctionDef(self, node: ast.FunctionDef) -> None:
                self.functions.append(node.name)
                self.generic_visit(node)
                self.functions.pop()

            visit_AsyncFunctionDef = visit_FunctionDef

            def visit_Call(self, node: ast.Call) -> None:
                name = _call_name(node)
                if name in NoSolverSafetyTests.PROCESS_CALLS:
                    function = self.functions[-1] if self.functions else "<module>"
                    found.append((self.path, function, name, node))
                self.generic_visit(node)

        for path in sorted(TEST_ROOT.rglob("*.py")):
            tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
            Visitor(path).visit(tree)

        self.assertEqual(len(found), 2, [(str(p), f, n) for p, f, n, _ in found])
        by_boundary = {
            (path.name, function, name): call
            for path, function, name, call in found
        }
        self.assertEqual(
            set(by_boundary),
            {
                ("ostram_regression.py", "_git", "subprocess.run"),
                ("test_canonical_cli.py", "_run_smoke", "subprocess.run"),
            },
        )

        call = by_boundary[("ostram_regression.py", "_git", "subprocess.run")]
        self.assertTrue(call.args)
        command = call.args[0]
        self.assertIsInstance(command, ast.List)
        self.assertGreaterEqual(len(command.elts), 2)
        self.assertEqual(command.elts[0].value, "git")
        self.assertEqual(command.elts[1].value, "-C")
        shell_keywords = [kw for kw in call.keywords if kw.arg == "shell"]
        self.assertEqual(shell_keywords, [])

        smoke = by_boundary[
            ("test_canonical_cli.py", "_run_smoke", "subprocess.run")
        ]
        self.assertEqual(len(smoke.args), 1)
        smoke_command = smoke.args[0]
        self.assertIsInstance(smoke_command, ast.List)
        self.assertEqual(len(smoke_command.elts), 3)
        executable, no_bytecode, forwarded = smoke_command.elts
        self.assertIsInstance(executable, ast.Attribute)
        self.assertIsInstance(executable.value, ast.Name)
        self.assertEqual((executable.value.id, executable.attr), ("sys", "executable"))
        self.assertIsInstance(no_bytecode, ast.Constant)
        self.assertEqual(no_bytecode.value, "-B")
        self.assertIsInstance(forwarded, ast.Starred)
        self.assertIsInstance(forwarded.value, ast.Name)
        self.assertEqual(forwarded.value.id, "arguments")
        smoke_keywords = {keyword.arg: keyword.value for keyword in smoke.keywords}
        self.assertEqual(
            set(smoke_keywords),
            {"cwd", "env", "capture_output", "text", "timeout", "check"},
        )
        for keyword in ("capture_output", "text"):
            self.assertIs(smoke_keywords[keyword].value, True)
        self.assertEqual(smoke_keywords["timeout"].value, 30)
        self.assertIs(smoke_keywords["check"].value, False)

        cli_test = ast.parse(
            (TEST_ROOT / "test_canonical_cli.py").read_text(encoding="utf-8-sig")
        )
        smoke_invocations = [
            node
            for node in ast.walk(cli_test)
            if isinstance(node, ast.Call) and _call_name(node) == "_run_smoke"
        ]
        self.assertEqual(len(smoke_invocations), 2)
        literal_routes = []
        for invocation in smoke_invocations:
            self.assertIsInstance(invocation.args[0], ast.List)
            literal_routes.append(
                [element.value for element in invocation.args[0].elts]
            )
        self.assertEqual(
            literal_routes,
            [["-m", "ostram", "--help"], ["-m", "ostram", "unknown"]],
        )


if __name__ == "__main__":
    unittest.main()
