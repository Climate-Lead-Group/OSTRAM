from __future__ import annotations

import argparse
import importlib
import importlib.util
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import contextmanager, redirect_stderr, redirect_stdout
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import yaml

from ostram.paths import ProjectPaths
from ostram.pipeline.compilation.transforms import effects as b1_effects
from ostram.pipeline.compilation.transforms import planning as b1_planning


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
B1_ENTRYPOINT = REPO_ROOT / "ostram" / "pipeline" / "compilation" / "runner.py"


def _load_b1(label: str):
    module_name = f"ostram.pipeline.compilation._characterization_{label}"
    spec = importlib.util.spec_from_file_location(module_name, B1_ENTRYPOINT)
    if spec is None or spec.loader is None:
        raise AssertionError(f"could not load module spec for {B1_ENTRYPOINT}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    finally:
        sys.modules.pop(module_name, None)
    return module


def _implementation(module):
    """Return the implementation module before and after wrapper extraction."""
    return getattr(module, "_impl", module)


@contextmanager
def _b1_fixture(*scenarios: str):
    with tempfile.TemporaryDirectory() as temp:
        script_dir = Path(temp).resolve()
        entrypoint = script_dir / "runner.py"
        entrypoint.write_text("# fixture only\n", encoding="utf-8")
        compiler = script_dir / "compiler.py"
        compiler.write_text("# never executed\n", encoding="utf-8")
        config = script_dir / "Config_MOMF_T1_A.yaml"
        original = b"xtra_scen:\r\n  Main_Scenario: ORIGINAL\r\n"
        config.write_bytes(original)
        scenario_root = script_dir / "A1_Outputs"
        scenario_root.mkdir()
        for scenario in scenarios:
            (scenario_root / f"A1_Outputs_{scenario}").mkdir()
        yield SimpleNamespace(
            script_dir=script_dir,
            entrypoint=entrypoint,
            compiler=compiler,
            config=config,
            original=original,
            scenario_root=scenario_root,
            backup=config.with_suffix(".yaml.bak"),
        )


def _fixture_update(events: list[tuple[object, ...]]):
    def update(path: Path, scenario: str) -> None:
        events.append(("update", path, scenario))
        path.write_text(f"scenario: {scenario}\n", encoding="utf-8")

    return update


def _patch_fixture_paths(runner, fixture):
    paths = runner.B1Paths(
        script_dir=fixture.script_dir,
        config_path=fixture.config,
        compiler_path=fixture.compiler,
        scenarios_root=fixture.scenario_root,
    )
    return mock.patch.object(runner.B1Paths, "defaults", return_value=paths)


class B1ImportAndCliCharacterizationTests(unittest.TestCase):
    def test_import_is_silent_and_does_not_cross_process_or_file_boundaries(self) -> None:
        with (
            mock.patch.object(subprocess, "run") as process_run,
            mock.patch("shutil.copy2") as copy_file,
            mock.patch("shutil.move") as move_file,
            redirect_stdout(io.StringIO()) as stdout,
            redirect_stderr(io.StringIO()) as stderr,
        ):
            module = _load_b1("import_safety")

        process_run.assert_not_called()
        copy_file.assert_not_called()
        move_file.assert_not_called()
        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")
        self.assertTrue(callable(module.main))

    def test_cli_defaults_and_explicit_filter_are_unchanged(self) -> None:
        runner = _implementation(_load_b1("cli_values"))
        with mock.patch.object(sys, "argv", ["python -m ostram compile-inputs"]):
            defaults = runner.parse_cli_args()
        with mock.patch.object(
            sys,
            "argv",
            ["python -m ostram compile-inputs", "--scenarios", " C , A, C "],
        ):
            explicit = runner.parse_cli_args()

        self.assertEqual(vars(defaults), {"scenarios": None})
        self.assertEqual(vars(explicit), {"scenarios": " C , A, C "})

    def test_public_wrapper_forwards_explicit_orchestration_seams(self) -> None:
        module = _load_b1("wrapper_forwarding")
        cli_args = argparse.Namespace(scenarios="A")
        paths = module._impl.B1Paths(
            script_dir=B1_ENTRYPOINT.parent,
            config_path=B1_ENTRYPOINT.with_name("Config_MOMF_T1_A.yaml"),
            compiler_path=B1_ENTRYPOINT.with_name("compiler.py"),
            scenarios_root=B1_ENTRYPOINT.parent / "A1_Outputs",
        )

        with (
            mock.patch.object(module, "parse_cli_args", return_value=cli_args) as parse,
            mock.patch.object(
                module._impl.B1Paths,
                "defaults",
                return_value=paths,
            ) as resolve_paths,
            mock.patch.object(
                module._impl, "orchestrate", return_value=None
            ) as orchestrate,
        ):
            result = module.main()

        self.assertIsNone(result)
        parse.assert_called_once_with()
        resolve_paths.assert_called_once_with()
        orchestrate.assert_called_once_with(
            cli_args,
            paths,
            scenario_discoverer=module.list_scenario_suffixes,
            scenario_updater=module.update_main_scenario,
            compiler_runner=module.run_compiler,
        )

    def test_help_unknown_option_and_missing_value_keep_argparse_exit_codes(self) -> None:
        runner = _implementation(_load_b1("cli_exits"))
        cases = (
            (["python -m ostram compile-inputs", "--help"], 0),
            (["python -m ostram compile-inputs", "--unknown"], 2),
            (["python -m ostram compile-inputs", "--scenarios"], 2),
        )
        for argv, expected in cases:
            with self.subTest(argv=argv):
                with (
                    mock.patch.object(sys, "argv", argv),
                    redirect_stdout(io.StringIO()) as stdout,
                    redirect_stderr(io.StringIO()),
                    self.assertRaises(SystemExit) as raised,
                ):
                    runner.parse_cli_args()
                self.assertEqual(raised.exception.code, expected)
                if expected == 0:
                    help_text = stdout.getvalue()
                    self.assertIn("Run B1 compiler across scenarios", help_text)
                    self.assertIn("--scenarios SCENARIOS", help_text)


class B1ScenarioCharacterizationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.runner = _implementation(_load_b1(self.id().rsplit(".", 1)[-1]))

    def test_discovery_is_sorted_and_preserves_all_existing_exclusions(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
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
            (root / "a1_outputs_wrong_case").mkdir()

            self.assertEqual(self.runner.list_scenario_suffixes(root), list(included))

    def test_no_filter_runs_every_discovered_scenario_in_discovery_order(self) -> None:
        with _b1_fixture("C", "A", "B") as fixture:
            events: list[tuple[object, ...]] = []
            with (
                _patch_fixture_paths(self.runner, fixture),
                mock.patch.object(
                    self.runner,
                    "parse_cli_args",
                    return_value=argparse.Namespace(scenarios=None),
                ),
                mock.patch.object(
                    self.runner,
                    "update_main_scenario",
                    side_effect=_fixture_update(events),
                ),
                mock.patch.object(
                    self.runner,
                    "run_compiler",
                    side_effect=lambda path: events.append(("compile", path)) or 0,
                ),
                redirect_stdout(io.StringIO()) as stdout,
            ):
                result = self.runner.main()

            self.assertIsNone(result)
            self.assertEqual(
                [event[2] for event in events if event[0] == "update"],
                ["A", "B", "C"],
            )
            self.assertEqual(len([event for event in events if event[0] == "compile"]), 3)
            self.assertNotIn("Scenario filter active", stdout.getvalue())

    def test_filter_uses_requested_order_and_collapses_requested_duplicates(self) -> None:
        with _b1_fixture("C", "A", "B") as fixture:
            scenarios: list[str] = []

            def update(path: Path, scenario: str) -> None:
                scenarios.append(scenario)
                path.write_text(f"scenario: {scenario}\n", encoding="utf-8")

            with (
                _patch_fixture_paths(self.runner, fixture),
                mock.patch.object(
                    self.runner,
                    "parse_cli_args",
                    return_value=argparse.Namespace(scenarios=" C, A, C "),
                ),
                mock.patch.object(self.runner, "update_main_scenario", side_effect=update),
                mock.patch.object(self.runner, "run_compiler", return_value=0) as compiler,
                redirect_stdout(io.StringIO()) as stdout,
            ):
                self.runner.main()

            self.assertEqual(scenarios, ["C", "A"])
            self.assertEqual(compiler.call_count, 2)
            self.assertIn("[INFO] Scenario filter active: ['C', 'A']", stdout.getvalue())

    def test_unknown_filter_preserves_unknown_order_and_duplicates_and_exits_one(self) -> None:
        with _b1_fixture("A", "B") as fixture:
            with (
                _patch_fixture_paths(self.runner, fixture),
                mock.patch.object(
                    self.runner,
                    "parse_cli_args",
                    return_value=argparse.Namespace(
                        scenarios="Missing, A, Missing, Other"
                    ),
                ),
                mock.patch.object(self.runner.shutil, "copy2") as copy_file,
                mock.patch.object(self.runner, "run_compiler") as compiler,
                redirect_stdout(io.StringIO()) as stdout,
                self.assertRaises(SystemExit) as raised,
            ):
                self.runner.main()

            self.assertEqual(raised.exception.code, 1)
            copy_file.assert_not_called()
            compiler.assert_not_called()
            output = stdout.getvalue()
            self.assertIn("['Missing', 'Missing', 'Other']", output)
            self.assertIn("Discovered: ['A', 'B']", output)

    def test_truthy_empty_filter_runs_nothing_but_still_backs_up_and_restores(self) -> None:
        with _b1_fixture("A", "B") as fixture:
            with (
                _patch_fixture_paths(self.runner, fixture),
                mock.patch.object(
                    self.runner,
                    "parse_cli_args",
                    return_value=argparse.Namespace(scenarios=" , , "),
                ),
                mock.patch.object(self.runner, "update_main_scenario") as update,
                mock.patch.object(self.runner, "run_compiler") as compiler,
                redirect_stdout(io.StringIO()) as stdout,
            ):
                result = self.runner.main()

            self.assertIsNone(result)
            update.assert_not_called()
            compiler.assert_not_called()
            self.assertEqual(fixture.config.read_bytes(), fixture.original)
            self.assertFalse(fixture.backup.exists())
            self.assertIn("[INFO] Scenario filter active: []", stdout.getvalue())
            self.assertIn("[INFO] All done.", stdout.getvalue())

    def test_empty_string_is_no_filter_and_runs_all_discovered_scenarios(self) -> None:
        with _b1_fixture("B", "A") as fixture:
            events: list[tuple[object, ...]] = []
            with (
                _patch_fixture_paths(self.runner, fixture),
                mock.patch.object(
                    self.runner,
                    "parse_cli_args",
                    return_value=argparse.Namespace(scenarios=""),
                ),
                mock.patch.object(
                    self.runner,
                    "update_main_scenario",
                    side_effect=_fixture_update(events),
                ),
                mock.patch.object(self.runner, "run_compiler", return_value=0),
                redirect_stdout(io.StringIO()),
            ):
                self.runner.main()

            self.assertEqual(
                [event[2] for event in events if event[0] == "update"], ["A", "B"]
            )

    def test_no_discovery_exits_zero_before_filter_validation_or_backup(self) -> None:
        with _b1_fixture() as fixture:
            with (
                _patch_fixture_paths(self.runner, fixture),
                mock.patch.object(
                    self.runner,
                    "parse_cli_args",
                    return_value=argparse.Namespace(scenarios="Missing"),
                ),
                mock.patch.object(self.runner.shutil, "copy2") as copy_file,
                redirect_stdout(io.StringIO()) as stdout,
                self.assertRaises(SystemExit) as raised,
            ):
                self.runner.main()

            self.assertEqual(raised.exception.code, 0)
            copy_file.assert_not_called()
            self.assertIn("Nothing to do", stdout.getvalue())


class B1CommandBoundaryCharacterizationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.runner = _implementation(_load_b1(self.id().rsplit(".", 1)[-1]))

    def test_compiler_command_uses_current_interpreter_exact_tokens_cwd_and_inherited_env(self) -> None:
        with _b1_fixture("A") as fixture:
            inherited = {"EXISTING": "kept", "PYTHONHASHSEED": "0"}
            interpreter = str(fixture.script_dir / "python with spaces.exe")
            with (
                mock.patch.dict(os.environ, inherited, clear=True),
                mock.patch.object(self.runner.sys, "executable", interpreter),
                mock.patch.object(
                    self.runner.subprocess,
                    "run",
                    return_value=SimpleNamespace(returncode=7),
                ) as process_run,
            ):
                result = self.runner.run_compiler(fixture.script_dir)

            self.assertEqual(result, 7)
            process_run.assert_called_once_with(
                [interpreter, "-B", "-m", "ostram.pipeline.compilation.compiler"],
                cwd=str(fixture.script_dir),
            )
            self.assertNotIn("env", process_run.call_args.kwargs)
            self.assertNotIn("shell", process_run.call_args.kwargs)
            self.assertNotIn("check", process_run.call_args.kwargs)

    def test_missing_compiler_fails_before_command_runner(self) -> None:
        with _b1_fixture("A") as fixture:
            fixture.compiler.unlink()
            with mock.patch.object(self.runner.subprocess, "run") as process_run:
                with self.assertRaisesRegex(FileNotFoundError, "Missing script"):
                    self.runner.run_compiler(fixture.script_dir)
            process_run.assert_not_called()

    def test_missing_config_and_missing_compiler_preflights_keep_exit_one(self) -> None:
        cases = ("config", "compiler")
        for missing in cases:
            with self.subTest(missing=missing), _b1_fixture("A") as fixture:
                (fixture.config if missing == "config" else fixture.compiler).unlink()
                with (
                    _patch_fixture_paths(self.runner, fixture),
                    mock.patch.object(
                        self.runner,
                        "parse_cli_args",
                        return_value=argparse.Namespace(scenarios=None),
                    ),
                    mock.patch.object(self.runner, "run_compiler") as compiler,
                    redirect_stdout(io.StringIO()) as stdout,
                    self.assertRaises(SystemExit) as raised,
                ):
                    self.runner.main()
                self.assertEqual(raised.exception.code, 1)
                compiler.assert_not_called()
                self.assertIn("[ERROR]", stdout.getvalue())


class B1IsolatedBoundaryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.runner = _implementation(_load_b1(self.id().rsplit(".", 1)[-1]))

    def test_explicit_paths_are_stored_without_caller_cwd_lookup(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            expected_root = (Path(temp) / "nested").resolve()
            paths = self.runner.B1Paths(
                script_dir=expected_root,
                config_path=expected_root / "Config_MOMF_T1_A.yaml",
                compiler_path=expected_root / "compiler.py",
                scenarios_root=expected_root / "A1_Outputs",
            )

        self.assertEqual(paths.script_dir, expected_root)
        self.assertEqual(paths.config_path, expected_root / "Config_MOMF_T1_A.yaml")
        self.assertEqual(paths.compiler_path, expected_root / "compiler.py")
        self.assertEqual(paths.scenarios_root, expected_root / "A1_Outputs")

    def test_canonical_materialized_directory_handoff_uses_nested_a1_outputs(
        self,
    ) -> None:
        scenario = "A_Calibrated_BAU"
        workbook_name = "A-O_AR_Model_Base_Year.xlsx"
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            project = ProjectPaths(
                project_root=REPO_ROOT,
                workspace=(Path(temp) / "workspace").resolve(),
                layout="project",
            )
            expected_directory = (
                project.a1_outputs / f"A1_Outputs_{scenario}"
            )
            expected_directory.mkdir(parents=True)
            expected_workbook = expected_directory / workbook_name
            expected_workbook.write_bytes(b"materialized workbook fixture")

            with mock.patch.object(
                self.runner,
                "resolve_paths",
                return_value=project,
            ):
                paths = self.runner.B1Paths.defaults()

            params = yaml.safe_load(paths.config_path.read_text(encoding="utf-8"))
            params["xtra_scen"]["Main_Scenario"] = scenario
            plan = b1_planning.TransformPlan(
                params=params,
                base_year=params["base_year"],
                final_year=params["final_year"],
                time_range_vector=[],
                wide_param_header=params["sets"],
                other_setup_params=params["xtra_scen"],
                other_setup_params_timeslices=[],
            )
            selected_workbook = Path(
                plan.scenario_workbook("Print_Base_Year")
            ).resolve()

            self.assertEqual(paths.scenarios_root, project.a1_outputs)
            self.assertEqual(selected_workbook, expected_workbook)
            self.assertTrue(selected_workbook.is_file())
            self.assertEqual(selected_workbook.parent.parent, project.a1_outputs)
            self.assertEqual(selected_workbook.parts.count("A1_Outputs"), 1)
            self.assertEqual(
                selected_workbook.parent.name,
                f"A1_Outputs_{scenario}",
            )

    def test_scenario_resolution_is_pure_and_retains_diagnostics(self) -> None:
        discovered = ["A", "B", "C"]
        default = self.runner.resolve_scenarios(discovered, None)
        explicit = self.runner.resolve_scenarios(
            discovered, " C, Missing, A, C, Missing "
        )
        truthy_empty = self.runner.resolve_scenarios(discovered, " , ")

        self.assertEqual(default.selected, ("A", "B", "C"))
        self.assertEqual(default.requested, ())
        self.assertFalse(default.filter_active)
        self.assertEqual(explicit.selected, ("C", "A"))
        self.assertEqual(
            explicit.requested, ("C", "Missing", "A", "C", "Missing")
        )
        self.assertEqual(explicit.unknown, ("Missing", "Missing"))
        self.assertTrue(explicit.filter_active)
        self.assertEqual(truthy_empty.selected, ())
        self.assertTrue(truthy_empty.filter_active)

    def test_last_resort_regex_edge_case_is_preserved_not_silently_fixed(self) -> None:
        original = "xtra_scen:\n  Main_Scenario: ORIGINAL\n"
        self.assertEqual(
            self.runner.regex_update_main_scenario(original, "NEXT"),
            "xtra_scen:\n  Main_Scenario: 'NEXT'ORIGINAL\n",
        )

    def test_command_plan_and_injected_runner_are_separate(self) -> None:
        compiler = Path("root") / "compiler.py"
        cwd = Path("root")
        command = self.runner.build_compiler_command(
            interpreter="chosen-python", compiler_path=compiler, cwd=cwd
        )
        calls: list[tuple[object, ...]] = []

        def fake_runner(argv, *, cwd):
            calls.append((argv, cwd))
            return SimpleNamespace(returncode=6)

        return_code = self.runner.execute_command(
            command, command_runner=fake_runner
        )

        self.assertEqual(
            command.argv,
            ("chosen-python", "-B", "-m", "ostram.pipeline.compilation.compiler"),
        )
        self.assertEqual(command.cwd, cwd.resolve())
        self.assertEqual(return_code, 6)
        self.assertEqual(
            calls,
            [
                (
                    [
                        "chosen-python",
                        "-B",
                        "-m",
                        "ostram.pipeline.compilation.compiler",
                    ],
                    str(cwd.resolve()),
                )
            ],
        )

    def test_configuration_scope_exposes_backup_and_restores_on_body_exception(self) -> None:
        config = Path("Config_MOMF_T1_A.yaml")
        backup = Path("Config_MOMF_T1_A.yaml.bak")
        events: list[tuple[object, ...]] = []

        def copy_file(source, target):
            events.append(("copy", source, target))

        def move_file(source, target):
            events.append(("move", source, target))

        with self.assertRaisesRegex(RuntimeError, "body failed"):
            with self.runner.preserved_configuration(
                config,
                copy_file=copy_file,
                move_file=move_file,
                emit=lambda message: events.append(("emit", message)),
            ) as yielded_backup:
                self.assertEqual(yielded_backup, backup)
                events.append(("body",))
                raise RuntimeError("body failed")

        self.assertEqual(
            events,
            [
                ("copy", config, backup),
                ("emit", "[INFO] Backup created: Config_MOMF_T1_A.yaml.bak"),
                ("body",),
                ("move", str(backup), str(config)),
                ("emit", "\n[INFO] Restored original YAML from backup."),
            ],
        )

    def test_b1_production_path_has_only_the_compiler_process_boundary(self) -> None:
        wrapper_source = B1_ENTRYPOINT.read_text(encoding="utf-8-sig")
        helper_source = (B1_ENTRYPOINT.parent / "orchestrator.py").read_text(
            encoding="utf-8-sig"
        )
        compiler_source = (B1_ENTRYPOINT.parent / "compiler.py").read_text(
            encoding="utf-8-sig"
        )
        config_source = (
            REPO_ROOT / "config" / "compilation" / "Config_MOMF_T1_A.yaml"
        ).read_text(encoding="utf-8-sig")
        combined = wrapper_source + helper_source

        self.assertIn('script_dir / "compiler.py"', helper_source)
        self.assertIn("runner(list(command.argv), cwd=str(command.cwd))", helper_source)
        for retired_projection_dependency in (
            "Xtra_Proj",
            "A-Xtra_Projections.xlsx",
            "Projections_sheet",
            "Projections_control",
        ):
            self.assertNotIn(retired_projection_dependency, compiler_source)
            self.assertNotIn(retired_projection_dependency, config_source)
        for forbidden in (
            "python -m ostram run",
            "main_executer",
            "glpsol",
            "gurobi_cl",
            "cplex",
            "cbc",
        ):
            self.assertNotIn(forbidden, combined)

    def test_fresh_b1_extra_input_contract_contains_only_live_dependencies(
        self,
    ) -> None:
        compiler_path = B1_ENTRYPOINT.parent / "compiler.py"
        config_path = (
            REPO_ROOT / "config" / "compilation" / "Config_MOMF_T1_A.yaml"
        )
        producer_path = (
            REPO_ROOT / "ostram" / "pipeline" / "preparation" / "base_inputs.py"
        )
        effects_path = B1_ENTRYPOINT.parent / "transforms" / "effects.py"
        registry_path = REPO_ROOT / "config" / "scenarios" / "registry.json"
        templates = REPO_ROOT / "inputs" / "preparation" / "workbook_templates"

        compiler_source = compiler_path.read_text(encoding="utf-8-sig")
        config_source = config_path.read_text(encoding="utf-8-sig")
        producer_source = producer_path.read_text(encoding="utf-8-sig")
        effects_source = effects_path.read_text(encoding="utf-8-sig")
        params = yaml.safe_load(config_source)
        configured = {
            key: value
            for key, value in params.items()
            if str(key).startswith("Xtra_")
        }

        live_contract = {
            "Xtra_Emi": {
                "filename": "/A-Xtra_Emissions.xlsx",
                "template": "A-Xtra_Emissions.xlsx",
                "parsed": ("Emissions_ghg_df", "Emissions_ext_df"),
            },
            "Xtra_Storage": {
                "filename": "/A-Xtra_Storage.xlsx",
                "template": "A-Xtra_Storage.xlsx",
                "parsed": (
                    "xtra_storage_fixed_hori_param",
                    "xtra_storage_capitalcost",
                    "xtra_storage_technology_sto",
                ),
            },
        }
        self.assertEqual(
            configured,
            {
                key: contract["filename"]
                for key, contract in live_contract.items()
            },
        )

        retired_contract = {
            "Xtra_Proj": (
                "A-Xtra_Projections.xlsx",
                "Projections_sheet",
                "Projections_control",
            ),
            "Xtra_Battery": (
                "A-Xtra_Battery_Replacement.xlsx",
                "Battery_Replacement",
                "Battery_Replacement_df",
            ),
        }
        for key, retired_tokens in retired_contract.items():
            self.assertNotIn(key, configured)
            self.assertNotIn(key, config_source)
            self.assertNotIn(key, compiler_source)
            for token in retired_tokens:
                self.assertNotIn(token, config_source)
                self.assertNotIn(token, compiler_source)

        for key, contract in live_contract.items():
            filename = str(contract["template"])
            self.assertIn(f"extra_input('{key}')", compiler_source)
            self.assertIn(filename, producer_source)
            self.assertTrue((templates / filename).is_file())
            for parsed_object in contract["parsed"]:
                self.assertIn(parsed_object, compiler_source)

        self.assertNotIn("except FileNotFoundError", compiler_source)
        self.assertNotIn("except FileNotFoundError", effects_source)
        self.assertIn("return factory(path)", effects_source)
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            missing_root = Path(temp) / "fresh-extra-inputs"
            plan = b1_planning.TransformPlan(
                params={"A2_extra_inputs": str(missing_root), **configured},
                base_year=2023,
                final_year=2050,
                time_range_vector=[],
                wide_param_header=[],
                other_setup_params={},
                other_setup_params_timeslices=[],
            )
            for key in live_contract:
                with self.subTest(missing_active_key=key):
                    with self.assertRaises(FileNotFoundError):
                        b1_effects.open_workbook(plan.extra_input(key))

        registry = json.loads(registry_path.read_text(encoding="utf-8"))
        self.assertEqual(len(registry["decision_scenarios"]), 15)
        governed_sources = []
        for path in registry_path.parent.rglob("*"):
            if path.is_file() and path.suffix.lower() in {
                ".json",
                ".yaml",
                ".yml",
                ".csv",
            }:
                governed_sources.append(path.read_text(encoding="utf-8-sig"))
        combined_governed_source = "\n".join(governed_sources)
        for retired_or_legacy_token in (
            "Xtra_Proj",
            "A-Xtra_Projections.xlsx",
            "Xtra_Battery",
            "A-Xtra_Battery_Replacement.xlsx",
            "Projection.Mode",
            "Projection.Parameter",
            "Ref.km.BY",
        ):
            self.assertNotIn(retired_or_legacy_token, combined_governed_source)
        self.assertIn('"Projection.Mode": "User defined"', producer_source)
        self.assertIn('"Ref.km.BY": "not needed"', producer_source)
        self.assertFalse(params["Use_Transport"])


class B1ConfigurationAndFailureCharacterizationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.runner = _implementation(_load_b1(self.id().rsplit(".", 1)[-1]))

    @contextmanager
    def _main_context(self, fixture, scenarios: str | None):
        paths = self.runner.B1Paths(
            script_dir=fixture.script_dir,
            config_path=fixture.config,
            compiler_path=fixture.compiler,
            scenarios_root=fixture.scenario_root,
        )
        with (
            mock.patch.object(
                self.runner.B1Paths,
                "defaults",
                return_value=paths,
            ),
            mock.patch.object(
                self.runner,
                "parse_cli_args",
                return_value=argparse.Namespace(scenarios=scenarios),
            ),
        ):
            yield

    def test_predecessor_invocation_trace_and_status_order(self) -> None:
        with _b1_fixture("C", "A", "B") as fixture:
            events: list[tuple[object, ...]] = []
            real_copy = self.runner.shutil.copy2
            real_move = self.runner.shutil.move

            def copy_file(source, target):
                events.append(("copy2", source, target))
                return real_copy(source, target)

            def move_file(source, target):
                events.append(("move", source, target))
                return real_move(source, target)

            with (
                self._main_context(fixture, "C,A,C"),
                mock.patch.object(self.runner.shutil, "copy2", side_effect=copy_file),
                mock.patch.object(self.runner.shutil, "move", side_effect=move_file),
                mock.patch.object(
                    self.runner,
                    "update_main_scenario",
                    side_effect=_fixture_update(events),
                ),
                mock.patch.object(
                    self.runner,
                    "run_compiler",
                    side_effect=lambda path: events.append(("compile", path)) or 0,
                ),
                redirect_stdout(io.StringIO()) as stdout,
            ):
                result = self.runner.main()

            self.assertIsNone(result)
            self.assertEqual(
                events,
                [
                    ("copy2", fixture.config, fixture.backup),
                    ("update", fixture.config, "C"),
                    ("compile", fixture.script_dir),
                    ("update", fixture.config, "A"),
                    ("compile", fixture.script_dir),
                    ("move", str(fixture.backup), str(fixture.config)),
                ],
            )
            output = stdout.getvalue()
            markers = (
                "Scenario filter active",
                "Scenarios discovered",
                "Backup created",
                "Running scenario: C",
                "completed successfully for scenario 'C'",
                "Running scenario: A",
                "completed successfully for scenario 'A'",
                "Restored original YAML",
                "All done",
            )
            positions = [output.index(marker) for marker in markers]
            self.assertEqual(positions, sorted(positions))
            self.assertEqual(fixture.config.read_bytes(), fixture.original)
            self.assertFalse(fixture.backup.exists())

    def test_nonzero_compiler_exit_is_reported_continues_and_final_exit_is_zero(self) -> None:
        with _b1_fixture("A", "B", "C") as fixture:
            events: list[tuple[object, ...]] = []
            return_codes = iter((4, 0, 9))
            with (
                self._main_context(fixture, None),
                mock.patch.object(
                    self.runner,
                    "update_main_scenario",
                    side_effect=_fixture_update(events),
                ),
                mock.patch.object(self.runner, "run_compiler", side_effect=return_codes),
                redirect_stdout(io.StringIO()) as stdout,
                self.assertRaises(subprocess.CalledProcessError) as raised,
            ):
                self.runner.main()

            self.assertEqual(raised.exception.returncode, 4)
            self.assertEqual(
                [event[2] for event in events if event[0] == "update"],
                ["A"],
            )
            output = stdout.getvalue()
            self.assertIn("exited with code 4 for scenario 'A'", output)
            self.assertNotIn("Running scenario: B", output)
            self.assertNotIn("Running scenario: C", output)
            self.assertNotIn("[INFO] All done.", output)
            self.assertEqual(fixture.config.read_bytes(), fixture.original)

    def test_update_error_skips_only_that_compiler_and_continues(self) -> None:
        with _b1_fixture("A", "B", "C") as fixture:
            events: list[tuple[str, str]] = []

            def update(path: Path, scenario: str) -> None:
                events.append(("update", scenario))
                if scenario == "B":
                    raise ValueError("fixture update failure")
                path.write_text(f"scenario: {scenario}\n", encoding="utf-8")

            def compile_scenario(_path: Path) -> int:
                events.append(("compile", events[-1][1]))
                return 0

            with (
                self._main_context(fixture, None),
                mock.patch.object(self.runner, "update_main_scenario", side_effect=update),
                mock.patch.object(self.runner, "run_compiler", side_effect=compile_scenario),
                redirect_stdout(io.StringIO()) as stdout,
            ):
                result = self.runner.main()

            self.assertIsNone(result)
            self.assertEqual(
                events,
                [
                    ("update", "A"),
                    ("compile", "A"),
                    ("update", "B"),
                    ("update", "C"),
                    ("compile", "C"),
                ],
            )
            self.assertIn("Failed to update YAML for scenario 'B'", stdout.getvalue())
            self.assertEqual(fixture.config.read_bytes(), fixture.original)

    def test_unexpected_compiler_exception_restores_then_propagates_without_all_done(self) -> None:
        with _b1_fixture("A", "B") as fixture:
            with (
                self._main_context(fixture, None),
                mock.patch.object(
                    self.runner,
                    "update_main_scenario",
                    side_effect=lambda path, scenario: path.write_text(
                        f"scenario: {scenario}\n", encoding="utf-8"
                    ),
                ),
                mock.patch.object(
                    self.runner,
                    "run_compiler",
                    side_effect=RuntimeError("launch failed"),
                ),
                redirect_stdout(io.StringIO()) as stdout,
                self.assertRaisesRegex(RuntimeError, "launch failed"),
            ):
                self.runner.main()

            self.assertEqual(fixture.config.read_bytes(), fixture.original)
            self.assertFalse(fixture.backup.exists())
            self.assertIn("Restored original YAML", stdout.getvalue())
            self.assertNotIn("All done", stdout.getvalue())

    def test_keyboard_interrupt_restores_then_propagates(self) -> None:
        with _b1_fixture("A") as fixture:
            with (
                self._main_context(fixture, None),
                mock.patch.object(
                    self.runner,
                    "update_main_scenario",
                    side_effect=lambda path, scenario: path.write_text(
                        f"scenario: {scenario}\n", encoding="utf-8"
                    ),
                ),
                mock.patch.object(
                    self.runner, "run_compiler", side_effect=KeyboardInterrupt
                ),
                redirect_stdout(io.StringIO()),
                self.assertRaises(KeyboardInterrupt),
            ):
                self.runner.main()

            self.assertEqual(fixture.config.read_bytes(), fixture.original)
            self.assertFalse(fixture.backup.exists())

    def test_restore_failure_is_warned_swallowed_and_still_prints_all_done(self) -> None:
        with _b1_fixture("A") as fixture:
            with (
                self._main_context(fixture, None),
                mock.patch.object(
                    self.runner,
                    "update_main_scenario",
                    side_effect=lambda path, scenario: path.write_text(
                        f"scenario: {scenario}\n", encoding="utf-8"
                    ),
                ),
                mock.patch.object(self.runner, "run_compiler", return_value=0),
                mock.patch.object(
                    self.runner.shutil,
                    "move",
                    side_effect=OSError("restore blocked"),
                ),
                redirect_stdout(io.StringIO()) as stdout,
            ):
                result = self.runner.main()

            self.assertIsNone(result)
            self.assertNotEqual(fixture.config.read_bytes(), fixture.original)
            self.assertTrue(fixture.backup.exists())
            output = stdout.getvalue()
            self.assertIn("Could not restore YAML from backup: restore blocked", output)
            self.assertIn(f"Backup still available at: {fixture.backup}", output)
            self.assertIn("[INFO] All done.", output)

    def test_backup_creation_failure_is_outside_restore_scope(self) -> None:
        with _b1_fixture("A") as fixture:
            with (
                self._main_context(fixture, None),
                mock.patch.object(
                    self.runner.shutil,
                    "copy2",
                    side_effect=OSError("backup blocked"),
                ),
                mock.patch.object(self.runner.shutil, "move") as move_file,
                mock.patch.object(self.runner, "update_main_scenario") as update,
                mock.patch.object(self.runner, "run_compiler") as compiler,
                redirect_stdout(io.StringIO()),
                self.assertRaisesRegex(OSError, "backup blocked"),
            ):
                self.runner.main()

            move_file.assert_not_called()
            update.assert_not_called()
            compiler.assert_not_called()
            self.assertEqual(fixture.config.read_bytes(), fixture.original)


if __name__ == "__main__":
    unittest.main()
