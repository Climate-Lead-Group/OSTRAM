from __future__ import annotations

import ast
import builtins
import importlib
import io
import os
import pickle
import subprocess
import sys
import unittest
from collections.abc import Iterable
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock

import numpy as np
import pandas as pd
from pandas.testing import assert_frame_equal


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
B1_COMPILER = REPO_ROOT / "ostram" / "pipeline" / "compilation" / "compiler.py"
TRANSFORM_PACKAGE = "ostram.pipeline.compilation.transforms"
TRANSFORM_MODULES = (
    "planning",
    "tables",
    "effects",
    "validation",
    "delivery",
)

STRUCTURE_KEYS = (
    "YEAR",
    "TECHNOLOGY",
    "TIMESLICE",
    "FUEL",
    "EMISSION",
    "MODE_OF_OPERATION",
    "REGION",
    "DAYTYPE",
    "DAILYTIMEBRACKET",
    "SEASON",
    "STORAGE",
)


def _source() -> str:
    return B1_COMPILER.read_text(encoding="utf-8-sig")


def _tree() -> ast.Module:
    return ast.parse(_source(), filename=str(B1_COMPILER))


def _execute_nodes(
    nodes: Iterable[ast.stmt], namespace: dict[str, object]
) -> dict[str, object]:
    module = ast.Module(body=list(nodes), type_ignores=[])
    ast.fix_missing_locations(module)
    exec(compile(module, str(B1_COMPILER), "exec"), namespace)
    return namespace


def _qualified_name(node: ast.AST) -> str | None:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        prefix = _qualified_name(node.value)
        if prefix is not None:
            return f"{prefix}.{node.attr}"
    return None


def _top_level_node_containing_call(qualified_name: str) -> ast.stmt:
    matches = [
        node
        for node in _tree().body
        if any(
            isinstance(child, ast.Call)
            and _qualified_name(child.func) == qualified_name
            for child in ast.walk(node)
        )
    ]
    if len(matches) != 1:
        raise AssertionError(
            f"expected one B1 node calling {qualified_name}, got {matches!r}"
        )
    return matches[0]


def _candidate_modules() -> dict[str, object]:
    return {
        name: importlib.import_module(f"{TRANSFORM_PACKAGE}.{name}")
        for name in TRANSFORM_MODULES
    }


class _WorkbookFixture:
    def __init__(self, sheet_names: list[str], frame: pd.DataFrame) -> None:
        self.sheet_names = sheet_names
        self.frame = frame
        self.parse_calls: list[str] = []

    def parse(self, sheet_name: str) -> pd.DataFrame:
        self.parse_calls.append(sheet_name)
        return self.frame.copy(deep=True)


def _run_system_parameter_block(
    workbook: _WorkbookFixture,
    *,
    years: list[int],
    main: dict[object, pd.DataFrame] | None = None,
    additional: dict[object, pd.DataFrame] | None = None,
) -> tuple[dict[str, object], str]:
    tables = _candidate_modules()["tables"]
    namespace: dict[str, object] = {
        "pd": pd,
        "_tables": tables,
        "Parametrization": workbook,
        "time_range_vector": years,
        "other_setup_params": {"Main_Scenario": "MAIN", "Region": "GLOBAL"},
        "overall_param_df_dict": {} if main is None else main,
        "overall_param_df_dict_ndp": {} if additional is None else additional,
    }
    stdout = io.StringIO()
    with redirect_stdout(stdout):
        _execute_nodes(
            [_top_level_node_containing_call("_tables.build_system_parameter_rows")],
            namespace,
        )
    return namespace, stdout.getvalue()


class _FakeWriter:
    def __init__(self, recorder: "_DeliveryRecorder", path: object, engine: str) -> None:
        self.recorder = recorder
        self.path = path
        self.engine = engine

    def close(self) -> None:
        self.recorder.events.append(("excel_close", self.path))


class _DeliveryRecorder:
    def __init__(
        self,
        *,
        fail_excel_sheet: str | None = None,
        fail_csv_name: str | None = None,
    ) -> None:
        self.events: list[tuple[object, ...]] = []
        self.excel_frames: list[tuple[object, str, pd.DataFrame]] = []
        self.csv_frames: list[tuple[object, pd.DataFrame]] = []
        self.fail_excel_sheet = fail_excel_sheet
        self.fail_csv_name = fail_csv_name

    def excel_writer(self, path: object, *, engine: str) -> _FakeWriter:
        self.events.append(("excel_open", path, engine))
        return _FakeWriter(self, path, engine)

    def to_excel(
        self,
        frame: pd.DataFrame,
        writer: _FakeWriter,
        *,
        sheet_name: str,
        index: bool,
    ) -> None:
        self.events.append(("excel_write", writer.path, sheet_name, index))
        self.excel_frames.append((writer.path, sheet_name, frame.copy(deep=True)))
        if sheet_name == self.fail_excel_sheet:
            raise OSError(f"fixture Excel failure: {sheet_name}")

    def to_csv(
        self,
        frame: pd.DataFrame,
        path: object,
        *,
        index: bool,
        header: bool,
    ) -> None:
        self.events.append(("csv_write", path, index, header))
        self.csv_frames.append((path, frame.copy(deep=True)))
        if os.path.basename(os.fspath(path)) == self.fail_csv_name:
            raise OSError(f"fixture CSV failure: {self.fail_csv_name}")

    def makedirs(self, path: object, *, exist_ok: bool) -> None:
        self.events.append(("mkdir", path, exist_ok))


def _delivery_fixture(
    *, recorder: _DeliveryRecorder | None = None
) -> tuple[dict[str, object], _DeliveryRecorder, pd.DataFrame]:
    planning = _candidate_modules()["planning"]
    effect_recorder = _DeliveryRecorder() if recorder is None else recorder
    columns4 = list(STRUCTURE_KEYS)
    setup = {
        "Main_Scenario": "MAIN",
        "Other_Scenarios": ["C", "A", "C"],
        "Timeslices": ["L2", "L1"],
        "Mode_of_Operation": [2, 1],
        "Region": "GLOBAL",
        "Season": [2],
        "DayType": [1],
        "DailyTimeBracket": [3],
        "Storage": ["STO1"],
    }
    params = {
        "A1_outputs": "A1ROOT",
        "xtra_scen": setup,
        "Print_Dem_Completed": "/Demand_COMPLETED.xlsx",
        "initial_year": "2023",
        "A_O_Dem": "Demand",
        "Print_Paramet_Completed": "/Param_COMPLETED.xlsx",
        "Print_Paramet_Natural_Completed": "/Param_Natural_COMPLETED.xlsx",
        "Print_Proj_Completed": "/AR_COMPLETED.xlsx",
        "columns4": columns4,
        "Print_A2_Struct_List": "A2_Structure_Lists.xlsx",
        "lists": "Lists",
        "A2_output_main_scen": "A2_MAIN_ROOT",
        "A2_output": "A2_OTHER_ROOT",
    }
    transform_plan = planning.TransformPlan(
        params=params,
        base_year="2023",
        final_year="2024",
        time_range_vector=[2023, 2024],
        wide_param_header=None,
        other_setup_params=setup,
        other_setup_params_timeslices=setup["Timeslices"],
    )

    demand = pd.DataFrame({"2023": ["1.23456"], "Other": [2.34567]})
    parameter_z = pd.DataFrame({"raw": [1.23456]})
    parameter_a = pd.DataFrame({"raw": [2.34567]})
    natural_b = pd.DataFrame({"raw": [3.45678]})
    natural_a = pd.DataFrame({"raw": [4.56789]})
    projection_secondary = pd.DataFrame({"raw": [5.67891]})
    projection_primary = pd.DataFrame({"raw": [6.78912]})

    z_rows = pd.DataFrame(
        {
            "PARAMETER": ["ZParam", "ZParam"],
            "Scenario": ["MAIN", "MAIN"],
            "TECHNOLOGY": ["TECH", "TECH"],
            "YEAR": [2023, 2023],
            "Value": [1.0, np.nan],
        }
    )
    a_rows = pd.DataFrame(
        {
            "PARAMETER": ["AParam", "AParam", "AParam"],
            "Scenario": ["MAIN", "OTHER", np.nan],
            "TECHNOLOGY": ["A", "B", "C"],
            "YEAR": [2023, 2024, 2025],
            "Value": [2.0, 3.0, 4.0],
        }
    )
    dropped_nan_key = pd.DataFrame({"Value": [999.0]})
    ndp_only = pd.DataFrame({"Scenario": ["MAIN"], "Value": [777.0]})

    namespace: dict[str, object] = {
        "params": params,
        "transform_plan": transform_plan,
        "Demand_df_new": demand,
        "params_dict_new": {"ZSheet": parameter_z, "ASheet": parameter_a},
        "params_dict_new_natural": {
            "NaturalB": natural_b,
            "NaturalA": natural_a,
        },
        "AR_Base_proj_df_new": {
            "Secondary": projection_secondary,
            "Primary": projection_primary,
        },
        "time_range_vector": [2023, 2024],
        "All_Tech_list": ["TECH2", "TECH1", "TECH1"],
        "other_setup_params": setup,
        "All_Fuel_list": ["FUEL1"],
        "emissions_list": ["CO2"],
        "overall_param_df_dict": {
            "ZParam": z_rows,
            np.nan: dropped_nan_key,
            "AParam": a_rows,
        },
        "overall_param_df_dict_ndp": {"NDPOnly": ndp_only},
    }
    return namespace, effect_recorder, demand


def _run_delivery(
    namespace: dict[str, object], recorder: _DeliveryRecorder
) -> dict[str, object]:
    modules = _candidate_modules()
    tables = modules["tables"]
    effects = modules["effects"]
    delivery = modules["delivery"]
    params = namespace["params"]
    plan = namespace["transform_plan"]
    setup = namespace["other_setup_params"]

    namespace["Demand_df_new"] = effects.write_completed_demand_workbook(
        plan.scenario_workbook("Print_Dem_Completed"),
        namespace["Demand_df_new"],
        params["initial_year"],
        params["A_O_Dem"],
        writer_factory=recorder.excel_writer,
        write_frame=recorder.to_excel,
    )
    effects.write_sheet_mapping_workbook(
        plan.scenario_workbook("Print_Paramet_Completed"),
        namespace["params_dict_new"],
        writer_factory=recorder.excel_writer,
        write_frame=recorder.to_excel,
    )
    effects.write_sheet_mapping_workbook(
        plan.scenario_workbook("Print_Paramet_Natural_Completed"),
        namespace["params_dict_new_natural"],
        writer_factory=recorder.excel_writer,
        write_frame=recorder.to_excel,
    )
    effects.write_sheet_mapping_workbook(
        plan.scenario_workbook("Print_Proj_Completed"),
        namespace["AR_Base_proj_df_new"],
        writer_factory=recorder.excel_writer,
        write_frame=recorder.to_excel,
    )
    structure = tables.build_structure_tables(
        namespace["time_range_vector"],
        namespace["All_Tech_list"],
        namespace["All_Fuel_list"],
        namespace["emissions_list"],
        setup,
        params,
    )
    namespace["df_structure"] = structure.table
    namespace["structure_dict"] = structure.values
    effects.write_structure_workbook(
        plan.structure_workbook(),
        structure.table,
        params["lists"],
        writer_factory=recorder.excel_writer,
        write_frame=recorder.to_excel,
    )
    main_tables, additional_tables = delivery.clean_parameter_tables(
        namespace["overall_param_df_dict"]
    )
    namespace["overall_param_df_dict"] = main_tables
    namespace["overall_param_df_dict_ndp"] = additional_tables
    delivery.deliver_main_csvs(
        plan.main_output_root(),
        setup["Main_Scenario"],
        main_tables,
        structure.values,
        makedirs=recorder.makedirs,
        dataframe_factory=pd.DataFrame,
        csv_writer=recorder.to_csv,
    )
    if setup["Other_Scenarios"]:
        delivery.deliver_additional_csvs(
            plan.additional_output_root(),
            setup["Other_Scenarios"],
            setup["Main_Scenario"],
            additional_tables,
            structure.values,
            makedirs=recorder.makedirs,
            dataframe_factory=pd.DataFrame,
            csv_writer=recorder.to_csv,
        )
    return namespace


class B1CandidateImportSafetyTests(unittest.TestCase):
    def test_all_candidate_modules_import_without_io_process_or_cwd_effects(self) -> None:
        saved_modules = {
            name: module
            for name, module in list(sys.modules.items())
            if name == TRANSFORM_PACKAGE or name.startswith(TRANSFORM_PACKAGE + ".")
        }
        for name in saved_modules:
            sys.modules.pop(name, None)

        stdout = io.StringIO()
        stderr = io.StringIO()
        cwd_before = os.getcwd()
        patches = (
            mock.patch.object(builtins, "open", side_effect=AssertionError("open")),
            mock.patch.object(pd, "ExcelFile", side_effect=AssertionError("ExcelFile")),
            mock.patch.object(
                pd, "ExcelWriter", side_effect=AssertionError("ExcelWriter")
            ),
            mock.patch.object(pd, "read_csv", side_effect=AssertionError("read_csv")),
            mock.patch.object(
                pd.DataFrame, "to_excel", side_effect=AssertionError("to_excel")
            ),
            mock.patch.object(
                pd.DataFrame, "to_csv", side_effect=AssertionError("to_csv")
            ),
            mock.patch.object(os, "makedirs", side_effect=AssertionError("makedirs")),
            mock.patch.object(os, "chdir", side_effect=AssertionError("chdir")),
            mock.patch.object(pickle, "load", side_effect=AssertionError("pickle")),
            mock.patch.object(
                subprocess, "run", side_effect=AssertionError("subprocess.run")
            ),
            mock.patch.object(
                subprocess, "Popen", side_effect=AssertionError("subprocess.Popen")
            ),
        )
        started: list[mock._patch] = []
        try:
            for patcher in patches:
                patcher.start()
                started.append(patcher)
            with redirect_stdout(stdout), redirect_stderr(stderr):
                imported = _candidate_modules()
        finally:
            for patcher in reversed(started):
                patcher.stop()
            for name in list(sys.modules):
                if name == TRANSFORM_PACKAGE or name.startswith(
                    TRANSFORM_PACKAGE + "."
                ):
                    sys.modules.pop(name, None)
            sys.modules.update(saved_modules)

        self.assertEqual(tuple(imported), TRANSFORM_MODULES)
        self.assertEqual(stdout.getvalue(), "")
        self.assertEqual(stderr.getvalue(), "")
        self.assertEqual(os.getcwd(), cwd_before)

    def test_b1_import_node_uses_only_package_relative_transforms(self) -> None:
        import_nodes = [
            node
            for node in _tree().body
            if isinstance(node, ast.ImportFrom)
            and node.level == 1
            and node.module == "transforms"
        ]
        self.assertEqual(len(import_nodes), 5)
        self.assertEqual(
            [node.names[0].asname for node in import_nodes],
            ["_delivery", "_effects", "_planning", "_tables", "_validation"],
        )


class B1CompilerExtractionContractTests(unittest.TestCase):
    def test_compiler_delegates_each_extracted_boundary_in_predecessor_order(self) -> None:
        tree = _tree()
        calls = sorted(
            (
                node.lineno,
                node.col_offset,
                _qualified_name(node.func),
            )
            for node in ast.walk(tree)
            if isinstance(node, ast.Call)
        )
        selected_names = {
            "_effects.read_config",
            "_planning.build_transform_plan",
            "_validation.validate_demand_timeslices",
            "_validation.validate_capacity_timeslices",
            "_validation.validate_yearsplit_timeslices",
            "_validation.validate_daysplit_time_brackets",
            "_tables.build_system_parameter_rows",
            "_effects.write_completed_demand_workbook",
            "_effects.write_sheet_mapping_workbook",
            "_tables.build_structure_tables",
            "_effects.write_structure_workbook",
            "_delivery.clean_parameter_tables",
            "_delivery.deliver_main_csvs",
            "_delivery.deliver_additional_csvs",
        }
        actual = [name for _line, _column, name in calls if name in selected_names]
        self.assertEqual(
            actual,
            [
                "_effects.read_config",
                "_planning.build_transform_plan",
                "_validation.validate_demand_timeslices",
                "_validation.validate_capacity_timeslices",
                "_validation.validate_yearsplit_timeslices",
                "_validation.validate_daysplit_time_brackets",
                "_tables.build_system_parameter_rows",
                "_effects.write_completed_demand_workbook",
                "_effects.write_sheet_mapping_workbook",
                "_effects.write_sheet_mapping_workbook",
                "_effects.write_sheet_mapping_workbook",
                "_tables.build_structure_tables",
                "_effects.write_structure_workbook",
                "_delivery.clean_parameter_tables",
                "_delivery.deliver_main_csvs",
                "_delivery.deliver_additional_csvs",
            ],
        )

        aliases = {
            alias.asname
            for node in tree.body
            if isinstance(node, ast.ImportFrom)
            and node.level == 1
            and node.module == "transforms"
            for alias in node.names
        }
        self.assertTrue(
            {"_planning", "_tables", "_effects", "_validation", "_delivery"}
            <= aliases
        )
        self.assertFalse(
            any(
                isinstance(node, ast.FunctionDef)
                and node.name == "normalize_year_like_columns"
                for node in tree.body
            )
        )


class B1ConfigurationAndSourcePathCharacterizationTests(unittest.TestCase):
    def test_config_is_cwd_relative_and_preserves_setup_mutation(self) -> None:
        modules = _candidate_modules()
        effects = modules["effects"]
        planning = modules["planning"]
        loaded = {
            "base_year": "2023",
            "final_year": "2025",
            "sets": ["PARAMETER", "Value"],
            "xtra_scen": {
                "Region": "GLOBAL",
                "Timeslices": ["L3", "L1", "L2"],
                "Main_Scenario": "MAIN",
            },
        }
        opened: list[tuple[object, ...]] = []

        class OpenFixture:
            def __call__(self, *args, **kwargs):
                opened.append((*args, kwargs))
                return io.StringIO("fixture config")

        cwd_before = os.getcwd()
        params = effects.read_config(
            planning.CONFIG_PATH,
            opener=OpenFixture(),
            loader=lambda _stream: loaded,
        )
        plan = planning.build_transform_plan(params)

        self.assertEqual(opened, [("Config_MOMF_T1_A.yaml", "r", {})])
        self.assertEqual(os.getcwd(), cwd_before)
        self.assertEqual(plan.time_range_vector, [2023, 2024, 2025])
        self.assertEqual(plan.wide_param_header, ["PARAMETER", "Value"])
        self.assertEqual(
            loaded["xtra_scen"]["Timeslices"], ["L1", "L2", "L3"]
        )
        self.assertEqual(
            plan.other_setup_params,
            {
                "Region": "GLOBAL",
                "Timeslices": ["L1", "L2", "L3"],
                "Main_Scenario": "MAIN",
            },
        )
        self.assertIsNot(plan.other_setup_params, loaded["xtra_scen"])
        self.assertIs(
            plan.other_setup_params["Timeslices"],
            loaded["xtra_scen"]["Timeslices"],
        )

    def test_config_and_plan_failures_propagate_without_recovery(self) -> None:
        modules = _candidate_modules()
        effects = modules["effects"]
        planning = modules["planning"]

        with self.assertRaisesRegex(OSError, "fixture config failure"):
            effects.read_config(
                planning.CONFIG_PATH,
                opener=mock.Mock(side_effect=OSError("fixture config failure")),
            )
        with self.assertRaises(KeyError):
            planning.build_transform_plan(
                {"base_year": 2023, "final_year": 2024, "sets": []}
            )
        with self.assertRaises(ValueError):
            planning.build_transform_plan(
                {
                    "base_year": "not-a-year",
                    "final_year": 2024,
                    "sets": [],
                    "xtra_scen": {"Timeslices": []},
                }
            )

    def test_workbook_and_csv_source_paths_keep_exact_formulas_and_order(self) -> None:
        modules = _candidate_modules()
        effects = modules["effects"]
        planning = modules["planning"]
        calls: list[tuple[str, object]] = []

        def excel_file(path: object) -> object:
            calls.append(("excel", path))
            return object()

        def read_csv(path: object) -> object:
            calls.append(("csv", path))
            return pd.DataFrame({"VALUE": ["CO2"]})

        params = {
            "base_year": 2023,
            "final_year": 2023,
            "sets": [],
            "A1_outputs": "A1ROOT",
            "xtra_scen": {
                "Main_Scenario": "SCENARIO",
                "Timeslices": [],
            },
            "Print_Base_Year": "/Base.xlsx",
            "Print_Proj": "/Projection.xlsx",
            "Print_Demand": "/Demand.xlsx",
            "A2_extra_inputs": "EXTRA",
            "Xtra_Proj": "/ExtraProjection.xlsx",
            "Xtra_Battery": "/Battery.xlsx",
            "Print_Paramet": "/Parametrization.xlsx",
            "Xtra_Emi": "/Emissions.xlsx",
            "Use_OG_module": True,
            "Xtra_Storage": "/Storage.xlsx",
        }
        plan = planning.build_transform_plan(params)
        effects.open_workbook(
            plan.scenario_workbook("Print_Base_Year"), factory=excel_file
        )
        effects.open_workbook(
            plan.scenario_workbook("Print_Proj"), factory=excel_file
        )
        effects.open_workbook(
            plan.scenario_workbook("Print_Demand"), factory=excel_file
        )
        effects.open_workbook(plan.extra_input("Xtra_Proj"), factory=excel_file)
        effects.open_workbook(plan.extra_input("Xtra_Battery"), factory=excel_file)
        effects.open_workbook(
            plan.scenario_workbook("Print_Paramet"), factory=excel_file
        )
        effects.open_workbook(plan.extra_input("Xtra_Emi"), factory=excel_file)
        effects.read_csv(plan.og_emissions_csv(), reader=read_csv)
        effects.open_workbook(plan.extra_input("Xtra_Storage"), factory=excel_file)

        scenario_root = "A1ROOT_SCENARIO"
        self.assertEqual(
            calls,
            [
                ("excel", os.path.join("A1ROOT", scenario_root + "/Base.xlsx")),
                (
                    "excel",
                    os.path.join("A1ROOT", scenario_root + "/Projection.xlsx"),
                ),
                ("excel", os.path.join("A1ROOT", scenario_root + "/Demand.xlsx")),
                ("excel", "EXTRA/ExtraProjection.xlsx"),
                ("excel", "EXTRA/Battery.xlsx"),
                (
                    "excel",
                    os.path.join(
                        "A1ROOT", scenario_root + "/Parametrization.xlsx"
                    ),
                ),
                ("excel", "EXTRA/Emissions.xlsx"),
                (
                    "csv",
                    str(REPO_ROOT / "inputs" / "osemosys_global" / "EMISSION.csv"),
                ),
                ("excel", "EXTRA/Storage.xlsx"),
            ],
        )

    def test_pickle_boundary_preserves_binary_open_and_native_failures(self) -> None:
        effects = _candidate_modules()["effects"]
        handle = object()
        calls: list[tuple[object, ...]] = []

        def opener(*args):
            calls.append(("open", *args))
            return handle

        def loader(stream):
            calls.append(("load", stream))
            return {"fleet": 1}

        self.assertEqual(
            effects.load_pickle("A1ROOT_SCENARIO/Fleet_Groups.pkl", opener=opener, loader=loader),
            {"fleet": 1},
        )
        self.assertEqual(
            calls,
            [
                ("open", "A1ROOT_SCENARIO/Fleet_Groups.pkl", "rb"),
                ("load", handle),
            ],
        )
        with self.assertRaisesRegex(OSError, "fixture pickle failure"):
            effects.load_pickle(
                "missing.pkl",
                opener=mock.Mock(side_effect=OSError("fixture pickle failure")),
            )


class B1ValidationCharacterizationTests(unittest.TestCase):
    @staticmethod
    def _capture_abort(function, *args) -> tuple[list[str], SystemExit]:
        messages: list[str] = []

        def stop() -> None:
            raise SystemExit()

        with unittest.TestCase().assertRaises(SystemExit) as raised:
            function(*args, emit=messages.append, stop=stop)
        return messages, raised.exception

    def test_demand_and_capacity_validation_messages_and_precedence(self) -> None:
        validation = _candidate_modules()["validation"]
        cases = (
            (
                validation.validate_demand_timeslices,
                (["L1"], "Some", ["L2"]),
                validation.DEMAND_TIMESLICE_MISMATCH,
            ),
            (
                validation.validate_demand_timeslices,
                (["L1"], "Some", []),
                validation.DEMAND_TIMESLICE_DEFINITION_ERROR,
            ),
            (
                validation.validate_demand_timeslices,
                (["L1"], "All", ["L1"]),
                validation.DEMAND_TIMESLICE_DEFINITION_ERROR,
            ),
            (
                validation.validate_capacity_timeslices,
                (["L1"], "Some", ["L2"]),
                validation.CAPACITY_TIMESLICE_MISMATCH,
            ),
            (
                validation.validate_capacity_timeslices,
                (["L1"], "Some", []),
                validation.CAPACITY_TIMESLICE_DEFINITION_ERROR,
            ),
            (
                validation.validate_capacity_timeslices,
                (["L1"], "All", ["L1"]),
                validation.CAPACITY_TIMESLICE_DEFINITION_ERROR,
            ),
        )
        for function, args, expected_message in cases:
            with self.subTest(function=function.__name__, args=args):
                messages, error = self._capture_abort(function, *args)
                self.assertEqual(messages, [expected_message])
                self.assertIsNone(error.code)

        for function in (
            validation.validate_demand_timeslices,
            validation.validate_capacity_timeslices,
        ):
            messages: list[str] = []
            function(["L1"], "Some", ["L1"], emit=messages.append)
            function(["L1"], "All", [], emit=messages.append)
            self.assertEqual(messages, [])

    def test_yearsplit_and_daysplit_keep_legacy_failure_boundaries(self) -> None:
        validation = _candidate_modules()["validation"]
        messages, error = self._capture_abort(
            validation.validate_yearsplit_timeslices, "Some", []
        )
        self.assertEqual(messages, [validation.YEARSPLIT_TIMESLICE_ERROR])
        self.assertIsNone(error.code)

        messages, error = self._capture_abort(
            validation.validate_daysplit_time_brackets, [1], []
        )
        self.assertEqual(messages, [validation.DAYSPLIT_TIME_BRACKET_ERROR])
        self.assertIsNone(error.code)

        messages = []
        validation.validate_yearsplit_timeslices(
            "Some", ["L1"], emit=messages.append
        )
        validation.validate_daysplit_time_brackets(
            [1], ["L1"], emit=messages.append
        )
        validation.validate_daysplit_time_brackets(
            [1, 2], [], emit=messages.append
        )
        self.assertEqual(messages, [])

    def test_default_abort_prints_and_uses_bare_system_exit(self) -> None:
        validation = _candidate_modules()["validation"]
        stdout = io.StringIO()
        with redirect_stdout(stdout), self.assertRaises(SystemExit) as raised:
            validation.validate_yearsplit_timeslices("Some", [])
        self.assertIsNone(raised.exception.code)
        self.assertEqual(
            stdout.getvalue(), validation.YEARSPLIT_TIMESLICE_ERROR + "\n"
        )


class B1YearColumnNormalizationCharacterizationTests(unittest.TestCase):
    def test_headers_normalize_without_changing_values_index_or_dtypes(self) -> None:
        normalize = _candidate_modules()["tables"].normalize_year_like_columns
        frame = pd.DataFrame(
            [[1, 2.5, np.nan, "x", 5, 6, 7]],
            index=pd.Index([9], name="row"),
            columns=[
                2023,
                np.int64(2024),
                2025.0,
                2025.5,
                " 002026 ",
                " year ",
                "",
            ],
        )
        original = frame.copy(deep=True)

        actual = normalize(frame)

        expected = original.copy(deep=True)
        expected.columns = [
            "2023",
            "2024",
            "2025",
            2025.5,
            "2026",
            " year ",
            "",
        ]
        assert_frame_equal(actual, expected)
        assert_frame_equal(frame, original)
        self.assertIsNot(actual, frame)

    def test_noop_identity_and_duplicate_normalized_columns_are_preserved(self) -> None:
        normalize = _candidate_modules()["tables"].normalize_year_like_columns
        no_op = pd.DataFrame({"year": pd.Series([1], dtype="int64")})
        self.assertIs(normalize(no_op), no_op)

        duplicates = pd.DataFrame(
            [[1.0, np.nan, 3.0]], columns=[2023, "2023", " 2023 "]
        )
        actual = normalize(duplicates)
        self.assertEqual(list(actual.columns), ["2023", "2023", "2023"])
        self.assertEqual(actual.iloc[0, 0], 1.0)
        self.assertTrue(pd.isna(actual.iloc[0, 1]))
        self.assertEqual(actual.iloc[0, 2], 3.0)
        self.assertEqual(list(duplicates.columns), [2023, "2023", " 2023 "])


class B1SystemParameterCharacterizationTests(unittest.TestCase):
    def test_rows_years_missing_values_and_duplicates_match_predecessor(self) -> None:
        frame = pd.DataFrame(
            [
                ["ReserveMargin", 1.23456, np.nan, 1.0],
                ["ReserveMargin", 2.0, 3.33336, np.nan],
                ["EmissionPenalty", np.nan, 4.44446, 5.0],
            ],
            index=[7, 3, 9],
            columns=["Parameter", 2023, 2024.0, " 2025 "],
        )
        old_main = pd.DataFrame({"old": [1]})
        old_additional = pd.DataFrame({"old": [2]})
        main = {"ReserveMargin": old_main, "Keep": pd.DataFrame({"x": [1]})}
        additional = {
            "ReserveMargin": old_additional,
            "Keep": pd.DataFrame({"x": [2]}),
        }
        workbook = _WorkbookFixture(["Capacities", "System Parameters"], frame)

        namespace, output = _run_system_parameter_block(
            workbook,
            years=[2023, 2024, 2025],
            main=main,
            additional=additional,
        )

        expected_reserve = pd.DataFrame(
            [
                ["ReserveMargin", "MAIN", "GLOBAL", 2023, 1.2346],
                ["ReserveMargin", "MAIN", "GLOBAL", 2025, 1.0],
                ["ReserveMargin", "MAIN", "GLOBAL", 2023, 2.0],
                ["ReserveMargin", "MAIN", "GLOBAL", 2024, 3.3334],
            ],
            columns=["PARAMETER", "Scenario", "REGION", "YEAR", "Value"],
        )
        expected_penalty = pd.DataFrame(
            [
                ["EmissionPenalty", "MAIN", "GLOBAL", 2024, 4.4445],
                ["EmissionPenalty", "MAIN", "GLOBAL", 2025, 5.0],
            ],
            columns=["PARAMETER", "Scenario", "REGION", "YEAR", "Value"],
            index=[4, 5],
        )
        actual_main = namespace["overall_param_df_dict"]
        actual_additional = namespace["overall_param_df_dict_ndp"]
        assert_frame_equal(actual_main["ReserveMargin"], expected_reserve)
        assert_frame_equal(actual_main["EmissionPenalty"], expected_penalty)
        assert_frame_equal(actual_additional["ReserveMargin"], expected_reserve)
        assert_frame_equal(actual_additional["EmissionPenalty"], expected_penalty)
        self.assertIsNot(
            actual_main["ReserveMargin"], actual_additional["ReserveMargin"]
        )
        self.assertEqual(workbook.parse_calls, ["System Parameters"])
        self.assertEqual(
            output,
            "   Loaded system parameters: ['ReserveMargin', "
            "'EmissionPenalty']\n",
        )

    def test_absent_and_all_missing_sheets_preserve_diagnostics_and_state(self) -> None:
        sentinel = pd.DataFrame({"Value": [1.0]})
        no_sheet = _WorkbookFixture(["Capacities"], pd.DataFrame())
        namespace, output = _run_system_parameter_block(
            no_sheet, years=[2023], main={"Keep": sentinel}
        )
        self.assertIs(namespace["overall_param_df_dict"]["Keep"], sentinel)
        self.assertEqual(no_sheet.parse_calls, [])
        self.assertEqual(
            output,
            "   NOTE: No System Parameters sheet found — ReserveMargin not included.\n",
        )

        all_missing = _WorkbookFixture(
            ["System Parameters"],
            pd.DataFrame({"Parameter": ["ReserveMargin"], 2023: [np.nan]}),
        )
        namespace, output = _run_system_parameter_block(
            all_missing, years=[2023], main={"Keep": sentinel}
        )
        self.assertEqual(list(namespace["overall_param_df_dict"]), ["Keep"])
        self.assertEqual(
            output,
            "   WARNING: System Parameters sheet found but no valid rows.\n",
        )

    def test_bad_value_and_missing_year_fail_before_dictionary_update(self) -> None:
        cases = (
            (
                pd.DataFrame({"Parameter": ["ReserveMargin"], 2023: ["bad"]}),
                [2023],
                ValueError,
            ),
            (
                pd.DataFrame({"Parameter": ["ReserveMargin"], 2023: [1.0]}),
                [2023, 2024],
                KeyError,
            ),
        )
        for frame, years, error_type in cases:
            with self.subTest(error=error_type.__name__):
                sentinel = pd.DataFrame({"old": [1]})
                main = {"ReserveMargin": sentinel}
                workbook = _WorkbookFixture(["System Parameters"], frame)
                with self.assertRaises(error_type):
                    _run_system_parameter_block(
                        workbook, years=years, main=main
                    )
                self.assertIs(main["ReserveMargin"], sentinel)


class B1StructureCharacterizationTests(unittest.TestCase):
    def test_structure_padding_order_duplicates_and_dtypes_are_preserved(self) -> None:
        namespace, _recorder, _demand = _delivery_fixture()
        tables = _candidate_modules()["tables"]
        result = tables.build_structure_tables(
            namespace["time_range_vector"],
            namespace["All_Tech_list"],
            namespace["All_Fuel_list"],
            namespace["emissions_list"],
            namespace["other_setup_params"],
            namespace["params"],
        )

        structure = result.table
        structure_dict = result.values
        self.assertEqual(tuple(structure.columns), STRUCTURE_KEYS)
        self.assertEqual(tuple(structure_dict), STRUCTURE_KEYS)
        self.assertEqual(structure_dict["TECHNOLOGY"], ["TECH2", "TECH1", "TECH1"])
        self.assertEqual(structure_dict["YEAR"], [2023, 2024, ""])
        self.assertEqual(structure_dict["TIMESLICE"], ["L2", "L1", ""])
        self.assertEqual(structure_dict["REGION"], ["GLOBAL", "", ""])
        self.assertEqual(structure["YEAR"].dtype, object)
        self.assertEqual(structure["TECHNOLOGY"].dtype, object)


class B1OutputDeliveryCharacterizationTests(unittest.TestCase):
    def test_exact_workbook_csv_scenario_and_writer_order(self) -> None:
        namespace, recorder, demand = _delivery_fixture()
        original_main = {
            key: value.copy(deep=True)
            for key, value in namespace["overall_param_df_dict"].items()
            if not pd.isna(key)
        }

        _run_delivery(namespace, recorder)

        params = namespace["params"]
        scenario_root = "A1ROOT_MAIN"
        demand_path = os.path.join(
            "A1ROOT", scenario_root + "/Demand_COMPLETED.xlsx"
        )
        param_path = os.path.join(
            "A1ROOT", scenario_root + "/Param_COMPLETED.xlsx"
        )
        natural_path = os.path.join(
            "A1ROOT", scenario_root + "/Param_Natural_COMPLETED.xlsx"
        )
        projection_path = os.path.join(
            "A1ROOT", scenario_root + "/AR_COMPLETED.xlsx"
        )
        structure_path = params["Print_A2_Struct_List"]
        main_output = os.path.join("A2_MAIN_ROOT", "MAIN")

        expected: list[tuple[object, ...]] = [
            ("excel_open", demand_path, "xlsxwriter"),
            ("excel_write", demand_path, "Demand", False),
            ("excel_close", demand_path),
            ("excel_open", param_path, "xlsxwriter"),
            ("excel_write", param_path, "ZSheet", False),
            ("excel_write", param_path, "ASheet", False),
            ("excel_close", param_path),
            ("excel_open", natural_path, "xlsxwriter"),
            ("excel_write", natural_path, "NaturalB", False),
            ("excel_write", natural_path, "NaturalA", False),
            ("excel_close", natural_path),
            ("excel_open", projection_path, "xlsxwriter"),
            ("excel_write", projection_path, "Secondary", False),
            ("excel_write", projection_path, "Primary", False),
            ("excel_close", projection_path),
            ("excel_open", structure_path, "xlsxwriter"),
            ("excel_write", structure_path, "Lists", False),
            ("excel_close", structure_path),
            ("mkdir", main_output, True),
            ("csv_write", os.path.join(main_output, "ZParam.csv"), False, True),
            ("csv_write", os.path.join(main_output, "AParam.csv"), False, True),
        ]
        expected.extend(
            (
                "csv_write",
                os.path.join(main_output, f"{name}.csv"),
                False,
                True,
            )
            for name in sorted(STRUCTURE_KEYS)
        )
        for scenario in ("C", "A", "C"):
            output_root = os.path.join("A2_OTHER_ROOT", scenario)
            expected.append(("mkdir", output_root, True))
            expected.extend(
                (
                    "csv_write",
                    os.path.join(output_root, f"{name}.csv"),
                    False,
                    True,
                )
                for name in ("AParam", "ZParam")
            )
            expected.extend(
                (
                    "csv_write",
                    os.path.join(output_root, f"{name}.csv"),
                    False,
                    True,
                )
                for name in sorted(STRUCTURE_KEYS)
            )
        self.assertEqual(len(expected), 74)
        self.assertEqual(len(recorder.events), 74)
        self.assertEqual(recorder.events, expected)

        demand_written = next(
            frame
            for path, sheet, frame in recorder.excel_frames
            if path == demand_path and sheet == "Demand"
        )
        self.assertEqual(demand_written.loc[0, "2023"], 1.2346)
        self.assertEqual(demand_written.loc[0, "Other"], 2.3457)
        self.assertEqual(demand["2023"].dtype, np.dtype("float64"))
        self.assertEqual(demand.loc[0, "Other"], 2.34567)

        csv_by_path: dict[object, list[pd.DataFrame]] = {}
        for path, frame in recorder.csv_frames:
            csv_by_path.setdefault(path, []).append(frame)
        self.assertNotIn(os.path.join(main_output, "nan.csv"), csv_by_path)
        self.assertFalse(any("NDPOnly" in os.fspath(path) for path in csv_by_path))
        assert_frame_equal(
            csv_by_path[os.path.join(main_output, "ZParam.csv")][0],
            original_main["ZParam"],
        )
        scenario_a = csv_by_path[
            os.path.join("A2_OTHER_ROOT", "A", "AParam.csv")
        ][0]
        self.assertEqual(scenario_a["Scenario"].iloc[0], "A")
        self.assertEqual(scenario_a["Scenario"].iloc[1], "OTHER")
        self.assertTrue(pd.isna(scenario_a["Scenario"].iloc[2]))
        assert_frame_equal(
            namespace["overall_param_df_dict"]["AParam"],
            original_main["AParam"],
        )

    def test_excel_failure_propagates_before_close_or_later_writes(self) -> None:
        recorder = _DeliveryRecorder(fail_excel_sheet="Demand")
        namespace, recorder, _demand = _delivery_fixture(recorder=recorder)

        with self.assertRaisesRegex(OSError, "fixture Excel failure: Demand"):
            _run_delivery(namespace, recorder)

        demand_path = os.path.join(
            "A1ROOT", "A1ROOT_MAIN/Demand_COMPLETED.xlsx"
        )
        self.assertEqual(
            recorder.events,
            [
                ("excel_open", demand_path, "xlsxwriter"),
                ("excel_write", demand_path, "Demand", False),
            ],
        )

    def test_csv_failure_keeps_partial_writes_and_stops_extra_scenarios(self) -> None:
        recorder = _DeliveryRecorder(fail_csv_name="FUEL.csv")
        namespace, recorder, _demand = _delivery_fixture(recorder=recorder)

        with self.assertRaisesRegex(OSError, "fixture CSV failure: FUEL.csv"):
            _run_delivery(namespace, recorder)

        csv_paths = [
            event[1] for event in recorder.events if event[0] == "csv_write"
        ]
        main_output = os.path.join("A2_MAIN_ROOT", "MAIN")
        self.assertEqual(
            csv_paths,
            [
                os.path.join(main_output, "ZParam.csv"),
                os.path.join(main_output, "AParam.csv"),
                os.path.join(main_output, "DAILYTIMEBRACKET.csv"),
                os.path.join(main_output, "DAYTYPE.csv"),
                os.path.join(main_output, "EMISSION.csv"),
                os.path.join(main_output, "FUEL.csv"),
            ],
        )
        self.assertFalse(
            any(
                event[0] == "mkdir" and event[1] == os.path.join("A2_OTHER_ROOT", "C")
                for event in recorder.events
            )
        )
        self.assertFalse(any(event[0] in {"unlink", "rmtree"} for event in recorder.events))


class B1TransformProcessSafetyTests(unittest.TestCase):
    def test_compiler_source_has_no_b2_matrix_solver_or_process_boundary(self) -> None:
        sources = {"compiler.py": _source()}
        package_root = REPO_ROOT / "ostram" / "pipeline" / "compilation" / "transforms"
        for module_name in ("__init__", *TRANSFORM_MODULES):
            path = package_root / f"{module_name}.py"
            sources[path.name] = path.read_text(encoding="utf-8")

        for name, source in sources.items():
            with self.subTest(source=name):
                tree = ast.parse(source, filename=name)
                imported_roots: set[str] = set()
                for node in tree.body:
                    if isinstance(node, ast.Import):
                        imported_roots.update(
                            alias.name.split(".", 1)[0] for alias in node.names
                        )
                    elif isinstance(node, ast.ImportFrom) and node.module:
                        imported_roots.add(node.module.split(".", 1)[0])
                self.assertTrue(
                    imported_roots.isdisjoint(
                        {"subprocess", "multiprocessing", "asyncio"}
                    ),
                    imported_roots,
                )
                identifiers = {
                    node.id for node in ast.walk(tree) if isinstance(node, ast.Name)
                }
                self.assertTrue(
                    identifiers.isdisjoint(
                        {
                            "main_executer",
                            "create_matrix",
                            "invoke_solver_command",
                            "run_otoole_conversion",
                        }
                    ),
                    identifiers,
                )
                for forbidden in (
                    "python -m ostram run",
                    "glpsol",
                    "gurobi_cl",
                    "cplex",
                    "cbc",
                ):
                    self.assertNotIn(forbidden, source)


if __name__ == "__main__":
    unittest.main()
