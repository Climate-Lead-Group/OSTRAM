from __future__ import annotations

import ast
import io
import os
import sys
import types
import unittest
from collections.abc import Iterable
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

import numpy as np
import pandas as pd
from pandas.testing import assert_frame_equal


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
B1_COMPILER = REPO_ROOT / "t1_confection" / "B1_Compiler.py"

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


def _top_level_node_at(line_number: int) -> ast.stmt:
    matches = [node for node in _tree().body if node.lineno == line_number]
    if len(matches) != 1:
        raise AssertionError(
            f"expected one B1 top-level node at line {line_number}, got {matches!r}"
        )
    return matches[0]


def _top_level_nodes_between(first_line: int, last_line: int) -> list[ast.stmt]:
    return [
        node
        for node in _tree().body
        if node.lineno >= first_line and node.end_lineno <= last_line
    ]


def _execute_nodes(
    nodes: Iterable[ast.stmt], namespace: dict[str, object]
) -> dict[str, object]:
    module = ast.Module(body=list(nodes), type_ignores=[])
    ast.fix_missing_locations(module)
    exec(compile(module, str(B1_COMPILER), "exec"), namespace)
    return namespace


def _predecessor_normalizer():
    functions = [
        node
        for node in _tree().body
        if isinstance(node, ast.FunctionDef)
        and node.name == "normalize_year_like_columns"
    ]
    if len(functions) != 1:
        raise AssertionError(f"unexpected normalizer definitions: {functions!r}")
    namespace: dict[str, object] = {"np": np}
    _execute_nodes(functions, namespace)
    return namespace["normalize_year_like_columns"]


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
    namespace: dict[str, object] = {
        "pd": pd,
        "Parametrization": workbook,
        "normalize_year_like_columns": _predecessor_normalizer(),
        "time_range_vector": years,
        "other_setup_params": {"Main_Scenario": "MAIN", "Region": "GLOBAL"},
        "overall_param_df_dict": {} if main is None else main,
        "overall_param_df_dict_ndp": {} if additional is None else additional,
    }
    stdout = io.StringIO()
    with redirect_stdout(stdout):
        _execute_nodes([_top_level_node_at(1587)], namespace)
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
        "pd": types.SimpleNamespace(
            ExcelWriter=effect_recorder.excel_writer,
            DataFrame=pd.DataFrame,
            isna=pd.isna,
        ),
        "os": types.SimpleNamespace(
            path=os.path,
            makedirs=effect_recorder.makedirs,
        ),
        "params": params,
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
    def fake_to_excel(
        frame: pd.DataFrame,
        writer: _FakeWriter,
        *,
        sheet_name: str,
        index: bool,
    ) -> None:
        recorder.to_excel(frame, writer, sheet_name=sheet_name, index=index)

    def fake_to_csv(
        frame: pd.DataFrame,
        path: object,
        *,
        index: bool,
        header: bool,
    ) -> None:
        recorder.to_csv(frame, path, index=index, header=header)

    nodes = _top_level_nodes_between(1623, 1770)
    with (
        mock.patch.object(pd.DataFrame, "to_excel", new=fake_to_excel),
        mock.patch.object(pd.DataFrame, "to_csv", new=fake_to_csv),
    ):
        return _execute_nodes(nodes, namespace)


class B1ConfigurationAndSourcePathCharacterizationTests(unittest.TestCase):
    def test_config_is_cwd_relative_and_preserves_setup_mutation(self) -> None:
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

        namespace: dict[str, object] = {
            "open": OpenFixture(),
            "yaml": types.SimpleNamespace(safe_load=lambda _stream: loaded),
            "pd": pd,
        }
        _execute_nodes(_top_level_nodes_between(39, 58), namespace)

        self.assertEqual(opened, [("Config_MOMF_T1_A.yaml", "r", {})])
        self.assertEqual(namespace["time_range_vector"], [2023, 2024, 2025])
        self.assertEqual(
            loaded["xtra_scen"]["Timeslices"], ["L1", "L2", "L3"]
        )
        self.assertEqual(
            namespace["other_setup_params"],
            {
                "Region": "GLOBAL",
                "Timeslices": ["L1", "L2", "L3"],
                "Main_Scenario": "MAIN",
            },
        )

    def test_workbook_and_csv_source_paths_keep_exact_formulas_and_order(self) -> None:
        calls: list[tuple[str, object]] = []

        def excel_file(path: object) -> object:
            calls.append(("excel", path))
            return object()

        def read_csv(path: object) -> object:
            calls.append(("csv", path))
            return pd.DataFrame({"VALUE": ["CO2"]})

        params = {
            "A1_outputs": "A1ROOT",
            "xtra_scen": {"Main_Scenario": "SCENARIO"},
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
        namespace: dict[str, object] = {
            "pd": types.SimpleNamespace(ExcelFile=excel_file, read_csv=read_csv),
            "os": os,
            "params": params,
        }
        selected_lines = (64, 65, 342, 363, 533, 547, 1189, 1196, 1347)
        _execute_nodes(
            [_top_level_node_at(line) for line in selected_lines], namespace
        )

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
                ("csv", os.path.join("OG_csvs_inputs", "EMISSION.csv")),
                ("excel", "EXTRA/Storage.xlsx"),
            ],
        )


class B1YearColumnNormalizationCharacterizationTests(unittest.TestCase):
    def test_headers_normalize_without_changing_values_index_or_dtypes(self) -> None:
        normalize = _predecessor_normalizer()
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
        normalize = _predecessor_normalizer()
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
        nodes = [
            node
            for node in _top_level_nodes_between(1661, 1706)
            if not 1690 <= node.lineno <= 1692
        ]
        _execute_nodes(nodes, namespace)

        structure = namespace["df_structure"]
        structure_dict = namespace["structure_dict"]
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
        tree = _tree()
        imported_roots = {
            alias.name.split(".", 1)[0]
            for node in tree.body
            if isinstance(node, (ast.Import, ast.ImportFrom))
            for alias in node.names
        }
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
        source = _source()
        for forbidden in (
            "B2_Executing_OG_Model.py",
            "glpsol",
            "gurobi_cl",
            "cplex",
            "cbc",
        ):
            self.assertNotIn(forbidden, source)


if __name__ == "__main__":
    unittest.main()
