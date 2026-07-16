from __future__ import annotations

import csv
import importlib.util
import io
import json
import os
import runpy
import sys
import tempfile
import unittest
from contextlib import contextmanager, redirect_stdout
from pathlib import Path
from unittest import mock


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
T1_ROOT = REPO_ROOT / "t1_confection"
ANALYSIS_ROOT = REPO_ROOT / "tools" / "analysis"
VISUALIZATION_ROOT = ANALYSIS_ROOT / "visualization"


def _has_modules(*names: str) -> bool:
    return all(importlib.util.find_spec(name) is not None for name in names)


def _load_module(path: Path, label: str):
    module_name = f"_ostram_analysis_test_{label}"
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"could not load {path}")
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


@unittest.skipUnless(_has_modules("pandas", "numpy"), "pandas and NumPy are optional")
class ConcatenatorCompatibilityTests(unittest.TestCase):
    def _fixture(self, root: Path) -> Path:
        search = root / "Executables"
        scenario = search / "Scenario_A_0"
        scenario.mkdir(parents=True)
        (scenario / "model_Input.csv").write_text(
            "REGION,YEAR,TECHNOLOGY,TotalAnnualMinCapacityInvestment\n"
            "R1,2024,TECH_A,2\n"
            "R1,2023,TECH_A,1\n",
            encoding="utf-8",
        )
        (scenario / "model_Pre_processed_output.csv").write_text(
            "REGION,YEAR,TECHNOLOGY,TotalCapacityAnnual\n"
            "R1,2024,TECH_A,7\n",
            encoding="utf-8",
        )
        return search

    def _run(self, entrypoint: Path, search: Path, output_root: Path) -> str:
        output_root.mkdir()
        argv = [
            str(entrypoint),
            "--search-dir",
            str(search),
            "--output",
            str(output_root / "combined.csv"),
            "--inputs-file",
            str(output_root / "inputs.csv"),
            "--outputs-file",
            str(output_root / "outputs.csv"),
        ]
        stream = io.StringIO()
        with (
            mock.patch.object(sys, "argv", argv),
            _working_directory(output_root),
            redirect_stdout(stream),
        ):
            runpy.run_path(str(entrypoint), run_name="__main__")
        return stream.getvalue()

    def test_old_and_canonical_entrypoints_write_identical_stacked_outputs(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp)
            search = self._fixture(root)
            canonical_root = root / "canonical"
            wrapper_root = root / "wrapper"

            canonical_stdout = self._run(
                ANALYSIS_ROOT / "concat_all_scenarios.py", search, canonical_root
            )
            wrapper_stdout = self._run(
                T1_ROOT / "concat_all_scenarios_2.py", search, wrapper_root
            )

            for name in ("combined.csv", "inputs.csv", "outputs.csv"):
                self.assertEqual(
                    (canonical_root / name).read_bytes(),
                    (wrapper_root / name).read_bytes(),
                )
            self.assertEqual(
                canonical_stdout.replace(str(canonical_root), "<OUTPUT>"),
                wrapper_stdout.replace(str(wrapper_root), "<OUTPUT>"),
            )

            with (canonical_root / "combined.csv").open(newline="", encoding="utf-8") as handle:
                rows = list(csv.DictReader(handle))
            self.assertEqual(len(rows), 3)
            investments = [
                row
                for row in rows
                if row["TotalAnnualMinCapacityInvestment"]
            ]
            self.assertEqual(
                [row["AccumulatedTotalAnnualMinCapacityInvestment"] for row in investments],
                ["1.0", "3.0"],
            )


@unittest.skipUnless(_has_modules("pandas", "numpy"), "pandas and NumPy are optional")
class SensitivityAnalysisTests(unittest.TestCase):
    def test_metrics_and_old_path_wrapper_match_from_an_unrelated_cwd(self) -> None:
        import numpy as np
        import pandas as pd

        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp, _working_directory(Path(temp)):
            canonical = _load_module(
                ANALYSIS_ROOT / "analyse_sensitivity.py", "sensitivity"
            )
            wrapper = runpy.run_path(
                str(T1_ROOT / "analyse_sensitivity.py"),
                run_name="_ostram_sensitivity_wrapper_test",
            )

        self.assertEqual(canonical.T1_ROOT, T1_ROOT)
        self.assertEqual(wrapper["T1_ROOT"], T1_ROOT)

        columns = canonical.USECOLS
        defaults = {column: np.nan for column in columns}

        def row(**values):
            return {
                **defaults,
                "Scenario": "B_Opt_Clipped",
                "REGION": "R1",
                **values,
            }

        fixture = pd.DataFrame(
            [
                row(
                    YEAR=2050,
                    TECHNOLOGY="PWRCOABGDXX",
                    EMISSION="CO2",
                    TotalDiscountedCost=100,
                    DiscountedCapitalInvestment=10,
                    DiscountedCapitalInvestmentStorage=2,
                    AnnualVariableOperatingCost=3,
                    TotalCapacityAnnual=5,
                    AnnualEmissions=7,
                    ProductionByTechnologyAnnual=36,
                ),
                row(
                    YEAR=2050,
                    TECHNOLOGY="PWRSPVBGDXX",
                    TotalDiscountedCost=50,
                    DiscountedCapitalInvestment=20,
                    AnnualVariableOperatingCost=1,
                    TotalCapacityAnnual=4,
                    ProductionByTechnologyAnnual=18,
                ),
                row(
                    YEAR=2050,
                    TECHNOLOGY="TRNBGDXXINDEA",
                    ProductionByTechnologyAnnual=36,
                ),
                row(
                    YEAR=2050,
                    TECHNOLOGY="PWRBCKBGDXX",
                    ProductionByTechnologyAnnual=3.6,
                ),
            ],
            columns=columns,
        )

        expected = canonical.metrics_for(fixture, 20.0)
        self.assertEqual(wrapper["metrics_for"](fixture, 20.0), expected)
        self.assertEqual(expected["System cost (NPV) [M USD]"], 150.0)
        self.assertEqual(expected["BGD domestic gen 2050 [TWh]"], 15.0)
        self.assertEqual(expected["BGD net imports 2050 [TWh]"], 5.0)
        self.assertEqual(expected["BGD domestic share 2050 [%]"], 75.0)
        self.assertEqual(expected["Cross-border trade 2050 [TWh]"], 10.0)
        self.assertEqual(expected["Backstop generation [TWh]"], 1.0)

        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp)
            combined = root / "combined.csv"
            ceilings = root / "ceilings.json"
            baseline = root / "baseline.json"
            output_csv = root / "comparison.csv"
            output_txt = root / "report.txt"
            fixture.to_csv(combined, index=False)
            ceilings.write_text(
                json.dumps({"ceilings_gw": {}}), encoding="utf-8"
            )
            baseline.write_text(
                json.dumps({"bgd_demand_PJ": {"2050": 72.0}}),
                encoding="utf-8",
            )
            canonical.CEIL_JSON = ceilings
            canonical.BASE_JSON = baseline
            canonical.OUT_CSV = output_csv
            canonical.OUT_TXT = output_txt
            argv = ["analyse_sensitivity.py", "--combined", str(combined)]
            with (
                mock.patch.object(sys, "argv", argv),
                redirect_stdout(io.StringIO()),
            ):
                self.assertEqual(canonical.main(), 0)

            with output_csv.open(newline="", encoding="utf-8") as handle:
                comparison = {
                    row["metric"]: float(row["B_Opt_Clipped"])
                    for row in csv.DictReader(handle)
                }
            report = output_txt.read_text(encoding="utf-8")
            self.assertEqual(comparison["System cost (NPV) [M USD]"], 150.0)
            self.assertEqual(comparison["BGD net imports 2050 [TWh]"], 5.0)
            self.assertIn("OSTRAM PHASE-B SENSITIVITY ANALYSIS", report)
            self.assertIn("B_Opt_Clipped", report)
            self.assertEqual(
                {path.name for path in root.iterdir()},
                {
                    "combined.csv",
                    "ceilings.json",
                    "baseline.json",
                    "comparison.csv",
                    "report.txt",
                },
            )

    def test_ws4_collection_preserves_metrics_and_marks_missing_scenarios(self) -> None:
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp, _working_directory(Path(temp)):
            module = _load_module(
                T1_ROOT / "sensitivity_expansion" / "analyse_ws4_vs_phaseB.py",
                "ws4",
            )
            outputs = (
                Path(temp)
                / "Executables"
                / "B_Opt_Clipped_0"
                / "Outputs"
            )
            outputs.mkdir(parents=True)
            (outputs / "TotalDiscountedCost.csv").write_text(
                "VALUE\n100\n", encoding="utf-8"
            )
            (outputs / "AnnualEmissions.csv").write_text(
                "YEAR,VALUE\n2050,7\n", encoding="utf-8"
            )
            (outputs / "TotalCapacityAnnual.csv").write_text(
                "YEAR,TECHNOLOGY,VALUE\n2050,PWRCOABGDXX,5\n"
                "2050,PWRSPVBGDXX,4\n",
                encoding="utf-8",
            )
            (outputs / "ProductionByTechnologyAnnual.csv").write_text(
                "YEAR,TECHNOLOGY,VALUE\n2050,PWRCOABGDXX,36\n"
                "2050,PWRBCKBGDXX,0\n",
                encoding="utf-8",
            )
            module.EXEC = Path(temp) / "Executables"
            rows = module.collect()

        self.assertEqual(module.REPO, T1_ROOT)
        self.assertEqual(rows["B_Opt_Clipped"]["syscost"], 100.0)
        self.assertEqual(rows["B_Opt_Clipped"]["co2"], 7.0)
        self.assertEqual(rows["B_Opt_Clipped"]["coal"], 5.0)
        self.assertEqual(rows["B_Opt_Clipped"]["bgd"], 36.0)
        self.assertIsNone(rows["A_Calibrated_BAU"])


@unittest.skipUnless(
    _has_modules("pandas", "numpy", "matplotlib"),
    "plot fixture requires optional matplotlib",
)
class FigureReproductionTests(unittest.TestCase):
    def test_a1_a6_fixture_outputs_are_confined_and_have_expected_values(self) -> None:
        import numpy as np
        import pandas as pd

        module = _load_module(ANALYSIS_ROOT / "reproduce_A1_A6.py", "a1_a6")
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp)
            csv_path = root / "combined.csv"
            out_dir = root / "figures"
            records = []
            for scenario, scale in (("A_Calibrated_BAU", 1.0), ("B_Optimised_VRE", 2.0)):
                for year in (2023, 2024):
                    records.extend(
                        [
                            {
                                "Scenario": scenario,
                                "REGION": "R1",
                                "YEAR": year,
                                "TECHNOLOGY": "PWRCOABGDXX",
                                "FUEL": "ELCBGDXX02",
                                "TotalDiscountedCost": 100 * scale,
                                "TotalCapacityAnnual": 5 * scale,
                                "ProductionByTechnologyAnnual": 30 * scale,
                            },
                            {
                                "Scenario": scenario,
                                "REGION": "R1",
                                "YEAR": year,
                                "TECHNOLOGY": "PWRSPVBGDXX",
                                "FUEL": "ELCBGDXX02",
                                "TotalDiscountedCost": np.nan,
                                "TotalCapacityAnnual": 4 * scale,
                                "ProductionByTechnologyAnnual": 10 * scale,
                            },
                            {
                                "Scenario": scenario,
                                "REGION": "R1",
                                "YEAR": year,
                                "TECHNOLOGY": "PWRPETBGDXX",
                                "FUEL": "ELCBGDXX02",
                                "TotalDiscountedCost": np.nan,
                                "TotalCapacityAnnual": 3 * scale,
                                "ProductionByTechnologyAnnual": np.nan,
                            },
                            {
                                "Scenario": scenario,
                                "REGION": "R1",
                                "YEAR": year,
                                "TECHNOLOGY": "PWRSDSBGDXX",
                                "FUEL": "ELCBGDXX02",
                                "TotalDiscountedCost": np.nan,
                                "TotalCapacityAnnual": 2 * scale,
                                "ProductionByTechnologyAnnual": np.nan,
                            },
                            {
                                "Scenario": scenario,
                                "REGION": "R1",
                                "YEAR": year,
                                "TECHNOLOGY": "PWRLDSBGDXX",
                                "FUEL": "ELCBGDXX02",
                                "TotalDiscountedCost": np.nan,
                                "TotalCapacityAnnual": 1 * scale,
                                "ProductionByTechnologyAnnual": np.nan,
                            },
                        ]
                    )
            pd.DataFrame(records).to_csv(csv_path, index=False)
            module.CSV_PATH = str(csv_path)
            module.OUT_DIR = str(out_dir)
            with redirect_stdout(io.StringIO()):
                module.main()

            self.assertEqual(
                {path.name for path in out_dir.iterdir()},
                {
                    "A1_annual_discounted_cost.png",
                    "A2_cumulative_discounted_cost.png",
                    "A3_coal_share.png",
                    "A4_vre_share.png",
                    "A5_petroleum_oil_capacity.png",
                    "A6_storage_capacity.png",
                    "A1_A6_series_audit.csv",
                },
            )
            audit = pd.read_csv(out_dir / "A1_A6_series_audit.csv")
            selected = audit[
                (audit["metric"] == "A1_annual_cost_USD_B")
                & (audit["Scenario"] == "A-CalBAU")
                & (audit["YEAR"] == 2023)
            ]
            self.assertEqual(selected["value"].tolist(), [0.1])
            self.assertEqual(
                {path.name for path in root.iterdir()},
                {"combined.csv", "figures"},
            )


@unittest.skipUnless(_has_modules("pandas"), "pandas is optional")
class VisualizationUtilityTests(unittest.TestCase):
    def test_aggregated_dashboard_writes_a_deterministic_shape_to_fixture_dir(self) -> None:
        import pandas as pd

        module = _load_module(
            VISUALIZATION_ROOT / "Z_AUX_generate_interactive_dashboards_aggregated.py",
            "aggregated_dashboard",
        )
        frame = pd.DataFrame(
            [
                {
                    "Scenario": "Case",
                    "YEAR": 2030,
                    "TECHNOLOGY": "PWRSPVBGDXX",
                    "ProductionByTechnology": 12.0,
                    "TotalTechnologyAnnualActivityLowerLimit": 3.0,
                }
            ]
        )
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp, _working_directory(Path(temp)):
            with redirect_stdout(io.StringIO()):
                output = module.generate_interactive_dashboard(frame, "fixture.csv")
            output_path = Path(temp) / output
            html = output_path.read_text(encoding="utf-8")
            self.assertEqual({path.name for path in Path(temp).iterdir()}, {output})
        self.assertTrue(output.startswith("Dashboard_Interactive_Aggregated_fixture_"))
        self.assertIn("PWRSPVBGDXX", html)
        self.assertIn("Case", html)

    @unittest.skipUnless(_has_modules("yaml", "openpyxl"), "YAML and openpyxl are optional")
    def test_res_helpers_keep_t1_inputs_and_write_only_the_requested_html(self) -> None:
        module = _load_module(
            VISUALIZATION_ROOT / "Z_AUX_generate_RES_diagram.py", "res_diagram"
        )
        self.assertEqual(module.SCRIPT_DIR, T1_ROOT)
        links = [
            {
                "fuel_in": "MINCOABGD",
                "fuel_in_name": "Coal",
                "tech": "PWRCOABGDXX",
                "tech_name": "Coal power",
                "fuel_out": "ELCBGDXX01",
                "fuel_out_name": "Electricity",
                "mode": 1,
            }
        ]
        self.assertEqual(module.discover_regions(links), ["BGDXX"])
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            output = Path(temp) / "RES_Diagram.html"
            with redirect_stdout(io.StringIO()):
                module.generate_html(links, ["BGDXX"], output)
            self.assertEqual({path.name for path in Path(temp).iterdir()}, {output.name})
            self.assertIn("PWRCOABGDXX", output.read_text(encoding="utf-8"))

    def test_transmission_helpers_preserve_line_parsing_and_fixture_discovery(self) -> None:
        module = _load_module(
            VISUALIZATION_ROOT / "Z_AUX_generate_transmission_maps.py",
            "transmission_maps",
        )
        self.assertEqual(module.T1_ROOT, T1_ROOT)
        self.assertEqual(
            module.extract_from_to("TRNBGDXXINDEA"), ("BGDXX", "INDEA")
        )
        self.assertEqual(
            module.classify_flow_direction("ELCINDEA04", "BGDXX", "INDEA"),
            "a_to_b",
        )
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp)
            expected = root / "fixture_Combined_Inputs_Outputs.csv"
            expected.write_text("Scenario\nCase\n", encoding="utf-8")
            self.assertEqual(module.find_combined_csv(root), expected)

    @unittest.skipUnless(
        _has_modules("numpy", "plotly", "yaml", "openpyxl"),
        "visualization entrypoint smoke tests require optional dependencies",
    )
    def test_visualization_main_functions_are_safe_with_fixture_boundaries(self) -> None:
        import pandas as pd

        aggregated = _load_module(
            VISUALIZATION_ROOT / "Z_AUX_generate_interactive_dashboards_aggregated.py",
            "aggregated_main",
        )
        with (
            mock.patch.object(aggregated, "select_files_interactive", return_value=[]),
            redirect_stdout(io.StringIO()),
        ):
            self.assertIsNone(aggregated.main())

        res = _load_module(
            VISUALIZATION_ROOT / "Z_AUX_generate_RES_diagram.py", "res_main"
        )
        transmission = _load_module(
            VISUALIZATION_ROOT / "Z_AUX_generate_transmission_maps.py",
            "transmission_main",
        )
        with tempfile.TemporaryDirectory(dir=TEST_ROOT) as temp:
            root = Path(temp)
            res.SCRIPT_DIR = root
            with redirect_stdout(io.StringIO()):
                self.assertIsNone(res.main())
            self.assertEqual(list(root.iterdir()), [])

            transmission.T1_ROOT = root
            with self.assertRaisesRegex(FileNotFoundError, "Data file not found"):
                transmission.main()
            self.assertEqual(list(root.iterdir()), [])

            interconnections = _load_module(
                VISUALIZATION_ROOT / "Z_AUX_interconnections_dashboard.py",
                "interconnections_main",
            )
            interconnections.OUT_DIR = root / "Figures"
            interconnections.OUT_PATH = (
                interconnections.OUT_DIR / "interconnections_dashboard.html"
            )
            empty = pd.DataFrame()
            with (
                mock.patch.object(
                    interconnections,
                    "load_data",
                    return_value=(empty, empty, empty),
                ),
                mock.patch.object(
                    interconnections,
                    "build_annual",
                    return_value=pd.DataFrame([{"value": 1}]),
                ),
                mock.patch.object(
                    interconnections,
                    "build_seasonal",
                    return_value=empty,
                ),
                mock.patch.object(
                    interconnections,
                    "build_html",
                    return_value="<html>fixture</html>",
                ),
                redirect_stdout(io.StringIO()),
            ):
                interconnections.main()
            self.assertEqual(
                interconnections.OUT_PATH.read_text(encoding="utf-8"),
                "<html>fixture</html>",
            )

    @unittest.skipUnless(_has_modules("numpy", "plotly"), "NumPy and Plotly are optional")
    def test_interconnection_helpers_keep_t1_paths_and_annual_values(self) -> None:
        import pandas as pd

        module = _load_module(
            VISUALIZATION_ROOT / "Z_AUX_interconnections_dashboard.py",
            "interconnections",
        )
        self.assertEqual(module.HERE, T1_ROOT)
        self.assertEqual(module.parse_line("TRNBGDXXINDEA"), ("BGDXX", "INDEA"))
        self.assertIsNone(module.parse_line("TRNNLIINDEA"))
        fixture = pd.DataFrame(
            [
                {
                    "Scenario": "BAU",
                    "YEAR": 2030,
                    "TECHNOLOGY": "TRNBGDXXINDEA",
                    "ORIGIN": "BGDXX",
                    "DEST": "INDEA",
                    "PAIR": "BGDXX -> INDEA",
                    "ProductionByTechnologyAnnual": 10.0,
                    "TotalCapacityAnnual": 2.0,
                },
                {
                    "Scenario": "BAU",
                    "YEAR": 2030,
                    "TECHNOLOGY": "TRNBGDXXINDEA",
                    "ORIGIN": "BGDXX",
                    "DEST": "INDEA",
                    "PAIR": "BGDXX -> INDEA",
                    "ProductionByTechnologyAnnual": 5.0,
                    "TotalCapacityAnnual": 1.0,
                },
            ]
        )
        annual = module.build_annual(fixture)
        self.assertEqual(annual["energy_pj"].tolist(), [15.0])
        self.assertEqual(annual["capacity_gw"].tolist(), [3.0])


if __name__ == "__main__":
    unittest.main()
