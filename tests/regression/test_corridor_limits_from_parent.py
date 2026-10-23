"""Corridor scenario limits are computed in every run from the solved parent.

``B_Opt_TxCap150`` and ``B_Opt_TradeCap15`` read the solved ``B_Optimised_VRE``,
and ``A_Calibrated_BAU_Clipped_TxFreeze2026`` reads the solved
``A_Calibrated_BAU_Clipped`` of the same run. A missing parent result stops the run.
"""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import yaml
from openpyxl import Workbook

from ostram.pipeline.scenarios import scenario_edits as se
from ostram.pipeline.scenarios.registry import load_registry


REPO_ROOT = Path(__file__).resolve().parents[2]
SCENARIOS = REPO_ROOT / "config" / "scenarios"
YEARS = list(range(2023, 2051))
STUDY = list(range(2027, 2051))


def _sheet(rows):
    """Return a Secondary Techs-like worksheet holding ``rows`` (tech, param, value by year)."""
    wb = Workbook()
    ws = wb.active
    ws.append(["Tech", "Parameter", "Projection.Mode", *YEARS])
    for tech, param, values in rows:
        ws.append([tech, param, "User defined", *[values(y) for y in YEARS]])
    return ws


def _parent(folder: Path, tables: dict) -> Path:
    """Write ``Outputs/<name>.csv`` tables of a synthetic solved parent and return its folder."""
    outputs = folder / "Outputs"
    outputs.mkdir(parents=True)
    for name, text in tables.items():
        (outputs / f"{name}.csv").write_text(text, encoding="utf-8")
    return folder


def _apply(ws, edit, results):
    columns, year_columns = se.scan_columns(ws)
    log = {"scenario": results.scenario, "cells": [], "rows_created": [], "skipped": []}
    se.apply_edit(ws, columns, year_columns, edit, log, results=results)
    return columns, year_columns, log


def _row(ws, columns, year_columns, tech, param):
    r = se.find_row(ws, columns, tech, param)
    return {y: ws.cell(r, c).value for y, c in year_columns.items()}


class CorridorLimitTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_txcap150_floor_is_the_largest_base_year_capacity_of_the_parent(self) -> None:
        parent = _parent(self.dir / "B", {"TotalCapacityAnnual": (
            "REGION,TECHNOLOGY,YEAR,VALUE\n"
            "GLOBAL,TRNAAAXXBBBXX,2023,1.0\nGLOBAL,TRNAAAXXBBBXX,2025,3.0\n"
            "GLOBAL,TRNAAAXXBBBXX,2026,2.5\nGLOBAL,TRNAAAXXBBBXX,2030,9.0\n")})
        results = se.ParentResults("TxCap", {"B_Optimised_VRE": parent})
        ws = _sheet([
            ("TRNAAAXXBBBXX", "ResidualCapacity", lambda y: 2.4),
            ("TRNAAAXXBBBXX", "TotalAnnualMaxCapacity", lambda y: 9999),
            ("TRNAAAXXBBBXX", "TotalAnnualMinCapacityInvestment", lambda y: 5.0 if y == 2030 else 0.5),
        ])
        edit = {"sheet": "Secondary Techs", "tech": "TRNAAAXXBBBXX", "op": "set_to_residual_factor_floor",
                "factor": 1.5, "residual_year": 2023, "base_window_result": "B_Optimised_VRE",
                "base_window_years": [2023, 2026], "years": STUDY}
        cols, ycols, _ = _apply(ws, {**edit, "param": "TotalAnnualMaxCapacity"}, results)
        got = _row(ws, cols, ycols, "TRNAAAXXBBBXX", "TotalAnnualMaxCapacity")
        self.assertEqual(got[2026], 9999)
        self.assertEqual({got[y] for y in STUDY}, {3.6})  # max(1.5 x 2.4, 3.0)
        cols, ycols, _ = _apply(ws, {**edit, "param": "TotalAnnualMinCapacityInvestment", "clamp": True}, results)
        got = _row(ws, cols, ycols, "TRNAAAXXBBBXX", "TotalAnnualMinCapacityInvestment")
        # Headroom above the parent's base-year capacity: 3.6 - 3.0 = 0.6, used in order.
        self.assertEqual(got[2027], 0.5)
        self.assertAlmostEqual(got[2028], 0.1, places=12)
        self.assertEqual({got[y] for y in range(2029, 2051)}, {0.0})
        self.assertEqual(got[2024], 0.5)   # base years unchanged

    def test_txcap150_floor_from_the_parent_is_not_rounded(self) -> None:
        parent = _parent(self.dir / "B", {"TotalCapacityAnnual": (
            "REGION,TECHNOLOGY,YEAR,VALUE\nGLOBAL,TRNAAAXXBBBXX,2025,0.0018384966834968\n")})
        results = se.ParentResults("TxCap", {"B_Optimised_VRE": parent})
        ws = _sheet([("TRNAAAXXBBBXX", "ResidualCapacity", lambda y: 0.0),
                     ("TRNAAAXXBBBXX", "TotalAnnualMaxCapacity", lambda y: 9999)])
        cols, ycols, _ = _apply(ws, {
            "sheet": "Secondary Techs", "tech": "TRNAAAXXBBBXX", "param": "TotalAnnualMaxCapacity",
            "op": "set_to_residual_factor_floor", "factor": 1.5, "residual_year": 2023,
            "base_window_result": "B_Optimised_VRE", "base_window_years": [2023, 2026],
            "years": STUDY}, results)
        got = _row(ws, cols, ycols, "TRNAAAXXBBBXX", "TotalAnnualMaxCapacity")
        self.assertEqual({got[y] for y in STUDY}, {0.0018384966834968})

    def test_txfreeze_reads_end_2026_capacity_and_base_year_new_capacity(self) -> None:
        parent = _parent(self.dir / "AC", {
            "TotalCapacityAnnual": "REGION,TECHNOLOGY,YEAR,VALUE\nGLOBAL,TRNAAAXXBBBXX,2026,3.85\n",
            "NewCapacity": "REGION,TECHNOLOGY,YEAR,VALUE\nGLOBAL,TRNAAAXXBBBXX,2026,1.35\n",
        })
        results = se.ParentResults("TxFreeze", {"A_Calibrated_BAU_Clipped": parent})
        ws = _sheet([("TRNAAAXXBBBXX", "TotalAnnualMaxCapacity", lambda y: 9999),
                     ("TRNCCCXXDDDXX", "TotalAnnualMinCapacityInvestment", lambda y: 0.0)])
        cols, ycols, _ = _apply(ws, {
            "sheet": "Secondary Techs", "tech": "TRNAAAXXBBBXX", "param": "TotalAnnualMaxCapacity",
            "op": "from_result", "result": "A_Calibrated_BAU_Clipped", "variable": "TotalCapacityAnnual",
            "result_year": 2026, "years": STUDY, "create_if_absent": True}, results)
        got = _row(ws, cols, ycols, "TRNAAAXXBBBXX", "TotalAnnualMaxCapacity")
        self.assertEqual(got[2026], 9999)
        self.assertEqual({got[y] for y in STUDY}, {3.85})
        cols, ycols, log = _apply(ws, {
            "sheet": "Secondary Techs", "tech": "TRNAAAXXBBBXX", "param": "TotalAnnualMinCapacityInvestment",
            "op": "from_result", "result": "A_Calibrated_BAU_Clipped", "variable": "NewCapacity",
            "years": [2023, 2024, 2025, 2026], "create_if_absent": True}, results)
        got = _row(ws, cols, ycols, "TRNAAAXXBBBXX", "TotalAnnualMinCapacityInvestment")
        self.assertEqual([got[y] for y in (2023, 2024, 2025, 2026)], [0.0, 0.0, 0.0, 1.35])
        self.assertIsNone(got[2027])
        self.assertEqual(len(log["rows_created"]), 1)

    def test_from_result_writes_the_parent_value_exactly(self) -> None:
        parent = _parent(self.dir / "AC", {
            "NewCapacity": "REGION,TECHNOLOGY,YEAR,VALUE\nGLOBAL,TRNAAAXXBBBXX,2026,0.76\n"})
        results = se.ParentResults("TxFreeze", {"A_Calibrated_BAU_Clipped": parent})
        ws = _sheet([("TRNAAAXXBBBXX", "TotalAnnualMaxCapacityInvestment",
                      lambda y: 0.7599999999999999 if y == 2026 else 0.0)])
        cols, ycols, _ = _apply(ws, {
            "sheet": "Secondary Techs", "tech": "TRNAAAXXBBBXX", "param": "TotalAnnualMaxCapacityInvestment",
            "op": "from_result", "result": "A_Calibrated_BAU_Clipped", "variable": "NewCapacity",
            "years": [2023, 2024, 2025, 2026]}, results)
        self.assertEqual(_row(ws, cols, ycols, "TRNAAAXXBBBXX", "TotalAnnualMaxCapacityInvestment")[2026], 0.76)

    def test_tradecap15_splits_15_percent_of_demand_by_import_share(self) -> None:
        parent = _parent(self.dir / "B", {
            "Demand": ("REGION,TIMESLICE,FUEL,YEAR,VALUE\n"
                       "GLOBAL,S1,ELCBGDXX03,2027,60\nGLOBAL,S2,ELCBGDXX03,2027,40\n"
                       "GLOBAL,S1,ELCBGDXX03,2028,100\nGLOBAL,S1,ELCINDEA03,2027,999\n"),
            "ProductionByTechnologyAnnual": (
                "REGION,TECHNOLOGY,FUEL,YEAR,VALUE\n"
                "GLOBAL,TRNBGDXXINDEA,ELCBGDXX04,2027,30\n"
                "GLOBAL,TRNBGDXXINDEA,ELCINDEA04,2027,50\n"   # export, not counted
                "GLOBAL,TRNNPLXXBGDXX,ELCBGDXX04,2027,10\n"),
        })
        results = se.ParentResults("TradeCap", {"B_Optimised_VRE": parent})
        techs = ["TRNBGDXXINDEA", "TRNNPLXXBGDXX"]
        ws = _sheet([(t, "TotalTechnologyAnnualActivityUpperLimit", lambda y: None) for t in techs])
        cols, ycols, _ = _apply(ws, {
            "sheet": "Secondary Techs", "techs": techs, "param": "TotalTechnologyAnnualActivityUpperLimit",
            "op": "import_share_cap", "result": "B_Optimised_VRE", "share": 0.15,
            "demand_fuel": "ELCBGDXX03", "importer": "BGDXX", "years": [2027, 2028]}, results)
        a = _row(ws, cols, ycols, "TRNBGDXXINDEA", "TotalTechnologyAnnualActivityUpperLimit")
        b = _row(ws, cols, ycols, "TRNNPLXXBGDXX", "TotalTechnologyAnnualActivityUpperLimit")
        self.assertAlmostEqual(a[2027], 0.15 * 100 * 30 / 40, places=12)
        self.assertAlmostEqual(b[2027], 0.15 * 100 * 10 / 40, places=12)
        # 2028: the parent imports nothing, so every limit is 0.
        self.assertEqual((a[2028], b[2028]), (0.0, 0.0))
        self.assertIsNone(a[2029])

    def test_missing_parent_stops_naming_scenario_and_parent(self) -> None:
        registry = load_registry()
        with self.assertRaises(FileNotFoundError) as ctx:
            registry.derived_result_dependencies(
                ["B_Opt_TxCap150"], execution_workspace=self.dir / "none", environment={})
        self.assertIn("B_Opt_TxCap150", str(ctx.exception))
        self.assertIn("B_Optimised_VRE", str(ctx.exception))
        results = se.ParentResults("B_Opt_TradeCap15", {"B_Optimised_VRE": self.dir / "absent"})
        with self.assertRaises(se.ParentResultError) as ctx:
            results.demand("B_Optimised_VRE", "ELCBGDXX03", 2027)
        self.assertIn("B_Opt_TradeCap15", str(ctx.exception))
        self.assertIn("B_Optimised_VRE", str(ctx.exception))
        empty = _parent(self.dir / "empty", {"Demand": "REGION,TIMESLICE,FUEL,YEAR,VALUE\n"})
        with self.assertRaises(se.ParentResultError):
            se.ParentResults("X", {"B_Optimised_VRE": empty}).demand("B_Optimised_VRE", "ELCBGDXX03", 2027)
        with self.assertRaises(se.ParentResultError):
            se.ParentResults("X").value("B_Optimised_VRE", "NewCapacity", "T", 2023)

    def test_scenario_files_read_their_parents_and_type_no_earlier_run_numbers(self) -> None:
        registry = load_registry()
        expected = {
            "B_Opt_TxCap150": "B_Optimised_VRE",
            "B_Opt_TradeCap15": "B_Optimised_VRE",
            "A_Calibrated_BAU_Clipped_TxFreeze2026": "A_Calibrated_BAU_Clipped",
        }
        for name, parent in expected.items():
            with self.subTest(scenario=name):
                deps = registry.derived_by_name[name].dependencies
                self.assertEqual([d.scenario for d in deps], [parent])
                config = yaml.safe_load((SCENARIOS / name / "scenario.yaml").read_text(encoding="utf-8"))
                self.assertNotIn("earlier model run", config["description"])
                for edit in config["edits"]:
                    self.assertNotIn("base_window_floor", edit)
                    if edit["sheet"] == "Secondary Techs" and edit["param"] != "TotalTechnologyAnnualActivityUpperLimit":
                        if "values" in edit:
                            self.assertEqual(set(edit["values"].values()), {0})
        trade = yaml.safe_load((SCENARIOS / "B_Opt_TradeCap15" / "scenario.yaml").read_text(encoding="utf-8"))
        self.assertEqual(trade["edits"][0]["share"], 0.15)
        cap = yaml.safe_load((SCENARIOS / "B_Opt_TxCap150" / "scenario.yaml").read_text(encoding="utf-8"))
        self.assertEqual({e["factor"] for e in cap["edits"] if "factor" in e}, {1.5})


if __name__ == "__main__":
    unittest.main()
