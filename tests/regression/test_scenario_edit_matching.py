"""A scenario edit or VRE ceiling that matches nothing stops the run."""

from __future__ import annotations

import contextlib
import io
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from openpyxl import Workbook

from ostram.pipeline.scenarios import scenario_edits as edits


YEARS = list(range(2023, 2051))
HEADER = [
    "Tech.ID", "Tech", "Tech.Name", "Parameter.ID", "Parameter", "Unit",
    "Projection.Mode", "Projection.Parameter", *YEARS,
]
SCENARIO = "B_Opt_MyTest"


def _workbook():
    workbook = Workbook()
    secondary = workbook.active
    secondary.title = "Secondary Techs"
    secondary.append(HEADER)
    secondary.append([
        1, "PWRSOLBGDXX", "solar", 1, "TotalAnnualMaxCapacity", "GW",
        "User defined", 0, *[5.0] * len(YEARS),
    ])
    secondary.append([
        1, "PWRSOLBGDXX", "solar", 2, "CapitalCost", "USD", "User defined", 0,
        *[100.0] * len(YEARS),
    ])
    secondary.append([
        2, "PWRSPVBGDXX", "solar", 1, "TotalAnnualMaxCapacity", "GW",
        "User defined", 0, *[40.0] * len(YEARS),
    ])
    demand = workbook.create_sheet("Demand Techs")
    demand.append(HEADER)
    demand.append([
        3, "TRNNLIBGDXX", "backstop", 1, "ResidualCapacity", "GW",
        "User defined", 0, *[1.0] * len(YEARS),
    ])
    return workbook


def _log():
    return {"scenario": SCENARIO, "cells": [], "rows_created": [], "skipped": []}


def _apply(workbook, edit):
    sheet = workbook[edit["sheet"]]
    columns, years = edits.scan_columns(sheet)
    log = _log()
    edits.apply_edit(sheet, columns, years, edit, log)
    return log


class UnmatchedScenarioEditTests(unittest.TestCase):
    def assert_stops(self, edit, message):
        with self.assertRaises(edits.ScenarioEditError) as caught:
            _apply(_workbook(), edit)
        self.assertEqual(str(caught.exception), message)

    def test_misspelt_technology_stops_with_a_suggestion(self) -> None:
        self.assert_stops(
            {
                "sheet": "Secondary Techs", "tech": "PWRSOLBGXX",
                "param": "TotalAnnualMaxCapacity", "op": "set_flat", "value": 1.0,
            },
            f'Scenario {SCENARIO}: the edit to TotalAnnualMaxCapacity for '
            '"PWRSOLBGXX" changed nothing. No technology has that name. '
            'Did you mean PWRSOLBGDXX?',
        )

    def test_prefix_that_matches_nothing_stops(self) -> None:
        self.assert_stops(
            {
                "sheet": "Secondary Techs", "tech_prefix": "PWRSOLX",
                "param": "CapitalCost", "op": "multiply", "factor": 1.1,
            },
            f'Scenario {SCENARIO}: the edit to CapitalCost for "PWRSOLX" '
            'changed nothing. No technology starts with "PWRSOLX".',
        )

    def test_misspelt_parameter_stops_with_a_suggestion(self) -> None:
        self.assert_stops(
            {
                "sheet": "Secondary Techs", "tech": "PWRSOLBGDXX",
                "param": "TotalAnualMaxCapacity", "op": "set_flat", "value": 1.0,
            },
            f'Scenario {SCENARIO}: the edit to TotalAnualMaxCapacity for '
            '"PWRSOLBGDXX" changed nothing. Sheet "Secondary Techs" has no '
            'parameter "TotalAnualMaxCapacity". Did you mean '
            'TotalAnnualMaxCapacity?',
        )

    def test_create_if_absent_with_an_unknown_technology_stops(self) -> None:
        self.assert_stops(
            {
                "sheet": "Demand Techs", "tech": "TRNNLIBGXX",
                "param": "TotalAnnualMaxCapacity", "values": {"2030": 0.0},
                "create_if_absent": True,
            },
            f'Scenario {SCENARIO}: the edit to TotalAnnualMaxCapacity for '
            '"TRNNLIBGXX" changed nothing. No technology has that name. '
            'Did you mean TRNNLIBGDXX?',
        )

    def test_years_filter_with_no_matching_column_stops(self) -> None:
        self.assert_stops(
            {
                "sheet": "Secondary Techs", "tech": "PWRSOLBGDXX",
                "param": "CapitalCost", "op": "multiply", "factor": 1.5,
                "years": [2061, 2062],
            },
            f'Scenario {SCENARIO}: the edit to CapitalCost for "PWRSOLBGDXX" '
            'changed nothing. None of the years 2061, 2062 is a column in '
            'the workbook.',
        )

    def test_vre_ceiling_that_matches_nothing_stops(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            ceilings = Path(temp) / "full.yaml"
            ceilings.write_text(
                "vre_ceilings:\n"
                "  sheet: Secondary Techs\n"
                "  param: TotalAnnualMaxCapacity\n"
                "  ceilings_gw:\n"
                "    PWRSPVBGXX: 1.0\n",
                encoding="utf-8",
            )
            with self.assertRaises(edits.ScenarioEditError) as caught:
                edits.apply_ceiling_layer(_workbook(), _log(), ceiling_path=ceilings)
        self.assertEqual(
            str(caught.exception),
            f'Scenario {SCENARIO}: the VRE ceiling for "PWRSPVBGXX" changed '
            'nothing. No technology has that name. Did you mean PWRSPVBGDXX?',
        )

    def test_edit_that_matches_rows_but_changes_no_value_passes(self) -> None:
        log = _apply(
            _workbook(),
            {
                "sheet": "Secondary Techs", "tech": "PWRSOLBGDXX",
                "param": "TotalAnnualMaxCapacity", "op": "set_flat", "value": 5.0,
            },
        )
        self.assertEqual(log["cells"], [])
        self.assertEqual(log["skipped"], [])

    def test_create_if_absent_with_known_names_creates_the_row(self) -> None:
        log = _apply(
            _workbook(),
            {
                "sheet": "Demand Techs", "tech": "TRNNLIBGDXX",
                "param": "TotalAnnualMaxCapacity", "values": {"2030": 0.0},
                "create_if_absent": True,
            },
        )
        self.assertEqual(
            [(row["tech"], row["param"]) for row in log["rows_created"]],
            [("TRNNLIBGDXX", "TotalAnnualMaxCapacity")],
        )


class StandaloneStopMessageTests(unittest.TestCase):
    def test_main_prints_the_stop_message_without_a_traceback(self) -> None:
        message = f"Scenario {SCENARIO}: the edit changed nothing."
        stderr = io.StringIO()
        with (
            mock.patch.object(
                edits, "build_scenario",
                side_effect=edits.ScenarioEditError(message),
            ),
            mock.patch("sys.argv", ["scenario_edits", "--scenario", SCENARIO]),
            contextlib.redirect_stderr(stderr),
        ):
            code = edits.main()
        self.assertEqual(code, 1)
        self.assertEqual(stderr.getvalue(), message + "\n")
        self.assertNotIn("Traceback", stderr.getvalue())


if __name__ == "__main__":
    unittest.main()
