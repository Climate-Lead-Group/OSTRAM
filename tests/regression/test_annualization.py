"""Capital investment annualization (B2 optional post-processing).

The annualizer converts lump-sum ``CapitalInvestment`` (and
``CapitalInvestmentStorage``) into Capital Recovery Factor payment streams
using the lifetime and discount rate the model itself solved with, and
appends the result to the combined CSV as new rows without loading the
whole file into memory.
"""

from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

import numpy as np
import pandas as pd

from ostram.pipeline.execution import annualization as ann


def _crf(rate, life):
    return rate * (1 + rate) ** life / ((1 + rate) ** life - 1)


class CrfTests(unittest.TestCase):
    def test_crf_matches_closed_form(self) -> None:
        self.assertAlmostEqual(ann.calculate_crf(0.1, 20), _crf(0.1, 20), places=12)

    def test_zero_rate_is_straight_line(self) -> None:
        self.assertAlmostEqual(ann.calculate_crf(0.0, 10), 0.1, places=12)


class DatafileParameterTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.datafile = Path(self._tmp.name) / "S_0.txt"
        self.datafile.write_text(
            "\n".join(
                [
                    "param default 7 : DaysInDayType :=",
                    ";",
                    "param default 0.1 : DiscountRate :=",
                    "GLOBAL 0.08",
                    ";",
                    "param default 0.05 : DiscountRateStorage :=",
                    ";",
                    "param default 1 : OperationalLife :=",
                    "GLOBAL PWRA 20",
                    "GLOBAL PWRB 27.5",
                    ";",
                    "set EMISSION :=",
                    "CO2",
                    ";",
                ]
            ),
            encoding="utf-8",
        )

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_reads_default_and_region_override(self) -> None:
        table = ann.parse_datafile_parameter(self.datafile, "DiscountRate")
        self.assertEqual(table.default, 0.1)
        self.assertEqual(table.values, {("GLOBAL",): 0.08})

    def test_reads_default_only_block(self) -> None:
        table = ann.parse_datafile_parameter(self.datafile, "DiscountRateStorage")
        self.assertEqual(table.default, 0.05)
        self.assertEqual(table.values, {})

    def test_reads_two_key_rows(self) -> None:
        table = ann.parse_datafile_parameter(self.datafile, "OperationalLife")
        self.assertEqual(table.default, 1.0)
        self.assertEqual(table.values, {("GLOBAL", "PWRA"): 20.0, ("GLOBAL", "PWRB"): 27.5})

    def test_missing_parameter_returns_empty_table(self) -> None:
        table = ann.parse_datafile_parameter(self.datafile, "DiscountRateIdv")
        self.assertIsNone(table.default)
        self.assertEqual(table.values, {})

    def test_lookup_falls_back_to_default(self) -> None:
        table = ann.parse_datafile_parameter(self.datafile, "DiscountRate")
        self.assertEqual(table.lookup(("GLOBAL",)), 0.08)
        self.assertEqual(table.lookup(("OTHER",)), 0.1)


KEYS = ["Future", "Scenario", "REGION", "TECHNOLOGY"]


def _capex(rows):
    return pd.DataFrame(rows, columns=KEYS + ["YEAR", "value"])


def _lives(rows):
    return pd.DataFrame(rows, columns=KEYS + ["life"])


def _rates(rows):
    return pd.DataFrame(rows, columns=["Future", "Scenario", "REGION", "rate"])


class AnnualizeSeriesTests(unittest.TestCase):
    def test_single_investment_pays_crf_for_lifetime_years(self) -> None:
        result = ann.annualize_series(
            _capex([[0, "S", "GLOBAL", "PWRA", 2025, 100.0]]),
            _lives([[0, "S", "GLOBAL", "PWRA", 3]]),
            _rates([[0, "S", "GLOBAL", 0.1]]),
            horizon_end=2030,
        )
        payment = 100.0 * _crf(0.1, 3)
        self.assertEqual(list(result.columns), KEYS + ["YEAR", "annualized"])
        self.assertEqual(result["YEAR"].tolist(), [2025, 2026, 2027])
        np.testing.assert_allclose(result["annualized"].values, [payment] * 3)

    def test_payments_are_truncated_at_horizon(self) -> None:
        result = ann.annualize_series(
            _capex([[0, "S", "GLOBAL", "PWRA", 2025, 100.0]]),
            _lives([[0, "S", "GLOBAL", "PWRA", 30]]),
            _rates([[0, "S", "GLOBAL", 0.1]]),
            horizon_end=2026,
        )
        self.assertEqual(result["YEAR"].tolist(), [2025, 2026])

    def test_overlapping_investments_accumulate(self) -> None:
        result = ann.annualize_series(
            _capex(
                [
                    [0, "S", "GLOBAL", "PWRA", 2025, 100.0],
                    [0, "S", "GLOBAL", "PWRA", 2026, 50.0],
                ]
            ),
            _lives([[0, "S", "GLOBAL", "PWRA", 2]]),
            _rates([[0, "S", "GLOBAL", 0.1]]),
            horizon_end=2030,
        )
        crf = _crf(0.1, 2)
        self.assertEqual(result["YEAR"].tolist(), [2025, 2026, 2027])
        np.testing.assert_allclose(
            result["annualized"].values, [100 * crf, 150 * crf, 50 * crf]
        )

    def test_groups_are_independent_and_use_their_own_rate_and_life(self) -> None:
        result = ann.annualize_series(
            _capex(
                [
                    [0, "S", "GLOBAL", "PWRA", 2025, 100.0],
                    [0, "T", "GLOBAL", "PWRA", 2025, 100.0],
                ]
            ),
            _lives(
                [
                    [0, "S", "GLOBAL", "PWRA", 2],
                    [0, "T", "GLOBAL", "PWRA", 4],
                ]
            ),
            _rates([[0, "S", "GLOBAL", 0.1], [0, "T", "GLOBAL", 0.05]]),
            horizon_end=2030,
        )
        by_scenario = {k: g for k, g in result.groupby("Scenario")}
        self.assertEqual(by_scenario["S"]["YEAR"].tolist(), [2025, 2026])
        self.assertEqual(by_scenario["T"]["YEAR"].tolist(), [2025, 2026, 2027, 2028])
        np.testing.assert_allclose(by_scenario["S"]["annualized"].values, [100 * _crf(0.1, 2)] * 2)
        np.testing.assert_allclose(by_scenario["T"]["annualized"].values, [100 * _crf(0.05, 4)] * 4)

    def test_discounted_payments_recover_the_investment(self) -> None:
        result = ann.annualize_series(
            _capex([[0, "S", "GLOBAL", "PWRA", 2025, 100.0]]),
            _lives([[0, "S", "GLOBAL", "PWRA", 15]]),
            _rates([[0, "S", "GLOBAL", 0.0639]]),
            horizon_end=2050,
        )
        discounted = sum(
            value / (1.0639 ** (year - 2025 + 1))
            for year, value in zip(result["YEAR"], result["annualized"])
        )
        self.assertAlmostEqual(discounted, 100.0, places=9)

    def test_non_integer_lifetime_rounds_half_up(self) -> None:
        result = ann.annualize_series(
            _capex([[0, "S", "GLOBAL", "PWRA", 2025, 100.0]]),
            _lives([[0, "S", "GLOBAL", "PWRA", 2.5]]),
            _rates([[0, "S", "GLOBAL", 0.1]]),
            horizon_end=2030,
        )
        self.assertEqual(result["YEAR"].tolist(), [2025, 2026, 2027])
        np.testing.assert_allclose(result["annualized"].values, [100 * _crf(0.1, 3)] * 3)

    def test_missing_lifetime_is_an_error_naming_the_asset(self) -> None:
        with self.assertRaises(ValueError) as caught:
            ann.annualize_series(
                _capex([[0, "S", "GLOBAL", "PWRA", 2025, 100.0]]),
                _lives([]),
                _rates([[0, "S", "GLOBAL", 0.1]]),
                horizon_end=2030,
            )
        self.assertIn("PWRA", str(caught.exception))

    def test_missing_rate_is_an_error_naming_the_region(self) -> None:
        with self.assertRaises(ValueError) as caught:
            ann.annualize_series(
                _capex([[0, "S", "GLOBAL", "PWRA", 2025, 100.0]]),
                _lives([[0, "S", "GLOBAL", "PWRA", 2]]),
                _rates([]),
                horizon_end=2030,
            )
        self.assertIn("GLOBAL", str(caught.exception))

    def test_empty_capex_yields_empty_frame_with_columns(self) -> None:
        result = ann.annualize_series(_capex([]), _lives([]), _rates([]), horizon_end=2030)
        self.assertEqual(list(result.columns), KEYS + ["YEAR", "annualized"])
        self.assertEqual(len(result), 0)


COMBINED_COLUMNS = [
    "Future", "Scenario", "REGION", "YEAR", "TECHNOLOGY", "FUEL", "STORAGE",
    "CapitalInvestment", "CapitalInvestmentStorage", "NewCapacity",
    "OperationalLife", "OperationalLifeStorage",
]


class AnnualizeCombinedFileTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.executables = self.root / "Executables"
        (self.executables / "S_0").mkdir(parents=True)
        (self.executables / "S_0" / "S_0.txt").write_text(
            "param default 0.1 : DiscountRate :=\n;\n"
            "param default 0.05 : DiscountRateStorage :=\n;\n",
            encoding="utf-8",
        )
        rows = [
            # inputs block
            [0, "S", "GLOBAL", None, "PWRA", None, None, None, None, None, 2, None],
            [0, "S", "GLOBAL", None, None, None, "BAT", None, None, None, None, 3],
            # outputs block
            [0, "S", "GLOBAL", 2025.0, "PWRA", None, None, 100.0, None, None, None, None],
            [0, "S", "GLOBAL", 2025.0, "PWRA", "ELC", None, None, None, 1.5, None, None],
            [0, "S", "GLOBAL", 2030.0, "PWRA", None, None, None, None, 2.0, None, None],
            [0, "S", "GLOBAL", 2026.0, None, None, "BAT", None, 10.0, None, None, None],
        ]
        self.combined = self.root / "OSTRAM_Combined_Inputs_Outputs.csv"
        pd.DataFrame(rows, columns=COMBINED_COLUMNS).to_csv(self.combined, index=False)
        self.original = self.combined.read_text(encoding="utf-8")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _run(self, **kwargs):
        kwargs.setdefault("executables_dir", self.executables)
        kwargs.setdefault("verbose", False)
        return ann.annualize_capital_investment(self.combined, **kwargs)

    def test_appends_annualized_columns_and_rows_without_touching_existing_rows(self) -> None:
        self._run()
        updated = self.combined.read_text(encoding="utf-8")
        original_lines = self.original.splitlines()
        updated_lines = updated.splitlines()
        self.assertEqual(
            updated_lines[0],
            original_lines[0] + ",CapitalInvestmentAnnualized,CapitalInvestmentStorageAnnualized",
        )
        for old, new in zip(original_lines[1:], updated_lines[1:]):
            self.assertEqual(new, old + ",,")
        df = pd.read_csv(self.combined, low_memory=False)
        self.assertEqual(len(df), 6 + 2 + 3)

        tech = df[df["CapitalInvestmentAnnualized"].notna()]
        self.assertEqual(tech["YEAR"].tolist(), [2025.0, 2026.0])
        self.assertEqual(tech["TECHNOLOGY"].tolist(), ["PWRA", "PWRA"])
        self.assertTrue(tech["STORAGE"].isna().all())
        self.assertTrue(tech["FUEL"].isna().all())
        self.assertTrue(tech["CapitalInvestment"].isna().all())
        np.testing.assert_allclose(tech["CapitalInvestmentAnnualized"].values, [100 * _crf(0.1, 2)] * 2)

        storage = df[df["CapitalInvestmentStorageAnnualized"].notna()]
        self.assertEqual(storage["YEAR"].tolist(), [2026.0, 2027.0, 2028.0])
        self.assertEqual(storage["STORAGE"].tolist(), ["BAT"] * 3)
        self.assertTrue(storage["TECHNOLOGY"].isna().all())
        np.testing.assert_allclose(
            storage["CapitalInvestmentStorageAnnualized"].values, [10 * _crf(0.05, 3)] * 3
        )

    def test_horizon_is_the_last_year_present_in_the_file(self) -> None:
        self._run()
        df = pd.read_csv(self.combined, low_memory=False)
        self.assertEqual(df["YEAR"].max(), 2030.0)
        storage = df[df["CapitalInvestmentStorageAnnualized"].notna()]
        self.assertLessEqual(storage["YEAR"].max(), 2030.0)

    def test_returns_summary_counts(self) -> None:
        summary = self._run()
        self.assertEqual(summary["rows_read"], 6)
        self.assertEqual(summary["rows_appended"], 5)
        self.assertEqual(
            summary["columns_added"],
            ["CapitalInvestmentAnnualized", "CapitalInvestmentStorageAnnualized"],
        )

    def test_second_run_on_annualized_file_is_refused(self) -> None:
        self._run()
        with self.assertRaises(ValueError):
            self._run()
        # the refusal leaves the file untouched
        self.assertEqual(pd.read_csv(self.combined, low_memory=False).shape[0], 11)

    def test_missing_datafile_without_explicit_rate_is_an_error(self) -> None:
        (self.executables / "S_0" / "S_0.txt").unlink()
        with self.assertRaises(FileNotFoundError):
            self._run()
        self.assertEqual(self.combined.read_text(encoding="utf-8"), self.original)

    def test_explicit_rate_overrides_datafile_for_every_series(self) -> None:
        (self.executables / "S_0" / "S_0.txt").unlink()
        self._run(discount_rate=0.2)
        df = pd.read_csv(self.combined, low_memory=False)
        tech = df["CapitalInvestmentAnnualized"].dropna().values
        storage = df["CapitalInvestmentStorageAnnualized"].dropna().values
        np.testing.assert_allclose(tech, [100 * _crf(0.2, 2)] * 2)
        np.testing.assert_allclose(storage, [10 * _crf(0.2, 3)] * 3)

    def test_explicit_lifetime_overrides_operational_life(self) -> None:
        self._run(asset_lifetime=1)
        df = pd.read_csv(self.combined, low_memory=False)
        self.assertEqual(df["CapitalInvestmentAnnualized"].dropna().tolist(), [100 * _crf(0.1, 1)])
        self.assertEqual(len(df["CapitalInvestmentStorageAnnualized"].dropna()), 1)

    def test_file_without_capital_columns_is_left_unchanged(self) -> None:
        pd.DataFrame(
            [[0, "S", "GLOBAL", 2025.0, "PWRA", 1.0]],
            columns=["Future", "Scenario", "REGION", "YEAR", "TECHNOLOGY", "NewCapacity"],
        ).to_csv(self.combined, index=False)
        before = self.combined.read_text(encoding="utf-8")
        summary = self._run()
        self.assertEqual(summary["rows_appended"], 0)
        self.assertEqual(summary["columns_added"], [])
        self.assertEqual(self.combined.read_text(encoding="utf-8"), before)

    def test_appended_rows_use_the_file_line_terminator(self) -> None:
        self.combined.write_bytes(self.original.replace("\n", "\r\n").encode("utf-8"))
        self._run()
        raw = self.combined.read_bytes()
        self.assertNotIn(b"\r\r\n", raw)
        self.assertEqual(raw.count(b"\r\n"), raw.count(b"\n"))
        with open(self.combined, newline="") as handle:
            self.assertEqual(len(list(csv.reader(handle))), 12)


if __name__ == "__main__":
    unittest.main()
