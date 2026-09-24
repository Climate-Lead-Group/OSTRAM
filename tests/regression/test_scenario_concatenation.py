"""Streaming cross-scenario concatenation (B2 final post-processing).

The concatenator must produce the same three CSV files as the legacy
whole-campaign implementation while holding only one scenario in memory
at a time.  These tests pin the observable contract: column union and
order, per-scenario sorted blocks, the accumulated minimum-investment
column and the dated copies.
"""

from __future__ import annotations

import csv
import filecmp
import os
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from ostram.pipeline.execution import scenario_concatenation as concat


def _params(**overrides):
    params = {
        "executables": "Executables",
        "prefix_final_files": "OSTRAM_",
        "inputs_file": "Inputs.csv",
        "outputs_file": "Outputs.csv",
        "preprocess_data_name": "Pre_processed_",
        "output_files": "_output",
    }
    params.update(overrides)
    return params


def _write_scenario(root: Path, name: str, inputs: pd.DataFrame, outputs: pd.DataFrame | None):
    folder = root / "Executables" / name
    folder.mkdir(parents=True)
    inputs.to_csv(folder / f"{name}_Input.csv", index=False)
    if outputs is not None:
        outputs.to_csv(folder / f"Pre_processed_{name}_output.csv", index=False)


def _read(path):
    return pd.read_csv(path, low_memory=False)


class ScenarioConcatenationTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        (self.root / "Executables" / "Default").mkdir(parents=True)
        (self.root / "Executables" / "__pycache__").mkdir(parents=True)

        # Scenario B is created first on disk but must sort after A.
        _write_scenario(
            self.root,
            "B_0",
            inputs=pd.DataFrame(
                {
                    "REGION": ["GLOBAL", "GLOBAL", "GLOBAL"],
                    "TECHNOLOGY": ["PWRB", "PWRB", "PWRA"],
                    "YEAR": [2024.0, 2023.0, 2023.0],
                    "TotalAnnualMinCapacityInvestment": [2.0, 1.0, 5.0],
                    "OperationalLife": [None, None, None],
                }
            ),
            outputs=pd.DataFrame(
                {
                    "REGION": ["GLOBAL"],
                    "TECHNOLOGY": ["PWRB"],
                    "YEAR": [2023.0],
                    "NewCapacity": [0.5],
                    "OnlyInB": [7.0],
                }
            ),
        )
        _write_scenario(
            self.root,
            "A_0",
            inputs=pd.DataFrame(
                {
                    "REGION": ["GLOBAL", "GLOBAL"],
                    "TECHNOLOGY": ["PWRA", "PWRA"],
                    "YEAR": [2024.0, 2023.0],
                    "TotalAnnualMinCapacityInvestment": [3.0, None],
                    "CapitalCost": [10.0, 11.0],
                }
            ),
            outputs=pd.DataFrame(
                {
                    "REGION": ["GLOBAL", "GLOBAL"],
                    "TECHNOLOGY": ["PWRA", "PWRA"],
                    "YEAR": [2024.0, 2023.0],
                    "FUEL": [None, "ELC"],
                    "NewCapacity": [1.0, 2.0],
                }
            ),
        )

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _run(self, **overrides):
        return concat.concatenate_all_scenarios(str(self.root), _params(**overrides))

    def test_output_candidates_prefer_active_patch_chain_then_legacy_names(self) -> None:
        params = _params(storage_delay_active=True, reserve_margin_xlsx_active=True)
        self.assertEqual(
            concat.active_output_csv_candidates(params, "A_0"),
            [
                "Pre_processed_A_0_StorageDelayN5_RMCarefulXLSX_output.csv",
                "Pre_processed_A_0_output.csv",
                "Pre_processed_A_0_Output.csv",
            ],
        )

    def test_discovers_scenario_folders_in_sorted_order_skipping_service_entries(self) -> None:
        found = concat.discover_scenario_files(str(self.root), _params())
        self.assertEqual([item.name for item in found], ["A_0", "B_0"])
        self.assertEqual((found[0].scenario, found[0].future), ("A", "0"))
        self.assertTrue(found[0].input_path.endswith("A_0_Input.csv"))
        self.assertTrue(found[0].output_path.endswith("Pre_processed_A_0_output.csv"))

    def test_inputs_file_has_union_columns_metadata_first_and_sorted_rows(self) -> None:
        path_in, _, _ = self._run()
        df = _read(path_in)
        self.assertEqual(
            list(df.columns),
            [
                "Future",
                "Scenario",
                "REGION",
                "YEAR",
                "TECHNOLOGY",
                "CapitalCost",
                "OperationalLife",
                "TotalAnnualMinCapacityInvestment",
            ],
        )
        self.assertEqual(
            df[["Scenario", "TECHNOLOGY", "YEAR"]].values.tolist(),
            [
                ["A", "PWRA", 2023.0],
                ["A", "PWRA", 2024.0],
                ["B", "PWRA", 2023.0],
                ["B", "PWRB", 2023.0],
                ["B", "PWRB", 2024.0],
            ],
        )
        self.assertTrue((df["Future"] == 0).all())

    def test_outputs_file_keeps_scenario_specific_columns_and_fuel_key(self) -> None:
        _, path_out, _ = self._run()
        df = _read(path_out)
        self.assertEqual(
            list(df.columns),
            ["Future", "Scenario", "REGION", "YEAR", "TECHNOLOGY", "FUEL", "NewCapacity", "OnlyInB"],
        )
        self.assertEqual(df["Scenario"].tolist(), ["A", "A", "B"])
        self.assertTrue(df.loc[df["Scenario"] == "A", "OnlyInB"].isna().all())
        self.assertEqual(df.loc[df["Scenario"] == "B", "OnlyInB"].tolist(), [7.0])

    def test_combined_file_accumulates_min_investment_within_scenario_and_technology(self) -> None:
        _, _, path_comb = self._run()
        df = _read(path_comb)
        self.assertEqual(df.columns[-1], "AccumulatedTotalAnnualMinCapacityInvestment")
        self.assertEqual(len(df), 5 + 3)
        rows = df[df["TotalAnnualMinCapacityInvestment"].notna()]
        got = {
            (r.Scenario, r.TECHNOLOGY, r.YEAR): r.AccumulatedTotalAnnualMinCapacityInvestment
            for r in rows.itertuples()
        }
        self.assertEqual(
            got,
            {
                ("A", "PWRA", 2024.0): 3.0,
                ("B", "PWRA", 2023.0): 5.0,
                ("B", "PWRB", 2023.0): 1.0,
                ("B", "PWRB", 2024.0): 3.0,
            },
        )
        self.assertTrue(df.loc[df["TotalAnnualMinCapacityInvestment"].isna(), df.columns[-1]].isna().all())

    def test_combined_file_omits_accumulated_column_when_parameter_absent(self) -> None:
        for name in ("A_0", "B_0"):
            csv_path = self.root / "Executables" / name / f"{name}_Input.csv"
            frame = _read(csv_path).drop(columns=["TotalAnnualMinCapacityInvestment"])
            frame.to_csv(csv_path, index=False)
        _, _, path_comb = self._run()
        self.assertNotIn("AccumulatedTotalAnnualMinCapacityInvestment", _read(path_comb).columns)

    def test_dated_copies_are_byte_identical_and_named_after_today(self) -> None:
        path_in, path_out, path_comb = self._run()
        today = pd.Timestamp.today().date().isoformat()
        for path in (path_in, path_out):
            dated = path.replace(".csv", f"_{today}.csv")
            self.assertTrue(os.path.exists(dated), dated)
            self.assertTrue(filecmp.cmp(path, dated, shallow=False))
        self.assertFalse(os.path.exists(path_comb.replace(".csv", f"_{today}.csv")))
        self.assertTrue(path_comb.endswith("OSTRAM_Combined_Inputs_Outputs.csv"))

    def test_scenario_without_output_csv_contributes_inputs_only(self) -> None:
        os.remove(self.root / "Executables" / "B_0" / "Pre_processed_B_0_output.csv")
        path_in, path_out, path_comb = self._run()
        self.assertEqual(_read(path_out)["Scenario"].tolist(), ["A", "A"])
        self.assertEqual(_read(path_in)["Scenario"].tolist(), ["A", "A", "B", "B", "B"])
        self.assertEqual(len(_read(path_comb)), 7)

    def test_empty_executables_returns_no_paths(self) -> None:
        for name in ("A_0", "B_0"):
            folder = self.root / "Executables" / name
            for child in folder.iterdir():
                child.unlink()
            folder.rmdir()
        self.assertEqual(self._run(), (None, None, None))

    def test_written_csv_has_no_index_column_and_one_header(self) -> None:
        path_in, _, _ = self._run()
        with open(path_in, newline="") as handle:
            rows = list(csv.reader(handle))
        self.assertEqual(rows[0][0], "Future")
        self.assertEqual(sum(1 for row in rows if row[0] == "Future"), 1)
        self.assertEqual(len(rows), 6)


if __name__ == "__main__":
    unittest.main()
