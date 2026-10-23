"""The base-year pin table holds no values stored from an earlier run.

The rows labelled ``BENCHMARK_SUPPORTED`` with lineage
``ACCEPTED_BASE_YEAR_PIN_2023_2026`` held 2023 to 2026 values of
``TotalAnnualMaxCapacity``, ``TotalAnnualMaxCapacityInvestment`` and
``TotalAnnualMinCapacityInvestment`` stored from an earlier run. They were
removed. The ``RETIRED_SOLVER_REPLAY`` rows stay because they only clear
cells that an earlier replay had set.
"""
from __future__ import annotations

import csv
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
PIN_CSV = REPO_ROOT / "config" / "scenarios" / "rules" / "pwr_min_2023_2026_pin.csv"

STORED_CLASSIFICATION = "BENCHMARK_SUPPORTED"
STORED_LINEAGE = "ACCEPTED_BASE_YEAR_PIN_2023_2026"
RETIRED = "RETIRED_SOLVER_REPLAY"


def _rows() -> list[dict[str, str]]:
    with PIN_CSV.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        fields = set(reader.fieldnames or ())
        rows = list(reader)
    missing = {"authority_classification", "authority_lineage_class"} - fields
    if missing or not rows:
        raise AssertionError(f"{PIN_CSV} has no rows or lacks {sorted(missing)}")
    return rows


class PinHoldsNoStoredRunOutputTests(unittest.TestCase):
    def test_no_row_carries_the_stored_run_labels(self) -> None:
        stored = [
            row["source_rule_id"]
            for row in _rows()
            if row["authority_classification"] == STORED_CLASSIFICATION
            or row["authority_lineage_class"] == STORED_LINEAGE
        ]
        if stored:
            self.fail(f"{len(stored)} rows stored from an earlier run, e.g. {stored[:3]}")

    def test_retired_solver_replay_rows_only_clear_cells(self) -> None:
        setting = [
            row["source_rule_id"]
            for row in _rows()
            if row["authority_classification"] == RETIRED
            and (
                row["required_root_present"] != "false"
                or row["required_root_value"] != ""
                or row["required_root_state"] != "ABSENT"
            )
        ]
        if setting:
            self.fail(f"{len(setting)} {RETIRED} rows set a value, e.g. {setting[:3]}")


if __name__ == "__main__":
    unittest.main()
