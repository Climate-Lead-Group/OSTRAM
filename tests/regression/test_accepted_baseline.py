from __future__ import annotations

import copy
import sys
import unittest
from pathlib import Path


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
sys.path.insert(0, str(TEST_ROOT))

import accepted_baseline as baseline
import ostram_regression as regression


class AcceptedBaselineTests(unittest.TestCase):
    def test_portable_record_is_exact_and_canonical(self) -> None:
        result = baseline.validate_repository(REPO_ROOT)
        self.assertEqual(result["scenario_count"], 15)
        self.assertEqual(
            tuple(result["scenario_order"]),
            tuple(item[0] for item in baseline.EXPECTED_COMPILED_INPUTS),
        )
        self.assertEqual(
            result["protected_manifest_sha256"],
            baseline.EXPECTED_MANIFEST_SHA256,
        )

    def test_exact_filenames_sizes_and_hashes_are_bound(self) -> None:
        record = baseline.load_accepted_record()
        actual = tuple(
            (item["scenario"], item["size_bytes"], item["sha256"])
            for item in record["scenarios"]
        )
        self.assertEqual(actual, baseline.EXPECTED_COMPILED_INPUTS)
        for item in record["scenarios"]:
            scenario = item["scenario"]
            expected_filename = (
                f"Pre_processed_{scenario}_0_"
                "StorageDelayN5_OpenBCK_RMCarefulXLSX.txt"
            )
            self.assertEqual(item["filename"], expected_filename)
            self.assertEqual(
                item["relative_path"],
                f"t1_confection/Executables/{scenario}_0/{expected_filename}",
            )

    def test_record_validator_rejects_identity_drift(self) -> None:
        record = baseline.load_accepted_record()
        changed = copy.deepcopy(record)
        changed["scenarios"][0]["size_bytes"] += 1
        with self.assertRaisesRegex(
            baseline.BaselineValidationError, "byte-count drift"
        ):
            baseline.validate_record(changed)

    def test_report_lineage_is_distinct_and_byte_preserved(self) -> None:
        baseline.validate_report_lineage(REPO_ROOT)

    def test_narrow_ignore_rules_and_scope(self) -> None:
        baseline.validate_ignore_rules(REPO_ROOT)
        probes = (
            "t1_confection/Executables/Probe_0/run_output.log",
            "t1_confection/Executables/Probe_0/_validation_report.csv",
            "t1_confection/Executables/Probe_0/arbitrary.log",
            "t1_confection/Executables/Probe_0/Pre_processed_Probe_0.txt",
            "t1_confection/Config_MOMF_T1_A.yaml",
            "tests/regression/reports/accepted_compiled_solver_baseline_15.json",
        )
        output = regression._git(
            REPO_ROOT,
            "check-ignore",
            "-v",
            "--no-index",
            *probes,
        )
        self.assertIsNotNone(output)
        matched: dict[str, str] = {}
        for line in output.splitlines():
            source, path = line.split("\t", 1)
            pattern = source.rsplit(":", 1)[-1]
            matched[path.replace("\\", "/")] = pattern

        self.assertEqual(
            matched[probes[0]], baseline.EXPECTED_IGNORE_RULES[0]
        )
        self.assertEqual(
            matched[probes[1]], baseline.EXPECTED_IGNORE_RULES[1]
        )
        for probe in probes[2:]:
            self.assertNotIn(
                matched.get(probe),
                baseline.EXPECTED_IGNORE_RULES,
                probe,
            )


if __name__ == "__main__":
    unittest.main()
