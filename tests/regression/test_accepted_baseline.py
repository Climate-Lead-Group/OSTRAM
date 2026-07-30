from __future__ import annotations

import copy
import csv
import hashlib
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
sys.path.insert(0, str(TEST_ROOT))

import accepted_baseline as baseline


def _git(*args: str) -> str | None:
    result = subprocess.run(
        ["git", "-C", str(REPO_ROOT), *args],
        check=False,
        capture_output=True,
        text=True,
    )
    if result.returncode not in (0, 1):
        raise AssertionError(result.stderr)
    return result.stdout if result.stdout else None


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

    def test_governed_manifest_binds_root_and_derived_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = root / "STAGE_2_GOVERNED_COMPARATOR_MANIFEST.csv"
            rows = []
            for index, scenario in enumerate(
                baseline.canonical_scenarios(), start=1
            ):
                payload = f"{scenario} value {index}\n".encode("utf-8")
                filename = (
                    f"Pre_processed_{scenario}_0_"
                    "StorageDelayN5_OpenBCK_RMCarefulXLSX.txt"
                )
                output = (
                    root
                    / "t1_confection"
                    / "Executables"
                    / f"{scenario}_0"
                    / filename
                )
                output.parent.mkdir(parents=True)
                output.write_bytes(payload)
                rows.append(
                    {
                        "Scenario": scenario,
                        "AuthorityClass": (
                            baseline.GOVERNED_ROOT_AUTHORITY
                            if scenario in baseline.DECISION_ROOTS
                            else baseline.GOVERNED_DERIVED_AUTHORITY
                        ),
                        "SHA256": hashlib.sha256(payload).hexdigest(),
                        "ByteSize": len(payload),
                        "LineCount": 1,
                        "Provenance": "fixture root plus declared rules",
                    }
                )
            with manifest.open("w", encoding="utf-8", newline="") as stream:
                writer = csv.DictWriter(
                    stream, fieldnames=baseline.GOVERNED_MANIFEST_COLUMNS
                )
                writer.writeheader()
                writer.writerows(rows)

            loaded = baseline.load_governed_manifest(manifest)
            self.assertEqual(len(loaded), 15)
            self.assertEqual(
                len(baseline.validate_governed_output_files(root, loaded)),
                15,
            )

            rows[1]["AuthorityClass"] = baseline.GOVERNED_ROOT_AUTHORITY
            with manifest.open("w", encoding="utf-8", newline="") as stream:
                writer = csv.DictWriter(
                    stream, fieldnames=baseline.GOVERNED_MANIFEST_COLUMNS
                )
                writer.writeheader()
                writer.writerows(rows)
            with self.assertRaisesRegex(
                baseline.BaselineValidationError, "authority class drift"
            ):
                baseline.load_governed_manifest(manifest)

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
        output = _git(
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
