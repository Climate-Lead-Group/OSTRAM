from __future__ import annotations

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

    def test_narrow_ignore_rules_and_scope(self) -> None:
        result = baseline.validate_repository(REPO_ROOT)
        self.assertEqual(result["root_scenarios"], list(baseline.EXPECTED_ROOT_SCENARIOS))
        self.assertEqual(
            result["scenario_order"],
            list(baseline.EXPECTED_DECISION_SCENARIOS),
        )

        probes = (
            "t1_confection/A1_Outputs/Probe/A-O_Demand.xlsx",
            "t1_confection/A2_Output_Params/Probe/VariableCost.csv",
            "t1_confection/A2_Outputs_Params_otoole/Probe/VariableCost.csv",
            "t1_confection/Executables/Probe_0/arbitrary.log",
            "t1_confection/Outputs/Probe.csv",
            "t1_confection/Config_MOMF_T1_A.yaml",
            "tests/regression/accepted_baseline.py",
        )
        output = _git("check-ignore", "-v", "--no-index", *probes)
        self.assertIsNotNone(output)
        matched: dict[str, str] = {}
        for line in output.splitlines():
            source, path = line.split("\t", 1)
            pattern = source.rsplit(":", 1)[-1]
            matched[path.replace("\\", "/")] = pattern

        for probe, rule in zip(probes[:5], baseline.EXPECTED_IGNORE_RULES):
            self.assertEqual(matched[probe], rule)
        for probe in probes[5:]:
            self.assertNotIn(probe, matched)


if __name__ == "__main__":
    unittest.main()
