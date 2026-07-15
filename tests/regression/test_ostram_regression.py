from __future__ import annotations

import hashlib
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))

import ostram_regression as regression


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
FIXTURES = TEST_ROOT / "fixtures"
BASELINE = TEST_ROOT / "baselines" / "5ce4e66480e1-static-nosolver"


class NormalizationTests(unittest.TestCase):
    def test_normalization_drops_generated_index_sorts_and_normalizes_numbers(self) -> None:
        left = "Unnamed: 0,TECHNOLOGY,YEAR,VALUE\r\n1,B,2030.0,-0.0\r\n0,A,2029,1.250000000000000\r\n"
        right = "TECHNOLOGY,YEAR,VALUE\nA,2029.0,1.25\nB,2030,0\n"
        self.assertEqual(
            regression.normalize_csv_text(left).payload,
            regression.normalize_csv_text(right).payload,
        )

    def test_unicode_is_normalized(self) -> None:
        composed = "TECHNOLOGY,VALUE\nCaf\u00e9,1\n"
        decomposed = "TECHNOLOGY,VALUE\nCafe\u0301,1.0\n"
        self.assertEqual(
            regression.normalize_csv_text(composed).payload,
            regression.normalize_csv_text(decomposed).payload,
        )

    def test_invalid_numeric_is_rejected(self) -> None:
        with self.assertRaisesRegex(regression.RegressionError, "non-finite"):
            regression.normalize_csv_text("TECHNOLOGY,VALUE\nA,nan\n")

    def test_blank_parameter_values_are_omitted_only_when_requested(self) -> None:
        text = "TECHNOLOGY,YEAR,Value\nA,2030,\nB,2030,2\n"
        with self.assertRaisesRegex(regression.RegressionError, "invalid VALUE"):
            regression.normalize_csv_text(text)
        table = regression.normalize_csv_text(text, omit_blank_values=True)
        self.assertEqual(table.payload, b"TECHNOLOGY,YEAR,Value\nB,2030,2\n")

    def test_duplicate_key_is_rejected(self) -> None:
        with self.assertRaisesRegex(regression.RegressionError, "duplicate key"):
            regression.normalize_csv_text("TECHNOLOGY,YEAR,VALUE\nA,2030,1\nA,2030.0,2\n")

    def test_single_value_column_is_a_set_key(self) -> None:
        table = regression.normalize_csv_text('VALUE\n2\n""\n1\n""\n')
        self.assertEqual(table.key_columns, ("VALUE",))
        self.assertEqual(table.payload, b"VALUE\n1\n2\n")


class DiscoveryTests(unittest.TestCase):
    def test_exact_twenty_discovery(self) -> None:
        inventory = regression.load_scenarios()
        result = regression.discover_scenarios(REPO_ROOT, inventory)
        self.assertTrue(regression.discovery_passes(result))
        self.assertEqual(len(result["expected"]), 20)

    def test_missing_and_unexpected_are_reported(self) -> None:
        inventory = [{"name": f"S{i:02d}"} for i in range(20)]
        expected = {item["name"] for item in inventory}
        with mock.patch.object(
            regression,
            "_scenario_dirs",
            side_effect=[expected - {"S19"}, expected | {"EXTRA"}, set(), set()],
        ):
            result = regression.discover_scenarios(REPO_ROOT, inventory)
        self.assertEqual(result["missing_a1"], {"S19"})
        self.assertEqual(result["unexpected_configs"], {"EXTRA"})
        self.assertFalse(regression.discovery_passes(result))

    def test_cleanup_scope_preserves_twenty_and_accepts_sixteen(self) -> None:
        inventory = regression.load_scenarios()
        selected = regression.scenarios_for_scope(inventory, "cleanup-acceptance")
        excluded = {item["name"] for item in inventory if not item["cleanup_acceptance"]}
        self.assertEqual(len(inventory), 20)
        self.assertEqual(len(selected), 16)
        self.assertEqual(
            excluded,
            {"B_Opt_LinkFreeze", "B_Opt_SolarHi10", "B_Opt_TradeCap30", "B_Opt_TradeCap50"},
        )

    def test_cleanup_acceptance_discovery_is_complete(self) -> None:
        inventory = regression.load_scenarios()
        selected = regression.scenarios_for_scope(inventory, "cleanup-acceptance")
        result = regression.discover_scenarios(REPO_ROOT, selected)
        self.assertTrue(regression.cleanup_acceptance_discovery_passes(result))
        self.assertEqual(result["missing_a2"], set())
        self.assertEqual(result["missing_otoole"], set())


class CleanupAcceptanceGateTests(unittest.TestCase):
    def test_committed_evidence_passes_cleanup_acceptance(self) -> None:
        report = regression.evaluate_cleanup_acceptance(BASELINE)
        self.assertTrue(report["ok"])
        self.assertEqual(report["preservation_scenario_count"], 20)
        self.assertEqual(report["cleanup_acceptance_scenario_count"], 16)
        self.assertEqual(report["static_comparison_summary"], {"exact": 62, "normalized-exact": 2})
        self.assertEqual(report["solver_execution"], "not-performed")

    def test_missing_accepted_artifact_fails_gate(self) -> None:
        inventory = regression.load_scenarios()
        coverage = regression._read_csv_dicts(BASELINE / "coverage.csv")
        comparisons = regression._read_csv_dicts(BASELINE / "comparisons.csv")
        changed = [dict(row) for row in coverage]
        next(row for row in changed if row["scenario"] == "BAU")["working_tracked_a2"] = "False"
        report = regression.cleanup_acceptance_report(inventory, changed, comparisons)
        self.assertFalse(report["ok"])
        self.assertIn("BAU: required field working_tracked_a2 is not present", report["failures"])


class HashAndComparisonTests(unittest.TestCase):
    def test_streaming_hash(self) -> None:
        path = regression.DEFAULT_SCENARIOS
        self.assertEqual(
            regression.sha256_file(path, chunk_size=7),
            hashlib.sha256(path.read_bytes()).hexdigest(),
        )

    def test_missing_file_detection(self) -> None:
        report = regression.required_files_report(TEST_ROOT, ["scenarios.yaml", "missing.csv"])
        self.assertFalse(report["ok"])
        self.assertEqual(report["missing"], ["missing.csv"])

    def test_numeric_equivalence_is_distinct_from_exact_hash(self) -> None:
        left = FIXTURES / "numeric_left.csv"
        right = FIXTURES / "numeric_right.csv"
        result = regression.compare_csv_files(left, right, absolute=1e-8, relative=1e-8)
        self.assertEqual(result.status, "numeric-equivalent/hash-drift")
        self.assertTrue(result.passed)

    def test_comparison_reports_missing_extra_and_drift(self) -> None:
        base = [
            {"scenario": "A", "stage": "a2", "path": "one.csv", "raw_sha256": "1", "normalized_sha256": "n1"},
            {"scenario": "A", "stage": "a2", "path": "missing.csv", "raw_sha256": "2", "normalized_sha256": "n2"},
        ]
        candidate = [
            {"scenario": "A", "stage": "a2", "path": "one.csv", "raw_sha256": "3", "normalized_sha256": "n3"},
            {"scenario": "A", "stage": "a2", "path": "extra.csv", "raw_sha256": "4", "normalized_sha256": "n4"},
        ]
        report = regression.compare_hash_records(base, candidate)
        self.assertFalse(report["ok"])
        self.assertEqual(report["missing"], [("", "A", "a2", "missing.csv")])
        self.assertEqual(report["extra"], [("", "A", "a2", "extra.csv")])
        self.assertEqual(report["normalized_drift"], [("", "A", "a2", "one.csv")])

    def test_porcelain_parser_preserves_first_path_character(self) -> None:
        status = " M cleanroom_tests/cleanroom_check.py\n?? tests/\n"
        self.assertEqual(
            regression.parse_porcelain_paths(status),
            ["cleanroom_tests/cleanroom_check.py", "tests/"],
        )

    def test_pre_edit_backups_are_excluded_from_exact_comparison(self) -> None:
        self.assertTrue(regression.excluded_from_exact_comparison("A-O_Parametrization_PRE_BAND_backup.xlsx"))
        self.assertFalse(regression.excluded_from_exact_comparison("A-O_Parametrization.xlsx"))


class XlsxNormalizationTests(unittest.TestCase):
    def test_restrictions_source_timestamp_is_normalized(self) -> None:
        workbook = (
            b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            b'<sheets><sheet name="Restrictions" sheetId="1" r:id="rId1"/></sheets></workbook>'
        )
        rels = (
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'
        )
        sheet = (
            b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
            b'<row r="1"><c r="A1" t="inlineStr"><is><t>source_run_timestamp</t></is></c></row>'
            b'<row r="2"><c r="A2" t="inlineStr"><is><t>2026-01-01T01:02:03</t></is></c></row>'
            b'</sheetData></worksheet>'
        )
        members = {
            "xl/workbook.xml": workbook,
            "xl/_rels/workbook.xml.rels": rels,
            "xl/worksheets/sheet1.xml": sheet,
        }
        normalized = regression.normalize_xlsx_members(members)
        self.assertIn(b"NORMALIZED", normalized["xl/worksheets/sheet1.xml"])
        self.assertNotIn(b"2026-01-01", normalized["xl/worksheets/sheet1.xml"])


class InventoryValidationTests(unittest.TestCase):
    def test_duplicate_inventory_names_are_rejected(self) -> None:
        with self.assertRaisesRegex(regression.RegressionError, "duplicate"):
            regression.load_scenarios(FIXTURES / "duplicate_scenarios.yaml")


if __name__ == "__main__":
    unittest.main()
