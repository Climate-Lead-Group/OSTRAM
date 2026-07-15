from __future__ import annotations

import ast
import csv
import unittest
from collections import Counter
from pathlib import Path


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
ARCHIVE_ROOT = REPO_ROOT / "docs" / "archive"
LEGACY_TOOLS = ARCHIVE_ROOT / "legacy-tools"
FIXTURES = TEST_ROOT / "fixtures"


class ArchivedEntrypointTests(unittest.TestCase):
    def test_obsolete_legacy_tools_are_archived_and_parse(self) -> None:
        expected = {
            REPO_ROOT / "t1_confection" / "concat_all_scenarios.py":
                LEGACY_TOOLS / "concat_all_scenarios_merge.py",
            REPO_ROOT / "t1_confection" / "Z_AUX_united_regions.py":
                LEGACY_TOOLS / "Z_AUX_united_regions.py",
        }
        for former, archived in expected.items():
            self.assertFalse(former.exists(), former)
            self.assertTrue(archived.is_file(), archived)
            ast.parse(archived.read_text(encoding="utf-8-sig"), filename=str(archived))

    def test_obsolete_legacy_tools_have_no_production_import(self) -> None:
        obsolete_modules = {
            "concat_all_scenarios",
            "set_final_v18_interconnector_values",
            "Z_AUX_fix_excel_profiles",
            "Z_AUX_united_regions",
        }
        for path in REPO_ROOT.rglob("*.py"):
            if ARCHIVE_ROOT in path.parents:
                continue
            tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
            imported = set()
            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    imported.update(alias.name.split(".")[0] for alias in node.names)
                elif isinstance(node, ast.ImportFrom) and node.module:
                    imported.add(node.module.split(".")[0])
            self.assertTrue(obsolete_modules.isdisjoint(imported), path)

    def test_legacy_merge_fixture_demonstrates_row_multiplication(self) -> None:
        def read_rows(name: str) -> list[dict[str, str]]:
            with (FIXTURES / name).open(encoding="utf-8", newline="") as handle:
                return list(csv.DictReader(handle))

        left = read_rows("legacy_concat_input.csv")
        right = read_rows("legacy_concat_output.csv")
        common = tuple(column for column in left[0] if column in right[0])
        left_counts = Counter(tuple(row[column] for column in common) for row in left)
        right_counts = Counter(tuple(row[column] for column in common) for row in right)
        legacy_merge_rows = sum(
            left_counts[key] * right_counts[key]
            for key in left_counts.keys() & right_counts.keys()
        )

        self.assertEqual(legacy_merge_rows, 6)
        self.assertEqual(len(left) + len(right), 5)
        self.assertGreater(legacy_merge_rows, len(left) + len(right))

        archived = (LEGACY_TOOLS / "concat_all_scenarios_merge.py").read_text(
            encoding="utf-8-sig"
        )
        maintained = (REPO_ROOT / "t1_confection" / "concat_all_scenarios_2.py").read_text(
            encoding="utf-8-sig"
        )
        self.assertIn("pd.merge", archived)
        self.assertIn("pd.concat", maintained)

    def test_stale_workbook_writer_is_archived_and_fails_closed(self) -> None:
        archived_path = LEGACY_TOOLS / "Z_AUX_fix_excel_profiles.py"
        stub_path = REPO_ROOT / "t1_confection" / "Z_AUX_fix_excel_profiles.py"
        archived = archived_path.read_text(encoding="utf-8-sig")
        stub = stub_path.read_text(encoding="utf-8-sig")

        ast.parse(archived, filename=str(archived_path))
        ast.parse(stub, filename=str(stub_path))
        self.assertIn("import openpyxl", archived)
        self.assertIn("wb.save(excel_path)", archived)
        self.assertIn("return 2", stub)
        self.assertIn("disabled", stub)
        self.assertNotIn("openpyxl", stub)
        self.assertNotIn("load_workbook", stub)
        self.assertNotIn("shutil", stub)
        self.assertNotIn(".save(", stub)

    def test_ws3_template_writer_is_archived_and_fails_closed(self) -> None:
        archived_path = (
            ARCHIVE_ROOT / "ws3-ws4" / "scripts" / "set_final_v18_interconnector_values.py"
        )
        stub_path = (
            REPO_ROOT / "ws3_transmission_audit" / "set_final_v18_interconnector_values.py"
        )
        archived = archived_path.read_text(encoding="utf-8-sig")
        stub = stub_path.read_text(encoding="utf-8-sig")

        ast.parse(archived, filename=str(archived_path))
        ast.parse(stub, filename=str(stub_path))
        self.assertIn("OSTRAM_ws3_workcopy", archived)
        self.assertIn("shutil.copy", archived)
        self.assertIn("wb.save(V18)", archived)
        self.assertIn("return 2", stub)
        self.assertIn("disabled", stub)
        self.assertNotIn("OSTRAM_ws3_workcopy", stub)
        self.assertNotIn("openpyxl", stub)
        self.assertNotIn("shutil", stub)
        self.assertNotIn(".save(", stub)

        retained_audits = (
            "audit_transmission_values.py",
            "compute_internal_tx_residuals.py",
            "verify_base_consistency.py",
        )
        for name in retained_audits:
            self.assertTrue((REPO_ROOT / "ws3_transmission_audit" / name).is_file(), name)


if __name__ == "__main__":
    unittest.main()
