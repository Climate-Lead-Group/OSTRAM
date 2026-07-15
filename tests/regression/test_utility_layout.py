from __future__ import annotations

import ast
import unittest
from pathlib import Path


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
ANALYSIS_ROOT = REPO_ROOT / "tools" / "analysis"
LEGACY_ROOT = REPO_ROOT / "t1_confection"
UTILITY_NAMES = (
    "check_combined.py",
    "ostram_scenario_analysis.py",
    "ostram_trn_plotter.py",
    "slice_by_country.py",
)
CORE_ENTRY_POINTS = (
    "run.py",
    "t1_confection/A0_generate_tech_country_matrix.py",
    "t1_confection/A1_Pre_processing_OG_csvs.py",
    "t1_confection/A2_AddTx.py",
    "t1_confection/A3_process.py",
    "t1_confection/B1_Run_Compiler.py",
    "t1_confection/B1_Compiler.py",
    "t1_confection/B2_Executing_OG_Model.py",
)


class UtilityLayoutTests(unittest.TestCase):
    def test_targets_and_compatibility_wrappers_exist_and_parse(self) -> None:
        for name in UTILITY_NAMES:
            target = ANALYSIS_ROOT / name
            wrapper = LEGACY_ROOT / name
            self.assertTrue(target.is_file(), target)
            self.assertTrue(wrapper.is_file(), wrapper)
            ast.parse(target.read_text(encoding="utf-8-sig"), filename=str(target))
            ast.parse(wrapper.read_text(encoding="utf-8"), filename=str(wrapper))

    def test_wrappers_delegate_to_the_matching_analysis_target(self) -> None:
        for name in UTILITY_NAMES:
            source = (LEGACY_ROOT / name).read_text(encoding="utf-8")
            self.assertIn('"tools" / "analysis"', source)
            self.assertIn(f'"{name}"', source)
            self.assertIn("runpy.run_path", source)

    def test_core_entry_points_do_not_reference_moved_utilities(self) -> None:
        for relative in CORE_ENTRY_POINTS:
            source = (REPO_ROOT / relative).read_text(encoding="utf-8-sig")
            for name in UTILITY_NAMES:
                self.assertNotIn(name, source, f"{relative} references {name}")


if __name__ == "__main__":
    unittest.main()
