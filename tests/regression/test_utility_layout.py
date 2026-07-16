from __future__ import annotations

import ast
import unittest
from pathlib import Path


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
ANALYSIS_ROOT = REPO_ROOT / "tools" / "analysis"
LEGACY_ROOT = REPO_ROOT / "t1_confection"
UTILITY_PATHS = (
    ("tools/analysis/check_combined.py", "t1_confection/check_combined.py"),
    (
        "tools/analysis/ostram_scenario_analysis.py",
        "t1_confection/ostram_scenario_analysis.py",
    ),
    ("tools/analysis/ostram_trn_plotter.py", "t1_confection/ostram_trn_plotter.py"),
    ("tools/analysis/slice_by_country.py", "t1_confection/slice_by_country.py"),
    ("tools/analysis/analyse_sensitivity.py", "t1_confection/analyse_sensitivity.py"),
    (
        "tools/analysis/concat_all_scenarios.py",
        "t1_confection/concat_all_scenarios_2.py",
    ),
    ("tools/analysis/reproduce_A1_A6.py", "t1_confection/reproduce_A1_A6.py"),
    (
        "tools/analysis/visualization/Z_AUX_generate_interactive_dashboards_aggregated.py",
        "t1_confection/Z_AUX_generate_interactive_dashboards_aggregated.py",
    ),
    (
        "tools/analysis/visualization/Z_AUX_generate_RES_diagram.py",
        "t1_confection/Z_AUX_generate_RES_diagram.py",
    ),
    (
        "tools/analysis/visualization/Z_AUX_generate_transmission_maps.py",
        "t1_confection/Z_AUX_generate_transmission_maps.py",
    ),
    (
        "tools/analysis/visualization/Z_AUX_interconnections_dashboard.py",
        "t1_confection/Z_AUX_interconnections_dashboard.py",
    ),
)
MOVED_OLD_NAMES = tuple(Path(wrapper).name for _, wrapper in UTILITY_PATHS)
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
LEGACY_RUNNER_NAMES = ("run_directions.bat", "run_sensitivities.bat")
ROOT_LEGACY_RUNNER_NAMES = (
    "run_baselines.bat",
    "run_directions.bat",
    "run_sensitivities.bat",
)


class UtilityLayoutTests(unittest.TestCase):
    def test_targets_and_compatibility_wrappers_exist_and_parse(self) -> None:
        for target_relative, wrapper_relative in UTILITY_PATHS:
            target = REPO_ROOT / target_relative
            wrapper = REPO_ROOT / wrapper_relative
            self.assertTrue(target.is_file(), target)
            self.assertTrue(wrapper.is_file(), wrapper)
            ast.parse(target.read_text(encoding="utf-8-sig"), filename=str(target))
            ast.parse(wrapper.read_text(encoding="utf-8"), filename=str(wrapper))

    def test_wrappers_delegate_to_the_matching_analysis_target(self) -> None:
        for target_relative, wrapper_relative in UTILITY_PATHS:
            source = (REPO_ROOT / wrapper_relative).read_text(encoding="utf-8")
            target = Path(target_relative)
            self.assertIn('"tools"', source)
            self.assertIn('"analysis"', source)
            if "visualization" in target.parts:
                self.assertIn('"visualization"', source)
            self.assertIn(f'"{target.name}"', source)
            self.assertIn("runpy.run_path", source)

    def test_core_entry_points_do_not_reference_moved_utilities(self) -> None:
        for relative in CORE_ENTRY_POINTS:
            source = (REPO_ROOT / relative).read_text(encoding="utf-8-sig")
            for name in MOVED_OLD_NAMES:
                self.assertNotIn(name, source, f"{relative} references {name}")

    def test_obsolete_nested_runners_are_archived_and_fail_closed(self) -> None:
        archive = REPO_ROOT / "docs" / "archive" / "legacy-runners" / "t1_confection"
        for name in LEGACY_RUNNER_NAMES:
            archived = (archive / name).read_text(encoding="utf-8-sig").lower()
            stub = (LEGACY_ROOT / name).read_text(encoding="utf-8-sig").lower()
            self.assertIn("ostram_clean", archived)
            self.assertIn("exit /b 2", stub)
            self.assertIn("obsolete runner is disabled", stub)
            self.assertNotIn("b2_executing_og_model.py", stub)
            self.assertNotIn("python ", stub)

    def test_obsolete_root_runners_are_archived_and_fail_closed(self) -> None:
        archive = REPO_ROOT / "docs" / "archive" / "legacy-runners" / "root"
        for name in ROOT_LEGACY_RUNNER_NAMES:
            archived = (archive / name).read_text(encoding="utf-8-sig").lower()
            stub = (REPO_ROOT / name).read_text(encoding="utf-8-sig").lower()
            self.assertIn("b2_executing_og_model.py", archived)
            self.assertIn("exit /b 2", stub)
            self.assertIn("obsolete root runner is disabled", stub)
            self.assertNotIn("b2_executing_og_model.py", stub)
            self.assertNotIn("python ", stub)


if __name__ == "__main__":
    unittest.main()
