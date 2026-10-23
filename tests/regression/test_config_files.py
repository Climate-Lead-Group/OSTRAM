"""Every file under config/ is expected, and every scenario entry is registered.

The list holds the delivered files. The folder of a scenario registered later,
config/scenarios/<name>/, may hold any files.
"""

from __future__ import annotations

from pathlib import Path
import unittest

import yaml

from tests.regression.setup_file import SETUP_FILE


REPO_ROOT = Path(__file__).resolve().parents[2]
CONFIG = REPO_ROOT / "config"
SHARED_SCENARIO_FOLDERS = ("rules",)

# The files under config/, relative to the repository root, sorted.
EXPECTED_CONFIG_FILES = (
    "config/compilation/Config_MOMF_T1_A.yaml",
    "config/execution/Config_MOMF_T1_AB.yaml",
    "config/preparation/Config_country_codes.yaml",
    "config/preparation/Config_region_consolidation.yaml",
    SETUP_FILE.relative_to(REPO_ROOT).as_posix(),
    "config/setups/registry.json",
    "config/scenarios/A_Calibrated_BAU/bau_calibration.yaml",
    "config/scenarios/A_Calibrated_BAU/build_cap_rule.yaml",
    "config/scenarios/A_Calibrated_BAU/relax_interconnectors.yaml",
    "config/scenarios/A_Calibrated_BAU/retirement_schedule.yaml",
    "config/scenarios/A_Calibrated_BAU_Clipped/scenario.yaml",
    "config/scenarios/A_Calibrated_BAU_Clipped_TxFreeze2026/scenario.yaml",
    "config/scenarios/B_Opt_Clipped/scenario.yaml",
    "config/scenarios/B_Opt_Clipped_GrossImportCap15/scenario.yaml",
    "config/scenarios/B_Opt_DirBidir/scenario.yaml",
    "config/scenarios/B_Opt_DirBidir/set_interconnector_direction.yaml",
    "config/scenarios/B_Opt_DirContractual/scenario.yaml",
    "config/scenarios/B_Opt_DirContractual/set_interconnector_direction.yaml",
    "config/scenarios/B_Opt_IndiaCosts/scenario.yaml",
    "config/scenarios/B_Opt_IndiaCostsFuel/scenario.yaml",
    "config/scenarios/B_Opt_SolarCapex130/scenario.yaml",
    "config/scenarios/B_Opt_SolarCapexHi/scenario.yaml",
    "config/scenarios/B_Opt_SolarCapexSpike/scenario.yaml",
    "config/scenarios/B_Opt_TradeCap15/scenario.yaml",
    "config/scenarios/B_Opt_TxCap150/scenario.yaml",
    "config/scenarios/B_Optimised_VRE/build_cap_rule.yaml",
    "config/scenarios/B_Optimised_VRE/relax_interconnectors.yaml",
    "config/scenarios/B_Optimised_VRE/retirement_schedule.yaml",
    "config/scenarios/B_Optimised_VRE/storage_floors.yaml",
    "config/scenarios/C_Target_VRE/bau_calibration.yaml",
    "config/scenarios/C_Target_VRE/build_cap_rule.yaml",
    "config/scenarios/C_Target_VRE/relax_interconnectors.yaml",
    "config/scenarios/C_Target_VRE/retirement_schedule.yaml",
    "config/scenarios/C_Target_VRE/set_vre_targets.yaml",
    "config/scenarios/C_Target_VRE/storage_floors.yaml",
    "config/scenarios/C_Target_VRE_Clipped/scenario.yaml",
    "config/scenarios/rules/pwr_min_2023_2026_pin.csv",
    "config/scenarios/technology_types.csv",
)


def _added_scenario_folders() -> set[str]:
    """Return config/scenarios/<name>/ of every registered scenario with no listed file."""
    registry = yaml.safe_load(SETUP_FILE.read_text(encoding="utf-8"))["scenario_registry"]
    listed = {path.split("/")[2] for path in EXPECTED_CONFIG_FILES if path.startswith("config/scenarios/")}
    return {
        f"config/scenarios/{name}/"
        for name in registry["decision_scenarios"]
        if name not in listed
    }


class ConfigFileTests(unittest.TestCase):
    def test_every_file_under_config_is_on_the_list(self) -> None:
        added = _added_scenario_folders()
        on_disk = sorted(
            path.relative_to(REPO_ROOT).as_posix()
            for path in CONFIG.rglob("*")
            if path.is_file()
        )
        on_disk = [
            path for path in on_disk
            if not any(path.startswith(folder) for folder in added)
        ]
        problems = [
            f"{path} is not on EXPECTED_CONFIG_FILES. Add it to the list in "
            "this test, or delete it."
            for path in on_disk
            if path not in EXPECTED_CONFIG_FILES
        ]
        if problems:
            self.fail("\n".join(problems))

    def test_every_file_on_the_list_exists(self) -> None:
        problems = [
            f"{path} is on EXPECTED_CONFIG_FILES but missing. Restore it, or "
            "delete it from the list in this test."
            for path in EXPECTED_CONFIG_FILES
            if not (REPO_ROOT / path).is_file()
        ]
        if problems:
            self.fail("\n".join(problems))

    def test_every_scenario_entry_has_a_registered_scenario(self) -> None:
        profile = yaml.safe_load(
            SETUP_FILE.read_text(encoding="utf-8")
        )
        execution = yaml.safe_load(
            (CONFIG / "execution" / "Config_MOMF_T1_AB.yaml").read_text(encoding="utf-8")
        )
        registry = profile["scenario_registry"]
        registered = {entry["name"] for entry in registry["root_scenarios"]}
        registered |= set(registry["decision_scenarios"])
        entries = [
            ("folder config/scenarios/{}", folder.name)
            for folder in sorted((CONFIG / "scenarios").iterdir())
            if folder.is_dir()
            and folder.name not in SHARED_SCENARIO_FOLDERS
            and any(path.is_file() for path in folder.rglob("*"))
        ]
        entries += [
            ("south_asia.yaml policies.historical_trade_roots entry {}", name)
            for name in profile["policies"].get("historical_trade_roots", [])
        ]
        entries += [
            ("Config_MOMF_T1_AB.yaml cplex_numerical_emphasis_scenarios entry {}", name)
            for name in execution.get("cplex_numerical_emphasis_scenarios", [])
        ]
        problems = [
            f"{label.format(name)} has no registered scenario. Register the "
            "scenario in scenario_registry in config/setups/south_asia.yaml, or "
            "delete the entry."
            for label, name in entries
            if name not in registered
        ]
        if problems:
            self.fail("\n".join(problems))


if __name__ == "__main__":
    unittest.main()
