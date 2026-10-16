"""Every file under config/ is expected, and every scenario entry is registered."""

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
    "config/scenarios/B_Opt_DirBidir/set_interconnector_direction.yaml",
    "config/scenarios/B_Opt_DirContractual/set_interconnector_direction.yaml",
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
    "config/scenarios/rules/internal_tx_residuals.csv",
    "config/scenarios/rules/pwr_min_2023_2026_pin.csv",
    "config/scenarios/technology_types.csv",
)


class ConfigFileTests(unittest.TestCase):
    def test_every_file_under_config_is_on_the_list(self) -> None:
        on_disk = sorted(
            path.relative_to(REPO_ROOT).as_posix()
            for path in CONFIG.rglob("*")
            if path.is_file()
        )
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
        registered |= {entry["name"] for entry in registry["derived_scenarios"]}
        entries = [
            ("south_asia.yaml scenario_inputs block {}", name)
            for name in profile["scenario_inputs"]
        ]
        entries += [
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
