from __future__ import annotations

import json
import os
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest import mock

from openpyxl import Workbook

from ostram.pipeline.preparation import configuration as country_config
from ostram.pipeline.scenarios.materializer import (
    MaterializationPaths,
    REQUIRED_AO_FILES,
    materialize_scenarios,
)
from ostram.pipeline.scenarios.registry import (
    CANONICAL_SCENARIOS,
    DECISION_SCENARIOS,
    ROOT_SCENARIOS,
    ensure_root_output_directories,
    load_registry,
)
from ostram.pipeline.scenarios import apply_patches


REPO_ROOT = Path(__file__).resolve().parents[2]
PREPARATION_WORKSPACE = REPO_ROOT / "workspace" / "preparation"


class ScenarioRegistryTests(unittest.TestCase):
    def test_country_yaml_loader_is_script_anchored_sorted_and_cached(self) -> None:
        original_cache = country_config._cached_config
        self.addCleanup(
            setattr,
            country_config,
            "_cached_config",
            original_cache,
        )
        with tempfile.TemporaryDirectory() as temp:
            config_path = Path(temp) / "Config_country_codes.yaml"
            config_path.write_text(
                "country_data:\n"
                "  ZZZ: {english_name: Zed, ostram_name: Zeta}\n"
                "  AAA: {english_name: Aye, ostram_name: Alfa}\n"
                "countries: [ZZZ, AAA]\n",
                encoding="utf-8",
            )
            with mock.patch.object(country_config, "CONFIG_PATH", config_path):
                country_config._cached_config = None
                self.assertEqual(country_config.get_countries(), ["AAA", "ZZZ"])
                self.assertEqual(
                    country_config.get_model_countries_list(),
                    ["ZZZ", "AAA"],
                )

                config_path.write_text("country_data: {BBB: {}}\n", encoding="utf-8")
                self.assertEqual(country_config.get_countries(), ["AAA", "ZZZ"])
                country_config._cached_config = None
                self.assertEqual(country_config.get_countries(), ["BBB"])

        country_config._cached_config = None
        self.assertEqual(
            country_config.CONFIG_PATH,
            REPO_ROOT / "config" / "preparation" / "Config_country_codes.yaml",
        )

    def test_exact_roots_decision_order_bases_and_overlays(self) -> None:
        registry = load_registry()
        self.assertEqual(registry.root_names, ROOT_SCENARIOS)
        self.assertEqual(registry.decision_scenarios, DECISION_SCENARIOS)
        self.assertEqual(registry.scenario_names, CANONICAL_SCENARIOS)

        derived = registry.derived_by_name
        self.assertEqual(
            derived["A_Calibrated_BAU_Clipped"].base_scenario,
            "A_Calibrated_BAU",
        )
        self.assertEqual(
            derived["C_Target_VRE_Clipped"].base_scenario,
            "C_Target_VRE",
        )
        b_derived = [
            scenario
            for scenario in registry.derived
            if scenario.name not in {
                "A_Calibrated_BAU_Clipped",
                "C_Target_VRE_Clipped",
            }
        ]
        self.assertTrue(
            all(
                scenario.base_scenario == "B_Optimised_VRE"
                for scenario in b_derived
            )
        )
        overlays = {
            scenario.name
            for scenario in registry.derived
            if scenario.direction_overlay is not None
        }
        self.assertEqual(
            overlays,
            {"B_Opt_DirBidir", "B_Opt_DirContractual"},
        )

    def test_selection_is_canonical_and_fail_closed(self) -> None:
        registry = load_registry()
        self.assertEqual(
            registry.select(
                "C_Target_VRE_Clipped,A_Calibrated_BAU,B_Opt_TradeCap15"
            ),
            (
                "A_Calibrated_BAU",
                "B_Opt_TradeCap15",
                "C_Target_VRE_Clipped",
            ),
        )
        with self.assertRaisesRegex(ValueError, "duplicate"):
            registry.select("BAU,BAU")
        with self.assertRaisesRegex(ValueError, "unknown"):
            registry.select("Not_A_Scenario")

    def test_required_roots_and_c_on_a_result_dependency(self) -> None:
        registry = load_registry()
        selected = registry.select(
            "A_Calibrated_BAU_Clipped,B_Opt_TradeCap15,"
            "C_Target_VRE_Clipped"
        )
        self.assertEqual(
            registry.required_roots(selected),
            (
                "A_Calibrated_BAU",
                "B_Optimised_VRE",
                "C_Target_VRE",
            ),
        )
        with tempfile.TemporaryDirectory() as temp:
            seed = Path(temp) / "a_output.csv"
            seed.write_text("fixture\n", encoding="utf-8")
            environment = {
                "OSTRAM_A_CALIBRATED_BAU_RESULT": str(seed),
            }
            dependencies = registry.result_dependencies(
                ("C_Target_VRE",),
                execution_workspace=Path(temp),
                environment=environment,
            )
            self.assertEqual(
                dependencies["OSTRAM_A_CALIBRATED_BAU_RESULT"],
                seed.resolve(),
            )

    def test_a1_preparation_creates_only_four_root_directories(self) -> None:
        registry = load_registry()
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "A1_Outputs"
            ensure_root_output_directories(output, registry)
            self.assertEqual(
                sorted(path.name for path in output.iterdir()),
                sorted(f"A1_Outputs_{name}" for name in ROOT_SCENARIOS),
            )

    def test_patcher_uses_declared_non_b_base_and_rejects_conflict(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            outputs = root / "A1_Outputs"
            configs = root / "configs"
            source = outputs / "A1_Outputs_A_Calibrated_BAU"
            source.mkdir(parents=True)
            workbook = Workbook()
            workbook.active.title = "fixture"
            workbook.save(source / apply_patches.PARAM_FILE)
            patch_dir = configs / "A_Calibrated_BAU_Clipped"
            patch_dir.mkdir(parents=True)
            (patch_dir / "patches.json").write_text(
                json.dumps(
                    {
                        "scenario": "A_Calibrated_BAU_Clipped",
                        "base_scenario": "A_Calibrated_BAU",
                        "apply_vre_ceiling_layer": False,
                        "edits": [],
                    }
                ),
                encoding="utf-8",
            )

            log = apply_patches.build_scenario(
                "A_Calibrated_BAU_Clipped",
                a1_outputs=outputs,
                configs=configs,
                authority_path=root / "unused.xlsx",
            )
            self.assertEqual(log["source"], "A_Calibrated_BAU")
            self.assertTrue(
                (
                    outputs
                    / "A1_Outputs_A_Calibrated_BAU_Clipped"
                    / apply_patches.PARAM_FILE
                ).is_file()
            )
            with self.assertRaisesRegex(ValueError, "conflicts"):
                apply_patches.build_scenario(
                    "A_Calibrated_BAU_Clipped",
                    source="B_Optimised_VRE",
                    a1_outputs=outputs,
                    configs=configs,
                    authority_path=root / "unused.xlsx",
                )

    def test_materializer_routes_one_exact_root_selection(self) -> None:
        registry = load_registry()
        with tempfile.TemporaryDirectory() as temp:
            temp_root = Path(temp)
            outputs = temp_root / "A1_Outputs"
            snapshot = outputs / "_post_a2_snapshot_BAU"
            snapshot.mkdir(parents=True)
            for filename in REQUIRED_AO_FILES:
                (snapshot / filename).write_bytes(b"fixture")

            paths = MaterializationPaths(
                preparation_workspace=PREPARATION_WORKSPACE,
                a1_outputs=outputs,
                a3_entrypoint=(
                    REPO_ROOT / "ostram" / "pipeline" / "scenarios" / "transform.py"
                ),
                a3_process=(
                    REPO_ROOT / "ostram" / "pipeline" / "scenarios" / "transformations"
                ),
                soasia=REPO_ROOT / "inputs" / "scenarios" / "OSTRAM_Scenario_Inputs.xlsx",
            )

            def materialize_root(
                root: str,
                environment: dict[str, str],
            ) -> None:
                del environment
                shutil.copytree(
                    outputs / f"_post_a2_snapshot_{root}",
                    outputs / f"A1_Outputs_{root}",
                )

            record = materialize_scenarios(
                ("BAU",),
                paths=paths,
                registry=registry,
                environment=dict(os.environ),
                root_materializer=materialize_root,
                direction_applier=lambda *_: {},
            )
            self.assertEqual(record["selected_scenarios"], ["BAU"])
            self.assertEqual(record["required_roots"], ["BAU"])
            self.assertEqual(record["derived"], [])
            self.assertFalse(record["solver_invoked"])


if __name__ == "__main__":
    unittest.main()
