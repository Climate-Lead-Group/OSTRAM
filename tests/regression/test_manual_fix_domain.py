"""Exercise the real A3 manual-fix stage, including missing in-domain rows."""
from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook

from ostram.paths import PROFILE_AUTHORITIES_ENV
from ostram.pipeline.scenarios.transformations.patch_ao_c2a import load_taxonomy


ROOT = Path(__file__).resolve().parents[2]
TRAINING_TAXONOMY = ROOT / "examples/unescap/config/scenarios/technology_types.csv"


class ManualFixDomainTests(unittest.TestCase):
    def run_stage(self, base_copies: int) -> subprocess.CompletedProcess[str]:
        with tempfile.TemporaryDirectory() as tmp:
            stage = Path(tmp)
            source = stage / "wvaligned_outputs"
            source.mkdir()
            tech = "PWRNGSBGDXX"
            fixtures = {
                "A-O_Parametrization": (
                    "Secondary Techs", ["Tech", "Tech.Name"], [[tech, "old name"]]),
                "A-O_Demand": (
                    "Demand_Projection", ["Fuel/Tech", 2023, 2050],
                    [["ELCBGDXX03", 1, 1]]),
                "A-O_AR_Model_Base_Year": (
                    "Secondary",
                    ["Tech", "Tech.Name", "Mode.Operation", "Fuel.I", "Fuel.I.Name"],
                    [[tech, "old name", 1, "GASIND", "old fuel"]] * base_copies),
                "A-O_AR_Projections": (
                    "Secondary",
                    ["Tech", "Tech.Name", "Mode.Operation", "Direction", "Fuel",
                     "Fuel.Name", 2023, 2050],
                    [[tech, "old name", 1, "Input", "GASIND", "old fuel", 2.33765, 2.33765]]),
            }
            for stem, (sheet, header, rows) in fixtures.items():
                wb = Workbook()
                ws = wb.active
                ws.title = sheet
                ws.append(header)
                for row in rows:
                    ws.append(row)
                wb.save(source / f"{stem}_wvaligned.xlsx")
                wb.close()
            env = {key: value for key, value in os.environ.items()
                   if not key.startswith("OSTRAM_")}
            env.update(
                OSTRAM_STAGE_WORKDIR=str(stage),
                OSTRAM_PROJECT_ROOT=str(ROOT),
                OSTRAM_WORKSPACE=str(stage / "workspace"),
                PYTHONUTF8="1",
                PYTHONDONTWRITEBYTECODE="1",
            )
            env[PROFILE_AUTHORITIES_ENV] = json.dumps(
                {"interconnector_taxonomy": str(TRAINING_TAXONOMY)})
            return subprocess.run(
                [sys.executable, "-B", "-m",
                 "ostram.pipeline.scenarios.transformations.apply_manual_fixes"],
                cwd=ROOT, env=env, capture_output=True, text=True, encoding="utf-8",
                timeout=60,
            )

    def test_reduced_domain_retains_all_four_bangladesh_substitutions(self) -> None:
        result = self.run_stage(1)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("Substitutions applied: 4 / 4", result.stdout)
        self.assertEqual(result.stdout.count("OUT_OF_DOMAIN  "), 16)
        self.assertIn("ALL TESTS PASSED", result.stdout)

    def test_missing_or_duplicate_in_domain_target_fails_closed(self) -> None:
        for copies in (0, 2):
            with self.subTest(copies=copies):
                result = self.run_stage(copies)
                self.assertNotEqual(result.returncode, 0, result.stdout)
                self.assertIn("FAIL  sub locatable: AR_Base/Secondary PWRNGSBGDXX", result.stdout)
                self.assertNotIn("OUT_OF_DOMAIN  AR_Base/Secondary PWRNGSBGDXX", result.stdout)

    def test_full_taxonomy_retains_every_existing_substitution_target(self) -> None:
        full = load_taxonomy(ROOT / "config/scenarios/technology_types.csv")
        self.assertTrue({"PWRNGSBGDXX", "PWRGEOINDNO", "PWRNGSMDVXX",
                         "PWROILMDVXX", "PWROILNPLXX"}.issubset(full))
        self.assertEqual(
            {code for code in full if code.startswith("PWRSHP")},
            {"PWRSHPINDEA", "PWRSHPINDNE", "PWRSHPINDNO", "PWRSHPINDSO", "PWRSHPINDWE"},
        )
        training = load_taxonomy(TRAINING_TAXONOMY)
        self.assertEqual({code for code in training if code.startswith("PWRSHP")},
                         {"PWRSHPINDEA"})


if __name__ == "__main__":
    unittest.main()
