"""Late A3 technology emissions survive without importing its seed engineering data."""
from __future__ import annotations

from contextlib import redirect_stdout
import io
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import pandas as pd

from ostram.pipeline.preparation import base_inputs


ROOT = Path(__file__).resolve().parents[2]
TRAINING = ROOT / "examples/unescap"


class ExtensionEmissionTests(unittest.TestCase):
    def filter(self, preserve: bool):
        paths = SimpleNamespace(
            ao_decisions=TRAINING / "config/scenarios/ao_extension_decisions.csv",
            interconnector_taxonomy=TRAINING / "config/scenarios/technology_types.csv",
        )
        technology = ["PWRNGSBGDXX", "PWRCCGBGDXX01", "PWRNGSMDVXX", "PWRNGSINDEA"]
        data = {
            "EmissionActivityRatio": pd.DataFrame({
                "TECHNOLOGY": technology,
                "VALUE": [0.13113936, 0.1, 0.1, 0.13075],
            }),
            "InputActivityRatio": pd.DataFrame({
                "TECHNOLOGY": technology, "VALUE": [99.0, 99.0, 99.0, 2.59925],
            }),
            "CapitalCost": pd.DataFrame({
                "TECHNOLOGY": technology, "VALUE": [99.0] * 4,
            }),
        }
        matrix = dict(enabled=True, matrix_filtering_enabled=True,
                      matrix={"NGS": {"BGD": False, "MDV": False, "IND": True}})
        with patch.object(base_inputs, "_PROJECT_PATHS", paths), \
                patch.object(base_inputs, "profile_policy", return_value=preserve), \
                redirect_stdout(io.StringIO()):
            return base_inputs.filter_by_tech_country_matrix(data, matrix)

    def test_declared_training_emissions_survive_without_seed_activity_or_costs(self):
        filtered = self.filter(True)
        self.assertEqual(filtered["EmissionActivityRatio"].TECHNOLOGY.tolist(),
                         ["PWRNGSBGDXX", "PWRNGSINDEA"])
        self.assertEqual(filtered["EmissionActivityRatio"].VALUE.iloc[0], 0.13113936)
        for parameter in ("InputActivityRatio", "CapitalCost"):
            self.assertEqual(filtered[parameter].TECHNOLOGY.tolist(), ["PWRNGSINDEA"])

    def test_default_full_profile_keeps_existing_filter_semantics(self):
        for frame in self.filter(False).values():
            self.assertEqual(frame.TECHNOLOGY.tolist(), ["PWRNGSINDEA"])


if __name__ == "__main__":
    unittest.main()
