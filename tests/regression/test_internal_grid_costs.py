"""The internal grid costs are computed in every run from their sourced inputs.

The shared cost is ESMAP's South Asia 400 kV line brought to 2023 with the
World Bank MUV index; renewable plants add the Green Energy Corridor Phase II
average. No derived value is typed in the configuration or the rule code.
"""
from __future__ import annotations

import inspect
import unittest
from pathlib import Path

import yaml

from ostram.pipeline.scenarios.rules import apply_internal_transmission as itx


REPO_ROOT = Path(__file__).resolve().parents[2]
CONFIG = REPO_ROOT / "config" / "preparation" / "Config_country_codes.yaml"
DERIVED = ("base_capital_cost", "base_fixed_cost", "re_add_on", "re_capex_multiplier")


class InternalGridCostTests(unittest.TestCase):
    def test_tracked_inputs_give_the_expected_costs(self) -> None:
        knobs = itx.read_config(CONFIG)
        self.assertAlmostEqual(knobs["base_capital_cost"], 109.536, delta=1e-6)
        self.assertAlmostEqual(knobs["base_fixed_cost"], 3.28608, delta=1e-6)
        self.assertAlmostEqual(knobs["re_add_on"], 76.517128202, delta=1e-6)
        self.assertAlmostEqual(knobs["re_capex_multiplier"], 1.698556896, delta=1e-6)
        self.assertAlmostEqual(knobs["base_capital_cost"] * knobs["re_capex_multiplier"], 186.053128202, delta=1e-6)
        self.assertAlmostEqual(knobs["base_fixed_cost"] * knobs["re_capex_multiplier"], 5.581593846, delta=1e-6)

    def test_rule_writes_the_computed_costs(self) -> None:
        knobs = itx.read_config(CONFIG)
        values = itx.build_target_values(knobs, {}, ["INDNO"], parts=("costs",))
        self.assertAlmostEqual(values["PWRTRNINDNO"]["CapitalCost"], 109.536, delta=1e-6)
        self.assertAlmostEqual(values["RNWTRNINDNO"]["CapitalCost"], 186.053128202, delta=1e-6)
        self.assertAlmostEqual(values["TRNNLIINDNO"]["FixedCost"], 3.28608, delta=1e-6)
        self.assertAlmostEqual(values["RNWRPOINDNO"]["FixedCost"], 5.581593846, delta=1e-6)

    def test_no_derived_value_is_typed(self) -> None:
        block = yaml.safe_load(CONFIG.read_text(encoding="utf-8"))[itx.CONFIG_BLOCK]
        for key in DERIVED:
            self.assertNotIn(key, block)
            self.assertNotIn(key, itx.DEFAULTS)
        self.assertEqual(set(itx.COST_INPUTS) - set(block), set())
        source = inspect.getsource(itx)
        for typed in ("131.15", "3.93", "109.536", "3.28608", "76.517", "1.6985", "186.053", "5.5815"):
            self.assertNotIn(typed, source)
            self.assertNotIn(typed, CONFIG.read_text(encoding="utf-8").split("internal_transmission:")[1]
                             .split("enable_dsptrn")[0].replace("76.52 USD2023/kW", ""))

    def test_missing_input_stops_the_rule(self) -> None:
        block = yaml.safe_load(CONFIG.read_text(encoding="utf-8"))[itx.CONFIG_BLOCK]
        block.pop("muv_2022")
        with self.assertRaises(ValueError) as caught:
            itx.internal_grid_costs(block)
        self.assertIn("muv_2022", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
