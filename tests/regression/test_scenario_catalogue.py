"""The scenario catalogue page matches south_asia.yaml and the scenario registry."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import re
import unittest

import yaml

from tests.regression.setup_file import SETUP_FILE


REPO_ROOT = Path(__file__).resolve().parents[2]
GENERATOR = REPO_ROOT / "tools" / "generate_scenario_catalogue.py"
REGENERATE = "python tools/generate_scenario_catalogue.py"


def _generator():
    spec = importlib.util.spec_from_file_location("generate_scenario_catalogue", GENERATOR)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _registered_scenarios() -> list[str]:
    profile = yaml.safe_load(
        SETUP_FILE.read_text(encoding="utf-8")
    )
    registry = profile["scenario_registry"]
    return list(registry["support_scenarios"]) + list(registry["decision_scenarios"])


class ScenarioCatalogueTests(unittest.TestCase):
    def test_generated_block_is_current(self) -> None:
        generator = _generator()
        page = generator.PAGE.read_text(encoding="utf-8")
        self.assertEqual(
            generator.current_block(page),
            generator.generated_block(),
            f"The generated block of {generator.PAGE.relative_to(REPO_ROOT).as_posix()} "
            f"is out of date. Run `{REGENERATE}` from the repository root and "
            "commit the page.",
        )

    def test_origin_table_lists_exactly_the_registered_scenarios(self) -> None:
        generator = _generator()
        page = generator.PAGE.read_text(encoding="utf-8")
        origin = page.split("## Origin of each scenario", 1)[1]
        listed = re.findall(r"^\| `([A-Za-z0-9_]+)` \|", origin, flags=re.MULTILINE)
        self.assertEqual(
            sorted(listed),
            sorted(_registered_scenarios()),
            "The origin table of the scenario catalogue must have one row per "
            "registered scenario. Add or remove the rows by hand, then run "
            f"`{REGENERATE}` from the repository root for the generated block.",
        )
        self.assertEqual(len(listed), len(set(listed)), "duplicate rows in the origin table")


if __name__ == "__main__":
    unittest.main()
