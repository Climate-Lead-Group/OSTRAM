"""The scenario catalogue page matches south_asia.yaml and the scenario registry.

The page documents the delivered scenarios, the rows of its origin table. A
scenario registered after them needs no change to the page, and the generator
leaves it out until its row is added to the table.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
import re
import tempfile
import unittest
from types import SimpleNamespace
from unittest import mock

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


def _origin_rows(page: str) -> list[str]:
    origin = page.split("## Origin of each scenario", 1)[1]
    return re.findall(r"^\| `([A-Za-z0-9_]+)` \|", origin, flags=re.MULTILINE)


class ScenarioCatalogueTests(unittest.TestCase):
    def test_generated_block_is_current(self) -> None:
        generator = _generator()
        page = generator.PAGE.read_text(encoding="utf-8")
        listed = set(_origin_rows(page))
        documented = [name for name in _registered_scenarios() if name in listed]
        self.assertEqual(
            generator.current_block(page),
            generator.generated_block(documented),
            f"The generated block of {generator.PAGE.relative_to(REPO_ROOT).as_posix()} "
            f"is out of date. Run `{REGENERATE}` from the repository root and "
            "commit the page.",
        )

    def test_origin_table_lists_the_first_registered_scenarios(self) -> None:
        generator = _generator()
        page = generator.PAGE.read_text(encoding="utf-8")
        listed = _origin_rows(page)
        self.assertEqual(
            sorted(listed),
            sorted(_registered_scenarios()[:len(listed)]),
            "The origin table of the scenario catalogue must have one row per "
            "delivered scenario, and the delivered scenarios must be the first "
            "registered ones. Add or remove the rows by hand, then run "
            f"`{REGENERATE}` from the repository root for the generated block.",
        )
        self.assertEqual(len(listed), len(set(listed)), "duplicate rows in the origin table")

    def test_generator_leaves_out_a_registered_scenario_without_a_row(self) -> None:
        generator = _generator()
        page = generator.PAGE.read_text(encoding="utf-8")
        real = generator.resolve_paths(project_root=REPO_ROOT)
        setup = yaml.safe_load(SETUP_FILE.read_text(encoding="utf-8"))
        setup["scenario_registry"]["decision_scenarios"].append("Z_Unlisted_Test_Scenario")
        with tempfile.TemporaryDirectory() as folder:
            copy = Path(folder) / "setup.yaml"
            copy.write_text(yaml.safe_dump(setup, sort_keys=False), encoding="utf-8")
            paths = SimpleNamespace(
                scenario_registry=copy, scenario_config_root=real.scenario_config_root
            )
            with mock.patch.object(generator, "resolve_paths", return_value=paths):
                block = generator.generated_block()
        self.assertNotIn("Z_Unlisted_Test_Scenario", block)
        self.assertEqual(generator.current_block(page), block)


if __name__ == "__main__":
    unittest.main()
