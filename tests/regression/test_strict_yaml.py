"""YAML files are read strictly: a repeated key or an unknown scenario key stops the run."""
from __future__ import annotations

import io
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest import mock

import yaml

from ostram import __main__ as cli
from ostram import yaml_io
from ostram.pipeline.scenarios.registry import (
    ScenarioFileError,
    load_registry,
    load_scenario_file,
)
from tests.regression.test_run_orchestration import LauncherHarness, _load_launcher


REPO_ROOT = Path(__file__).resolve().parents[2]
SCENARIO = "T_GrossImportCap20BGD"
HEAD = (
    f"scenario: {SCENARIO}\n"
    "base_scenario: B_Optimised_VRE\n"
    "description: Test scenario.\n"
    "apply_vre_ceiling_layer: true\n"
    "edits: []\n"
)
CAP_20_BGD = (
    "gross_import_cap:\n"
    "  countries:\n"
    "  - BGD\n"
    "  fraction: 0.20\n"
    "  first_year: 2027\n"
    "  last_year: 2050\n"
)
CAP_15_SIX = (
    "gross_import_cap:\n"
    "  countries:\n"
    "  - BGD\n"
    "  - BTN\n"
    "  - IND\n"
    "  - LKA\n"
    "  - MDV\n"
    "  - NPL\n"
    "  fraction: 0.15\n"
    "  first_year: 2027\n"
    "  last_year: 2050\n"
)


def _tracked_yaml_files() -> list[Path]:
    try:
        listed = subprocess.run(
            ["git", "ls-files", "*.yaml", "*.yml"],
            cwd=REPO_ROOT, capture_output=True, text=True, check=True,
        ).stdout.split()
    except (OSError, subprocess.CalledProcessError):
        return sorted(
            path for pattern in ("*.yaml", "*.yml") for path in REPO_ROOT.rglob(pattern)
            if not {"workspace", ".git"} & set(path.relative_to(REPO_ROOT).parts)
        )
    return [REPO_ROOT / name for name in listed]


class StrictReaderTests(unittest.TestCase):
    def _write(self, text: str) -> Path:
        folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, folder, ignore_errors=True)
        path = folder / "scenario.yaml"
        path.write_text(text, encoding="utf-8")
        return path

    def test_repeated_top_level_key_names_file_key_and_both_lines(self) -> None:
        path = self._write("a: 1\nb: 2\na: 3\n")
        with self.assertRaises(yaml_io.DuplicateKeyError) as caught:
            yaml_io.load_file(path)
        message = str(caught.exception)
        self.assertIn(str(path), message)
        self.assertIn("line 3: the key 'a' appears a second time (first at line 1)", message)

    def test_repeated_key_inside_a_nested_mapping_stops_loading(self) -> None:
        path = self._write("outer:\n  inner:\n    x: 1\n    y: 2\n    x: 3\n")
        with self.assertRaisesRegex(yaml_io.DuplicateKeyError, r"line 5: the key 'x'.*first at line 3"):
            yaml_io.load_file(path)
        with self.assertRaisesRegex(yaml_io.DuplicateKeyError, "the key 'x'"):
            yaml_io.safe_load("- {x: 1, x: 2}\n")

    def test_misspelled_block_name_suggests_the_known_key(self) -> None:
        path = self._write(HEAD + CAP_20_BGD.replace("gross_import_cap:", "gross_import_caps:"))
        with self.assertRaises(ScenarioFileError) as caught:
            load_scenario_file(path)
        message = str(caught.exception)
        self.assertIn(f"{path}, line 6: unknown key 'gross_import_caps'.", message)
        self.assertIn("Did you mean 'gross_import_cap'?", message)
        self.assertIn("Allowed keys: scenario, base_scenario,", message)

    def test_unknown_key_inside_the_cap_block_stops_loading_and_note_is_accepted(self) -> None:
        path = self._write(HEAD + CAP_20_BGD.replace("  fraction:", "  fractions:"))
        with self.assertRaises(ScenarioFileError) as caught:
            load_scenario_file(path)
        message = str(caught.exception)
        self.assertIn("line 9: unknown key 'fractions' in 'gross_import_cap'.", message)
        self.assertIn("Did you mean 'fraction'?", message)
        self.assertIn("Allowed keys: countries, fraction, first_year, last_year, note.", message)

        accepted = load_scenario_file(self._write(HEAD + CAP_20_BGD + "  note: Twenty per cent.\n"))
        self.assertEqual(accepted["gross_import_cap"]["note"], "Twenty per cent.")

    def test_every_tracked_yaml_file_reads_as_safe_load_reads_it(self) -> None:
        files = _tracked_yaml_files()
        self.assertTrue(files)
        for path in files:
            with self.subTest(path=path.relative_to(REPO_ROOT).as_posix()):
                text = path.read_text(encoding="utf-8")
                self.assertEqual(yaml_io.safe_load(text, source=path), yaml.safe_load(text))


class RunStopsAtRegistryLoadTests(unittest.TestCase):
    """A bad scenario file stops ``python -m ostram run`` before any build step."""

    def _run(self, scenario_text: str) -> tuple[int, str, list[str]]:
        temp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, temp, ignore_errors=True)
        (temp / "config" / "setups").mkdir(parents=True)
        shutil.copytree(REPO_ROOT / "config" / "scenarios", temp / "config" / "scenarios")
        registry_text = (REPO_ROOT / "config" / "setups" / "south_asia.yaml").read_text(encoding="utf-8")
        last = "  - C_Target_VRE_TxCostHigh\n"
        self.assertEqual(registry_text.count(last), 1)
        registry = temp / "config" / "setups" / "south_asia.yaml"
        registry.write_text(registry_text.replace(last, last + f"  - {SCENARIO}\n"), encoding="utf-8")
        folder = temp / "config" / "scenarios" / SCENARIO
        folder.mkdir()
        (folder / "scenario.yaml").write_text(scenario_text, encoding="utf-8")

        launcher = _load_launcher("strict_yaml")
        stderr = io.StringIO()
        with (
            LauncherHarness(launcher, snapshot_exists=False) as harness,
            mock.patch.object(launcher, "load_registry", lambda: load_registry(registry)),
            mock.patch.object(sys, "argv", ["python -m ostram run", "--skip-pull", "--compile-only"]),
            redirect_stdout(io.StringIO()),
            redirect_stderr(stderr),
        ):
            code = cli._invoke_run_guard(launcher)
        return code, stderr.getvalue(), [str(event[0]) for event in harness.events]

    def test_copied_cap_block_left_below_the_new_one_stops_the_run(self) -> None:
        code, stderr, events = self._run(HEAD + CAP_20_BGD + CAP_15_SIX)
        self.assertEqual(code, 1)
        self.assertIn(f"config/scenarios/{SCENARIO}/scenario.yaml, line 12", stderr.replace("\\", "/"))
        self.assertIn("line 12: the key 'gross_import_cap' appears a second time (first at line 6)", stderr)
        self.assertEqual(set(events), {"guess_env"})

    def test_misspelled_block_name_stops_the_run(self) -> None:
        code, stderr, events = self._run(HEAD + CAP_20_BGD.replace("gross_import_cap:", "gross_import_caps:"))
        self.assertEqual(code, 1)
        self.assertIn("line 6: unknown key 'gross_import_caps'. Did you mean 'gross_import_cap'?", stderr)
        self.assertEqual(set(events), {"guess_env"})

    def test_valid_new_scenario_reaches_the_build_steps(self) -> None:
        code, stderr, events = self._run(HEAD + CAP_20_BGD + "  note: Twenty per cent.\n")
        self.assertEqual(code, 0, stderr)
        self.assertIn("pipeline", events)


if __name__ == "__main__":
    unittest.main()
