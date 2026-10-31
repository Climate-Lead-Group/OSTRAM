"""Checks of the website generator, ``python -m ostram site``.

The site is built into temporary folders. The tests compare two builds byte for
byte, read every value the site shows back from the setup manifests and their
authority files, check every internal link and anchor and every external host,
and build once more with a fixture ``lac`` setup registered.
"""

from __future__ import annotations

import csv
import filecmp
import html
import json
from pathlib import Path
import posixpath
import re
import tempfile
import unittest

import yaml
from openpyxl import load_workbook

from ostram.paths import resolve_paths
from ostram.setups import load_manifest, load_setup
from ostram.site.blocks import format_value
from ostram.site.build import build_site

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = Path(__file__).resolve().parent / "fixtures" / "lac" / "setup.yaml"
ALLOWED_SCRIPT_HOSTS = {"cdn.jsdelivr.net"}


def _files(folder: Path) -> list[str]:
    return sorted(p.relative_to(folder).as_posix() for p in folder.rglob("*") if p.is_file())


def _pages(folder: Path) -> dict[str, str]:
    return {name: (folder / name).read_text(encoding="utf-8") for name in _files(folder) if name.endswith(".html")}


def _taxonomy(path: Path) -> list[tuple[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return [(row[0], row[1]) for row in list(csv.reader(handle))[1:] if len(row) > 1 and row[0]]


def expected_values(setup_id: str, paths, manifest=None) -> dict[str, object]:
    """Read the headline values of one setup from its files, independently of the generator."""

    manifest = manifest or load_setup(setup_id, paths=paths)
    sources = manifest.source_paths(paths, require_exists=False)
    expected: dict[str, object] = {}
    rows = _taxonomy(sources["interconnector_taxonomy"])
    expected["nodes"] = len({tech[6:] for category, tech in rows if category == "TRANSMISSION_PLANNED_NONRENO"})
    expected["interconnectors"] = sum(1 for category, _ in rows if category == "INTERCONNECTORS")
    expected["internal_grid_technologies"] = sum(
        1 for category, _ in rows
        if category.startswith("TRANSMISSION_") and category != "TRANSMISSION_SUM"
    )
    compilation = yaml.safe_load(sources["compilation_config"].read_text(encoding="utf-8"))
    expected["first_year"] = int(compilation["initial_year"])
    expected["last_year"] = int(compilation["final_year"])
    expected["years"] = expected["last_year"] - expected["first_year"] + 1
    text = sources["compilation_config"].read_text(encoding="utf-8")
    timeslices = re.search(r"^(\s*)Timeslices:\n((?:\1- .*\n)+)", text, re.MULTILINE)
    expected["timeslices"] = len(timeslices.group(2).strip().splitlines())
    registry_path = sources["scenario_registry"]
    raw = (
        json.loads(registry_path.read_text(encoding="utf-8"))
        if registry_path.suffix == ".json"
        else yaml.safe_load(registry_path.read_text(encoding="utf-8"))
    )
    registry = raw.get("scenario_registry", raw)
    expected["support_scenarios"] = len(registry["support_scenarios"])
    expected["decision_scenarios"] = len(registry["decision_scenarios"])
    expected["root_scenarios"] = len(registry["root_scenarios"])
    expected["scenarios"] = expected["support_scenarios"] + expected["decision_scenarios"]
    country = yaml.safe_load(sources["country_config"].read_text(encoding="utf-8"))
    for name, value in (country.get("internal_transmission") or {}).items():
        expected[f"internal.{name}"] = value
    return {f"{setup_id}.{key}": value for key, value in expected.items()}


class SiteBuildTest(unittest.TestCase):
    """One build of the repository's site, shared by the checks below."""

    @classmethod
    def setUpClass(cls):
        cls.paths = resolve_paths(project_root=ROOT)
        cls.tmp = tempfile.TemporaryDirectory()
        cls.out = Path(cls.tmp.name) / "first"
        cls.page_list = build_site(cls.out, paths=cls.paths)
        cls.pages = _pages(cls.out)
        cls.values = json.loads((cls.out / "values.json").read_text(encoding="utf-8"))

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_second_build_is_identical(self):
        second = Path(self.tmp.name) / "second"
        build_site(second, paths=self.paths)
        self.assertEqual(_files(self.out), _files(second))
        _, mismatch, errors = filecmp.cmpfiles(self.out, second, _files(self.out), shallow=False)
        self.assertEqual((mismatch, errors), ([], []))

    def test_every_page_of_the_list_is_written(self):
        for path in self.page_list:
            self.assertIn(path, self.pages)
        for path in ("index.html", "glossary.html", "applications/south_asia.html", "applications/lac.html"):
            self.assertIn(path, self.pages)

    def test_values_equal_the_authority_files(self):
        for setup_id in ("south_asia", "unescap_training"):
            for key, value in expected_values(setup_id, self.paths).items():
                with self.subTest(key=key):
                    self.assertIn(key, self.values)
                    self.assertEqual(self.values[key]["value"], value)
        manifest = load_setup("unescap_training", paths=self.paths)
        self.assertEqual(self.values["unescap_training.nodes"]["value"], len(manifest.metadata["country_regions"]))
        self.assertEqual(self.values["unescap_training.timeslices"]["value"], manifest.metadata["timeslices"]["count"])
        self.assertEqual(self.values["unescap_training.years"]["value"], manifest.metadata["year_range"]["count"])

    def test_repository_counts(self):
        self.assertEqual(self.values["repo.tests"]["value"], len(list((ROOT / "tests").rglob("test_*.py"))))
        self.assertEqual(self.values["repo.docs_pages"]["value"], len(list((ROOT / "docs").rglob("*.md"))))
        rules = [p for p in (ROOT / "ostram/pipeline/scenarios/rules").glob("*.py") if p.name != "__init__.py"]
        self.assertEqual(self.values["repo.rule_modules"]["value"], len(rules))
        self.assertEqual(self.values["repo.workflows"]["value"], len(list((ROOT / ".github/workflows").glob("*.yml"))))
        registry = json.loads((ROOT / "config/setups/registry.json").read_text(encoding="utf-8"))
        self.assertEqual(self.values["repo.setups"]["value"], len(registry["setups"]))
        scenarios = ROOT / "ostram/pipeline/scenarios"
        lines = sum(len(p.read_text(encoding="utf-8").splitlines()) for p in scenarios.rglob("*.py"))
        self.assertEqual(self.values["repo.stage.A3.lines"]["value"], lines)

    def test_workbook_and_manifest_values(self):
        manifest = load_setup("south_asia", paths=self.paths)
        raw = yaml.safe_load(manifest.path.read_text(encoding="utf-8"))
        for tech, ceiling in raw["vre_ceilings"]["ceilings_gw"].items():
            self.assertEqual(self.values[f"south_asia.vre_ceiling.{tech}"]["value"], float(ceiling))
        cap = raw["policies"]["link_import_cap"]
        self.assertEqual(self.values["south_asia.link_import_cap.first_year"]["value"], cap["first_year"])
        workbook = load_workbook(
            manifest.source_paths(self.paths, require_exists=False)["scenario_workbook"], read_only=True, data_only=True
        )
        rows = list(workbook["Transmission_Cost_Basis"].iter_rows(values_only=True))
        header = list(rows[0])
        checked = 0
        for row in rows[1:]:
            record = dict(zip(header, row))
            if isinstance(record["Technology"], str) and record["Technology"].startswith("TRN"):
                tech = record["Technology"]
                self.assertEqual(self.values[f"south_asia.corridor.{tech}.capital"]["value"], record["Capital_USD 2023_per_kW"])
                self.assertEqual(self.values[f"south_asia.corridor.{tech}.fixed"]["value"], record["Fixed_USD 2023_per_kW_year"])
                checked += 1
        workbook.close()
        self.assertEqual(checked, self.values["south_asia.interconnectors"]["value"])

    def test_every_value_on_a_page_matches_its_source(self):
        span = re.compile(r'<span class="v" data-v="([^"]+)"(?: data-f="([^"]*)")? title="[^"]*">([^<]*)</span>')
        seen = 0
        for name, text in self.pages.items():
            for key, kind, shown in span.findall(text):
                with self.subTest(page=name, key=key):
                    self.assertIn(key, self.values)
                    self.assertEqual(html.unescape(shown), format_value(self.values[key]["value"], kind))
                seen += 1
        self.assertGreater(seen, 100)

    def test_internal_links_and_anchors_resolve(self):
        ids = {name: set(re.findall(r'\sid="([^"]+)"', text)) for name, text in self.pages.items()}
        for name, text in self.pages.items():
            duplicates = [i for i in ids[name] if len(re.findall(rf'\sid="{re.escape(i)}"', text)) > 1]
            self.assertEqual(duplicates, [], name)
            for attribute, target in re.findall(r'\s(href|src)="([^"]+)"', text):
                target = html.unescape(target)
                if re.match(r"^[a-z][a-z0-9+.-]*:", target):
                    continue
                path, _, anchor = target.partition("#")
                resolved = posixpath.normpath(posixpath.join(posixpath.dirname(name), path)) if path else name
                with self.subTest(page=name, link=target):
                    self.assertTrue((self.out / resolved).is_file(), f"missing {resolved}")
                    if anchor and resolved.endswith(".html"):
                        self.assertIn(anchor, ids[resolved])

    def test_only_the_allowed_host_is_loaded(self):
        for name, text in self.pages.items():
            loaded = re.findall(r'<(?:script|img|iframe|source)[^>]*\ssrc="(https?://[^"/]+)', text)
            loaded += re.findall(r'<link[^>]*\shref="(https?://[^"/]+)', text)
            loaded += re.findall(r'import [^;]* from "(https?://[^"/]+)', text)
            for origin in loaded:
                self.assertIn(origin.split("//", 1)[1], ALLOWED_SCRIPT_HOSTS, name)
        for name in _files(self.out):
            if name.endswith((".css", ".js")):
                text = (self.out / name).read_text(encoding="utf-8")
                self.assertNotRegex(text, r"@import|url\(\s*['\"]?https?:", name)

    def test_no_em_dash_in_the_site_text(self):
        package = ROOT / "ostram" / "site"
        for path in sorted(package.rglob("*")):
            if path.is_file() and path.suffix in (".html", ".yaml", ".css", ".js", ".py"):
                self.assertNotIn("\u2014", path.read_text(encoding="utf-8"), path.name)

    def test_lac_is_a_placeholder_without_a_lac_setup(self):
        page = self.pages["applications/lac.html"]
        self.assertIn('class="placeholder"', page)
        self.assertNotIn('data-v="lac.', page)
        self.assertNotIn('data-v="lac.', self.pages["index.html"])


class LacFixtureTest(unittest.TestCase):
    """With a ``lac`` setup registered, the same templates fill its pages."""

    def test_lac_pages_fill_with_a_lac_setup(self):
        paths = resolve_paths(project_root=ROOT)
        with tempfile.TemporaryDirectory() as tmp:
            registry = Path(tmp) / "registry.json"
            fixture = FIXTURE.relative_to(ROOT).as_posix()
            registry.write_text(
                json.dumps(
                    {
                        "schema": "ostram-setup-registry-v1",
                        "setups": {
                            "south_asia": "project:config/setups/south_asia.yaml",
                            "lac": f"project:{fixture}",
                        },
                    }
                ),
                encoding="utf-8",
            )
            out = Path(tmp) / "site"
            build_site(out, paths=paths, registry_path=registry)
            pages = _pages(out)
            values = json.loads((out / "values.json").read_text(encoding="utf-8"))
            manifest = load_manifest(FIXTURE, expected_setup="lac")
            for key, value in expected_values("lac", paths, manifest).items():
                with self.subTest(key=key):
                    self.assertEqual(values[key]["value"], value)
            lac = pages["applications/lac.html"]
            self.assertNotIn('class="placeholder"', lac)
            self.assertIn('data-v="lac.nodes"', lac)
            self.assertIn('data-v="lac.nodes"', pages["index.html"])
            self.assertIn('id="families-lac"', pages["transmission/internal-grid.html"])
            self.assertIn('id="rpo-lac"', pages["transmission/repowering.html"])
            self.assertNotIn("applications/unescap_training.html", pages)


if __name__ == "__main__":
    unittest.main()
