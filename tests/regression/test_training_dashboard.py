"""Regression coverage for the profile-aware interactive training dashboard."""

from __future__ import annotations

import json
from pathlib import Path
import re
import tempfile
from types import SimpleNamespace
import unittest
from unittest import mock

import pandas as pd

from ostram.examples import _report
from ostram.reporting.training_dashboard import (
    build_dashboard_data,
    emission_serves_region,
    is_generation_tech,
    render_html,
    scenario_label,
    storage_prefixes,
)


# Deliberately non-UNESCAP scenario names, regions and emission codes: the
# dashboard must take every identity from the data/metadata, not constants.
METADATA = {
    "title": "Synthetic Training Profile",
    "country_regions": [
        {"region": "MMRXX", "label": "Myanmar"},
        {"region": "LKAXX", "label": "Sri Lanka"},
    ],
    "interconnectors": [{"technology": "TRNMMRXXLKAXX"}],
    "storage": ["SDSMMRXX01", "LDSLKAXX01"],
    "effective_values": {"interconnector_capacity_gw": 2.496},
    "year_range": {"start": 2023, "end": 2050},
}


def _row(scenario, year, tech=None, emission=None, **values):
    row = {
        "Scenario": scenario,
        "REGION": "GLOBAL",
        "YEAR": year,
        "TECHNOLOGY": tech,
        "EMISSION": emission,
    }
    row.update(values)
    return row


def _alpha_rows(solar_2050=20.0):
    rows = [
        _row("Alpha_Run", 2030, "PWRSPVMMRXX01",
             ProductionByTechnologyAnnual=10.0, TotalCapacityAnnual=4.0,
             NewCapacity=1.0, TotalAnnualMaxCapacityInvestment=3.0,
             CapitalInvestment=200.0),
        _row("Alpha_Run", 2050, "PWRSPVMMRXX01",
             ProductionByTechnologyAnnual=solar_2050, TotalCapacityAnnual=8.0,
             NewCapacity=2.0, TotalAnnualMaxCapacityInvestment=3.0),
        _row("Alpha_Run", 2030, "PWRCOALKAXX01",
             ProductionByTechnologyAnnual=30.0, TotalCapacityAnnual=6.0,
             NewCapacity=0.0, TotalAnnualMaxCapacityInvestment=2.0),
        _row("Alpha_Run", 2050, "PWRCOALKAXX01",
             ProductionByTechnologyAnnual=15.0, TotalCapacityAnnual=5.0,
             NewCapacity=0.0, TotalAnnualMaxCapacityInvestment=2.0,
             CapitalInvestment=300.0),
        # Storage, backstop and internal transmission must stay out of the
        # generation families and the VRE share denominator.
        _row("Alpha_Run", 2030, "PWRSDSMMRXX01", TotalCapacityAnnual=1.0),
        _row("Alpha_Run", 2050, "PWRSDSMMRXX01", TotalCapacityAnnual=2.0),
        _row("Alpha_Run", 2050, "PWRBCKMMRXX01",
             ProductionByTechnologyAnnual=99.0),
        _row("Alpha_Run", 2050, "PWRTRNMMRXX01",
             ProductionByTechnologyAnnual=50.0),
        _row("Alpha_Run", 2030, "TRNMMRXXLKAXX",
             TotalCapacityAnnual=2.496, ProductionByTechnologyAnnual=1.5),
        _row("Alpha_Run", 2050, "TRNMMRXXLKAXX",
             TotalCapacityAnnual=3.0, ProductionByTechnologyAnnual=2.5),
        _row("Alpha_Run", 2030, emission="CO2MMR", AnnualEmissions=5.0),
        _row("Alpha_Run", 2050, emission="CO2MMR", AnnualEmissions=3.0),
        _row("Alpha_Run", 2030, emission="CO2LKA", AnnualEmissions=7.0),
        _row("Alpha_Run", 2050, emission="CO2LKA", AnnualEmissions=4.0),
        _row("Alpha_Run", None, TotalDiscountedCost=1000.0),
        # Beta_Run only has input rows (no solver outputs): must stay
        # unavailable instead of rendering fabricated series.
        _row("Beta_Run", 2030, "PWRSPVMMRXX01",
             TotalAnnualMaxCapacityInvestment=3.0),
    ]
    # Exact duplicates of combined input/output rows must not double-count.
    rows.append(dict(rows[1]))
    rows.append(dict(rows[11]))
    rows.append(_row("Alpha_Run", None, TotalDiscountedCost=1000.0))
    return rows


def _write_csv(path: Path, rows) -> Path:
    pd.DataFrame(rows).to_csv(path, index=False)
    return path


def _build(tmp: Path, metadata=METADATA, snapshots=None):
    if snapshots is None:
        before = _write_csv(tmp / "baseline.csv", _alpha_rows())
        after = _write_csv(tmp / "edited.csv", _alpha_rows(solar_2050=25.0))
        snapshots = [("baseline", before), ("edited", after)]
    return build_dashboard_data(
        snapshots,
        profile_id="synthetic",
        manifest=tmp / "profile.yaml",
        workspace=tmp / "workspace",
        metadata=metadata,
    )


def _payload(document: str) -> dict:
    match = re.search(
        r'<script id="ostram-profile-data" type="application/json">(.*?)</script>',
        document,
        re.S,
    )
    assert match is not None
    return json.loads(match.group(1))


class ClassifierTests(unittest.TestCase):
    def test_scenario_labels_are_derived_not_hardcoded(self) -> None:
        self.assertEqual(scenario_label("A_Calibrated_BAU"), "A · Calibrated BAU")
        self.assertEqual(scenario_label("B_Optimised_VRE"), "B · Optimised VRE")
        self.assertEqual(scenario_label("Training"), "Training")
        self.assertEqual(scenario_label("Alpha_Run"), "Alpha_Run")

    def test_storage_prefixes_come_from_profile_metadata(self) -> None:
        self.assertEqual(storage_prefixes(METADATA), ("PWRSDS", "PWRLDS"))
        self.assertEqual(storage_prefixes({"storage": ["XYZMMRXX01"]}), ("PWRXYZ",))
        self.assertEqual(storage_prefixes({}), ("PWRSDS", "PWRLDS"))
        self.assertEqual(storage_prefixes(None), ("PWRSDS", "PWRLDS"))

    def test_generation_classification_excludes_non_generation(self) -> None:
        storage = storage_prefixes(METADATA)
        self.assertTrue(is_generation_tech("PWRSPVMMRXX01", storage))
        self.assertFalse(is_generation_tech("PWRSDSMMRXX01", storage))
        self.assertFalse(is_generation_tech("PWRLDSLKAXX01", storage))
        self.assertFalse(is_generation_tech("PWRTRNMMRXX01", storage))
        self.assertFalse(is_generation_tech("PWRBCKMMRXX01", storage))
        self.assertFalse(is_generation_tech("TRNMMRXXLKAXX", storage))

    def test_emission_region_attribution_is_structural(self) -> None:
        self.assertTrue(emission_serves_region("CO2MMR", "MMRXX"))
        self.assertTrue(emission_serves_region("CO2LKA", "LKAXX"))
        self.assertTrue(emission_serves_region("CO2BGD", "BGDXX"))
        self.assertTrue(emission_serves_region("CO2IND", "INDEA"))
        self.assertFalse(emission_serves_region("CO2MMR", "LKAXX"))
        self.assertFalse(emission_serves_region("", "MMRXX"))


class AggregationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._tmp = tempfile.TemporaryDirectory()
        cls.data = _build(Path(cls._tmp.name))
        cls.alpha = cls.data["snapshots"]["baseline"]["Alpha_Run"]

    @classmethod
    def tearDownClass(cls) -> None:
        cls._tmp.cleanup()

    def test_generation_family_totals_without_duplicate_overcounting(self) -> None:
        system = self.alpha["System"]
        self.assertEqual(
            system["generation"][2050],
            {"Solar PV": 20.0, "Coal": 15.0},
        )
        self.assertEqual(system["generation"][2030], {"Solar PV": 10.0, "Coal": 30.0})
        self.assertEqual(self.alpha["MMRXX"]["generation"][2050], {"Solar PV": 20.0})
        self.assertEqual(self.alpha["LKAXX"]["generation"][2050], {"Coal": 15.0})

    def test_capacity_and_new_capacity_series(self) -> None:
        system = self.alpha["System"]
        self.assertEqual(system["capacity"][2050], {"Solar PV": 8.0, "Coal": 5.0})
        self.assertEqual(
            system["new_capacity"][2030], {"Solar PV": 1.0, "Coal": 0.0}
        )
        self.assertEqual(system["new_capacity"][2050]["Solar PV"], 2.0)

    def test_emissions_route_by_emission_code_and_sum_to_system(self) -> None:
        system = self.alpha["System"]["emissions"]
        mmr = self.alpha["MMRXX"]["emissions"]
        lka = self.alpha["LKAXX"]["emissions"]
        self.assertEqual(mmr, {2030: 5.0, 2050: 3.0})
        self.assertEqual(lka, {2030: 7.0, 2050: 4.0})
        for year in (2030, 2050):
            self.assertAlmostEqual(system[year], mmr[year] + lka[year], places=9)

    def test_storage_uses_profile_declared_families(self) -> None:
        self.assertEqual(self.alpha["System"]["storage"], {2030: 1.0, 2050: 2.0})
        self.assertEqual(self.alpha["MMRXX"]["storage"], {2030: 1.0, 2050: 2.0})
        self.assertEqual(self.alpha["LKAXX"]["storage"], {})

    def test_vre_share_only_counts_valid_generation(self) -> None:
        system = self.alpha["System"]["vre_share"]
        self.assertEqual(system[2050], round(20.0 / 35.0, 4))
        self.assertEqual(self.alpha["MMRXX"]["vre_share"][2050], 1.0)
        self.assertEqual(self.alpha["LKAXX"]["vre_share"][2050], 0.0)

    def test_costs_are_system_wide_and_deduplicated(self) -> None:
        expected = {"total_discounted": 1000.0, "capex": 500.0}
        self.assertEqual(self.alpha["System"]["cost"], expected)
        self.assertEqual(self.alpha["MMRXX"]["cost"], expected)
        self.assertEqual(self.alpha["LKAXX"]["cost"], expected)

    def test_interconnector_series_keep_exact_seed_value(self) -> None:
        for region in ("System", "MMRXX", "LKAXX"):
            block = self.alpha[region]["interconnectors"]
            self.assertEqual(
                block["TotalCapacityAnnual"], {2030: 2.496, 2050: 3.0}
            )
            self.assertEqual(
                block["ProductionByTechnologyAnnual"], {2030: 1.5, 2050: 2.5}
            )

    def test_lid_diagnostic_labels_use_profile_region_names(self) -> None:
        lid = self.alpha["System"]["lid_diagnostic"]
        self.assertIn("PWRSPVMMRXX01", lid)
        self.assertEqual(lid["PWRSPVMMRXX01"]["label"], "Solar PV — Myanmar")
        self.assertEqual(lid["PWRSPVMMRXX01"]["newcap"], {2030: 1.0, 2050: 2.0})
        self.assertEqual(lid["PWRSPVMMRXX01"]["lid"], {2030: 3.0, 2050: 3.0})
        self.assertEqual(lid["PWRCOALKAXX01"]["label"], "Coal — Sri Lanka")
        self.assertNotIn("PWRSPVMMRXX01", self.alpha["LKAXX"]["lid_diagnostic"])

    def test_scenarios_and_regions_come_from_data_and_metadata(self) -> None:
        self.assertEqual(
            [item["id"] for item in self.data["scenarios"]],
            ["Alpha_Run", "Beta_Run"],
        )
        self.assertEqual(
            self.data["regions"],
            [
                {"id": "System", "label": "System"},
                {"id": "MMRXX", "label": "Myanmar"},
                {"id": "LKAXX", "label": "Sri Lanka"},
            ],
        )
        self.assertEqual(self.data["year_range"], {"start": 2023, "end": 2050})
        self.assertEqual(self.data["snapshot_labels"], ["baseline", "edited"])
        self.assertEqual(
            self.data["effective_values"]["interconnector_capacity_gw"], 2.496
        )

    def test_input_only_scenario_stays_unavailable(self) -> None:
        beta = self.data["snapshots"]["baseline"]["Beta_Run"]
        self.assertFalse(beta["System"]["available"])
        self.assertEqual(beta["System"]["generation"], {})
        self.assertTrue(self.alpha["System"]["available"])


class EdgeCaseTests(unittest.TestCase):
    def test_missing_optional_columns_degrade_gracefully(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            tmp = Path(temp)
            rows = [
                {
                    "Scenario": "Alpha_Run",
                    "YEAR": 2030,
                    "TECHNOLOGY": "PWRSPVMMRXX01",
                    "ProductionByTechnologyAnnual": 10.0,
                }
            ]
            path = _write_csv(tmp / "thin.csv", rows)
            data = _build(tmp, snapshots=[("thin", path)])
            metrics = data["snapshots"]["thin"]["Alpha_Run"]["System"]
            self.assertTrue(metrics["available"])
            self.assertEqual(metrics["generation"], {2030: {"Solar PV": 10.0}})
            for key in ("capacity", "new_capacity", "emissions", "storage",
                        "interconnectors", "lid_diagnostic"):
                self.assertEqual(metrics[key], {}, key)
            self.assertEqual(
                metrics["cost"], {"total_discounted": 0.0, "capex": 0.0}
            )
            self.assertEqual(data["year_range"], {"start": 2023, "end": 2050})

    def test_single_snapshot_builds_and_renders(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            tmp = Path(temp)
            path = _write_csv(tmp / "only.csv", _alpha_rows())
            data = _build(tmp, snapshots=[("only", path)])
            self.assertEqual(data["snapshot_labels"], ["only"])
            document = render_html(data)
            self.assertIn("sel-before", document)
            self.assertEqual(_payload(document)["snapshot_labels"], ["only"])

    def test_year_range_falls_back_to_observed_years(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            tmp = Path(temp)
            metadata = {
                key: value for key, value in METADATA.items() if key != "year_range"
            }
            path = _write_csv(tmp / "obs.csv", _alpha_rows())
            data = _build(tmp, metadata=metadata, snapshots=[("obs", path)])
            self.assertEqual(data["year_range"], {"start": 2030, "end": 2050})

    def test_duplicate_snapshot_labels_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            tmp = Path(temp)
            path = _write_csv(tmp / "one.csv", _alpha_rows())
            with self.assertRaisesRegex(ValueError, "duplicate snapshot label"):
                _build(tmp, snapshots=[("dup", path), ("dup", path)])


class RendererTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._tmp = tempfile.TemporaryDirectory()
        cls.document = render_html(_build(Path(cls._tmp.name)))

    @classmethod
    def tearDownClass(cls) -> None:
        cls._tmp.cleanup()

    def test_controls_for_snapshot_scenario_and_region_modes(self) -> None:
        for control in (
            'id="mode-snap"', 'id="mode-scen"', 'id="sel-before"',
            'id="sel-after"', 'id="sel-scenario"', 'id="sel-snapshot"',
            'id="sel-scen-a"', 'id="sel-scen-b"', 'id="sel-region"',
        ):
            self.assertIn(control, self.document)
        self.assertIn("sel-lid-tech", self.document)

    def test_expected_chart_panels_and_kpi_cards(self) -> None:
        for marker in (
            "Generation mix by technology",
            "Installed capacity by technology",
            "New capacity investment by technology",
            "CO₂ emissions",
            "Storage capacity (GW)",
            "Interconnector capacity (GW)",
            "Interconnector trade flow (PJ)",
            "VRE share of generation",
            "Lid diagnostic",
            "System cost (NPV)",
            "Capital investment",
            "VRE share",
            "No results for",
        ):
            self.assertIn(marker, self.document)

    def test_report_surface_is_not_a_raw_json_dump(self) -> None:
        self.assertNotIn("<pre", self.document)
        self.assertIn("stackedArea", self.document)
        self.assertIn("lineChart", self.document)

    def test_embedded_profile_report_payload_round_trips(self) -> None:
        payload = _payload(self.document)
        self.assertEqual(payload["schema"], "ostram-profile-report-v1")
        self.assertEqual(payload["profile_id"], "synthetic")
        self.assertEqual(payload["snapshot_labels"], ["baseline", "edited"])
        self.assertEqual(
            payload["effective_values"]["interconnector_capacity_gw"], 2.496
        )
        self.assertIn("2.496", self.document)

    def test_output_is_offline_and_self_contained(self) -> None:
        for forbidden in ("http", "<link", " src=", "@import", "url(",
                          "fetch(", "XMLHttpRequest", "@font-face"):
            self.assertNotIn(forbidden, self.document, forbidden)
        self.assertIn("Climate Lead Group", self.document)

    def test_unescap_visual_language_is_preserved(self) -> None:
        for marker in ("--teal-dark:#00414D", "--teal:#0A595F",
                       "footer-logo-svg", "hero", "badge"):
            self.assertIn(marker, self.document)


class EscapingTests(unittest.TestCase):
    def test_hostile_scenario_names_and_titles_are_escaped(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            tmp = Path(temp)
            hostile = 'Evil</script><script>alert(1)_"quoted"'
            rows = [
                {
                    "Scenario": hostile,
                    "YEAR": 2030,
                    "TECHNOLOGY": "PWRSPVMMRXX01",
                    "ProductionByTechnologyAnnual": 1.0,
                }
            ]
            metadata = dict(METADATA)
            metadata["title"] = '<b>&Sneaky "Title"'
            path = _write_csv(tmp / "bad.csv", rows)
            document = render_html(_build(tmp, metadata=metadata,
                                          snapshots=[("bad", path)]))
            # Only the template's own two script closers may exist.
            self.assertEqual(document.count("</script>"), 2)
            self.assertNotIn("Evil</script>", document)
            # Outside the inert JSON payload, the hostile title must appear
            # only in HTML-escaped form.
            without_payload = re.sub(
                r'<script id="ostram-profile-data"[^>]*>.*?</script>',
                "",
                document,
                flags=re.S,
            )
            self.assertNotIn('<b>&Sneaky', without_payload)
            self.assertIn("&lt;b&gt;&amp;Sneaky &quot;Title&quot;", document)
            payload = _payload(document)
            self.assertIn(hostile, payload["snapshots"]["bad"])


class ReportRouteTests(unittest.TestCase):
    def _workspace(self, temp: str):
        workspace = Path(temp).resolve()
        execution = workspace / "execution"
        execution.mkdir()
        _write_csv(execution / "OSTRAM_Combined_Inputs_Outputs.csv", _alpha_rows())
        paths = SimpleNamespace(workspace=workspace, execution_workspace=execution)
        manifest = SimpleNamespace(
            profile_id="unescap",
            path=workspace / "profile.yaml",
            metadata=METADATA,
        )
        return workspace, execution, paths, manifest

    def test_capture_baseline_rebuild_and_multi_capture_report(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            workspace, execution, paths, manifest = self._workspace(temp)
            with mock.patch("ostram.examples.resolve_paths", return_value=paths):
                report = _report(manifest, "baseline")
                self.assertEqual(report, workspace / "reports" / "unescap.html")
                self.assertTrue(
                    (workspace / "reports" / "snapshots" / "baseline.csv").exists()
                )
                payload = _payload(report.read_text(encoding="utf-8"))
                self.assertEqual(payload["snapshot_labels"], ["baseline"])

                _write_csv(
                    execution / "OSTRAM_Combined_Inputs_Outputs.csv",
                    _alpha_rows(solar_2050=25.0),
                )
                _report(manifest, "exercise1")
                rebuilt = _report(manifest, None)
                self.assertEqual(rebuilt, workspace / "reports" / "unescap.html")
                payload = _payload(rebuilt.read_text(encoding="utf-8"))
                self.assertEqual(
                    sorted(payload["snapshot_labels"]), ["baseline", "exercise1"]
                )
                metrics = payload["snapshots"]["exercise1"]["Alpha_Run"]["System"]
                self.assertEqual(metrics["generation"]["2050"]["Solar PV"], 25.0)
            for produced in workspace.rglob("*.html"):
                self.assertTrue(produced.resolve().is_relative_to(workspace))

    def test_compare_route_writes_separate_interconnector_report(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            workspace, _execution, paths, manifest = self._workspace(temp)
            with mock.patch("ostram.examples.resolve_paths", return_value=paths):
                for label in ("forward", "reverse", "bidirectional"):
                    _report(manifest, label)
                compared = _report(manifest, None, "forward,reverse,bidirectional")
            self.assertEqual(
                compared,
                workspace / "reports" / "unescap-interconnector-comparison.html",
            )
            self.assertTrue(compared.resolve().is_relative_to(workspace))
            payload = _payload(compared.read_text(encoding="utf-8"))
            self.assertEqual(
                payload["snapshot_labels"], ["forward", "reverse", "bidirectional"]
            )


if __name__ == "__main__":
    unittest.main()
