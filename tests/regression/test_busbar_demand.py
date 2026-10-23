"""Demand rows measured at the bus bar sit on the far side of the internal loss.

The ``Measured_at`` column of ``Demand_Projection`` marks each demand row as
``bus-bar`` or ``consumption``. The loss stage writes a bus-bar row as the
workbook value x (1 - loss); imports reach the receiving node's ``ELC..01``, so
they cross its internal grid; the energy balance has no add-back; and the rules
that read demand read the published value.
"""
from __future__ import annotations

import re
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import yaml
from openpyxl import Workbook, load_workbook

from ostram.pipeline.execution import gross_import_cap
from ostram.pipeline.preparation import transmission
from ostram.pipeline.scenarios.rules import apply_internal_transmission as itx
from ostram.pipeline.scenarios.rules import apply_internal_tx_losses as losses


REPO_ROOT = Path(__file__).resolve().parents[2]
WORKBOOK = REPO_ROOT / "inputs" / "scenarios" / "OSTRAM_Scenario_Inputs.xlsx"
COUNTRY_CONFIG = REPO_ROOT / "config" / "preparation" / "Config_country_codes.yaml"
MODEL = REPO_ROOT / "model" / "osemosys_fast_preprocessed.txt"
BUS_BAR = {"ELCBGDXX03", "ELCBTNXX03", "ELCINDNO03", "ELCINDSO03", "ELCINDEA03",
           "ELCINDWE03", "ELCINDNE03", "ELCNPLXX03", "ELCLKAXX03"}


def _template(path: Path, rows, header=("Demand/Share", "Fuel/Tech", "Name", "Measured_at", 2023, 2024)):
    wb = Workbook()
    ws = wb.active
    ws.title = "Demand_Projection"
    ws.append(list(header))
    for row in rows:
        ws.append(list(row))
    wb.save(path)
    return path


class MeasuredAtTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_workbook_marks_nine_bus_bar_rows_and_maldives_consumption(self) -> None:
        flags = losses.read_measured_at(WORKBOOK)
        self.assertEqual(losses.bus_bar_fuels(WORKBOOK), BUS_BAR)
        self.assertEqual(flags["ELCMDVXX03"], "consumption")
        self.assertEqual(len(flags), 10)

    def test_missing_column_blank_or_unknown_value_raises(self) -> None:
        path = _template(self.dir / "no_column.xlsx", [("Demand", "ELCAAAXX03", "x", 1, 1)],
                         header=("Demand/Share", "Fuel/Tech", "Name", 2023, 2024))
        with self.assertRaises(ValueError) as ctx:
            losses.read_measured_at(path)
        self.assertIn("Measured_at", str(ctx.exception))
        for value in (None, "busbar", "sent out"):
            path = _template(self.dir / "bad.xlsx", [("Demand", "ELCAAAXX03", "x", value, 1, 1)])
            with self.assertRaises(ValueError) as ctx:
                losses.read_measured_at(path)
            self.assertIn("ELCAAAXX03", str(ctx.exception))

    def test_scenario_rows_replace_bau_rows(self) -> None:
        path = _template(self.dir / "scen.xlsx", [
            ("BAU", "Demand", "ELCAAAXX03", "x", "bus-bar", 1),
            ("S1", "Demand", "ELCAAAXX03", "x", "consumption", 1),
            ("S2", "Demand", "ELCAAAXX03", "x", "", 1),
        ], header=("scenario", "Demand/Share", "Fuel/Tech", "Name", "Measured_at", 2023))
        self.assertEqual(losses.read_measured_at(path), {"ELCAAAXX03": "bus-bar"})
        self.assertEqual(losses.read_measured_at(path, "S1"), {"ELCAAAXX03": "consumption"})


class LossStageDemandTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.input = self.dir / "stage5"
        self.input.mkdir()
        for name, sheet, header in ((losses.PROJ_FILE, "Demand Techs", ["Tech", "Direction", 2023]),
                                    (losses.BASE_FILE, "Demand Techs", ["Tech", "Value.Fuel.O"])):
            wb = Workbook()
            wb.active.title = sheet
            wb.active.append(header)
            wb.save(self.input / name)
        wb = Workbook()
        ws = wb.active
        ws.title = "Demand_Projection"
        ws.append(["Demand/Share", "Fuel/Tech", "Name", "2023", "2024"])
        ws.append(["Demand", "ELCAAAXX03", "x", 100.0, 200.0])
        ws.append(["Demand", "ELCBBBXX03", "x", 10.0, 20.0])
        wb.save(self.input / losses.DEMAND_FILE)
        self.template = _template(self.dir / "template.xlsx", [
            ("Demand", "ELCAAAXX03", "x", "bus-bar", 100.0, 200.0),
            ("Demand", "ELCBBBXX03", "x", "consumption", 10.0, 20.0)])
        self.config = self.dir / "Config_country_codes.yaml"
        self.config.write_text(yaml.safe_dump({"internal_transmission": {"transmission_loss": 0.03}}),
                               encoding="utf-8")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _demand(self):
        wb = load_workbook(self.input / losses.DEMAND_FILE, data_only=True)
        out = {r[1]: (r[3], r[4]) for r in wb.active.iter_rows(min_row=2, values_only=True)}
        wb.close()
        return out

    def test_bus_bar_rows_scaled_once_consumption_rows_kept(self) -> None:
        log = losses.run(self.input, skip_backup=True, config=self.config, template=self.template)
        self.assertEqual(log["demand"]["cells"], 2)
        self.assertEqual(self._demand(), {"ELCAAAXX03": (100.0 * 0.97, 200.0 * 0.97),
                                          "ELCBBBXX03": (10.0, 20.0)})
        again = losses.run(self.input, skip_backup=True, config=self.config, template=self.template)
        self.assertEqual(again["demand"]["cells"], 0)
        self.assertEqual(self._demand()["ELCAAAXX03"], (100.0 * 0.97, 200.0 * 0.97))

    def test_cell_neither_published_nor_scaled_raises(self) -> None:
        wb = load_workbook(self.input / losses.DEMAND_FILE)
        wb.active["D2"] = 123.0
        wb.save(self.input / losses.DEMAND_FILE)
        with self.assertRaises(ValueError) as ctx:
            losses.run(self.input, skip_backup=True, config=self.config, template=self.template)
        self.assertIn("ELCAAAXX03", str(ctx.exception))

    def test_setup_without_internal_transmission_block_keeps_demand(self) -> None:
        self.config.write_text("enable_dsptrn: true\n", encoding="utf-8")
        log = losses.run(self.input, skip_backup=True, config=self.config, template=self.template)
        self.assertIn("skipped", log["demand"])
        self.assertEqual(self._demand()["ELCAAAXX03"], (100.0, 200.0))


class ResidualPeakReadsPublishedDemandTests(unittest.TestCase):
    def test_bus_bar_demand_is_divided_by_one_minus_loss(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            dem = Workbook()
            ws = dem.active
            ws.title = "Demand_Projection"
            ws.append(["Demand/Share", "Fuel/Tech", "Projection.Mode", "2023"])
            ws.append(["Demand", "ELCAAAXX03", "User defined", 31.536 * 0.97])
            ws.append(["Demand", "ELCBBBXX03", "User defined", 31.536])
            ws = dem.create_sheet("Profiles")
            ws.append(["Timeslices", "Fuel/Tech", "Projection.Mode", "2023"])
            for fuel in ("ELCAAAXX03", "ELCBBBXX03"):
                ws.append(["S1", fuel, "User defined", 1.0])
            dem.save(folder / itx.DEMAND_FILE_NAME)
            par = Workbook()
            ws = par.active
            ws.title = "Yearsplit"
            ws.append(["Timeslices", "Parameter", "Projection.Mode", "2023"])
            ws.append(["S1", "YearSplit", "User defined", 1.0])
            ws = par.create_sheet("Capacities")
            ws.append(["Timeslices", "Tech", "Parameter", "Projection.Mode", "2023"])
            par.save(folder / itx.PARAM_FILE_NAME)
            inputs = itx.read_residual_inputs(folder, {"ELCAAAXX03"}, 0.03)
            # 31.536 PJ a year is 1 GW: the bus-bar row reads back the published 1 GW.
            a = itx.compute_node_residuals(inputs, "AAAXX", 1.0)
            b = itx.compute_node_residuals(inputs, "BBBXX", 1.0)
            self.assertAlmostEqual(a["peak_GW"], 1.0, places=12)
            self.assertAlmostEqual(b["peak_GW"], 1.0, places=12)


class GrossImportFractionTests(unittest.TestCase):
    def test_bus_bar_country_fraction_over_one_minus_loss_consumption_country_kept(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            scen = root / "scenarios" / "T_Cap"
            scen.mkdir(parents=True)
            (scen / "scenario.yaml").write_text(
                "scenario: T_Cap\ngross_import_cap:\n  countries: [BGD, MDV]\n  fraction: 0.15\n"
                "  first_year: 2030\n  last_year: 2030\n", encoding="utf-8")
            (root / "SpecifiedAnnualDemand.csv").write_text(
                "REGION,FUEL,YEAR,VALUE\nGLOBAL,ELCBGDXX03,2030,1.0\nGLOBAL,ELCMDVXX03,2030,1.0\n",
                encoding="utf-8")
            (root / "OutputActivityRatio.csv").write_text(
                "REGION,TECHNOLOGY,FUEL,MODE_OF_OPERATION,YEAR,VALUE\n", encoding="utf-8")
            paths = SimpleNamespace(scenario_config_root=root / "scenarios", scenario_workbook=WORKBOOK,
                                    country_config=COUNTRY_CONFIG)
            with mock.patch.object(gross_import_cap, "resolve_paths", return_value=paths):
                text = gross_import_cap.render_policy(root, "T_Cap")
        got = dict(re.findall(r"^(BGD|MDV) 2030 (\S+)$", text, re.M))
        self.assertEqual(float(got["BGD"]), 0.15 / 0.97)
        self.assertEqual(float(got["MDV"]), 0.15)


class ImportsCrossTheInternalGridTests(unittest.TestCase):
    def test_dsptrn_mode_2_delivers_to_elc01(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp) / "base.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Secondary"
            ws.append(["Mode.Operation", "Tech", "Tech.Name", "Fuel.I", "Fuel.I.Name", "Fuel.O", "Fuel.O.Name"])
            ws.append([1, "PWRCOABGDXX", "coal", "COABGD", "coal", "ELCBGDXX02", "x"])
            ws = wb.create_sheet("Demand Techs")
            ws.append(["Mode.Operation", "Fuel.I", "Fuel.I.Name", "Value.Fuel.I", "Unit.Fuel.I", "Tech",
                       "Tech.Name", "Fuel.O", "Fuel.O.Name", "Value.Fuel.O", "Unit.Fuel.O"])
            wb.save(base)
            transmission.process_base_year(base, [("BGD", "XX")], enable_dsptrn=True)
            wb = load_workbook(base, data_only=True)
            rows = list(wb["Demand Techs"].iter_rows(min_row=1, values_only=True))
            hdr = rows[0]
            dsp = {r[hdr.index("Mode.Operation")]: (r[hdr.index("Fuel.I")], r[hdr.index("Fuel.O")])
                   for r in rows[1:] if r[hdr.index("Tech")] == "DSPTRNBGDXX"}
            self.assertEqual(dsp, {1: ("ELCBGDXX02", "ELCBGDXX03"), 2: ("ELCBGDXX04", "ELCBGDXX01")})

            proj = Path(tmp) / "proj.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Secondary"
            ws.append(["Mode.Operation", "Tech", "Tech.Name", "Fuel", "Fuel.Name", "Direction", 2023])
            ws.append([1, "PWRCOABGDXX", "coal", "ELCBGDXX02", "x", "Output", 1])
            ws = wb.create_sheet("Demand Techs")
            ws.append(["Mode.Operation", "Tech", "Tech.Name", "Fuel", "Fuel.Name", "Direction",
                       "Projection.Mode", "Projection.Parameter", 2023])
            wb.save(proj)
            transmission.process_projections(proj, [("BGD", "XX")], enable_dsptrn=True)
            wb = load_workbook(proj, data_only=True)
            rows = list(wb["Demand Techs"].iter_rows(min_row=1, values_only=True))
            hdr = rows[0]
            dsp = {(r[hdr.index("Mode.Operation")], r[hdr.index("Direction")]): r[hdr.index("Fuel")]
                   for r in rows[1:] if r[hdr.index("Tech")] == "DSPTRNBGDXX"}
            self.assertEqual(dsp[(2, "Output")], "ELCBGDXX01")
            self.assertEqual(dsp[(1, "Output")], "ELCBGDXX03")


class NoAddBackTests(unittest.TestCase):
    def test_energy_balance_has_its_standard_form(self) -> None:
        text = MODEL.read_text(encoding="utf-8")
        for name in ("D4BusbarAccountingActive", "D4BusbarDemandFuels", "D4InternalLossLinks"):
            self.assertNotIn(name, text)
        line = next(l for l in text.splitlines() if l.startswith("s.t. EBa11_EnergyBalanceEachTS5"))
        supply, demand = line.split(">=")
        self.assertNotIn("if ", supply)
        self.assertTrue(supply.rstrip().endswith("*YearSplit[l,y]"))

    def test_runner_writes_no_add_back_parameter(self) -> None:
        source = (REPO_ROOT / "ostram" / "pipeline" / "execution" / "runner.py").read_text(encoding="utf-8")
        self.assertNotIn("D4BusbarAccountingActive", source)
        self.assertIn("param D4HistoricalTradeActive := {trade_active};", source)


if __name__ == "__main__":
    unittest.main()
