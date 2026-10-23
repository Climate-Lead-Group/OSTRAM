"""The import cap at existing link capacity on the two India-Nepal links.

The model file holds the constraint and its gate, the solver datafile of a
scenario with active trade bounds names both links in their India-to-Nepal
mode from 2027, and a scenario without active trade bounds gets no cap.
"""
from __future__ import annotations

import io
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import yaml

from ostram.pipeline.execution import gross_import_cap, link_import_cap, runner


REPO_ROOT = Path(__file__).resolve().parents[2]
MODEL = REPO_ROOT / "model" / "osemosys_fast_preprocessed.txt"
POLICY = yaml.safe_load((REPO_ROOT / "config/setups/south_asia.yaml").read_text(encoding="utf-8"))[
    "policies"]["link_import_cap"]
CONVERSION_FORMAT = REPO_ROOT / "ostram" / "resources" / "compilation" / "conversion_format.yaml"
EXPECTED = "set D4LinkImportCap := (TRNINDEANPLXX,1,2027) (TRNINDNONPLXX,1,2027);"


def _ratios(folder: Path) -> None:
    rows_in, rows_out = [], []
    for link, india in (("TRNINDEANPLXX", "INDEA"), ("TRNINDNONPLXX", "INDNO")):
        for year in (2027, 2050):
            rows_in += [f"GLOBAL,{link},ELC{india}03,1,{year},1.0", f"GLOBAL,{link},ELCNPLXX03,2,{year},1.0"]
            rows_out += [f"GLOBAL,{link},ELCNPLXX04,1,{year},0.96", f"GLOBAL,{link},ELC{india}04,2,{year},0.96"]
    head = "REGION,TECHNOLOGY,FUEL,MODE_OF_OPERATION,YEAR,VALUE\n"
    (folder / "InputActivityRatio.csv").write_text(head + "\n".join(rows_in) + "\n", encoding="utf-8")
    (folder / "OutputActivityRatio.csv").write_text(head + "\n".join(rows_out) + "\n", encoding="utf-8")


class LinkImportCapTests(unittest.TestCase):
    def test_policy_names_the_two_nepal_links_from_2027(self) -> None:
        self.assertEqual(POLICY, {"links": ["TRNINDEANPLXX", "TRNINDNONPLXX"], "from_country": "IND",
                                  "to_country": "NPL", "first_year": 2027})

    def test_model_holds_the_gated_constraint_and_keeps_the_trade_lines(self) -> None:
        text = MODEL.read_text(encoding="utf-8")
        lines = text.splitlines()
        self.assertEqual(sum(line.startswith("s.t. D4_I1_TRADE_") for line in lines), 50)
        declaration = [line for line in lines if line.startswith("set D4LinkImportCap ")]
        self.assertEqual(len(declaration), 1)
        self.assertIn("default {t in TECHNOLOGY, m in MODE_OF_OPERATION, y in YEAR: 1 = 0}", declaration[0])
        constraint = [line for line in lines if line.startswith("s.t. D4_LinkImportCap{")]
        self.assertEqual(len(constraint), 1)
        self.assertIn("D4HistoricalTradeActive = 1 and y >= y0", constraint[0])
        self.assertIn("<= ResidualCapacity[r,t,y]*CapacityToActivityUnit[r,t]*sum{l in TIMESLICE} "
                      "CapacityFactor[r,t,l,y]*YearSplit[l,y];", constraint[0])
        self.assertLess(text.index("param D4HistoricalTradeActive"), text.index("s.t. D4_LinkImportCap{"))

    def test_render_names_both_links_in_the_india_to_nepal_mode(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            _ratios(folder)
            with mock.patch.object(link_import_cap, "setup_policy", return_value=POLICY):
                self.assertEqual(link_import_cap.render_link_cap(folder, MODEL.read_text(encoding="utf-8")).strip(),
                                 EXPECTED)
                with self.assertRaises(ValueError):
                    link_import_cap.render_link_cap(folder, "end;\n")
            with mock.patch.object(link_import_cap, "setup_policy", return_value=None):
                self.assertEqual(link_import_cap.render_link_cap(folder, ""), "")

    def test_datafile_carries_the_cap_only_where_trade_bounds_are_active(self) -> None:
        for active in (1, 0):
            with self.subTest(active=active), tempfile.TemporaryDirectory() as temp:
                root = Path(temp).resolve()
                inputs = root / "A2_Outputs_Params_otoole"
                (inputs / "T_Link").mkdir(parents=True)
                _ratios(inputs / "T_Link")
                output = root / "Executables" / "T_Link_0" / "T_Link_0.txt"

                def convert(_input, out):
                    Path(out).write_text("end;\n", encoding="utf-8")

                context = mock.Mock()
                context.convert.side_effect = convert
                paths = SimpleNamespace(maintained_model=MODEL, scenario_config_root=root / "scenarios")
                params = {"executables": str(root / "Executables"), "Miscellaneous": str(CONVERSION_FORMAT.parent),
                          "otoole_config": CONVERSION_FORMAT.name}
                with (
                    mock.patch.object(runner, "resolve_paths", return_value=paths),
                    mock.patch.object(gross_import_cap, "resolve_paths", return_value=paths),
                    mock.patch.object(runner, "historical_trade_activation", return_value=active),
                    mock.patch.object(link_import_cap, "setup_policy", return_value=POLICY),
                    mock.patch.object(runner, "HERE", str(root), create=True),
                    mock.patch("otoole.input.Context", return_value=context),
                    redirect_stdout(io.StringIO()),
                ):
                    self.assertTrue(runner.run_otoole_conversion(str(inputs), "T_Link", params))
                text = output.read_text(encoding="utf-8")
                self.assertEqual(EXPECTED in text, bool(active))
                self.assertNotEqual("D4LinkImportCap" in text, not active)


if __name__ == "__main__":
    unittest.main()
