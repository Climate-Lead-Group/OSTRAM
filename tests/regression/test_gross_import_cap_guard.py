"""The datafile conversion stops when a gross import cap would be dropped."""
from __future__ import annotations

import io
from contextlib import redirect_stdout
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest import mock

from ostram.pipeline.execution import gross_import_cap, runner
from ostram.pipeline.scenarios.rules import apply_base_year_pin


REPO_ROOT = Path(__file__).resolve().parents[2]
CONVERSION_FORMAT = (
    REPO_ROOT / "ostram" / "resources" / "compilation" / "conversion_format.yaml"
)
SCENARIO = "T_ImportCapGuard"
CAP_BLOCK = (
    "gross_import_cap:\n"
    "  countries: [BGD]\n"
    "  fraction: 0.15\n"
    "  first_year: 2030\n"
    "  last_year: 2030\n"
)


def _convert(root: Path, scenario_text: str, context: mock.Mock):
    """Convert one scenario in ``root`` against a model file without the
    historical-trade parameter, with the otoole conversion replaced by ``context``."""
    model = root / "model.txt"
    model.write_text("param Demand default 0;\nend;\n", encoding="utf-8")
    scenario_dir = root / "scenarios" / SCENARIO
    scenario_dir.mkdir(parents=True)
    (scenario_dir / "scenario.yaml").write_text(scenario_text, encoding="utf-8")
    inputs = root / "A2_Outputs_Params_otoole"
    (inputs / SCENARIO).mkdir(parents=True)
    (inputs / SCENARIO / "SpecifiedAnnualDemand.csv").write_text(
        "REGION,FUEL,YEAR,VALUE\nGLOBAL,ELCBGDXX03,2030,1.0\n", encoding="utf-8"
    )
    (inputs / SCENARIO / "OutputActivityRatio.csv").write_text(
        "REGION,TECHNOLOGY,FUEL,MODE_OF_OPERATION,YEAR,VALUE\n"
        "GLOBAL,TRNBGDXXINDEA,ELCBGDXX01,2,2030,1.0\n",
        encoding="utf-8",
    )
    project_paths = SimpleNamespace(
        maintained_model=model,
        scenario_config_root=root / "scenarios",
        scenario_workbook=REPO_ROOT / "inputs" / "scenarios" / "OSTRAM_Scenario_Inputs.xlsx",
        country_config=REPO_ROOT / "config" / "preparation" / "Config_country_codes.yaml",
    )
    params = {
        "executables": str(root / "Executables"),
        "Miscellaneous": str(CONVERSION_FORMAT.parent),
        "otoole_config": CONVERSION_FORMAT.name,
    }
    with (
        mock.patch.object(runner, "resolve_paths", return_value=project_paths),
        mock.patch.object(gross_import_cap, "resolve_paths", return_value=project_paths),
        mock.patch.object(apply_base_year_pin, "calibration_root", return_value=None),
        mock.patch.object(runner, "HERE", str(root), create=True),
        mock.patch("otoole.input.Context", return_value=context),
        redirect_stdout(io.StringIO()),
    ):
        return runner.run_otoole_conversion(str(inputs), SCENARIO, params)


class GrossImportCapGuardTests(unittest.TestCase):
    def test_cap_needs_the_trade_parameter_in_the_model_file(self) -> None:
        with self.subTest("scenario with the block"), tempfile.TemporaryDirectory() as temp:
            root = Path(temp).resolve()
            context = mock.Mock()
            with self.assertRaises(ValueError) as caught:
                _convert(root, f"scenario: {SCENARIO}\n{CAP_BLOCK}", context)
            message = str(caught.exception)
            self.assertIn(SCENARIO, message)
            self.assertIn("D4HistoricalTradeActive", message)
            self.assertIn(str(root / "model.txt"), message)
            context.convert.assert_not_called()

        with self.subTest("scenario without the block"), tempfile.TemporaryDirectory() as temp:
            context = mock.Mock()
            self.assertTrue(_convert(Path(temp).resolve(), f"scenario: {SCENARIO}\n", context))
            context.convert.assert_called_once()


if __name__ == "__main__":
    unittest.main()
