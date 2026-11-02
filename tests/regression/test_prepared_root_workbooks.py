"""Setups that start from ready-made root workbooks, and the two B2 settings they use.

Covers the ``prepared_root_workbooks`` policy (A1/A2 seeding and the A3 path that
applies only restrictions and rule scripts), the ``storage_delay_cap_params`` key and
the ``datafile_float_format`` key.
"""

from __future__ import annotations

import io
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import pandas as pd

from ostram.pipeline.execution import runner
from ostram.pipeline.execution.model_file_adjustments import storage_delay
from ostram.pipeline.preparation.prepared_roots import ROOT_WORKBOOKS, seed_root_snapshots
from ostram.pipeline.scenarios import orchestrator


REPO_ROOT = Path(__file__).resolve().parents[2]


class SeedRootSnapshotsTests(unittest.TestCase):
    def test_every_root_gets_output_and_snapshot_copies(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "root_workbooks"
            extra = root / "extra"
            source.mkdir()
            extra.mkdir()
            for name in ROOT_WORKBOOKS:
                (source / name).write_text(name, encoding="utf-8")
            (extra / "A-Xtra_Storage.xlsx").write_text("storage", encoding="utf-8")
            (extra / "notes.txt").write_text("not copied", encoding="utf-8")
            a1 = root / "A1_Outputs"
            stale = a1 / "A1_Outputs_BAU"
            stale.mkdir(parents=True)
            (stale / "stale.txt").write_text("old", encoding="utf-8")

            seed_root_snapshots(source, extra, a1, root / "extra_inputs", ["BAU", "A"])

            for scenario in ("BAU", "A"):
                for folder in (f"A1_Outputs_{scenario}", f"_post_a2_snapshot_{scenario}"):
                    names = sorted(p.name for p in (a1 / folder).iterdir())
                    self.assertEqual(names, sorted(ROOT_WORKBOOKS))
            self.assertEqual(
                [p.name for p in (root / "extra_inputs").iterdir()], ["A-Xtra_Storage.xlsx"]
            )

    def test_missing_workbook_fails_before_writing(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "src").mkdir()
            (root / "extra").mkdir()
            with self.assertRaises(FileNotFoundError):
                seed_root_snapshots(root / "src", root / "extra", root / "a1", root / "x", ["BAU"])
            self.assertFalse((root / "a1").exists())


class PreparedA3PathTests(unittest.TestCase):
    def _run(self, prepared: bool):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        root = Path(temp.name)
        prep = root / "preparation"
        snapshot = prep / "A1_Outputs" / "_post_a2_snapshot_BAU"
        snapshot.mkdir(parents=True)
        for name in ROOT_WORKBOOKS:
            (snapshot / name).write_text(name, encoding="utf-8")
        process = root / "scenarios"
        process.mkdir()
        soasia = root / "scenario.xlsx"
        soasia.write_bytes(b"workbook")
        workdir = process / "_run"
        stages = {key: workdir / key for key in ("stage1", "stage1b", "stage2", "stage3", "stage5")}

        def build_workdir(parent, ts, rules, scenario):
            for path in stages.values():
                path.mkdir(parents=True, exist_ok=True)
            return {"wd": workdir, "s1": stages["stage1"], "s1b": stages["stage1b"],
                    "s2": stages["stage2"], "s3": stages["stage3"], "s5": stages["stage5"]}

        import shutil
        calls = []
        record = lambda name: (lambda *a, **k: calls.append(name))
        deps = orchestrator.A3Dependencies(
            resolve_scenario_config=lambda args, path: ("BAU", ["rule.py"], []),
            resolve_path=Path,
            build_workdir=build_workdir,
            materialize_scenario_template=record("stage_0"),
            stage_1_scripts_1_to_5=record("stage_1"),
            stage_1b=record("stage_1b"),
            stage_2_and_2_5=record("stage_2"),
            stage_3_fix_2=lambda *a: calls.append("stage_3") or stages["stage3"] / "p.xlsx",
            stage_4_consolidate=record("stage_4"),
            stage_4_5_apply_inherited_restrictions=record("stage_4_5"),
            stage_5_rules_scripts=record("stage_5"),
            stage_interconnector_costs=record("interconnector_costs"),
            stage_internal_transmission=record("internal_transmission"),
            stage_internal_tx_losses=record("internal_tx_losses"),
            stage_base_year_pin=record("base_year_pin"),
            stage_internal_tx_residuals=record("internal_tx_residuals"),
            stage_6_sync_og_to_ts20=record("sync_og"),
            stage_6_persist_restrictions=record("stage_6"),
            deliver_outputs=record("deliver"),
            remove_tree=shutil.rmtree,
            copy_tree=shutil.copytree,
            copy_file=shutil.copy,
            environment={},
            clock=lambda: 0.0,
            timestamp_now=lambda: SimpleNamespace(strftime=lambda fmt: "20261101_000000"),
            banner=lambda message: None,
            emit=lambda message: None,
        )
        paths = orchestrator.A3Paths(prep, process, soasia, prepared_workbooks=prepared)
        args = SimpleNamespace(soasia=None, input_dir=None, output_dir=None, keep_workdir=True)
        orchestrator.orchestrate_a3(args, paths, deps, ROOT_WORKBOOKS)
        return calls, stages["stage5"]

    def test_prepared_workbooks_skip_preparation_and_calibration_stages(self) -> None:
        calls, stage5 = self._run(prepared=True)
        self.assertEqual(calls, ["stage_4_5", "stage_5", "stage_6", "deliver"])
        self.assertEqual(sorted(p.name for p in stage5.iterdir()), sorted(ROOT_WORKBOOKS))

    def test_default_path_keeps_every_stage(self) -> None:
        calls, _ = self._run(prepared=False)
        self.assertEqual(
            calls,
            ["stage_0", "stage_1", "stage_1b", "stage_2", "stage_3", "stage_4", "stage_4_5",
             "stage_5", "interconnector_costs", "internal_transmission", "internal_tx_losses",
             "internal_tx_residuals", "sync_og", "stage_6", "deliver"],
        )


DATAFILE = """set REGION :=
GLOBAL
;
set YEAR :=
2023
2024
;
set STORAGE :=
SDSAAAXX01
;
set TECHNOLOGY :=
PWRSDSAAAXX
;
param default 0 : TechnologyToStorage :=
GLOBAL PWRSDSAAAXX SDSAAAXX01 1 1
;
param default -1 : TotalAnnualMaxCapacity :=
GLOBAL PWRSDSAAAXX 2023 0.5
GLOBAL PWRSDSAAAXX 2024 0.5
;
param default -1 : TotalAnnualMaxCapacityInvestment :=
GLOBAL PWRSDSAAAXX 2023 2
GLOBAL PWRSDSAAAXX 2024 2
;
end;
"""


class StorageDelayCapParamsTests(unittest.TestCase):
    def _block(self, lines, name):
        start, end = storage_delay.find_param_block(lines, name)
        return [line.strip() for line in lines[start + 1:end]]

    def test_default_rewrites_both_caps(self) -> None:
        lines = DATAFILE.splitlines(True)
        out, summary = storage_delay.adjust_data_lines(lines, 1, [], ["SDS"], "-1")
        self.assertIn("GLOBAL PWRSDSAAAXX 2023 0", self._block(out, "TotalAnnualMaxCapacity"))
        self.assertEqual(set(summary["cap_stats"]), set(storage_delay.TECH_CAP_PARAMS))

    def test_investment_only_keeps_the_stock_cap(self) -> None:
        lines = DATAFILE.splitlines(True)
        out, summary = storage_delay.adjust_data_lines(
            lines, 1, [], ["SDS"], "-1", cap_params=("TotalAnnualMaxCapacityInvestment",)
        )
        self.assertEqual(
            self._block(out, "TotalAnnualMaxCapacity"),
            ["GLOBAL PWRSDSAAAXX 2023 0.5", "GLOBAL PWRSDSAAAXX 2024 0.5"],
        )
        self.assertIn("GLOBAL PWRSDSAAAXX 2023 0", self._block(out, "TotalAnnualMaxCapacityInvestment"))
        self.assertEqual(list(summary["cap_stats"]), ["TotalAnnualMaxCapacityInvestment"])

    def test_unknown_cap_parameter_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            storage_delay.adjust_data_lines(
                DATAFILE.splitlines(True), 1, [], ["SDS"], "-1", cap_params=("CapitalCost",)
            )

    def test_runner_forwards_the_key(self) -> None:
        params = {
            "storage_delay_active": True,
            "storage_delay_cap_params": ["TotalAnnualMaxCapacityInvestment"],
            "preprocess_data_name": "Pre_processed_",
            "executables": "Executables",
            "osemosys_model": str(REPO_ROOT / "model" / "osemosys_fast_preprocessed.txt"),
        }
        stage = mock.Mock(return_value=SimpleNamespace(returncode=0, stdout="", stderr=""))
        with mock.patch.object(runner, "_run_stage_command", stage), \
                mock.patch.object(runner, "_resolve_storage_delay_model_output_path",
                                  return_value=Path("model_out.txt")):
            runner.run_storage_delay_adjustment(params, "BAU")
        command = stage.call_args.args[0]
        index = command.index("--cap-params")
        self.assertEqual(command[index + 1:index + 2], ["TotalAnnualMaxCapacityInvestment"])


class DatafileFloatFormatTests(unittest.TestCase):
    def _writer(self, params):
        captured = {}

        class Stop(Exception):
            pass

        def fake_context(reader, writer):
            captured["writer"] = writer
            raise Stop

        params = {
            "executables": tempfile.mkdtemp(),
            "Miscellaneous": str(REPO_ROOT / "ostram" / "resources" / "compilation"),
            "otoole_config": "conversion_format.yaml",
            **params,
        }
        with mock.patch("otoole.input.Context", fake_context),                 mock.patch.object(runner, "HERE", "", create=True):
            with self.assertRaises(Stop):
                runner.run_otoole_conversion(tempfile.mkdtemp(), "BAU", params)
        return captured["writer"]

    def _written(self, writer) -> str:
        frame = pd.DataFrame(
            {"REGION": ["GLOBAL"], "TECHNOLOGY": ["PWRX"], "YEAR": [2023], "VALUE": [1234.56789]}
        ).set_index(["REGION", "TECHNOLOGY", "YEAR"])
        handle = io.StringIO()
        writer._write_parameter(frame, "CapitalCost", handle, 0)
        return handle.getvalue()

    def test_default_is_full_precision(self) -> None:
        self.assertIn("1234.56789", self._written(self._writer({})))

    def test_stock_format_gives_six_significant_digits(self) -> None:
        text = self._written(self._writer({"datafile_float_format": "%g"}))
        self.assertIn("GLOBAL PWRX 2023 1234.57", text)


if __name__ == "__main__":
    unittest.main()
