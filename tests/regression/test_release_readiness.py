from __future__ import annotations

from contextlib import redirect_stdout
import html
from html.parser import HTMLParser
import io
from pathlib import Path
import re
import tempfile
from types import SimpleNamespace
import unittest
from unittest import mock

import pandas as pd

from ostram import __main__ as cli
from ostram.examples import _CAPTURE, _comparison_labels, _report
from ostram.pipeline.execution import runner as execution_runner
from ostram.pipeline.execution.orchestrator import validate_cbc_solution
from ostram.pipeline.preparation import merge_country_template as country_merge
from ostram.pipeline.preparation.scenario_country_sync import _default_ao


REPO_ROOT = Path(__file__).resolve().parents[2]
EXAMPLE_ROOT = REPO_ROOT / "examples" / "unescap"


class _Links(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.hrefs: list[str] = []
        self.ids: set[str] = set()

    def handle_starttag(self, tag, attrs) -> None:
        values = dict(attrs)
        if "id" in values:
            self.ids.add(values["id"])
        if tag == "a" and "href" in values:
            self.hrefs.append(values["href"])


def _parse_html(path: Path) -> _Links:
    parser = _Links()
    parser.feed(path.read_text(encoding="utf-8"))
    return parser


class PortableRuntimeTests(unittest.TestCase):
    def test_environment_executable_discovers_extensionless_posix_binary(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            env_root = Path(temp) / "env"
            python = env_root / "python"
            cbc = env_root / "bin" / "cbc"
            cbc.parent.mkdir(parents=True)
            python.touch()
            cbc.touch()
            with (
                mock.patch.object(execution_runner.sys, "executable", str(python)),
                mock.patch.object(execution_runner, "ensure_env_tool_paths"),
            ):
                self.assertEqual(
                    Path(execution_runner.get_env_executable("cbc")), cbc.resolve()
                )

    def test_missing_solver_fails_closed_without_shell_lookup_commands(self) -> None:
        with (
            mock.patch.object(execution_runner, "ensure_env_tool_paths"),
            mock.patch.object(execution_runner, "get_env_executable", return_value="cbc"),
            mock.patch.object(execution_runner.shutil, "which", return_value=None),
            self.assertRaisesRegex(FileNotFoundError, "solver executable 'cbc'"),
        ):
            execution_runner.check_enviro_variables("cbc")

    def test_cbc_solution_status_rejects_zero_exit_infeasibility(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            solution = Path(temp) / "model.sol"
            solution.write_text(
                "Infeasible - objective value 559384.17921019\n",
                encoding="utf-8",
            )
            with self.assertRaisesRegex(RuntimeError, "did not produce an optimal"):
                validate_cbc_solution(solution)
            solution.write_text(
                "Optimal - objective value 123.5\n",
                encoding="utf-8",
            )
            self.assertEqual(
                validate_cbc_solution(solution),
                "Optimal - objective value 123.5",
            )

    def test_affected_runtime_has_no_shell_or_drive_letter_assumptions(self) -> None:
        for relative in (
            "ostram/examples.py",
            "ostram/pipeline/execution/runner.py",
            "ostram/pipeline/preparation/scenario_country_sync.py",
            "ostram/reporting/training_dashboard.py",
        ):
            source = (REPO_ROOT / relative).read_text(encoding="utf-8")
            self.assertNotIn("shell=True", source, relative)
            self.assertNotIn("cmd.exe", source.lower(), relative)
            self.assertNotIn("powershell", source.lower(), relative)
            self.assertNotRegex(source, r"[A-Za-z]:[\\/]", relative)

    def test_country_modules_publish_no_legacy_direct_script_commands(self) -> None:
        for relative in (
            "ostram/pipeline/preparation/country_templates.py",
            "ostram/pipeline/preparation/country_validation.py",
        ):
            source = (REPO_ROOT / relative).read_text(encoding="utf-8")
            self.assertNotRegex(source, r"python\s+Z_[A-Za-z0-9_]+\.py", relative)


class CanonicalTrainingRouteTests(unittest.TestCase):
    def test_country_merge_targets_profile_authority_without_changing_full_default(self) -> None:
        paths = SimpleNamespace(
            osemosys_inputs=Path("prepared-authority"),
            preparation_workspace=Path("legacy-preparation"),
        )
        with mock.patch.object(country_merge, "active_profile_id", return_value="unescap"):
            self.assertEqual(
                country_merge._default_input_dir(paths),
                Path("prepared-authority"),
            )
        with mock.patch.object(country_merge, "active_profile_id", return_value="full"):
            self.assertEqual(
                country_merge._default_input_dir(paths),
                Path("legacy-preparation") / "og_csvs_inputs",
            )

    def test_profile_subcommand_help_needs_no_prepared_workspace(self) -> None:
        commands = (
            ["example", "report", "unescap", "--help"],
            ["--profile", "unescap", "country", "template", "--help"],
            ["--profile", "unescap", "country", "merge", "--help"],
            ["--profile", "unescap", "country", "validate", "--help"],
            ["--profile", "unescap", "scenario", "sync-country", "--help"],
        )
        with tempfile.TemporaryDirectory() as temp:
            workspace = Path(temp) / "never-created"
            for command in commands:
                with self.subTest(command=command), redirect_stdout(io.StringIO()):
                    with self.assertRaises(SystemExit) as raised:
                        cli.main(["--workspace", str(workspace), *command])
                    self.assertEqual(raised.exception.code, 0)
            self.assertFalse(workspace.exists())

    def test_default_country_sync_uses_canonical_bau_with_multiple_snapshots(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            outputs = Path(temp)
            for scenario in ("BAU", "A_Calibrated_BAU", "B_Optimised_VRE"):
                path = outputs / f"_post_a2_snapshot_{scenario}" / "A-O_Parametrization.xlsx"
                path.parent.mkdir()
                path.touch()
            selected = _default_ao(SimpleNamespace(a1_outputs=outputs))
            self.assertEqual(
                selected,
                outputs / "_post_a2_snapshot_BAU" / "A-O_Parametrization.xlsx",
            )

    def test_report_compare_labels_are_fail_closed(self) -> None:
        self.assertEqual(
            _comparison_labels("forward, reverse,bidirectional"),
            ("forward", "reverse", "bidirectional"),
        )
        for invalid in ("forward", "forward,forward", "forward,../reverse"):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                _comparison_labels(invalid)

    def test_capture_and_selected_comparison_use_only_workspace_results(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            workspace = Path(temp)
            execution = workspace / "execution"
            execution.mkdir()
            result = execution / "OSTRAM_Combined_Inputs_Outputs.csv"
            pd.DataFrame([
                {
                    "Scenario": "B_Optimised_VRE",
                    "REGION": "GLOBAL",
                    "YEAR": 2050,
                    "TECHNOLOGY": "TRNBGDXXINDEA",
                    "TotalCapacityAnnual": 2.496,
                }
            ]).to_csv(result, index=False)
            paths = SimpleNamespace(workspace=workspace, execution_workspace=execution)
            manifest = SimpleNamespace(
                profile_id="unescap",
                path=EXAMPLE_ROOT / "profile.yaml",
                metadata={
                    "country_regions": [
                        {"region": "BGDXX", "label": "Bangladesh"},
                        {"region": "INDEA", "label": "India East"},
                    ],
                    "interconnectors": [{"technology": "TRNBGDXXINDEA"}],
                },
            )
            with mock.patch("ostram.examples.resolve_paths", return_value=paths):
                _report(manifest, "forward")
                _report(manifest, "reverse")
                compared = _report(manifest, None, "forward,reverse")
                with self.assertRaises(FileExistsError):
                    _report(manifest, "forward")
            self.assertEqual(
                compared,
                workspace / "reports" / "unescap-interconnector-comparison.html",
            )
            html = compared.read_text(encoding="utf-8")
            self.assertIn('"forward"', html)
            self.assertIn('"reverse"', html)


class TrainingAssetTests(unittest.TestCase):
    def test_live_docs_contain_no_retired_or_invented_commands(self) -> None:
        live_documents = [
            EXAMPLE_ROOT / "README.md",
            EXAMPLE_ROOT / "AUTHORING_AND_ACCEPTANCE.md",
            *sorted((EXAMPLE_ROOT / "exercises").glob("*.html")),
        ]
        text = "\n".join(
            path.read_text(encoding="utf-8")
            for path in live_documents
        )
        for retired in (
            "--rebuild",
            "--force",
            "country populate-workbook",
            "scenario list",
            "E1+E2",
        ):
            self.assertNotIn(retired, text)

    def test_every_documented_capture_label_is_accepted_by_the_cli(self) -> None:
        documents = [
            EXAMPLE_ROOT / "README.md",
            EXAMPLE_ROOT / "AUTHORING_AND_ACCEPTANCE.md",
            *sorted((EXAMPLE_ROOT / "exercises").glob("*.html")),
        ]
        labels: list[str] = []
        for path in documents:
            source = path.read_text(encoding="utf-8")
            plain = html.unescape(re.sub(r"<[^>]+>", "", source))
            labels.extend(
                re.findall(r"--capture\s+[\"']?([A-Za-z0-9_.+-]+)", plain)
            )
        self.assertTrue(labels)
        self.assertEqual(
            [label for label in labels if not _CAPTURE.fullmatch(label)],
            [],
        )

    def test_all_internal_html_links_resolve_and_fragments_exist(self) -> None:
        parsed = {path.resolve(): _parse_html(path) for path in EXAMPLE_ROOT.rglob("*.html")}
        failures: list[str] = []
        for source, document in parsed.items():
            for href in document.hrefs:
                if not href or href.startswith(("http://", "https://", "mailto:", "/")):
                    continue
                target_text, _, fragment = href.partition("#")
                target = source if not target_text else (source.parent / target_text).resolve()
                if not target.exists():
                    failures.append(f"{source.relative_to(REPO_ROOT)} -> missing {href}")
                    continue
                if fragment and target.suffix.lower() == ".html":
                    target_doc = parsed.get(target) or _parse_html(target)
                    if fragment not in target_doc.ids:
                        failures.append(
                            f"{source.relative_to(REPO_ROOT)} -> missing fragment {href}"
                        )
        self.assertEqual(failures, [])


if __name__ == "__main__":
    unittest.main()
