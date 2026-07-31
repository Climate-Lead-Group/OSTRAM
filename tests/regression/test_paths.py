from __future__ import annotations

import os
from pathlib import Path
import tempfile
import unittest
import uuid
from unittest import mock

from ostram import paths as path_module


REPO_ROOT = Path(__file__).resolve().parents[2]


class ProjectPathResolutionTests(unittest.TestCase):
    def test_cli_project_root_precedes_environment_and_anchor(self) -> None:
        resolved = path_module.resolve_paths(
            project_root=REPO_ROOT,
            environ={
                path_module.PROJECT_ROOT_ENV: str(REPO_ROOT / "missing"),
                path_module.WORKSPACE_ENV: str(REPO_ROOT / "env-work"),
            },
        )
        self.assertEqual(resolved.project_root, REPO_ROOT)
        self.assertEqual(resolved.workspace, REPO_ROOT / "env-work")

    def test_environment_project_root_precedes_package_anchor(self) -> None:
        resolved = path_module.resolve_paths(
            environ={path_module.PROJECT_ROOT_ENV: str(REPO_ROOT)}
        )
        self.assertEqual(resolved.project_root, REPO_ROOT)

    def test_workspace_precedence_and_unicode_space_backslash_paths(self) -> None:
        cli_workspace = REPO_ROOT / "workspace path Ω" / "drive-C"
        resolved = path_module.resolve_paths(
            project_root=str(REPO_ROOT).replace("/", "\\"),
            workspace=cli_workspace,
            environ={path_module.WORKSPACE_ENV: str(REPO_ROOT / "ignored")},
        )
        self.assertEqual(resolved.workspace, cli_workspace.resolve())
        self.assertTrue(resolved.project_root.is_absolute())

    def test_default_workspace_is_lazy(self) -> None:
        resolved = path_module.resolve_paths(project_root=REPO_ROOT, environ={})
        stage = resolved.stage_workspace(
            "stage11-lazy-test", f"Scenario With Space {uuid.uuid4().hex}"
        )
        self.assertFalse(stage.exists())
        created = resolved.stage_workspace(stage.parent.name, stage.name, create=True)
        self.assertTrue(created.is_dir())

    def test_non_editable_install_without_bundle_fails_clearly(self) -> None:
        with mock.patch.object(
            path_module, "_source_anchor", return_value=Path("C:/site-packages")
        ):
            with self.assertRaisesRegex(
                path_module.ProjectResolutionError,
                "non-editable.*project bundle",
            ):
                path_module.resolve_paths(environ={})

    def test_invalid_explicit_bundle_fails_without_creating_directories(self) -> None:
        with tempfile.TemporaryDirectory(dir=REPO_ROOT / "workspace") as temp:
            candidate = Path(temp) / "not-a-project"
            with self.assertRaisesRegex(
                path_module.ProjectResolutionError, "valid OSTRAM project bundle"
            ):
                path_module.resolve_paths(project_root=candidate, environ={})
            self.assertFalse(candidate.exists())

    def test_representative_real_resources_are_read_absolute_and_read_only(self) -> None:
        resolved = path_module.resolve_paths(project_root=REPO_ROOT, environ={})
        before = {
            path: path.stat().st_mtime_ns
            for path in (
                resolved.scenario_registry,
                resolved.scenario_workbook,
                resolved.timeslice_workbook,
                resolved.compilation_config,
                resolved.execution_config,
                resolved.maintained_model,
            )
        }
        record = resolved.inspect_resources()
        after = {path: path.stat().st_mtime_ns for path in before}
        self.assertEqual(before, after)
        self.assertEqual(record["registry_schema"], "ostram-scenario-registry-v1")
        self.assertEqual(record["scenario_workbook_signature"], "504b0304")
        self.assertEqual(
            record["root_scenarios"],
            ["BAU", "A_Calibrated_BAU", "B_Optimised_VRE", "C_Target_VRE"],
        )
        for key in (
            "project_root",
            "workspace",
            "scenario_workbook",
            "execution_config",
            "maintained_model",
            "package_resources",
        ):
            self.assertTrue(Path(str(record[key])).is_absolute(), key)


if __name__ == "__main__":
    unittest.main()
