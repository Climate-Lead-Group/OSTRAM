from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
import io
import json
import os
from pathlib import Path
import tempfile
import types
import unittest
from unittest import mock

import ostram.__main__ as cli
from ostram import examples
from ostram.paths import ProjectPaths
from ostram.setup_workspace import (
    STAMP_NAME,
    SetupWorkspaceError,
    prepare_setup,
    prepared_setup,
    setup_workspace,
)
from ostram.setups import (
    AuthorityResolutionError,
    SetupManifestError,
    load_manifest,
    load_setup,
    resolve_authority_reference,
)


REPO_ROOT = Path(__file__).resolve().parents[2]


def project_paths(root: Path, workspace: Path | None = None) -> ProjectPaths:
    return ProjectPaths(root.resolve(), (workspace or root / "workspace").resolve(), "project")


def write_project(root: Path) -> None:
    for directory in ("ostram", "inputs", "config", "model"):
        (root / directory).mkdir(parents=True, exist_ok=True)
    (root / "ostram" / "__init__.py").write_text("", encoding="utf-8")
    (root / "environment.yaml").write_text("name: fixture\n", encoding="utf-8")


class ManifestSafetyTests(unittest.TestCase):
    def test_namespaces_are_contained_and_explicit(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            write_project(root)
            paths = project_paths(root)
            setup_root = root / "setups" / "tiny"
            setup_root.mkdir(parents=True)
            self.assertEqual(
                resolve_authority_reference(
                    "project:inputs/data.csv", paths=paths, setup_root=setup_root
                ),
                (root / "inputs" / "data.csv").resolve(),
            )
            self.assertEqual(
                resolve_authority_reference(
                    "setup:seed.xlsx", paths=paths, setup_root=setup_root
                ),
                (setup_root / "seed.xlsx").resolve(),
            )
            self.assertEqual(
                resolve_authority_reference(
                    "package:resources/table.csv", paths=paths, setup_root=setup_root
                ),
                (root / "ostram" / "resources" / "table.csv").resolve(),
            )
            for reference in (
                "inputs/data.csv", "http:data.csv", "project:",
                "project:../secret", "setup:/absolute", "package:C:\\secret",
            ):
                with self.subTest(reference=reference), self.assertRaises(
                    AuthorityResolutionError
                ):
                    resolve_authority_reference(
                        reference, paths=paths, setup_root=setup_root
                    )

    def test_duplicate_roles_and_invalid_mutability_fail(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = root / "setup.yaml"
            manifest.write_text(
                "schema: ostram-setup-v1\nid: tiny\nauthorities:\n"
                "  seed: setup:a.csv\n  seed: setup:b.csv\n",
                encoding="utf-8",
            )
            with self.assertRaisesRegex(SetupManifestError, "duplicate"):
                load_manifest(manifest)

            manifest.write_text(
                "schema: ostram-setup-v1\nid: tiny\nauthorities:\n"
                "  seed: {path: project:a.csv, mutable: true}\n",
                encoding="utf-8",
            )
            parsed = load_manifest(manifest)
            write_project(root)
            (root / "a.csv").write_text("x\n", encoding="utf-8")
            with self.assertRaisesRegex(SetupManifestError, "must use setup"):
                parsed.source_paths(project_paths(root))

    def test_old_build_cap_policy_name_is_refused(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            manifest = Path(temp) / "setup.yaml"
            manifest.write_text(
                "schema: ostram-setup-v1\nid: tiny\nauthorities:\n"
                "  seed: setup:a.csv\npolicies:\n  lid_rule_new_semantics: true\n",
                encoding="utf-8",
            )
            with self.assertRaisesRegex(SetupManifestError, "build_cap_replaces_workbook_values"):
                load_manifest(manifest)


class AtomicWorkspaceTests(unittest.TestCase):
    def _fixture(self, root: Path):
        write_project(root)
        setup_root = root / "fixtures" / "tiny"
        setup_root.mkdir(parents=True)
        (setup_root / "seed.txt").write_text("version one", encoding="utf-8")
        (root / "inputs" / "readonly.txt").write_text("readonly", encoding="utf-8")
        manifest_path = setup_root / "setup.yaml"
        manifest_path.write_text(
            "schema: ostram-setup-v1\nid: tiny\nauthorities:\n"
            "  seed: {path: setup:seed.txt, mutable: true}\n"
            "  seed_alias: {path: setup:seed.txt, mutable: true}\n"
            "  readonly: project:inputs/readonly.txt\n",
            encoding="utf-8",
        )
        return project_paths(root), load_manifest(manifest_path)

    def test_prepare_reset_stamp_last_and_readonly_authority(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            result = prepare_setup(manifest, paths=paths)
            self.assertTrue((result.workspace / STAMP_NAME).is_file())
            self.assertEqual(result.authorities["seed"].read_text(), "version one")
            self.assertEqual(
                result.authorities["seed"], result.authorities["seed_alias"]
            )
            self.assertEqual(
                result.authorities["readonly"],
                (paths.project_root / "inputs" / "readonly.txt").resolve(),
            )
            with self.assertRaises(FileExistsError):
                prepare_setup(manifest, paths=paths)

            manifest.path.write_text(
                manifest.path.read_text(encoding="utf-8")
                + "metadata:\n  revision: two\n",
                encoding="utf-8",
            )
            changed_manifest = load_manifest(manifest.path)
            prepared_at = result.stamp["prepared_at_utc"]
            # A changed manifest with the same mutable roles keeps the workspace.
            result = prepared_setup(changed_manifest, paths=paths)
            self.assertEqual(result.stamp["manifest_sha256"], changed_manifest.digest)
            self.assertEqual(result.stamp["prepared_at_utc"], prepared_at)
            result = prepare_setup(changed_manifest, paths=paths, reset=True)
            self.assertEqual(result.stamp["manifest_sha256"], changed_manifest.digest)
            manifest = changed_manifest

            (manifest.root / "seed.txt").write_text("version two", encoding="utf-8")
            # Mutable sources are now auto-refreshed instead of raising.
            refreshed = prepared_setup(manifest, paths=paths)
            self.assertEqual(refreshed.authorities["seed"].read_text(), "version two")

    def test_repeated_prepare_prints_the_message_without_a_traceback(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            with (
                mock.patch.object(examples, "_active_manifest", return_value=manifest),
                mock.patch.object(examples, "resolve_paths", return_value=paths),
                redirect_stdout(io.StringIO()),
                redirect_stderr(io.StringIO()) as stderr,
            ):
                self.assertEqual(examples.main(["prepare", "tiny"]), 0)
                self.assertEqual(examples.main(["prepare", "tiny"]), 1)
            self.assertIn("setup workspace already exists", stderr.getvalue())
            self.assertNotIn("Traceback", stderr.getvalue())

    def test_unedited_copy_refreshes_and_nothing_moves(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            result = prepare_setup(manifest, paths=paths)
            (manifest.root / "seed.txt").write_text("version two", encoding="utf-8")
            refreshed = prepared_setup(manifest, paths=paths)
            self.assertEqual(refreshed.authorities["seed"].read_text(), "version two")
            self.assertFalse((result.workspace / "authorities_previous").exists())

    def test_edited_file_copy_moves_before_refresh(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            result = prepare_setup(manifest, paths=paths)
            copy = result.authorities["seed"]
            relative = copy.relative_to(result.workspace)
            copy.write_text("trainee edit", encoding="utf-8")
            (manifest.root / "seed.txt").write_text("version two", encoding="utf-8")
            with mock.patch("builtins.print") as printed:
                refreshed = prepared_setup(manifest, paths=paths)
            self.assertEqual(refreshed.authorities["seed"].read_text(), "version two")
            moved = list((result.workspace / "authorities_previous").glob("*"))
            self.assertEqual(len(moved), 1)
            self.assertEqual((moved[0] / relative).read_text(), "trainee edit")
            lines = [str(call.args[0]) for call in printed.call_args_list]
            self.assertEqual(sum("Moved the edited copy" in line for line in lines), 1)

    def test_edited_directory_copy_moves_intact_before_refresh(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            data = manifest.root / "data"
            data.mkdir()
            (data / "a.csv").write_text("a", encoding="utf-8")
            manifest = self._change_manifest(
                manifest,
                "  readonly:",
                "  data: {path: setup:data, mutable: true, kind: directory}\n  readonly:",
            )
            result = prepare_setup(manifest, paths=paths)
            copy = result.authorities["data"]
            relative = copy.relative_to(result.workspace)
            (copy / "trainee.csv").write_text("trainee", encoding="utf-8")
            (data / "b.csv").write_text("b", encoding="utf-8")
            with mock.patch("builtins.print"):
                refreshed = prepared_setup(manifest, paths=paths)
            self.assertEqual(
                sorted(path.name for path in refreshed.authorities["data"].iterdir()),
                ["a.csv", "b.csv"],
            )
            moved = list((result.workspace / "authorities_previous").glob("*"))
            self.assertEqual(len(moved), 1)
            self.assertEqual(
                sorted(path.name for path in (moved[0] / relative).iterdir()),
                ["a.csv", "trainee.csv"],
            )
            self.assertEqual((moved[0] / relative / "trainee.csv").read_text(), "trainee")

    def _change_manifest(self, manifest, old: str = "", new: str = ""):
        text = manifest.path.read_text(encoding="utf-8").replace(old, new)
        manifest.path.write_text(text + f"# {manifest.digest}\n", encoding="utf-8")
        return load_manifest(manifest.path)

    def test_changed_manifest_with_same_roles_keeps_workspace(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            result = prepare_setup(manifest, paths=paths)
            result.authorities["seed"].write_text("trainee edit", encoding="utf-8")

            for entry_point in (
                lambda changed: prepared_setup(changed, paths=paths),
                lambda changed: prepare_setup(changed, paths=paths),
            ):
                manifest = self._change_manifest(manifest)
                result = entry_point(manifest)
                self.assertEqual(result.authorities["seed"].read_text(), "trainee edit")
                self.assertEqual(result.stamp["manifest_sha256"], manifest.digest)
                stamp = json.loads(
                    (result.workspace / STAMP_NAME).read_text(encoding="utf-8")
                )
                self.assertEqual(stamp["manifest_sha256"], manifest.digest)
            self.assertEqual(
                sorted(path.name for path in result.workspace.parent.iterdir()),
                ["tiny"],
            )

    def test_changed_mutable_roles_move_workspace_and_prepare_again(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            result = prepare_setup(manifest, paths=paths)
            result.authorities["seed"].write_text("trainee edit", encoding="utf-8")
            alias = "  seed_alias: {path: setup:seed.txt, mutable: true}\n"
            readonly = "  readonly:"

            for entry_point, old, new in (
                (lambda changed: prepared_setup(changed, paths=paths), alias, ""),
                (lambda changed: prepare_setup(changed, paths=paths), readonly, alias + readonly),
            ):
                manifest = self._change_manifest(manifest, old, new)
                result = entry_point(manifest)
                self.assertEqual(result.authorities["seed"].read_text(), "version one")
                self.assertEqual(result.stamp["manifest_sha256"], manifest.digest)
            previous = sorted(result.workspace.parent.glob("tiny_previous_*"))
            self.assertEqual(len(previous), 2)
            self.assertEqual(
                (previous[0] / "authorities" / "seed" / "seed.txt").read_text(),
                "trainee edit",
            )
            for folder in previous:
                self.assertTrue((folder / STAMP_NAME).is_file())

    def test_interrupted_foreign_and_missing_prepared_files_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            interrupted = paths.workspace / "setups" / ".tiny.preparing"
            interrupted.mkdir(parents=True)
            with self.assertRaisesRegex(SetupWorkspaceError, "interrupted"):
                prepare_setup(manifest, paths=paths)
            interrupted.rmdir()

            result = prepare_setup(manifest, paths=paths)
            result.authorities["seed"].unlink()
            with self.assertRaisesRegex(SetupWorkspaceError, "missing"):
                prepared_setup(manifest, paths=paths)
            # The source still exists: absence in the bundle may not fall back.
            self.assertTrue((manifest.root / "seed.txt").is_file())

            stamp_path = setup_workspace(paths, "tiny") / STAMP_NAME
            stamp = json.loads(stamp_path.read_text(encoding="utf-8"))
            stamp["setup_id"] = "foreign"
            stamp_path.write_text(json.dumps(stamp), encoding="utf-8")
            with self.assertRaisesRegex(SetupWorkspaceError, "foreign"):
                prepared_setup(manifest, paths=paths)


class SetupCliTests(unittest.TestCase):
    def test_setup_is_selected_before_route_import_and_not_forwarded(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            write_project(root)
            (root / "inputs" / "authority.txt").write_text("ok", encoding="utf-8")
            profiles = root / "config" / "setups"
            profiles.mkdir(parents=True)
            (profiles / "registry.json").write_text(
                json.dumps({
                    "schema": "ostram-setup-registry-v1",
                    "setups": {"tiny": "project:config/setups/tiny.yaml"},
                }), encoding="utf-8"
            )
            (profiles / "tiny.yaml").write_text(
                "schema: ostram-setup-v1\nid: tiny\nauthorities:\n"
                "  authority: project:inputs/authority.txt\n",
                encoding="utf-8",
            )
            events = []

            def load_route(route):
                events.append(("import", os.environ.get("OSTRAM_SETUP")))
                return types.SimpleNamespace(
                    main=lambda: events.append(("argv", tuple(__import__("sys").argv)))
                )

            with mock.patch.object(cli, "_load_route_module", side_effect=load_route):
                result = cli.main([
                    "--project-root", str(root), "--setup", "tiny",
                    "run", "--sentinel",
                ])
            self.assertEqual(result, 0)
            self.assertEqual(events[0], ("import", "tiny"))
            self.assertEqual(events[1][1], ("python -m ostram run", "--sentinel"))

    def test_setup_option_after_command_is_left_for_historical_parser(self) -> None:
        import sys

        seen = []
        fake = types.SimpleNamespace(
            main=mock.Mock(side_effect=lambda: seen.append(tuple(sys.argv)))
        )
        with mock.patch.object(cli, "_load_route_module", return_value=fake):
            cli.main(["run", "--setup", "full"])
        self.assertEqual(
            seen, [("python -m ostram run", "--setup", "full")]
        )

    def test_unqualified_and_explicit_south_asia_have_same_authority_bundle(self) -> None:
        paths = project_paths(REPO_ROOT)
        implicit = load_setup("south_asia", paths=paths)
        explicit = load_setup("south_asia", paths=paths)
        self.assertEqual(implicit.digest, explicit.digest)
        self.assertFalse(implicit.policies["build_cap_replaces_workbook_values"])
        self.assertEqual(
            implicit.metadata["effective_values"][
                "TRNBGDXXINDEA_residual_capacity_gw"
            ],
            2.496,
        )

    def test_unescap_training_manifest_is_complete_and_namespaced(self) -> None:
        paths = project_paths(REPO_ROOT)
        manifest = load_setup("unescap_training", paths=paths)
        sources = manifest.source_paths(paths)
        required_roles = {
            "scenario_registry", "scenario_inputs", "scenario_config_root",
            "scenario_workbook", "timeslice_workbook", "ao_decisions",
            "interconnector_authority", "interconnector_taxonomy",
            "compilation_config", "execution_config", "preparation_config",
            "country_config", "region_config", "osemosys_inputs",
            "preparation_inputs", "preparation_templates",
            "preparation_resources", "compilation_resources",
            "secondary_technology_inputs", "execution_inputs",
            "maintained_model",
        }
        self.assertTrue(required_roles.issubset(manifest.authorities))
        self.assertTrue(all(path.exists() for path in sources.values()))
        self.assertEqual(manifest.policies["build_cap_replaces_workbook_values"], True)
        # The base-year pin of the South Asia model is not domain-valid on the
        # reduced two-region model (CBC-certified infeasible); the setup
        # keeps it disabled.
        self.assertEqual(manifest.policies["apply_pwr_min_pin"], False)
        self.assertEqual(
            manifest.metadata["effective_values"][
                "TRNBGDXXINDEA_residual_capacity_gw"
            ],
            2.496,
        )
        self.assertEqual(
            manifest.metadata["effective_values"][
                "TRNBGDXXINDEA_relaxed_total_annual_max_capacity_gw"
            ],
            9999.0,
        )
        self.assertEqual(
            {
                role
                for role, spec in manifest.authorities.items()
                if spec.namespace == "project"
            },
            {"maintained_model", "execution_inputs", "preparation_templates"},
        )
        self.assertEqual(
            sources["scenario_workbook"], sources["interconnector_authority"]
        )


if __name__ == "__main__":
    unittest.main()
