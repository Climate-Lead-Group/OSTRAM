from __future__ import annotations

import json
import os
from pathlib import Path
import tempfile
import types
import unittest
from unittest import mock

import ostram.__main__ as cli
from ostram.paths import ProjectPaths
from ostram.profile_workspace import (
    STAMP_NAME,
    ProfileWorkspaceError,
    prepare_profile,
    prepared_profile,
    profile_workspace,
)
from ostram.profiles import (
    AuthorityResolutionError,
    ProfileManifestError,
    load_manifest,
    load_profile,
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
            profile_root = root / "profiles" / "tiny"
            profile_root.mkdir(parents=True)
            self.assertEqual(
                resolve_authority_reference(
                    "project:inputs/data.csv", paths=paths, profile_root=profile_root
                ),
                (root / "inputs" / "data.csv").resolve(),
            )
            self.assertEqual(
                resolve_authority_reference(
                    "profile:seed.xlsx", paths=paths, profile_root=profile_root
                ),
                (profile_root / "seed.xlsx").resolve(),
            )
            self.assertEqual(
                resolve_authority_reference(
                    "package:resources/table.csv", paths=paths, profile_root=profile_root
                ),
                (root / "ostram" / "resources" / "table.csv").resolve(),
            )
            for reference in (
                "inputs/data.csv", "http:data.csv", "project:",
                "project:../secret", "profile:/absolute", "package:C:\\secret",
            ):
                with self.subTest(reference=reference), self.assertRaises(
                    AuthorityResolutionError
                ):
                    resolve_authority_reference(
                        reference, paths=paths, profile_root=profile_root
                    )

    def test_duplicate_roles_and_invalid_mutability_fail(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = root / "profile.yaml"
            manifest.write_text(
                "schema: ostram-profile-v1\nid: tiny\nauthorities:\n"
                "  seed: profile:a.csv\n  seed: profile:b.csv\n",
                encoding="utf-8",
            )
            with self.assertRaisesRegex(ProfileManifestError, "duplicate"):
                load_manifest(manifest)

            manifest.write_text(
                "schema: ostram-profile-v1\nid: tiny\nauthorities:\n"
                "  seed: {path: project:a.csv, mutable: true}\n",
                encoding="utf-8",
            )
            parsed = load_manifest(manifest)
            write_project(root)
            (root / "a.csv").write_text("x\n", encoding="utf-8")
            with self.assertRaisesRegex(ProfileManifestError, "must use profile"):
                parsed.source_paths(project_paths(root))


class AtomicWorkspaceTests(unittest.TestCase):
    def _fixture(self, root: Path):
        write_project(root)
        profile_root = root / "fixtures" / "tiny"
        profile_root.mkdir(parents=True)
        (profile_root / "seed.txt").write_text("version one", encoding="utf-8")
        (root / "inputs" / "readonly.txt").write_text("readonly", encoding="utf-8")
        manifest_path = profile_root / "profile.yaml"
        manifest_path.write_text(
            "schema: ostram-profile-v1\nid: tiny\nauthorities:\n"
            "  seed: {path: profile:seed.txt, mutable: true}\n"
            "  seed_alias: {path: profile:seed.txt, mutable: true}\n"
            "  readonly: project:inputs/readonly.txt\n",
            encoding="utf-8",
        )
        return project_paths(root), load_manifest(manifest_path)

    def test_prepare_reset_stamp_last_and_readonly_authority(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            result = prepare_profile(manifest, paths=paths)
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
                prepare_profile(manifest, paths=paths)

            manifest.path.write_text(
                manifest.path.read_text(encoding="utf-8")
                + "metadata:\n  revision: two\n",
                encoding="utf-8",
            )
            changed_manifest = load_manifest(manifest.path)
            with self.assertRaisesRegex(ProfileWorkspaceError, "different manifest"):
                prepare_profile(changed_manifest, paths=paths)
            result = prepare_profile(changed_manifest, paths=paths, reset=True)
            self.assertEqual(result.stamp["manifest_sha256"], changed_manifest.digest)
            manifest = changed_manifest

            (manifest.root / "seed.txt").write_text("version two", encoding="utf-8")
            # Mutable sources are now auto-refreshed instead of raising.
            refreshed = prepared_profile(manifest, paths=paths)
            self.assertEqual(refreshed.authorities["seed"].read_text(), "version two")

    def test_interrupted_foreign_and_missing_prepared_files_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            paths, manifest = self._fixture(Path(temp))
            interrupted = paths.workspace / "profiles" / ".tiny.preparing"
            interrupted.mkdir(parents=True)
            with self.assertRaisesRegex(ProfileWorkspaceError, "interrupted"):
                prepare_profile(manifest, paths=paths)
            interrupted.rmdir()

            result = prepare_profile(manifest, paths=paths)
            result.authorities["seed"].unlink()
            with self.assertRaisesRegex(ProfileWorkspaceError, "missing"):
                prepared_profile(manifest, paths=paths)
            # The source still exists: absence in the bundle may not fall back.
            self.assertTrue((manifest.root / "seed.txt").is_file())

            stamp_path = profile_workspace(paths, "tiny") / STAMP_NAME
            stamp = json.loads(stamp_path.read_text(encoding="utf-8"))
            stamp["profile_id"] = "foreign"
            stamp_path.write_text(json.dumps(stamp), encoding="utf-8")
            with self.assertRaisesRegex(ProfileWorkspaceError, "foreign"):
                prepared_profile(manifest, paths=paths)


class ProfileCliTests(unittest.TestCase):
    def test_profile_is_selected_before_route_import_and_not_forwarded(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            write_project(root)
            (root / "inputs" / "authority.txt").write_text("ok", encoding="utf-8")
            profiles = root / "config" / "profiles"
            profiles.mkdir(parents=True)
            (profiles / "registry.json").write_text(
                json.dumps({
                    "schema": "ostram-profile-registry-v1",
                    "profiles": {"tiny": "project:config/profiles/tiny.yaml"},
                }), encoding="utf-8"
            )
            (profiles / "tiny.yaml").write_text(
                "schema: ostram-profile-v1\nid: tiny\nauthorities:\n"
                "  authority: project:inputs/authority.txt\n",
                encoding="utf-8",
            )
            events = []

            def load_route(route):
                events.append(("import", os.environ.get("OSTRAM_PROFILE")))
                return types.SimpleNamespace(
                    main=lambda: events.append(("argv", tuple(__import__("sys").argv)))
                )

            with mock.patch.object(cli, "_load_route_module", side_effect=load_route):
                result = cli.main([
                    "--project-root", str(root), "--profile", "tiny",
                    "run", "--sentinel",
                ])
            self.assertEqual(result, 0)
            self.assertEqual(events[0], ("import", "tiny"))
            self.assertEqual(events[1][1], ("python -m ostram run", "--sentinel"))

    def test_profile_option_after_command_is_left_for_historical_parser(self) -> None:
        import sys

        seen = []
        fake = types.SimpleNamespace(
            main=mock.Mock(side_effect=lambda: seen.append(tuple(sys.argv)))
        )
        with mock.patch.object(cli, "_load_route_module", return_value=fake):
            cli.main(["run", "--profile", "full"])
        self.assertEqual(
            seen, [("python -m ostram run", "--profile", "full")]
        )

    def test_unqualified_and_explicit_full_have_same_authority_bundle(self) -> None:
        paths = project_paths(REPO_ROOT)
        implicit = load_profile("full", paths=paths)
        explicit = load_profile("full", paths=paths)
        self.assertEqual(implicit.digest, explicit.digest)
        self.assertFalse(implicit.policies["lid_rule_new_semantics"])
        self.assertEqual(
            implicit.metadata["effective_values"][
                "TRNBGDXXINDEA_residual_capacity_gw"
            ],
            2.496,
        )

    def test_unescap_manifest_is_complete_and_namespaced(self) -> None:
        paths = project_paths(REPO_ROOT)
        manifest = load_profile("unescap", paths=paths)
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
        self.assertEqual(manifest.policies["lid_rule_new_semantics"], True)
        # The full-model WS-4 base-year pin is not domain-valid on the
        # reduced two-region model (CBC-certified infeasible); the profile
        # must keep it disabled and must not re-declare the rules digest.
        self.assertEqual(manifest.policies["apply_pwr_min_pin"], False)
        self.assertNotIn("pwr_min_pin_rules_sha256", manifest.policies)
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
