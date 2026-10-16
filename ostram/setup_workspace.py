"""Atomic preparation and validation of mutable setup workspaces."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import shutil
from typing import Mapping

from ostram.paths import ProjectPaths
from ostram.setups import SetupError, SetupManifest
from ostram.validation.setup_domain import validate_seed_domain


STAMP_NAME = ".ostram-setup.json"
STAMP_SCHEMA = "ostram-prepared-setup-v1"


class SetupWorkspaceError(SetupError):
    """Raised for interrupted, foreign, or stale prepared workspaces."""


@dataclass(frozen=True)
class PreparedSetup:
    """One prepared setup: its workspace, resolved authority paths and stamp."""

    setup_id: str
    workspace: Path
    authorities: Mapping[str, Path]
    stamp: Mapping[str, object]


def setups_root(paths: ProjectPaths) -> Path:
    """Return the folder of prepared setups, ``<workspace>/setups``."""

    return (paths.workspace / "setups").resolve()


def setup_workspace(paths: ProjectPaths, setup_id: str) -> Path:
    """Return the workspace of one setup, ``<workspace>/setups/<setup_id>``."""

    return (setups_root(paths) / setup_id).resolve()


def _staging_path(paths: ProjectPaths, setup_id: str) -> Path:
    return (setups_root(paths) / f".{setup_id}.preparing").resolve()


def _backup_path(paths: ProjectPaths, setup_id: str) -> Path:
    return (setups_root(paths) / f".{setup_id}.reset-backup").resolve()


def _read_stamp(target: Path) -> dict:
    stamp_path = target / STAMP_NAME
    if not target.is_dir() or not stamp_path.is_file():
        raise SetupWorkspaceError(
            f"interrupted or foreign setup workspace (stamp missing): {target}"
        )
    try:
        stamp = json.loads(stamp_path.read_text(encoding="utf-8"))
    except Exception as error:
        raise SetupWorkspaceError(f"invalid setup stamp {stamp_path}: {error}") from error
    if stamp.get("schema") != STAMP_SCHEMA:
        raise SetupWorkspaceError(f"unsupported setup workspace stamp: {stamp_path}")
    return stamp


def _validate_stamp(target: Path, manifest: SetupManifest) -> dict:
    stamp = _read_stamp(target)
    if stamp.get("setup_id") != manifest.setup_id:
        raise SetupWorkspaceError(
            f"foreign setup workspace at {target}: stamp declares "
            f"{stamp.get('setup_id')!r}, expected {manifest.setup_id!r}"
        )
    return stamp


def _copy_authority(source: Path, destination: Path, kind: str) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if kind == "directory":
        shutil.copytree(source, destination)
    else:
        shutil.copy2(source, destination)


def _authority_digest(source: Path, kind: str) -> str:
    digest = hashlib.sha256()
    files = [source] if kind == "file" else sorted(
        (path for path in source.rglob("*") if path.is_file()),
        key=lambda path: path.relative_to(source).as_posix(),
    )
    for path in files:
        relative = path.name if kind == "file" else path.relative_to(source).as_posix()
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    return digest.hexdigest()


def _copy_digest(candidate: Path, kind: str) -> str | None:
    """Return the digest of a workspace copy, or None when it cannot be read."""
    try:
        return _authority_digest(candidate, kind)
    except OSError:
        return None


def _keep_or_move_workspace(target: Path, stamp: dict, manifest: SetupManifest) -> bool:
    """Handle a workspace whose stamp records another manifest digest.

    When the manifest's mutable authority roles equal the stamp's, the
    workspace is kept, the stamp takes the new digest and the result is True.
    Otherwise the workspace moves to a dated folder beside it and the result
    is False.
    """

    mutable_roles = {
        role for role, spec in manifest.authorities.items() if spec.mutable
    }
    recorded = stamp.get("mutable_authorities")
    if isinstance(recorded, dict) and set(recorded) == mutable_roles:
        stamp["manifest_sha256"] = manifest.digest
        (target / STAMP_NAME).write_text(
            json.dumps(stamp, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        print(f"[setup] Kept {target} (manifest changed, same mutable authority roles)")
        return True
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
    previous = target.with_name(f"{manifest.setup_id}_previous_{timestamp}")
    os.replace(target, previous)
    print(f"[setup] Mutable authority roles changed; previous workspace moved to {previous}")
    return False


def prepare_setup(
    manifest: SetupManifest,
    *,
    paths: ProjectPaths,
    reset: bool = False,
) -> PreparedSetup:
    """Validate, stage, stamp-last, and atomically publish one setup bundle."""

    sources = manifest.source_paths(paths, require_exists=True)
    validate_seed_domain(manifest, sources.get("osemosys_inputs"))
    mutable = {
        role: spec for role, spec in manifest.authorities.items() if spec.mutable
    }
    root = setups_root(paths)
    target = setup_workspace(paths, manifest.setup_id)
    staging = _staging_path(paths, manifest.setup_id)
    backup = _backup_path(paths, manifest.setup_id)

    # Existing debris is evidence of an interrupted mutation; never guess.
    if staging.exists() or backup.exists():
        raise SetupWorkspaceError(
            f"interrupted preparation state exists for {manifest.setup_id!r}: "
            f"{staging if staging.exists() else backup}"
        )
    if target.exists():
        if reset:
            # A reset is the supported way to republish after an intentional
            # manifest edit, but it must never adopt a foreign or unstamped
            # workspace.
            stamp = _read_stamp(target)
            if stamp.get("setup_id") != manifest.setup_id:
                raise SetupWorkspaceError(
                    f"foreign setup workspace at {target}: stamp declares "
                    f"{stamp.get('setup_id')!r}, expected {manifest.setup_id!r}"
                )
        else:
            stamp = _validate_stamp(target, manifest)
            if stamp.get("manifest_sha256") == manifest.digest:
                raise FileExistsError(
                    f"setup workspace already exists: {target}; pass --reset to replace it"
                )
            if _keep_or_move_workspace(target, stamp, manifest):
                return prepared_setup(manifest, paths=paths)

    root.mkdir(parents=True, exist_ok=True)
    staging.mkdir()
    relative_authorities: dict[str, str] = {}
    source_digests: dict[str, str] = {}
    copied_sources: dict[tuple[Path, str], Path] = {}
    for role, spec in mutable.items():
        source = sources[role]
        source_digest = _authority_digest(source, spec.kind)
        source_key = (source.resolve(), spec.kind)
        destination = copied_sources.get(source_key)
        if destination is None:
            destination = staging / "authorities" / role / source.name
            _copy_authority(source, destination, spec.kind)
            copied_sources[source_key] = destination
        if _authority_digest(destination, spec.kind) != source_digest:
            raise SetupWorkspaceError(
                f"authority {role!r} changed while it was being prepared"
            )
        relative_authorities[role] = destination.relative_to(staging).as_posix()
        source_digests[role] = source_digest

    stamp = {
        "schema": STAMP_SCHEMA,
        "setup_id": manifest.setup_id,
        "manifest": str(manifest.path),
        "manifest_sha256": manifest.digest,
        "prepared_at_utc": datetime.now(timezone.utc).isoformat(),
        "mutable_authorities": relative_authorities,
        "mutable_source_sha256": source_digests,
    }
    # This is intentionally the final write inside staging.
    (staging / STAMP_NAME).write_text(
        json.dumps(stamp, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )

    if target.exists():
        os.replace(target, backup)
        try:
            os.replace(staging, target)
        except BaseException:
            os.replace(backup, target)
            raise
        shutil.rmtree(backup)
    else:
        os.replace(staging, target)

    return prepared_setup(manifest, paths=paths)


def prepared_setup(
    manifest: SetupManifest,
    *,
    paths: ProjectPaths,
) -> PreparedSetup:
    """Resolve one atomic runtime bundle without per-authority fallback."""

    mutable_roles = {
        role for role, spec in manifest.authorities.items() if spec.mutable
    }
    # Immutable setups preserve lazy historical command validation while the
    # complete role map still prevents fallback. Mutable setups must verify
    # their seed plus every read-only companion before activation.
    sources = manifest.source_paths(paths, require_exists=bool(mutable_roles))
    target = setup_workspace(paths, manifest.setup_id)
    if not mutable_roles:
        return PreparedSetup(
            manifest.setup_id,
            target,
            sources,
            {
                "schema": STAMP_SCHEMA,
                "setup_id": manifest.setup_id,
                "manifest_sha256": manifest.digest,
                "mutable_authorities": {},
                "mutable_source_sha256": {},
            },
        )

    stamp = _validate_stamp(target, manifest)
    if stamp.get("manifest_sha256") != manifest.digest and not _keep_or_move_workspace(
        target, stamp, manifest
    ):
        return prepare_setup(manifest, paths=paths)
    mapping = stamp.get("mutable_authorities")
    source_digests = stamp.get("mutable_source_sha256")
    if not isinstance(mapping, dict) or set(mapping) != mutable_roles:
        raise SetupWorkspaceError(
            f"prepared authority roles do not match manifest for {manifest.setup_id!r}"
        )
    if not isinstance(source_digests, dict) or set(source_digests) != mutable_roles:
        raise SetupWorkspaceError(
            f"prepared source digests do not match manifest for {manifest.setup_id!r}"
        )
    resolved = dict(sources)
    stamp_dirty = False
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
    for role, relative in mapping.items():
        if not isinstance(relative, str):
            raise SetupWorkspaceError(f"invalid prepared path for authority {role!r}")
        candidate = (target / relative).resolve()
        try:
            candidate.relative_to(target)
        except ValueError as error:
            raise SetupWorkspaceError(
                f"prepared authority {role!r} escapes setup workspace"
            ) from error
        spec = manifest.authorities[role]
        current_digest = _authority_digest(sources[role], spec.kind)
        if current_digest != source_digests[role]:
            # The source changed since preparation. The workspace copy is
            # replaced; a copy edited since preparation, or one whose digest
            # cannot be read, moves to authorities_previous first. The stamp
            # records the source digest, which equals the copy's digest at
            # preparation and after each refresh.
            copy_digest = _copy_digest(candidate, spec.kind) if candidate.exists() else None
            if copy_digest != current_digest:
                if copy_digest == source_digests[role]:
                    if spec.kind == "directory":
                        shutil.rmtree(candidate)
                    else:
                        candidate.unlink()
                elif candidate.exists():
                    moved_to = target / "authorities_previous" / timestamp / relative
                    moved_to.parent.mkdir(parents=True, exist_ok=True)
                    os.replace(candidate, moved_to)
                    print(f"[setup] Moved the edited copy of {role!r} to {moved_to}")
                _copy_authority(sources[role], candidate, spec.kind)
            source_digests[role] = current_digest
            stamp_dirty = True
            print(
                f"[setup] Refreshed mutable authority {role!r} "
                f"(source changed since preparation)"
            )
        exists = candidate.is_file() if spec.kind == "file" else candidate.is_dir()
        if not exists:
            raise SetupWorkspaceError(
                f"prepared authority {role!r} is missing: {candidate}"
            )
        resolved[role] = candidate

    if stamp_dirty:
        stamp["mutable_source_sha256"] = source_digests
        stamp["refreshed_at_utc"] = datetime.now(timezone.utc).isoformat()
        (target / STAMP_NAME).write_text(
            json.dumps(stamp, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )

    return PreparedSetup(manifest.setup_id, target, resolved, stamp)
