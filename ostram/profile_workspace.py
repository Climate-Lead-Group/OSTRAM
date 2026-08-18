"""Atomic preparation and validation of mutable profile workspaces."""

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
from ostram.profiles import ProfileError, ProfileManifest
from ostram.validation.profile import validate_seed_domain


STAMP_NAME = ".ostram-profile.json"
STAMP_SCHEMA = "ostram-prepared-profile-v1"


class ProfileWorkspaceError(ProfileError):
    """Raised for interrupted, foreign, or stale prepared workspaces."""


@dataclass(frozen=True)
class PreparedProfile:
    profile_id: str
    workspace: Path
    authorities: Mapping[str, Path]
    stamp: Mapping[str, object]


def profiles_root(paths: ProjectPaths) -> Path:
    return (paths.workspace / "profiles").resolve()


def profile_workspace(paths: ProjectPaths, profile_id: str) -> Path:
    return (profiles_root(paths) / profile_id).resolve()


def _staging_path(paths: ProjectPaths, profile_id: str) -> Path:
    return (profiles_root(paths) / f".{profile_id}.preparing").resolve()


def _backup_path(paths: ProjectPaths, profile_id: str) -> Path:
    return (profiles_root(paths) / f".{profile_id}.reset-backup").resolve()


def _read_stamp(target: Path) -> dict:
    stamp_path = target / STAMP_NAME
    if not target.is_dir() or not stamp_path.is_file():
        raise ProfileWorkspaceError(
            f"interrupted or foreign profile workspace (stamp missing): {target}"
        )
    try:
        stamp = json.loads(stamp_path.read_text(encoding="utf-8"))
    except Exception as error:
        raise ProfileWorkspaceError(f"invalid profile stamp {stamp_path}: {error}") from error
    if stamp.get("schema") != STAMP_SCHEMA:
        raise ProfileWorkspaceError(f"unsupported profile workspace stamp: {stamp_path}")
    return stamp


def _validate_stamp(target: Path, manifest: ProfileManifest) -> dict:
    stamp = _read_stamp(target)
    if stamp.get("profile_id") != manifest.profile_id:
        raise ProfileWorkspaceError(
            f"foreign profile workspace at {target}: stamp declares "
            f"{stamp.get('profile_id')!r}, expected {manifest.profile_id!r}"
        )
    if stamp.get("manifest_sha256") != manifest.digest:
        raise ProfileWorkspaceError(
            f"profile workspace at {target} was prepared from a different manifest"
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


def prepare_profile(
    manifest: ProfileManifest,
    *,
    paths: ProjectPaths,
    reset: bool = False,
) -> PreparedProfile:
    """Validate, stage, stamp-last, and atomically publish one profile bundle."""

    sources = manifest.source_paths(paths, require_exists=True)
    validate_seed_domain(manifest, sources.get("osemosys_inputs"))
    mutable = {
        role: spec for role, spec in manifest.authorities.items() if spec.mutable
    }
    root = profiles_root(paths)
    target = profile_workspace(paths, manifest.profile_id)
    staging = _staging_path(paths, manifest.profile_id)
    backup = _backup_path(paths, manifest.profile_id)

    # Existing debris is evidence of an interrupted mutation; never guess.
    if staging.exists() or backup.exists():
        raise ProfileWorkspaceError(
            f"interrupted preparation state exists for {manifest.profile_id!r}: "
            f"{staging if staging.exists() else backup}"
        )
    if target.exists():
        if reset:
            # A reset is the supported way to republish after an intentional
            # manifest edit, but it must never adopt a foreign or unstamped
            # workspace.
            stamp = _read_stamp(target)
            if stamp.get("profile_id") != manifest.profile_id:
                raise ProfileWorkspaceError(
                    f"foreign profile workspace at {target}: stamp declares "
                    f"{stamp.get('profile_id')!r}, expected {manifest.profile_id!r}"
                )
        else:
            _validate_stamp(target, manifest)
            raise FileExistsError(
                f"profile workspace already exists: {target}; pass --reset to replace it"
            )

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
            raise ProfileWorkspaceError(
                f"authority {role!r} changed while it was being prepared"
            )
        relative_authorities[role] = destination.relative_to(staging).as_posix()
        source_digests[role] = source_digest

    stamp = {
        "schema": STAMP_SCHEMA,
        "profile_id": manifest.profile_id,
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

    return prepared_profile(manifest, paths=paths)


def prepared_profile(
    manifest: ProfileManifest,
    *,
    paths: ProjectPaths,
) -> PreparedProfile:
    """Resolve one atomic runtime bundle without per-authority fallback."""

    mutable_roles = {
        role for role, spec in manifest.authorities.items() if spec.mutable
    }
    # Immutable profiles preserve lazy historical command validation while the
    # complete role map still prevents fallback. Mutable profiles must verify
    # their seed plus every read-only companion before activation.
    sources = manifest.source_paths(paths, require_exists=bool(mutable_roles))
    target = profile_workspace(paths, manifest.profile_id)
    if not mutable_roles:
        return PreparedProfile(
            manifest.profile_id,
            target,
            sources,
            {
                "schema": STAMP_SCHEMA,
                "profile_id": manifest.profile_id,
                "manifest_sha256": manifest.digest,
                "mutable_authorities": {},
                "mutable_source_sha256": {},
            },
        )

    stamp = _validate_stamp(target, manifest)
    mapping = stamp.get("mutable_authorities")
    source_digests = stamp.get("mutable_source_sha256")
    if not isinstance(mapping, dict) or set(mapping) != mutable_roles:
        raise ProfileWorkspaceError(
            f"prepared authority roles do not match manifest for {manifest.profile_id!r}"
        )
    if not isinstance(source_digests, dict) or set(source_digests) != mutable_roles:
        raise ProfileWorkspaceError(
            f"prepared source digests do not match manifest for {manifest.profile_id!r}"
        )
    resolved = dict(sources)
    stamp_dirty = False
    for role, relative in mapping.items():
        if not isinstance(relative, str):
            raise ProfileWorkspaceError(f"invalid prepared path for authority {role!r}")
        candidate = (target / relative).resolve()
        try:
            candidate.relative_to(target)
        except ValueError as error:
            raise ProfileWorkspaceError(
                f"prepared authority {role!r} escapes profile workspace"
            ) from error
        spec = manifest.authorities[role]
        current_digest = _authority_digest(sources[role], spec.kind)
        if current_digest != source_digests[role]:
            # Mutable source legitimately changed (e.g. student edited the
            # scenario workbook).  Refresh the workspace copy and update the
            # stamp instead of failing.
            if candidate.exists():
                if spec.kind == "directory":
                    shutil.rmtree(candidate)
                else:
                    candidate.unlink()
            _copy_authority(sources[role], candidate, spec.kind)
            source_digests[role] = current_digest
            stamp_dirty = True
            print(
                f"[profile] Refreshed mutable authority {role!r} "
                f"(source changed since preparation)"
            )
        exists = candidate.is_file() if spec.kind == "file" else candidate.is_dir()
        if not exists:
            raise ProfileWorkspaceError(
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

    return PreparedProfile(manifest.profile_id, target, resolved, stamp)
