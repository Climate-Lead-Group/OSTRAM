"""Fail-closed profile manifests and namespaced authority resolution.

Profiles are deliberately data-only.  Loading this module never imports a
pipeline route and never creates a workspace.  A manifest is resolved as one
bundle so a caller cannot accidentally combine authorities from two models.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
from typing import Any, Mapping

import yaml

from ostram.paths import ProjectPaths


PROFILE_ENV = "OSTRAM_PROFILE"
PROFILE_MANIFEST_ENV = "OSTRAM_PROFILE_MANIFEST"
PROFILE_WORKSPACE_ENV = "OSTRAM_PROFILE_WORKSPACE"
PROFILE_AUTHORITIES_ENV = "OSTRAM_PROFILE_AUTHORITIES"
PROFILE_POLICIES_ENV = "OSTRAM_PROFILE_POLICIES"
DEFAULT_PROFILE = "full"
PROFILE_SCHEMA = "ostram-profile-v1"
REGISTRY_SCHEMA = "ostram-profile-registry-v1"
_PROFILE_ID = re.compile(r"^[a-z0-9][a-z0-9_-]*$")
_ROLE = re.compile(r"^[a-z][a-z0-9_]*$")
_REFERENCE = re.compile(r"^([a-z]+):(.*)$", re.DOTALL)
SUPPORTED_NAMESPACES = frozenset({"profile", "project", "package"})


class ProfileError(RuntimeError):
    """Base error for invalid or unavailable profile state."""


class ProfileManifestError(ProfileError):
    """Raised when a profile registry or manifest violates its contract."""


class AuthorityResolutionError(ProfileError):
    """Raised when a namespaced authority reference is unsafe."""


class _UniqueKeyLoader(yaml.SafeLoader):
    pass


def _construct_unique_mapping(loader, node, deep=False):
    mapping: dict[Any, Any] = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        if key in mapping:
            raise ProfileManifestError(f"duplicate mapping key/authority role: {key!r}")
        mapping[key] = loader.construct_object(value_node, deep=deep)
    return mapping


_UniqueKeyLoader.add_constructor(
    yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
    _construct_unique_mapping,
)


def _load_mapping(path: Path) -> Mapping[str, Any]:
    try:
        raw = yaml.load(path.read_text(encoding="utf-8"), Loader=_UniqueKeyLoader)
    except ProfileError:
        raise
    except Exception as error:
        raise ProfileManifestError(f"could not parse {path}: {error}") from error
    if not isinstance(raw, Mapping):
        raise ProfileManifestError(f"{path} must contain a mapping")
    return raw


def _safe_relative(reference: str) -> PurePosixPath:
    if not reference or reference.strip() != reference:
        raise AuthorityResolutionError("authority path must be non-empty and unpadded")
    normalized = reference.replace("\\", "/")
    path = PurePosixPath(normalized)
    if path.is_absolute() or normalized.startswith("/"):
        raise AuthorityResolutionError(f"authority path must be relative: {reference!r}")
    if any(part in ("", ".", "..") for part in path.parts):
        raise AuthorityResolutionError(
            f"authority path contains traversal or empty components: {reference!r}"
        )
    if any(":" in part for part in path.parts):
        raise AuthorityResolutionError(f"authority path is not portable: {reference!r}")
    return path


def _contained(root: Path, relative: PurePosixPath) -> Path:
    root = root.resolve()
    candidate = root.joinpath(*relative.parts).resolve()
    try:
        candidate.relative_to(root)
    except ValueError as error:
        raise AuthorityResolutionError(
            f"authority escapes permitted root {root}: {relative}"
        ) from error
    return candidate


def resolve_authority_reference(
    reference: str,
    *,
    paths: ProjectPaths,
    profile_root: Path,
) -> Path:
    """Resolve one explicit ``scheme:path`` reference inside its allowed root."""

    if not isinstance(reference, str):
        raise AuthorityResolutionError("authority reference must be text")
    match = _REFERENCE.fullmatch(reference)
    if match is None:
        raise AuthorityResolutionError(
            f"authority reference requires an explicit namespace: {reference!r}"
        )
    scheme, value = match.groups()
    if scheme not in SUPPORTED_NAMESPACES:
        raise AuthorityResolutionError(f"unsupported authority namespace: {scheme!r}")
    relative = _safe_relative(value)
    roots = {
        "profile": Path(profile_root),
        "project": paths.project_root,
        "package": paths.package_root,
    }
    return _contained(roots[scheme], relative)


@dataclass(frozen=True)
class AuthoritySpec:
    role: str
    reference: str
    mutable: bool = False
    kind: str = "file"

    @property
    def namespace(self) -> str:
        match = _REFERENCE.fullmatch(self.reference)
        return match.group(1) if match else ""


@dataclass(frozen=True)
class ProfileManifest:
    profile_id: str
    path: Path
    authorities: Mapping[str, AuthoritySpec]
    metadata: Mapping[str, Any] = field(default_factory=dict)
    policies: Mapping[str, Any] = field(default_factory=dict)

    @property
    def root(self) -> Path:
        return self.path.parent

    @property
    def digest(self) -> str:
        return hashlib.sha256(self.path.read_bytes()).hexdigest()

    def source_paths(
        self,
        paths: ProjectPaths,
        *,
        require_exists: bool = True,
    ) -> dict[str, Path]:
        """Resolve and validate the complete source authority bundle."""

        resolved: dict[str, Path] = {}
        errors: list[str] = []
        for role, spec in self.authorities.items():
            try:
                target = resolve_authority_reference(
                    spec.reference,
                    paths=paths,
                    profile_root=self.root,
                )
                if spec.mutable and spec.namespace != "profile":
                    raise ProfileManifestError(
                        f"mutable authority {role!r} must use profile:, "
                        f"not {spec.namespace or 'an implicit namespace'}"
                    )
                if require_exists:
                    exists = target.is_file() if spec.kind == "file" else target.is_dir()
                    if not exists:
                        raise FileNotFoundError(
                            f"expected {spec.kind} does not exist: {target}"
                        )
                resolved[role] = target
            except Exception as error:
                errors.append(f"{role}: {error}")
        if errors:
            raise ProfileManifestError(
                f"profile {self.profile_id!r} authority bundle is invalid; "
                + "; ".join(errors)
            )
        return resolved

    def require_roles(self, *roles: str) -> None:
        missing = [role for role in roles if role not in self.authorities]
        if missing:
            raise ProfileManifestError(
                f"profile {self.profile_id!r} is missing authorities: {missing}"
            )


def _parse_authority(role: str, value: object) -> AuthoritySpec:
    if not _ROLE.fullmatch(role):
        raise ProfileManifestError(f"invalid authority role: {role!r}")
    if isinstance(value, str):
        return AuthoritySpec(role=role, reference=value)
    if not isinstance(value, Mapping):
        raise ProfileManifestError(f"authority {role!r} must be text or a mapping")
    unknown = set(value) - {"path", "mutable", "kind"}
    if unknown:
        raise ProfileManifestError(
            f"authority {role!r} has unsupported keys: {sorted(unknown)}"
        )
    reference = value.get("path")
    if not isinstance(reference, str):
        raise ProfileManifestError(f"authority {role!r} requires a text path")
    mutable = value.get("mutable", False)
    if not isinstance(mutable, bool):
        raise ProfileManifestError(f"authority {role!r} mutable must be boolean")
    kind = value.get("kind", "file")
    if kind not in ("file", "directory"):
        raise ProfileManifestError(
            f"authority {role!r} kind must be 'file' or 'directory'"
        )
    return AuthoritySpec(role=role, reference=reference, mutable=mutable, kind=kind)


def load_manifest(
    path: Path | str,
    *,
    expected_profile: str | None = None,
) -> ProfileManifest:
    manifest_path = Path(path).resolve()
    if not manifest_path.is_file():
        raise FileNotFoundError(f"profile manifest not found: {manifest_path}")
    raw = _load_mapping(manifest_path)
    if raw.get("schema") != PROFILE_SCHEMA:
        raise ProfileManifestError(
            f"unsupported profile schema in {manifest_path}: {raw.get('schema')!r}"
        )
    profile_id = raw.get("id")
    if not isinstance(profile_id, str) or not _PROFILE_ID.fullmatch(profile_id):
        raise ProfileManifestError(f"invalid profile id: {profile_id!r}")
    if expected_profile is not None and profile_id != expected_profile:
        raise ProfileManifestError(
            f"registry requested profile {expected_profile!r}, but manifest declares "
            f"{profile_id!r}"
        )
    authorities_raw = raw.get("authorities")
    if not isinstance(authorities_raw, Mapping) or not authorities_raw:
        raise ProfileManifestError("profile must declare a non-empty authorities mapping")
    authorities = {
        str(role): _parse_authority(str(role), value)
        for role, value in authorities_raw.items()
    }
    metadata = raw.get("metadata", {})
    policies = raw.get("policies", {})
    if not isinstance(metadata, Mapping) or not isinstance(policies, Mapping):
        raise ProfileManifestError("profile metadata and policies must be mappings")
    return ProfileManifest(
        profile_id=profile_id,
        path=manifest_path,
        authorities=authorities,
        metadata=dict(metadata),
        policies=dict(policies),
    )


def default_registry_path(paths: ProjectPaths) -> Path:
    project_registry = (paths.config_root / "profiles" / "registry.json").resolve()
    if project_registry.is_file():
        return project_registry
    # Compatibility project bundles created before profiles do not contain a
    # registry.  Use the installed registry as one indivisible contract while
    # still resolving its project: authorities against the selected bundle.
    return (Path(__file__).resolve().parent.parent / "config" / "profiles" / "registry.json").resolve()


def load_profile(
    profile_id: str,
    *,
    paths: ProjectPaths,
    registry_path: Path | str | None = None,
) -> ProfileManifest:
    if not _PROFILE_ID.fullmatch(profile_id):
        raise ProfileManifestError(f"invalid profile id: {profile_id!r}")
    registry = (
        default_registry_path(paths)
        if registry_path is None
        else Path(registry_path).resolve()
    )
    if not registry.is_file():
        raise FileNotFoundError(f"profile registry not found: {registry}")
    raw = _load_mapping(registry)
    if raw.get("schema") != REGISTRY_SCHEMA:
        raise ProfileManifestError(
            f"unsupported profile registry schema: {raw.get('schema')!r}"
        )
    profiles = raw.get("profiles")
    if not isinstance(profiles, Mapping):
        raise ProfileManifestError("profile registry requires a profiles mapping")
    if profile_id not in profiles:
        raise ProfileManifestError(
            f"unknown profile {profile_id!r}; available: {sorted(profiles)}"
        )
    entry = profiles[profile_id]
    reference = entry.get("manifest") if isinstance(entry, Mapping) else entry
    if not isinstance(reference, str):
        raise ProfileManifestError(
            f"profile registry entry {profile_id!r} requires a manifest reference"
        )
    # A registry is not itself a profile root, so profile: would be ambiguous.
    if reference.startswith("profile:"):
        raise ProfileManifestError("registry manifest references cannot use profile:")
    project_registry = (paths.config_root / "profiles" / "registry.json").resolve()
    if registry != project_registry and profile_id == DEFAULT_PROFILE:
        # The compatibility registry and full manifest travel together; the
        # manifest's own project: authorities still bind to ``paths``.
        manifest_path = (registry.parent / "full.yaml").resolve()
    else:
        manifest_path = resolve_authority_reference(
            reference,
            paths=paths,
            profile_root=registry.parent,
        )
    return load_manifest(manifest_path, expected_profile=profile_id)


def encoded_environment(
    manifest: ProfileManifest,
    *,
    authorities: Mapping[str, Path],
    profile_workspace: Path,
) -> dict[str, str]:
    """Return the complete child-process profile environment."""

    return {
        PROFILE_ENV: manifest.profile_id,
        PROFILE_MANIFEST_ENV: str(manifest.path),
        PROFILE_WORKSPACE_ENV: str(Path(profile_workspace).resolve()),
        PROFILE_AUTHORITIES_ENV: json.dumps(
            {role: str(path.resolve()) for role, path in authorities.items()},
            sort_keys=True,
        ),
        PROFILE_POLICIES_ENV: json.dumps(dict(manifest.policies), sort_keys=True),
    }


def active_profile_id(environ: Mapping[str, str] | None = None) -> str:
    environment = os.environ if environ is None else environ
    return environment.get(PROFILE_ENV, DEFAULT_PROFILE)


def active_policies(environ: Mapping[str, str] | None = None) -> dict[str, Any]:
    environment = os.environ if environ is None else environ
    raw = environment.get(PROFILE_POLICIES_ENV)
    if not raw:
        return {}
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as error:
        raise ProfileError(f"invalid {PROFILE_POLICIES_ENV}: {error}") from error
    if not isinstance(value, dict):
        raise ProfileError(f"{PROFILE_POLICIES_ENV} must encode an object")
    return value


def profile_policy(
    name: str,
    default: Any = None,
    *,
    environ: Mapping[str, str] | None = None,
) -> Any:
    return active_policies(environ).get(name, default)
