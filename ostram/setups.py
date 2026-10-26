"""Fail-closed setup manifests and namespaced authority resolution.

Setups are deliberately data-only.  Loading this module never imports a
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

from ostram import yaml_io
from ostram.paths import ProjectPaths


SETUP_ENV = "OSTRAM_SETUP"
SETUP_MANIFEST_ENV = "OSTRAM_SETUP_MANIFEST"
SETUP_WORKSPACE_ENV = "OSTRAM_SETUP_WORKSPACE"
SETUP_AUTHORITIES_ENV = "OSTRAM_SETUP_AUTHORITIES"
SETUP_POLICIES_ENV = "OSTRAM_SETUP_POLICIES"
DEFAULT_SETUP = "south_asia"
SETUP_SCHEMA = "ostram-setup-v1"
REGISTRY_SCHEMA = "ostram-setup-registry-v1"
_SETUP_ID = re.compile(r"^[a-z0-9][a-z0-9_-]*$")
_ROLE = re.compile(r"^[a-z][a-z0-9_]*$")
_REFERENCE = re.compile(r"^([a-z]+):(.*)$", re.DOTALL)
SUPPORTED_NAMESPACES = frozenset({"setup", "project", "package"})


class SetupError(RuntimeError):
    """Base error for invalid or unavailable setup state."""


class SetupManifestError(SetupError):
    """Raised when a setup registry or manifest violates its contract."""


class AuthorityResolutionError(SetupError):
    """Raised when a namespaced authority reference is unsafe."""


def _load_mapping(path: Path) -> Mapping[str, Any]:
    try:
        raw = yaml_io.load_file(path)
    except yaml_io.DuplicateKeyError as error:
        raise SetupManifestError(f"duplicate mapping key/authority role: {error}") from error
    except Exception as error:
        raise SetupManifestError(f"could not parse {path}: {error}") from error
    if not isinstance(raw, Mapping):
        raise SetupManifestError(f"{path} must contain a mapping")
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
    setup_root: Path,
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
        "setup": Path(setup_root),
        "project": paths.project_root,
        "package": paths.package_root,
    }
    return _contained(roots[scheme], relative)


@dataclass(frozen=True)
class AuthoritySpec:
    """One authority of a setup manifest.

    ``reference`` is a ``scheme:path`` reference, ``mutable`` marks an authority
    that is copied into the setup workspace, and ``kind`` is ``file`` or
    ``directory``.
    """

    role: str
    reference: str
    mutable: bool = False
    kind: str = "file"

    @property
    def namespace(self) -> str:
        """Return the scheme of the reference, or an empty string when it has none."""

        match = _REFERENCE.fullmatch(self.reference)
        return match.group(1) if match else ""


@dataclass(frozen=True)
class SetupManifest:
    """One loaded setup manifest with its authorities, metadata and policies."""

    setup_id: str
    path: Path
    authorities: Mapping[str, AuthoritySpec]
    metadata: Mapping[str, Any] = field(default_factory=dict)
    policies: Mapping[str, Any] = field(default_factory=dict)

    @property
    def root(self) -> Path:
        """Return the folder that holds the manifest file."""

        return self.path.parent

    @property
    def digest(self) -> str:
        """Return the SHA-256 hex digest of the manifest file bytes."""

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
                    setup_root=self.root,
                )
                if spec.mutable and spec.namespace != "setup":
                    raise SetupManifestError(
                        f"mutable authority {role!r} must use setup:, "
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
            raise SetupManifestError(
                f"setup {self.setup_id!r} authority bundle is invalid; "
                + "; ".join(errors)
            )
        return resolved

    def require_roles(self, *roles: str) -> None:
        """Check that the manifest declares an authority for every role in ``roles``.

        Raises:
            SetupManifestError: If any role has no authority.
        """

        missing = [role for role in roles if role not in self.authorities]
        if missing:
            raise SetupManifestError(
                f"setup {self.setup_id!r} is missing authorities: {missing}"
            )


def _parse_authority(role: str, value: object) -> AuthoritySpec:
    if not _ROLE.fullmatch(role):
        raise SetupManifestError(f"invalid authority role: {role!r}")
    if isinstance(value, str):
        return AuthoritySpec(role=role, reference=value)
    if not isinstance(value, Mapping):
        raise SetupManifestError(f"authority {role!r} must be text or a mapping")
    unknown = set(value) - {"path", "mutable", "kind"}
    if unknown:
        raise SetupManifestError(
            f"authority {role!r} has unsupported keys: {sorted(unknown)}"
        )
    reference = value.get("path")
    if not isinstance(reference, str):
        raise SetupManifestError(f"authority {role!r} requires a text path")
    mutable = value.get("mutable", False)
    if not isinstance(mutable, bool):
        raise SetupManifestError(f"authority {role!r} mutable must be boolean")
    kind = value.get("kind", "file")
    if kind not in ("file", "directory"):
        raise SetupManifestError(
            f"authority {role!r} kind must be 'file' or 'directory'"
        )
    return AuthoritySpec(role=role, reference=reference, mutable=mutable, kind=kind)


def load_manifest(
    path: Path | str,
    *,
    expected_setup: str | None = None,
) -> SetupManifest:
    """Load and validate one setup manifest file.

    Raises:
        FileNotFoundError: If ``path`` is not a file.
        SetupManifestError: If the schema, id, authorities, metadata or
            policies are invalid, or if the id differs from ``expected_setup``.
    """

    manifest_path = Path(path).resolve()
    if not manifest_path.is_file():
        raise FileNotFoundError(f"setup manifest not found: {manifest_path}")
    raw = _load_mapping(manifest_path)
    if raw.get("schema") != SETUP_SCHEMA:
        raise SetupManifestError(
            f"unsupported setup schema in {manifest_path}: {raw.get('schema')!r}"
        )
    setup_id = raw.get("id")
    if not isinstance(setup_id, str) or not _SETUP_ID.fullmatch(setup_id):
        raise SetupManifestError(f"invalid setup id: {setup_id!r}")
    if expected_setup is not None and setup_id != expected_setup:
        raise SetupManifestError(
            f"registry requested setup {expected_setup!r}, but manifest declares "
            f"{setup_id!r}"
        )
    authorities_raw = raw.get("authorities")
    if not isinstance(authorities_raw, Mapping) or not authorities_raw:
        raise SetupManifestError("setup must declare a non-empty authorities mapping")
    authorities = {
        str(role): _parse_authority(str(role), value)
        for role, value in authorities_raw.items()
    }
    metadata = raw.get("metadata", {})
    policies = raw.get("policies", {})
    if not isinstance(metadata, Mapping) or not isinstance(policies, Mapping):
        raise SetupManifestError("setup metadata and policies must be mappings")
    if "lid_rule_new_semantics" in policies:
        raise SetupManifestError(
            f"{manifest_path}: the policy lid_rule_new_semantics is called "
            "build_cap_replaces_workbook_values; rename it"
        )
    return SetupManifest(
        setup_id=setup_id,
        path=manifest_path,
        authorities=authorities,
        metadata=dict(metadata),
        policies=dict(policies),
    )


def default_registry_path(paths: ProjectPaths) -> Path:
    """Return the path of the setup registry.

    This is the project's ``config/setups/registry.json`` when it exists,
    else the ``config/setups/registry.json`` beside the installed package.
    """

    project_registry = (paths.config_root / "setups" / "registry.json").resolve()
    if project_registry.is_file():
        return project_registry
    # Compatibility project bundles created before setups do not contain a
    # registry.  Use the installed registry as one indivisible contract while
    # still resolving its project: authorities against the selected bundle.
    return (Path(__file__).resolve().parent.parent / "config" / "setups" / "registry.json").resolve()


def load_setup(
    setup_id: str,
    *,
    paths: ProjectPaths,
    registry_path: Path | str | None = None,
) -> SetupManifest:
    """Load the manifest of ``setup_id`` through the setup registry.

    The registry is ``registry_path``, or :func:`default_registry_path` when it
    is ``None``. When the installed registry is used for the ``full`` setup,
    the manifest is the ``south_asia.yaml`` beside that registry.

    Raises:
        FileNotFoundError: If the registry or the manifest file is missing.
        SetupManifestError: If the id, the registry or the manifest is invalid,
            or if the registry lists no such setup.
        AuthorityResolutionError: If the manifest reference is unsafe.
    """

    if not _SETUP_ID.fullmatch(setup_id):
        raise SetupManifestError(f"invalid setup id: {setup_id!r}")
    registry = (
        default_registry_path(paths)
        if registry_path is None
        else Path(registry_path).resolve()
    )
    if not registry.is_file():
        raise FileNotFoundError(f"setup registry not found: {registry}")
    raw = _load_mapping(registry)
    if raw.get("schema") != REGISTRY_SCHEMA:
        raise SetupManifestError(
            f"unsupported setup registry schema: {raw.get('schema')!r}"
        )
    setups = raw.get("setups")
    if not isinstance(setups, Mapping):
        raise SetupManifestError("setup registry requires a setups mapping")
    if setup_id not in setups:
        raise SetupManifestError(
            f"unknown setup {setup_id!r}; available: {sorted(setups)}"
        )
    entry = setups[setup_id]
    reference = entry.get("manifest") if isinstance(entry, Mapping) else entry
    if not isinstance(reference, str):
        raise SetupManifestError(
            f"setup registry entry {setup_id!r} requires a manifest reference"
        )
    # A registry is not itself a setup root, so setup: would be ambiguous.
    if reference.startswith("setup:"):
        raise SetupManifestError("registry manifest references cannot use setup:")
    project_registry = (paths.config_root / "setups" / "registry.json").resolve()
    if registry != project_registry and setup_id == DEFAULT_SETUP:
        # The compatibility registry and default manifest travel together; the
        # manifest's own project: authorities still bind to ``paths``.
        manifest_path = (registry.parent / PurePosixPath(reference).name).resolve()
    else:
        manifest_path = resolve_authority_reference(
            reference,
            paths=paths,
            setup_root=registry.parent,
        )
    return load_manifest(manifest_path, expected_setup=setup_id)


def encoded_environment(
    manifest: SetupManifest,
    *,
    authorities: Mapping[str, Path],
    setup_workspace: Path,
) -> dict[str, str]:
    """Return the complete child-process setup environment."""

    return {
        SETUP_ENV: manifest.setup_id,
        SETUP_MANIFEST_ENV: str(manifest.path),
        SETUP_WORKSPACE_ENV: str(Path(setup_workspace).resolve()),
        SETUP_AUTHORITIES_ENV: json.dumps(
            {role: str(path.resolve()) for role, path in authorities.items()},
            sort_keys=True,
        ),
        SETUP_POLICIES_ENV: json.dumps(dict(manifest.policies), sort_keys=True),
    }


def active_setup_id(environ: Mapping[str, str] | None = None) -> str:
    """Return the setup id in ``OSTRAM_SETUP``, or ``south_asia`` when it is unset."""

    environment = os.environ if environ is None else environ
    return environment.get(SETUP_ENV, DEFAULT_SETUP)


def active_policies(environ: Mapping[str, str] | None = None) -> dict[str, Any]:
    """Return the setup policies decoded from ``OSTRAM_SETUP_POLICIES``.

    An unset or empty value gives an empty dict.

    Raises:
        SetupError: If the value is not a JSON object.
    """

    environment = os.environ if environ is None else environ
    raw = environment.get(SETUP_POLICIES_ENV)
    if not raw:
        return {}
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as error:
        raise SetupError(f"invalid {SETUP_POLICIES_ENV}: {error}") from error
    if not isinstance(value, dict):
        raise SetupError(f"{SETUP_POLICIES_ENV} must encode an object")
    return value


def setup_policy(
    name: str,
    default: Any = None,
    *,
    environ: Mapping[str, str] | None = None,
) -> Any:
    """Return the active setup policy ``name``, or ``default`` when it is absent."""

    return active_policies(environ).get(name, default)
