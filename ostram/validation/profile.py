"""Profile-level validation and reportable effective values."""

from __future__ import annotations

import csv
from decimal import Decimal, InvalidOperation
import hashlib
import os
from pathlib import Path
import re
from typing import Any, Mapping

from ostram.paths import ProjectPaths, resolve_paths
from ostram.profiles import (
    PROFILE_MANIFEST_ENV,
    ProfileError,
    ProfileManifest,
    active_profile_id,
    load_manifest,
)


_DOMAIN_ALGORITHM = "sha256-sorted-values-v1"
_SHA256 = re.compile(r"^[0-9a-f]{64}$")


class ProfileDomainError(ProfileError):
    """Raised when a seed or compiled domain drifts from its profile contract."""


def _domain_contract(manifest: ProfileManifest) -> Mapping[str, Any] | None:
    contract = manifest.metadata.get("domain_contract")
    if contract is None:
        return None
    if not isinstance(contract, Mapping):
        raise ProfileDomainError("metadata.domain_contract must be a mapping")
    if contract.get("membership_hash") != _DOMAIN_ALGORITHM:
        raise ProfileDomainError(
            "metadata.domain_contract.membership_hash must be "
            f"{_DOMAIN_ALGORITHM!r}"
        )
    return contract


def _domain_digest(values: set[str]) -> str:
    payload = "".join(f"{value}\n" for value in sorted(values)).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _read_domain_set(root: Path, set_name: str) -> set[str]:
    path = Path(root) / f"{set_name}.csv"
    if not path.is_file():
        raise ProfileDomainError(f"missing {set_name} authority: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        if reader.fieldnames is None or "VALUE" not in reader.fieldnames:
            raise ProfileDomainError(f"{path} requires a VALUE column")
        values = [str(row["VALUE"]).strip() for row in reader]
    if any(not value for value in values):
        raise ProfileDomainError(f"{path} contains a blank VALUE")
    duplicates = sorted({value for value in values if values.count(value) > 1})
    if duplicates:
        raise ProfileDomainError(f"{path} contains duplicate members: {duplicates}")
    return set(values)


def _expected_stage_set(
    contract: Mapping[str, Any],
    stage: str,
    set_name: str,
) -> tuple[int, str]:
    stages = contract.get("stages")
    if not isinstance(stages, Mapping):
        raise ProfileDomainError("domain_contract.stages must be a mapping")
    stage_contract = stages.get(stage)
    if not isinstance(stage_contract, Mapping):
        raise ProfileDomainError(f"domain_contract.stages.{stage} must be a mapping")
    set_contract = stage_contract.get(set_name)
    if not isinstance(set_contract, Mapping):
        raise ProfileDomainError(
            f"domain_contract.stages.{stage}.{set_name} must be a mapping"
        )
    count = set_contract.get("count")
    digest = set_contract.get("sha256")
    if not isinstance(count, int) or count < 0:
        raise ProfileDomainError(f"invalid {stage} {set_name} count: {count!r}")
    if not isinstance(digest, str) or _SHA256.fullmatch(digest) is None:
        raise ProfileDomainError(f"invalid {stage} {set_name} sha256: {digest!r}")
    return count, digest


def _validate_members(
    values: set[str],
    *,
    contract: Mapping[str, Any],
    stage: str,
    set_name: str,
) -> dict[str, Any]:
    expected_count, expected_digest = _expected_stage_set(
        contract, stage, set_name
    )
    actual_digest = _domain_digest(values)
    if len(values) != expected_count or actual_digest != expected_digest:
        raise ProfileDomainError(
            f"{stage} {set_name} domain mismatch: expected "
            f"count={expected_count}, sha256={expected_digest}; got "
            f"count={len(values)}, sha256={actual_digest}"
        )
    return {"count": len(values), "sha256": actual_digest}


def _text_list(value: object, label: str) -> list[str]:
    if not isinstance(value, list) or any(
        not isinstance(item, str) or not item for item in value
    ):
        raise ProfileDomainError(f"{label} must be a list of non-empty strings")
    if len(value) != len(set(value)):
        raise ProfileDomainError(f"{label} contains duplicate values")
    return list(value)


def _project_seed_technologies(
    seed: set[str], contract: Mapping[str, Any]
) -> set[str]:
    reconciliation = contract.get("technology_reconciliation")
    if not isinstance(reconciliation, Mapping):
        raise ProfileDomainError("domain_contract.technology_reconciliation must be a mapping")
    expected_keys = {
        "drop_before_normalization",
        "normalize_pwr01_suffix",
        "remove_after_normalization",
        "add_after_normalization",
    }
    if set(reconciliation) != expected_keys:
        raise ProfileDomainError(
            "technology_reconciliation keys must be exactly "
            f"{sorted(expected_keys)}"
        )
    if reconciliation.get("normalize_pwr01_suffix") is not True:
        raise ProfileDomainError("technology reconciliation must enable PWR01 normalization")

    working = set(seed)
    pre_drop = set(
        _text_list(
            reconciliation["drop_before_normalization"],
            "technology_reconciliation.drop_before_normalization",
        )
    )
    missing = pre_drop - working
    if missing:
        raise ProfileDomainError(
            f"technology reconciliation pre-normalization removals are absent: {sorted(missing)}"
        )
    working -= pre_drop

    normalized = [
        value[:-2]
        if value.startswith("PWR") and len(value) == 13 and value.endswith("01")
        else value
        for value in working
    ]
    if len(normalized) != len(set(normalized)):
        collisions = sorted(
            {value for value in normalized if normalized.count(value) > 1}
        )
        raise ProfileDomainError(
            f"technology reconciliation creates duplicate identities: {collisions}"
        )
    working = set(normalized)

    post_remove = set(
        _text_list(
            reconciliation["remove_after_normalization"],
            "technology_reconciliation.remove_after_normalization",
        )
    )
    missing = post_remove - working
    if missing:
        raise ProfileDomainError(
            f"technology reconciliation post-normalization removals are absent: {sorted(missing)}"
        )
    working -= post_remove

    additions = set(
        _text_list(
            reconciliation["add_after_normalization"],
            "technology_reconciliation.add_after_normalization",
        )
    )
    overlap = additions & working
    if overlap:
        raise ProfileDomainError(
            f"technology reconciliation additions already exist: {sorted(overlap)}"
        )
    return working | additions


def _generated_delta(
    contract: Mapping[str, Any], set_name: str
) -> set[str]:
    delta = contract.get("generated_delta")
    if not isinstance(delta, Mapping) or set(delta) != {"TECHNOLOGY", "FUEL"}:
        raise ProfileDomainError(
            "domain_contract.generated_delta must contain exactly TECHNOLOGY and FUEL"
        )
    return set(_text_list(delta[set_name], f"generated_delta.{set_name}"))


def validate_seed_domain(
    manifest: ProfileManifest,
    osemosys_inputs: Path | None,
) -> dict[str, Any] | None:
    """Validate the pre-preparation domain and its count-preserving projection."""

    contract = _domain_contract(manifest)
    if contract is None:
        return None
    if osemosys_inputs is None:
        raise ProfileDomainError("domain contract requires the osemosys_inputs authority")

    technology = _read_domain_set(osemosys_inputs, "TECHNOLOGY")
    fuel = _read_domain_set(osemosys_inputs, "FUEL")
    seed_result = {
        "TECHNOLOGY": _validate_members(
            technology,
            contract=contract,
            stage="seed",
            set_name="TECHNOLOGY",
        ),
        "FUEL": _validate_members(
            fuel,
            contract=contract,
            stage="seed",
            set_name="FUEL",
        ),
    }
    projected_technology = _project_seed_technologies(technology, contract)
    projected_result = {
        "TECHNOLOGY": _validate_members(
            projected_technology,
            contract=contract,
            stage="projected_seed",
            set_name="TECHNOLOGY",
        ),
        "FUEL": _validate_members(
            fuel,
            contract=contract,
            stage="projected_seed",
            set_name="FUEL",
        ),
    }
    return {
        "stage": "seed",
        "seed": seed_result,
        "projected_seed": projected_result,
    }


def validate_compiled_domain(
    manifest: ProfileManifest,
    *,
    osemosys_inputs: Path,
    compiled_root: Path,
) -> dict[str, Any] | None:
    """Reject every compiled addition/removal outside the declared generated delta."""

    contract = _domain_contract(manifest)
    if contract is None:
        return None
    seed_result = validate_seed_domain(manifest, osemosys_inputs)
    assert seed_result is not None

    seed_technology = _read_domain_set(osemosys_inputs, "TECHNOLOGY")
    seed_fuel = _read_domain_set(osemosys_inputs, "FUEL")
    projected = {
        "TECHNOLOGY": _project_seed_technologies(seed_technology, contract),
        "FUEL": seed_fuel,
    }
    compiled = {
        "TECHNOLOGY": _read_domain_set(compiled_root, "TECHNOLOGY"),
        "FUEL": _read_domain_set(compiled_root, "FUEL"),
    }
    compiled_result: dict[str, Any] = {}
    delta_result: dict[str, list[str]] = {}
    for set_name in ("TECHNOLOGY", "FUEL"):
        declared_delta = _generated_delta(contract, set_name)
        overlap = projected[set_name] & declared_delta
        if overlap:
            raise ProfileDomainError(
                f"generated_delta.{set_name} already exists in projected seed: {sorted(overlap)}"
            )
        expected = projected[set_name] | declared_delta
        additions = compiled[set_name] - projected[set_name]
        removals = projected[set_name] - compiled[set_name]
        unexpected_additions = compiled[set_name] - expected
        missing_additions = declared_delta - compiled[set_name]
        if removals or unexpected_additions or missing_additions or additions != declared_delta:
            raise ProfileDomainError(
                f"compiled {set_name} delta mismatch: expected additions="
                f"{sorted(declared_delta)}, got additions={sorted(additions)}, "
                f"removals={sorted(removals)}, unexpected_additions="
                f"{sorted(unexpected_additions)}, missing_additions={sorted(missing_additions)}"
            )
        compiled_result[set_name] = _validate_members(
            compiled[set_name],
            contract=contract,
            stage="compiled",
            set_name=set_name,
        )
        delta_result[set_name] = sorted(additions)

    return {
        "stage": "compiled",
        "compiled": compiled_result,
        "generated_delta": delta_result,
        "seed": seed_result,
    }


def validate_active_compiled_domain(
    compiled_root: Path,
    *,
    environ: Mapping[str, str] | None = None,
) -> dict[str, Any] | None:
    """Apply the active profile contract before crossing the solver boundary."""

    environment = os.environ if environ is None else environ
    manifest_path = environment.get(PROFILE_MANIFEST_ENV)
    if not manifest_path:
        return None
    manifest = load_manifest(
        manifest_path,
        expected_profile=active_profile_id(environment),
    )
    if _domain_contract(manifest) is None:
        return None
    paths = resolve_paths(environ=environment)
    return validate_compiled_domain(
        manifest,
        osemosys_inputs=paths.osemosys_inputs,
        compiled_root=compiled_root,
    )


def validate_profile(
    manifest: ProfileManifest,
    *,
    paths: ProjectPaths,
) -> dict[str, Any]:
    authorities = manifest.source_paths(paths, require_exists=True)
    effective = dict(manifest.metadata.get("effective_values", {}))
    for name, value in effective.items():
        try:
            number = Decimal(str(value))
        except (InvalidOperation, ValueError) as error:
            raise ValueError(f"effective value {name!r} is not numeric: {value!r}") from error
        if not number.is_finite():
            raise ValueError(f"effective value {name!r} is not finite")
    domain = validate_seed_domain(manifest, authorities.get("osemosys_inputs"))
    return {
        "profile_id": manifest.profile_id,
        "manifest": str(manifest.path),
        "manifest_sha256": manifest.digest,
        "authorities": {name: str(Path(path)) for name, path in authorities.items()},
        "policies": dict(manifest.policies),
        "effective_values": effective,
        "domain": domain,
    }
