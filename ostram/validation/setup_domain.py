"""Setup-level validation and reportable effective values."""

from __future__ import annotations

import csv
from decimal import Decimal, InvalidOperation
import os
from pathlib import Path
from typing import Any, Mapping

from ostram.paths import ProjectPaths, resolve_paths
from ostram.setups import (
    SETUP_MANIFEST_ENV,
    SetupError,
    SetupManifest,
    active_setup_id,
    load_manifest,
)


class SetupDomainError(SetupError):
    """Raised when a seed or compiled domain drifts from its setup contract."""


def _domain_contract(manifest: SetupManifest) -> Mapping[str, Any] | None:
    contract = manifest.metadata.get("domain_contract")
    if contract is None:
        return None
    if not isinstance(contract, Mapping):
        raise SetupDomainError("metadata.domain_contract must be a mapping")
    return contract


def _read_domain_set(root: Path, set_name: str) -> set[str]:
    path = Path(root) / f"{set_name}.csv"
    if not path.is_file():
        raise SetupDomainError(f"missing {set_name} authority: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        if reader.fieldnames is None or "VALUE" not in reader.fieldnames:
            raise SetupDomainError(f"{path} requires a VALUE column")
        values = [str(row["VALUE"]).strip() for row in reader]
    if any(not value for value in values):
        raise SetupDomainError(f"{path} contains a blank VALUE")
    duplicates = sorted({value for value in values if values.count(value) > 1})
    if duplicates:
        raise SetupDomainError(f"{path} contains duplicate members: {duplicates}")
    return set(values)


def _text_list(value: object, label: str) -> list[str]:
    if not isinstance(value, list) or any(
        not isinstance(item, str) or not item for item in value
    ):
        raise SetupDomainError(f"{label} must be a list of non-empty strings")
    if len(value) != len(set(value)):
        raise SetupDomainError(f"{label} contains duplicate values")
    return list(value)


def _project_seed_technologies(
    seed: set[str], contract: Mapping[str, Any]
) -> set[str]:
    reconciliation = contract.get("technology_reconciliation")
    if not isinstance(reconciliation, Mapping):
        raise SetupDomainError("domain_contract.technology_reconciliation must be a mapping")
    expected_keys = {
        "drop_before_normalization",
        "normalize_pwr01_suffix",
        "remove_after_normalization",
        "add_after_normalization",
    }
    if set(reconciliation) != expected_keys:
        raise SetupDomainError(
            "technology_reconciliation keys must be exactly "
            f"{sorted(expected_keys)}"
        )
    if reconciliation.get("normalize_pwr01_suffix") is not True:
        raise SetupDomainError("technology reconciliation must enable PWR01 normalization")

    working = set(seed)
    pre_drop = set(
        _text_list(
            reconciliation["drop_before_normalization"],
            "technology_reconciliation.drop_before_normalization",
        )
    )
    missing = pre_drop - working
    if missing:
        raise SetupDomainError(
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
        raise SetupDomainError(
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
        raise SetupDomainError(
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
        raise SetupDomainError(
            f"technology reconciliation additions already exist: {sorted(overlap)}"
        )
    return working | additions


def _generated_delta(
    contract: Mapping[str, Any], set_name: str
) -> set[str]:
    delta = contract.get("generated_delta")
    if not isinstance(delta, Mapping) or set(delta) != {"TECHNOLOGY", "FUEL"}:
        raise SetupDomainError(
            "domain_contract.generated_delta must contain exactly TECHNOLOGY and FUEL"
        )
    return set(_text_list(delta[set_name], f"generated_delta.{set_name}"))


def validate_seed_domain(
    manifest: SetupManifest,
    osemosys_inputs: Path | None,
) -> dict[str, Any] | None:
    """Validate the pre-preparation domain and its count-preserving projection."""

    contract = _domain_contract(manifest)
    if contract is None:
        return None
    if osemosys_inputs is None:
        raise SetupDomainError("domain contract requires the osemosys_inputs authority")

    technology = _read_domain_set(osemosys_inputs, "TECHNOLOGY")
    fuel = _read_domain_set(osemosys_inputs, "FUEL")
    seed_result = {
        "TECHNOLOGY": {"count": len(technology)},
        "FUEL": {"count": len(fuel)},
    }
    projected_technology = _project_seed_technologies(technology, contract)
    projected_result = {
        "TECHNOLOGY": {"count": len(projected_technology)},
        "FUEL": {"count": len(fuel)},
    }
    return {
        "stage": "seed",
        "seed": seed_result,
        "projected_seed": projected_result,
    }


def validate_compiled_domain(
    manifest: SetupManifest,
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
            raise SetupDomainError(
                f"generated_delta.{set_name} already exists in projected seed: {sorted(overlap)}"
            )
        expected = projected[set_name] | declared_delta
        additions = compiled[set_name] - projected[set_name]
        removals = projected[set_name] - compiled[set_name]
        unexpected_additions = compiled[set_name] - expected
        missing_additions = declared_delta - compiled[set_name]
        if removals or unexpected_additions or missing_additions or additions != declared_delta:
            raise SetupDomainError(
                f"compiled {set_name} delta mismatch: expected additions="
                f"{sorted(declared_delta)}, got additions={sorted(additions)}, "
                f"removals={sorted(removals)}, unexpected_additions="
                f"{sorted(unexpected_additions)}, missing_additions={sorted(missing_additions)}"
            )
        compiled_result[set_name] = {"count": len(compiled[set_name])}
        delta_result[set_name] = sorted(additions)

    return {
        "stage": "compiled",
        "compiled": compiled_result,
        "generated_delta": delta_result,
        "seed": seed_result,
    }


def _domain_expanded(
    manifest: SetupManifest,
    osemosys_inputs: Path,
) -> bool:
    """Return True if the actual seed domain is larger than ``seed_set_sizes``.

    This happens after ``country merge`` adds a new country's technologies
    and fuels to the seed CSVs.  The domain has legitimately expanded and
    the contract (which was authored for the original country set) no longer
    applies.
    """
    sizes = manifest.metadata.get("seed_set_sizes")
    try:
        for set_name in ("TECHNOLOGY", "FUEL"):
            expected_count = sizes[set_name]
            actual = _read_domain_set(osemosys_inputs, set_name)
            if len(actual) > expected_count:
                return True
    except (SetupDomainError, KeyError, TypeError):
        pass
    return False


def validate_active_compiled_domain(
    compiled_root: Path,
    *,
    environ: Mapping[str, str] | None = None,
) -> dict[str, Any] | None:
    """Apply the active setup contract before crossing the solver boundary.

    When the seed domain has expanded beyond the contract (e.g. after a
    ``country merge``), the contract no longer reflects reality and strict
    validation is skipped with a diagnostic message.  This lets B2 proceed
    without requiring the user to edit ``setup.yaml`` or delete the
    workspace.
    """

    environment = os.environ if environ is None else environ
    manifest_path = environment.get(SETUP_MANIFEST_ENV)
    if not manifest_path:
        return None
    manifest = load_manifest(
        manifest_path,
        expected_setup=active_setup_id(environment),
    )
    contract = _domain_contract(manifest)
    if contract is None:
        return None

    paths = resolve_paths(environ=environment)

    # If the seed domain grew (country merge added members), the contract
    # is stale.  Skip validation so B2 can proceed.
    if _domain_expanded(manifest, paths.osemosys_inputs):
        print(
            "[setup-domain] Domain has expanded beyond the contract "
            "(country merge detected). Skipping domain validation."
        )
        return None

    return validate_compiled_domain(
        manifest,
        osemosys_inputs=paths.osemosys_inputs,
        compiled_root=compiled_root,
    )


def validate_setup(
    manifest: SetupManifest,
    *,
    paths: ProjectPaths,
) -> dict[str, Any]:
    """Validate one setup and return its report.

    The check requires every authority to exist, every effective value in the
    manifest metadata to be a finite number and the seed domain to meet the
    setup's domain contract when it declares one. The report holds the
    setup id, the manifest path and digest, the authority paths, the
    policies, the effective values and the domain check result.

    Raises:
        SetupManifestError: If an authority is invalid or missing.
        SetupDomainError: If the seed domain breaks the domain contract.
        ValueError: If an effective value is not a finite number.
    """

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
        "setup_id": manifest.setup_id,
        "manifest": str(manifest.path),
        "manifest_sha256": manifest.digest,
        "authorities": {name: str(Path(path)) for name, path in authorities.items()},
        "policies": dict(manifest.policies),
        "effective_values": effective,
        "domain": domain,
    }
