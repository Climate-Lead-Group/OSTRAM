#!/usr/bin/env python3
"""Validate governed compiled inputs without invoking a solver."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parents[1]
INVENTORY_PATH = REPO_ROOT / "config" / "scenarios" / "registry.json"

EXPECTED_ROOT_SCENARIOS = (
    "BAU",
    "A_Calibrated_BAU",
    "B_Optimised_VRE",
    "C_Target_VRE",
)
EXPECTED_DECISION_SCENARIOS = (
    "A_Calibrated_BAU",
    "A_Calibrated_BAU_Clipped",
    "B_Optimised_VRE",
    "B_Opt_Clipped",
    "B_Opt_DirBidir",
    "B_Opt_DirContractual",
    "B_Opt_IndiaCosts",
    "B_Opt_IndiaCostsFuel",
    "B_Opt_SolarCapex130",
    "B_Opt_SolarCapexHi",
    "B_Opt_SolarCapexSpike",
    "B_Opt_TradeCap15",
    "B_Opt_TxCap150",
    "C_Target_VRE",
    "C_Target_VRE_Clipped",
)
EXPECTED_IGNORE_RULES = ("/workspace/",)
GOVERNED_MANIFEST_COLUMNS = (
    "Scenario",
    "AuthorityClass",
    "SHA256",
    "ByteSize",
    "LineCount",
    "Provenance",
)
GOVERNED_ROOT_AUTHORITY = "RETAINED_FROZEN_ROOT_IDENTITY"
GOVERNED_DERIVED_AUTHORITY = "GOVERNED_ROOT_PLUS_DECLARED_RULES"
DECISION_ROOTS = frozenset(
    {"A_Calibrated_BAU", "B_Optimised_VRE", "C_Target_VRE"}
)
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")


class BaselineValidationError(ValueError):
    """The governed compiled-input contract is not exact."""


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise BaselineValidationError(message)


def _load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise BaselineValidationError(f"cannot load {path}: {exc}") from exc
    _require(isinstance(value, dict), f"{path} must contain a JSON object")
    return value


def scenario_contract(
    inventory_path: Path = INVENTORY_PATH,
) -> tuple[tuple[str, ...], tuple[str, ...]]:
    inventory = _load_json(inventory_path)
    roots = inventory.get("root_scenarios")
    decisions = inventory.get("decision_scenarios")
    _require(isinstance(roots, list), "scenario registry must contain root_scenarios")
    _require(
        isinstance(decisions, list),
        "scenario registry must contain decision_scenarios",
    )
    _require(
        all(isinstance(item, dict) and isinstance(item.get("name"), str) for item in roots),
        "root_scenarios must contain named objects",
    )
    root_names = tuple(item["name"] for item in roots)
    decision_names = tuple(decisions)
    _require(root_names == EXPECTED_ROOT_SCENARIOS, "root scenario contract drift")
    _require(
        decision_names == EXPECTED_DECISION_SCENARIOS,
        "accepted scenario order or membership drift",
    )
    _require(len(set(decision_names)) == 15, "accepted scenarios must be unique")
    return root_names, decision_names


def canonical_scenarios(inventory_path: Path = INVENTORY_PATH) -> tuple[str, ...]:
    return scenario_contract(inventory_path)[1]


def load_governed_manifest(
    path: Path,
    *,
    inventory_path: Path = INVENTORY_PATH,
) -> tuple[dict[str, Any], ...]:
    """Load the post-reconciliation compiled-input acceptance authority."""
    _require(path.is_file(), f"governed comparator manifest is missing: {path}")
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as stream:
            reader = csv.DictReader(stream)
            _require(
                tuple(reader.fieldnames or ()) == GOVERNED_MANIFEST_COLUMNS,
                "unexpected governed comparator manifest columns",
            )
            raw_rows = list(reader)
    except (OSError, csv.Error) as exc:
        raise BaselineValidationError(f"cannot load {path}: {exc}") from exc

    expected_names = canonical_scenarios(inventory_path)
    _require(len(raw_rows) == 15, "governed manifest must contain exactly 15 rows")
    _require(
        tuple(row["Scenario"] for row in raw_rows) == expected_names,
        "governed scenario order or membership drift",
    )

    rows: list[dict[str, Any]] = []
    for row in raw_rows:
        scenario = row["Scenario"]
        expected_authority = (
            GOVERNED_ROOT_AUTHORITY
            if scenario in DECISION_ROOTS
            else GOVERNED_DERIVED_AUTHORITY
        )
        _require(
            row["AuthorityClass"] == expected_authority,
            f"{scenario}: authority class drift",
        )
        _require(
            SHA256_RE.fullmatch(row["SHA256"]) is not None,
            f"{scenario}: invalid SHA-256",
        )
        try:
            byte_size = int(row["ByteSize"])
            line_count = int(row["LineCount"])
        except ValueError as exc:
            raise BaselineValidationError(
                f"{scenario}: byte size and line count must be integers"
            ) from exc
        _require(byte_size > 0, f"{scenario}: byte size must be positive")
        _require(line_count > 0, f"{scenario}: line count must be positive")
        _require(bool(row["Provenance"].strip()), f"{scenario}: provenance is required")
        rows.append(
            {
                "scenario": scenario,
                "authority_class": row["AuthorityClass"],
                "sha256": row["SHA256"],
                "byte_size": byte_size,
                "line_count": line_count,
                "provenance": row["Provenance"],
            }
        )
    return tuple(rows)


def validate_ignore_rules(repo_root: Path = REPO_ROOT) -> None:
    lines = (repo_root / ".gitignore").read_text(encoding="utf-8").splitlines()
    for rule in EXPECTED_IGNORE_RULES:
        _require(lines.count(rule) == 1, f"ignore rule must occur exactly once: {rule}")


def validate_repository(repo_root: Path = REPO_ROOT) -> dict[str, Any]:
    inventory_path = repo_root / INVENTORY_PATH.relative_to(REPO_ROOT)
    roots, decisions = scenario_contract(inventory_path)
    validate_ignore_rules(repo_root)
    return {
        "root_scenario_count": len(roots),
        "root_scenarios": list(roots),
        "scenario_count": len(decisions),
        "scenario_order": list(decisions),
        "generated_roots_ignored": list(EXPECTED_IGNORE_RULES),
    }


def validate_governed_output_files(
    output_root: Path,
    manifest_rows: tuple[dict[str, Any], ...],
) -> tuple[Path, ...]:
    """Require generated decision inputs to match the governed manifest."""
    matched: list[Path] = []
    for item in manifest_rows:
        scenario = item["scenario"]
        filename = (
            f"Pre_processed_{scenario}_0_"
            "StorageDelayN5_OpenBCK_RMCarefulXLSX.txt"
        )
        path = (
            output_root
            / "workspace"
            / "execution"
            / "Executables"
            / f"{scenario}_0"
            / filename
        )
        _require(path.is_file(), f"governed output is missing: {path}")
        _require(path.stat().st_size == item["byte_size"], f"size drift: {path}")
        _require(sha256_file(path) == item["sha256"], f"SHA-256 drift: {path}")
        with path.open("rb") as stream:
            line_count = sum(1 for _ in stream)
        _require(line_count == item["line_count"], f"line-count drift: {path}")
        matched.append(path)
    return tuple(matched)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--governed-manifest",
        type=Path,
        help="Authenticated Stage 2 governed comparator manifest.",
    )
    parser.add_argument(
        "--outputs-root",
        type=Path,
        help="Fresh disposable regeneration root to validate.",
    )
    args = parser.parse_args()
    if args.outputs_root is not None and args.governed_manifest is None:
        parser.error("--outputs-root requires --governed-manifest")

    result = validate_repository()
    if args.governed_manifest is not None:
        rows = load_governed_manifest(args.governed_manifest)
        result["governed_manifest_sha256"] = sha256_file(args.governed_manifest)
        result["governed_scenario_count"] = len(rows)
        if args.outputs_root is not None:
            result["matched_output_files"] = len(
                validate_governed_output_files(args.outputs_root, rows)
            )
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
