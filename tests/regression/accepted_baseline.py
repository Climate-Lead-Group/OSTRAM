#!/usr/bin/env python3
"""Read-only validation for the portable accepted 15-scenario baseline."""

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
BASELINE_PATH = HERE / "reports" / "accepted_compiled_solver_baseline_15.json"
INVENTORY_PATH = REPO_ROOT / "t1_confection" / "scenario_registry.json"
PRE_CORRECTION_REPORT = (
    HERE / "reports" / "pre_correction_41a54e5_compiled_input_equivalence_15.json"
)
OBSOLETE_REPORT = HERE / "reports" / "final_compiled_input_equivalence_15.json"

EXPECTED_BASELINE_SHA256 = (
    "4f864f0b65c7838b70e5cd18e44679c190669f8f25ed6013447ab01beb0ed67a"
)
EXPECTED_PRE_CORRECTION_SIZE = 3_670
EXPECTED_PRE_CORRECTION_SHA256 = (
    "85c8489e65e8028dd5955dbbf204f2222796799d132ba979cf238e60008b8286"
)
EXPECTED_PRE_CORRECTION_COMMIT = "41a54e51fd5a0776569b4900c44c624f09cc1f09"
EXPECTED_REFERENCE_COMMIT = "8dd3361a1fc7f2c9ea4df51d5b2d0e50e0ce8554"
EXPECTED_REFERENCE_TREE = "65cbef7e977084b0a45f7bd8fca958d69ca916ce"
EXPECTED_CORRECTION_COMMIT = "d295dcccca6c62e88f74484d7b8201e950881c3f"
EXPECTED_MANIFEST_SHA256 = (
    "778b4706522bc2b29911e74d5b31d24355c84cbe4c0c7d11d1c9680b2ddc9916"
)
EXPECTED_IGNORE_RULES = (
    "/t1_confection/A1_Outputs/",
    "/t1_confection/A2_Output_Params/",
    "/t1_confection/A2_Outputs_Params_otoole/",
    "/t1_confection/Executables/",
    "/t1_confection/Outputs/",
)
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

# Scenario, byte count, SHA-256. Filenames and paths are derived exactly from
# the scenario name and checked below.
EXPECTED_COMPILED_INPUTS = (
    ("A_Calibrated_BAU", 3_347_081, "3012716e1e042518ca01222a325ae9314ea44ba18a23380a9577a19822ea2d7d"),
    ("A_Calibrated_BAU_Clipped", 3_349_191, "9729e51499f1f698d72f25b470260948c7f61c742e6372b980fe6700a4f94ea6"),
    ("B_Optimised_VRE", 3_358_808, "336602360676a41b69811fd671ee0bc2309296370efc0c22c5af458ad0f4708c"),
    ("B_Opt_Clipped", 3_360_735, "8a81ce55bc6cf6bb9a6b45ce41b1e15cdfffc894b226578152107de21e275106"),
    ("B_Opt_DirBidir", 3_360_735, "8a81ce55bc6cf6bb9a6b45ce41b1e15cdfffc894b226578152107de21e275106"),
    ("B_Opt_DirContractual", 3_341_583, "db71d809bf22615c3853b4e877ca99b382f3cae135b6a989ecb7e2840162f125"),
    ("B_Opt_IndiaCosts", 3_357_675, "9d68372f13bfb40766c3d5f36f5829b5566a0ead7ef03cdfb692cf4b2e4a1154"),
    ("B_Opt_IndiaCostsFuel", 3_358_083, "5262fd3fb2e097396c39b34a7da436661bff977a7bbd49164fcadae1dc96c49f"),
    ("B_Opt_SolarCapex130", 3_361_071, "bb054bca2d64eeee22d85d58255b600bd2f8e769e5ec49840c41958474b6e1a9"),
    ("B_Opt_SolarCapexHi", 3_361_064, "21620a29fd8f0187301a3ce3e223741063161b52be776012e90d548d6253db19"),
    ("B_Opt_SolarCapexSpike", 3_360_749, "8cf20bdc0cdb389fde33a5f76c6104fc417758fe8bf8baa71b3049d93238d0fb"),
    ("B_Opt_TradeCap15", 3_365_324, "14b83c8624594c8aa2bfa26166084b84a50677657f1d1cba63bb78e6abf78db4"),
    ("B_Opt_TxCap150", 3_367_145, "d3fd23f7ed87eacdf38f18deaa8435e54bda24e7c105d8c3ae64dcc6179aeaaf"),
    ("C_Target_VRE", 3_380_564, "0544dea388ea3209d8e1a9b260ad288dc3e70447f7f44e091a53478b7a1a9457"),
    ("C_Target_VRE_Clipped", 3_382_283, "7e885b6146e121ac2f86181589cef70738d83f1c5457c982b9a6152a6b233851"),
)

SHA256_RE = re.compile(r"^[0-9a-f]{64}$")


class BaselineValidationError(ValueError):
    """The portable accepted-baseline contract is not exact."""


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


def canonical_scenarios(inventory_path: Path = INVENTORY_PATH) -> tuple[str, ...]:
    inventory = _load_json(inventory_path)
    scenarios = inventory.get("decision_scenarios")
    _require(
        isinstance(scenarios, list),
        "scenario registry must contain decision_scenarios",
    )
    return tuple(scenarios)


def validate_record(
    record: dict[str, Any],
    *,
    inventory_path: Path = INVENTORY_PATH,
) -> tuple[dict[str, Any], ...]:
    _require(
        record.get("schema") == "ostram-portable-accepted-compiled-solver-baseline-v1",
        "unexpected accepted-baseline schema",
    )
    _require(record.get("scenario_count") == 15, "scenario_count must be 15")

    lineage = record.get("lineage")
    _require(isinstance(lineage, dict), "lineage must be an object")
    _require(lineage.get("reference_commit") == EXPECTED_REFERENCE_COMMIT, "reference commit drift")
    _require(lineage.get("reference_tree") == EXPECTED_REFERENCE_TREE, "reference tree drift")
    _require(
        lineage.get("accepted_correction_commit") == EXPECTED_CORRECTION_COMMIT,
        "accepted correction drift",
    )
    _require(
        lineage.get("accepted_correction_tree") == EXPECTED_REFERENCE_TREE,
        "accepted correction tree drift",
    )
    _require(
        lineage.get("protected_manifest_sha256") == EXPECTED_MANIFEST_SHA256,
        "protected manifest digest drift",
    )

    scenarios = record.get("scenarios")
    _require(isinstance(scenarios, list), "scenarios must be a list")
    _require(len(scenarios) == 15, "record must contain exactly 15 scenarios")
    names = tuple(item.get("scenario") for item in scenarios)
    expected_names = tuple(item[0] for item in EXPECTED_COMPILED_INPUTS)
    _require(names == expected_names, "accepted scenario order or membership drift")
    _require(names == canonical_scenarios(inventory_path), "canonical inventory order drift")
    _require(len(set(names)) == 15, "accepted scenarios must be unique")

    for item, (scenario, size_bytes, sha256) in zip(
        scenarios, EXPECTED_COMPILED_INPUTS, strict=True
    ):
        filename = (
            f"Pre_processed_{scenario}_0_"
            "StorageDelayN5_OpenBCK_RMCarefulXLSX.txt"
        )
        relative_path = (
            f"t1_confection/Executables/{scenario}_0/{filename}"
        )
        _require(item.get("filename") == filename, f"{scenario}: filename drift")
        _require(item.get("relative_path") == relative_path, f"{scenario}: path drift")
        _require(item.get("size_bytes") == size_bytes, f"{scenario}: byte-count drift")
        _require(item.get("sha256") == sha256, f"{scenario}: SHA-256 drift")
        _require(SHA256_RE.fullmatch(sha256) is not None, f"{scenario}: invalid SHA-256")
        _require(item.get("primal_feasible") is True, f"{scenario}: primal-feasible flag drift")

    seed = record.get("accepted_a_combined_output_seed")
    _require(isinstance(seed, dict), "accepted A seed must be an object")
    _require(seed.get("size_bytes") == 44_743_620, "accepted A seed size drift")
    _require(
        seed.get("sha256")
        == "762a7b926f91710846dc37e474747f5d670aed3d8746d7b74117ee978e645f5a",
        "accepted A seed digest drift",
    )
    return tuple(scenarios)


def load_accepted_record(
    path: Path = BASELINE_PATH,
    *,
    inventory_path: Path = INVENTORY_PATH,
) -> dict[str, Any]:
    _require(path.is_file(), f"accepted baseline is missing: {path}")
    _require(
        sha256_file(path) == EXPECTED_BASELINE_SHA256,
        "portable accepted-baseline file bytes drift",
    )
    record = _load_json(path)
    validate_record(record, inventory_path=inventory_path)
    return record


def load_governed_manifest(
    path: Path,
    *,
    inventory_path: Path = INVENTORY_PATH,
) -> tuple[dict[str, Any], ...]:
    """Load the compact post-reconciliation acceptance authority."""
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


def validate_report_lineage(
    repo_root: Path = REPO_ROOT,
    record: dict[str, Any] | None = None,
) -> None:
    record = record or load_accepted_record()
    report = repo_root / PRE_CORRECTION_REPORT.relative_to(REPO_ROOT)
    obsolete = repo_root / OBSOLETE_REPORT.relative_to(REPO_ROOT)
    _require(report.is_file(), "pre-correction report is missing")
    _require(not obsolete.exists(), "obsolete final report name must remain absent")
    _require(report.stat().st_size == EXPECTED_PRE_CORRECTION_SIZE, "pre-correction size drift")
    _require(
        sha256_file(report) == EXPECTED_PRE_CORRECTION_SHA256,
        "pre-correction report bytes drift",
    )
    old = _load_json(report)
    _require(
        old.get("verified_source_commit") == EXPECTED_PRE_CORRECTION_COMMIT,
        "pre-correction source lineage drift",
    )
    _require(
        record["lineage"]["reference_commit"] != old["verified_source_commit"],
        "accepted and pre-correction records must retain distinct source lineage",
    )


def validate_ignore_rules(repo_root: Path = REPO_ROOT) -> None:
    lines = (repo_root / ".gitignore").read_text(encoding="utf-8").splitlines()
    for rule in EXPECTED_IGNORE_RULES:
        _require(lines.count(rule) == 1, f"ignore rule must occur exactly once: {rule}")


def validate_repository(repo_root: Path = REPO_ROOT) -> dict[str, Any]:
    record_path = repo_root / BASELINE_PATH.relative_to(REPO_ROOT)
    inventory_path = repo_root / INVENTORY_PATH.relative_to(REPO_ROOT)
    record = load_accepted_record(record_path, inventory_path=inventory_path)
    validate_report_lineage(repo_root, record)
    validate_ignore_rules(repo_root)
    return {
        "scenario_count": len(record["scenarios"]),
        "scenario_order": [item["scenario"] for item in record["scenarios"]],
        "baseline_sha256": EXPECTED_BASELINE_SHA256,
        "protected_manifest_sha256": EXPECTED_MANIFEST_SHA256,
    }


def validate_output_files(
    output_root: Path,
    record: dict[str, Any] | None = None,
) -> tuple[Path, ...]:
    """Require all 15 solver-consumed text targets to match raw accepted bytes."""
    record = record or load_accepted_record()
    matched: list[Path] = []
    for item in record["scenarios"]:
        path = output_root / item["relative_path"]
        _require(path.is_file(), f"accepted output is missing: {path}")
        _require(path.stat().st_size == item["size_bytes"], f"size drift: {path}")
        _require(sha256_file(path) == item["sha256"], f"SHA-256 drift: {path}")
        matched.append(path)
    return tuple(matched)


def validate_governed_output_files(
    output_root: Path,
    manifest_rows: tuple[dict[str, Any], ...],
) -> tuple[Path, ...]:
    """Require generated decision inputs to match the governed compact manifest."""
    matched: list[Path] = []
    for item in manifest_rows:
        scenario = item["scenario"]
        filename = (
            f"Pre_processed_{scenario}_0_"
            "StorageDelayN5_OpenBCK_RMCarefulXLSX.txt"
        )
        path = (
            output_root
            / "t1_confection"
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
        "--outputs-root",
        type=Path,
        help="Optionally verify the exact 15 generated targets under this root.",
    )
    parser.add_argument(
        "--governed-manifest",
        type=Path,
        help=(
            "Use the post-reconciliation governed CSV manifest. When combined "
            "with --outputs-root, validates all 15 freshly generated targets."
        ),
    )
    args = parser.parse_args()
    result = validate_repository()
    if args.governed_manifest is not None:
        rows = load_governed_manifest(args.governed_manifest)
        result["governed_manifest_sha256"] = sha256_file(args.governed_manifest)
        result["governed_scenario_count"] = len(rows)
        if args.outputs_root is not None:
            result["matched_output_files"] = len(
                validate_governed_output_files(args.outputs_root, rows)
            )
    elif args.outputs_root is not None:
        # The legacy JSON remains available only for historical-source checks.
        result["matched_output_files"] = len(validate_output_files(args.outputs_root))
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
