"""Materialize the generation calibration and retained controls.

Only explicit complete keys in the frozen table may change a workbook cell.
Historical evidence limits follow the documented 0.98/1.02 convention;
Nepal fallback is lower-only. The 2026 transition is lower-only at 0.95,
except two waste-mapping waivers (2023-2026 lower limits zero).
ABSENT operations retire replay values and expressly removed upper bounds.
Original observations and unresolved source cells remain unchanged.
Bangladesh 2023 coal/gas/combined-oil/hydro references are provisional
fiscal-average estimates: (FY2022/23 + FY2023/24) / 2, used as calendar-2023
proxies, not observed calendar generation. BPDB Annual Reports, both printed
p.14 / PDF p.15, Energy Generation by Fuel Type:
FY2022/23: https://objectstorage.ap-dcc-gazipur-1.oraclecloud15.com/n/axvjbnqprylg/b/V2Ministry/o/office-bpdb/2024/12/651fabb73895432c99a7f6760e1b381d.pdf
FY2023/24 archived report: https://www.energytransitionbd.org/_files/ugd/315ccb_6d364c5d297344e9a02aeab743af34a1.pdf
Pairs in GWh: coal (10081,19138), gas (46013,46579), combined oil
(20650,11797), hydro (610,825). Bounds remain reference * 0.0036 * 0.98/1.02.
SUPPORT.xlsx Current reference sheet keeps the original observations,
both saved-source paths/hashes and the averaging method separately.
The original 787 India/Sri Lanka capacity and investment controls are retained.
The original observations and the supporting evidence are not read in production.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import os
import shutil
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from ostram.paths import resolve_paths
from ostram.profiles import profile_policy
from ostram.pipeline.scenarios.rules.config_keys import check_keys

from openpyxl import load_workbook


PARAM_FILE = "A-O_Parametrization.xlsx"
SHEETS = ("Primary Techs", "Secondary Techs", "Capacities")
GENERATION_YEARS = frozenset({2023, 2024, 2025, 2026})
CAPACITY_YEARS = frozenset(range(2023, 2051))
PIN_ROOT_SCENARIOS = frozenset(
    {"A_Calibrated_BAU", "B_Optimised_VRE", "C_Target_VRE"}
)
RULES_CSV = (
    resolve_paths().scenario_config_root / "rules" / "pwr_min_2023_2026_pin.csv"
)
BACKUP_TAG = "_PRE_PWR_MIN_PIN_"

P_MAX_CAP = "TotalAnnualMaxCapacity"
P_MAX_INV = "TotalAnnualMaxCapacityInvestment"
P_MIN_INV = "TotalAnnualMinCapacityInvestment"
P_ACTIVITY_LOWER = "TotalTechnologyAnnualActivityLowerLimit"
P_ACTIVITY_UPPER = "TotalTechnologyAnnualActivityUpperLimit"
P_RESIDUAL_CAPACITY = "ResidualCapacity"
ALLOWED_PARAMETERS = frozenset(
    {
        P_MAX_CAP,
        P_MAX_INV,
        P_MIN_INV,
        P_ACTIVITY_LOWER,
        P_ACTIVITY_UPPER,
        P_RESIDUAL_CAPACITY,
    }
)
ACTIVITY_PARAMETERS = frozenset({P_ACTIVITY_LOWER, P_ACTIVITY_UPPER})
CAPACITY_PARAMETERS = frozenset({P_RESIDUAL_CAPACITY, P_MAX_INV, P_MIN_INV})
ALLOWED_COUNTRIES = frozenset({"BGD", "BTN", "IND", "LKA", "NPL"})
ORDERED_INDICES = ("REGION", "TECHNOLOGY", "YEAR")
GENERATION_OPERATION = "GENERATION_PIN"
CAPACITY_OPERATION = "CAPACITY_CHRONOLOGY"
EXPECTED_FIELDS = (
    "source_rule_id",
    "parameter",
    "ordered_parameter_indices",
    "region",
    "technology",
    "semantic_technology_group",
    "canonical_country",
    "year",
    "required_root_present",
    "required_root_value",
    "required_root_state",
    "verified_physical_unit",
    "verified_compiled_unit",
    "root_scenarios_with_actual_change",
    "authority_classification",
    "authority_lineage_class",
    "timeslice",
)


@dataclass(frozen=True)
class PinRule:
    """One validated row of the base-year pin rules CSV."""
    source_rule_id: str
    parameter: str
    region: str
    technology: str
    semantic_technology_group: str
    canonical_country: str
    year: int
    present: bool
    value: Decimal
    state: str
    physical_unit: str
    compiled_unit: str
    root_scenarios: tuple[str, ...]
    authority_classification: str
    authority_lineage_class: str
    operation_kind: str
    timeslice: str = ""

    @property
    def complete_key(self) -> tuple[str, str, str, int, str]:
        """Return the (region, technology, parameter, year, timeslice) key of the rule."""
        return self.region, self.technology, self.parameter, self.year, self.timeslice


def calibration_root(scenario: str) -> str | None:
    """Resolve audited authority through the canonical scenario ancestry."""
    if not profile_policy("apply_pwr_min_pin", True):
        return None
    from ostram.pipeline.scenarios.registry import load_registry
    registry = load_registry()
    seen = set()
    while scenario in registry.derived_by_name:
        if scenario in seen:
            raise ValueError(f"Cyclic calibration ancestry: {scenario}")
        seen.add(scenario)
        scenario = registry.derived_by_name[scenario].base_scenario
    return scenario if scenario in PIN_ROOT_SCENARIOS else None


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _parse_decimal(raw: str, *, label: str) -> Decimal:
    try:
        value = Decimal(raw)
    except (InvalidOperation, ValueError) as error:
        raise ValueError(f"{label} is not a decimal: {raw!r}") from error
    if not value.is_finite():
        raise ValueError(f"{label} must be finite: {raw!r}")
    return value


def _validate_units(
    parameter: str,
    technology: str,
    physical_unit: str,
    compiled_unit: str,
) -> None:
    if parameter == "CapacityFactor":
        expected = ("fraction", "fraction")
    elif parameter in ACTIVITY_PARAMETERS:
        expected = ("PJ_per_year_activity", "PJ_per_year")
    elif technology.startswith("MIN"):
        expected = (
            "native_supply_capacity_unit",
            "native_supply_capacity_unit",
        )
    else:
        expected = ("GW", "GW")
    actual = (physical_unit, compiled_unit)
    if actual != expected:
        raise ValueError(
            f"invalid unit pair for {technology}/{parameter}: "
            f"{actual!r}, expected {expected!r}"
        )


def _parse_rule(row: dict[str, str], row_number: int) -> PinRule:
    label = f"rules row {row_number}"
    timeslice = row.get("timeslice", "")
    expected_indices = ("REGION", "TECHNOLOGY", "TIMESLICE", "YEAR") if row['parameter'] == 'CapacityFactor' else ORDERED_INDICES
    if tuple(json.loads(row["ordered_parameter_indices"])) != expected_indices:
        raise ValueError(f"{label}: invalid complete indices")
    if row['parameter'] != 'CapacityFactor' and timeslice:
        raise ValueError(f"{label}: unexpected timeslice on annual parameter")
    parameter, technology, country = row["parameter"], row["technology"], row["canonical_country"]
    year = int(row["year"])
    if row["region"] != "GLOBAL" or parameter not in ALLOWED_PARAMETERS | {"CapacityFactor"}:
        raise ValueError(f"{label}: invalid region/parameter")
    if country not in ALLOWED_COUNTRIES or "MDV" in technology:
        raise ValueError(f"{label}: unsupported country")
    if not technology.startswith(("PWR", "MIN", "TRN")):
        raise ValueError(f"{label}: unsupported technology")
    if technology.startswith("TRN") and technology not in {"TRNINDNOINDWE", "TRNINDEAINDWE"}:
        raise ValueError(f"{label}: transmission correction is outside the reviewed scope")
    if row["required_root_present"] not in {"true", "false"}:
        raise ValueError(f"{label}: invalid presence")
    present = row["required_root_present"] == "true"
    value = _parse_decimal(row["required_root_value"], label=label) if present else Decimal(0)
    if not present and row["required_root_value"] != "":
        raise ValueError(f"{label}: absence must have a blank value")
    expected_state = ("ZERO" if value == 0 else "POSITIVE") if present else "ABSENT"
    if value < 0 or row["required_root_state"] != expected_state:
        raise ValueError(f"{label}: invalid state/value")
    scenarios = tuple(row["root_scenarios_with_actual_change"].split(";"))
    if not scenarios or len(set(scenarios)) != len(scenarios) or not set(scenarios).issubset(PIN_ROOT_SCENARIOS):
        raise ValueError(f"{label}: invalid scenarios")
    identity = row["source_rule_id"]
    operation = identity.split("::", 1)[0]
    authority = (row["authority_classification"], row["authority_lineage_class"])
    expected_id = f"{operation}::{parameter}::GLOBAL::{technology}::{year}"
    if operation == "PWR_MIN_PIN":
        if country not in {"IND", "LKA"} or parameter in ACTIVITY_PARAMETERS or not present or year not in GENERATION_YEARS:
            raise ValueError(f"{label}: inherited replay activity cannot be retained")
        operation_kind = "RETAINED_CONTROL"
    elif operation == "D4_R3_GENERATION":
        if parameter not in ACTIVITY_PARAMETERS or not present or year not in GENERATION_YEARS:
            raise ValueError(f"{label}: invalid R3 generation operation")
        if year == 2026 and (parameter != P_ACTIVITY_LOWER or country == "NPL"):
            raise ValueError(f"{label}: 2026 permits only the allowed non-Nepal lower transition")
        operation_kind = GENERATION_OPERATION
    elif operation == "D4_R3_REMOVE":
        if present or parameter not in ACTIVITY_PARAMETERS or year not in GENERATION_YEARS:
            raise ValueError(f"{label}: invalid activity retirement")
        operation_kind = "ACTIVITY_RETIREMENT"
    elif operation == "D4_PROFILE_INPUT":
        if parameter != 'CapacityFactor' or not present or year not in GENERATION_YEARS or not timeslice or not 0 <= value <= 1:
            raise ValueError(f"{label}: invalid time-slice profile correction")
        expected_id = f"{operation}::{parameter}::GLOBAL::{technology}::{timeslice}::{year}"
        operation_kind = CAPACITY_OPERATION
    elif operation == "D4_PHYSICAL_INPUT":
        if not technology.startswith("PWR") or parameter not in {P_RESIDUAL_CAPACITY, P_MAX_INV, P_MIN_INV} or year not in CAPACITY_YEARS or not present:
            raise ValueError(f"{label}: invalid physical input correction")
        operation_kind = CAPACITY_OPERATION
    elif operation == "D4_INPUT_CORRECTION":
        if parameter != P_MAX_CAP or technology not in {"TRNINDNOINDWE", "TRNINDEAINDWE"} or year not in CAPACITY_YEARS or not present:
            raise ValueError(f"{label}: invalid input correction")
        operation_kind = CAPACITY_OPERATION
    elif operation == "D4_R2_RESIDUAL":
        expected_id = f"{operation}::{technology}::{year}"
        if country not in {"BGD", "BTN"} or parameter != P_RESIDUAL_CAPACITY or year not in CAPACITY_YEARS or not present:
            raise ValueError(f"{label}: invalid residual operation")
        operation_kind = CAPACITY_OPERATION
    elif operation == "D4_R2_CAPACITY_FREEZE":
        expected_id = f"{operation}::{technology}::{parameter}::{year}"
        if country not in {"BGD", "BTN"} or parameter not in {P_MAX_INV, P_MIN_INV} or year not in GENERATION_YEARS or not present:
            raise ValueError(f"{label}: invalid investment freeze")
        operation_kind = CAPACITY_OPERATION
    else:
        raise ValueError(f"{label}: unsupported operation")
    if identity != expected_id:
        raise ValueError(f"{label}: ID does not match complete key")
    _validate_units(parameter, technology, row["verified_physical_unit"], row["verified_compiled_unit"])
    return PinRule(identity, parameter, "GLOBAL", technology, row["semantic_technology_group"], country,
                   year, present, value, expected_state, row["verified_physical_unit"], row["verified_compiled_unit"],
                   scenarios, *authority, operation_kind, timeslice)


def load_pin_rules(
    rules_csv: Path | str = RULES_CSV,
) -> tuple[PinRule, ...]:
    """Load and fail-close validate the complete source-rule allowlist."""
    path = Path(rules_csv)
    if not path.is_file():
        raise FileNotFoundError(path)
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        fields = tuple(reader.fieldnames or ())
        check_keys(fields, EXPECTED_FIELDS, path, "header")
        if fields != EXPECTED_FIELDS:
            raise ValueError(
                f"rule header mismatch: {fields!r} != {EXPECTED_FIELDS!r}"
            )
        rules = tuple(
            _parse_rule(row, row_number)
            for row_number, row in enumerate(reader, start=2)
        )
    ids = [rule.source_rule_id for rule in rules]
    keys = [rule.complete_key for rule in rules]
    if len(ids) != len(set(ids)):
        raise ValueError("duplicate source_rule_id in pin rules")
    if len(keys) != len(set(keys)):
        raise ValueError("duplicate complete source key in pin rules")
    return rules


def _headers(worksheet) -> dict[object, int]:
    headers: dict[object, int] = {}
    for column in range(1, worksheet.max_column + 1):
        value = worksheet.cell(row=1, column=column).value
        if value is None:
            continue
        if value in headers:
            raise ValueError(
                f"{worksheet.title!r} has duplicate header {value!r}"
            )
        headers[value] = column
    return headers


def _year_columns(headers: dict[object, int]) -> dict[int, int]:
    result: dict[int, int] = {}
    for raw, column in headers.items():
        year: int | None = None
        if isinstance(raw, int) and not isinstance(raw, bool):
            year = raw
        elif isinstance(raw, str) and raw.strip().isdigit():
            year = int(raw.strip())
        if year is None:
            continue
        if year in result:
            raise ValueError(f"duplicate year header {year}")
        result[year] = column
    return result


def _cell_decimal(value: object, *, label: str) -> Decimal:
    if value is None or value == "" or isinstance(value, bool):
        raise ValueError(f"{label} is not an explicit numeric value")
    if isinstance(value, str) and value.startswith("="):
        raise ValueError(f"{label} is a formula")
    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError) as error:
        raise ValueError(f"{label} is not numeric: {value!r}") from error
    if not result.is_finite():
        raise ValueError(f"{label} is not finite")
    return result


def _excel_number(value: Decimal) -> int | float:
    if value == value.to_integral_value():
        return int(value)
    result = float(value)
    if not math.isfinite(result):
        raise ValueError(f"cannot serialize non-finite value {value}")
    return result


def make_backup(input_dir: Path) -> Path:
    """Copy the input folder to a timestamped `_PRE_PWR_MIN_PIN_` sibling and return its path."""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = input_dir.parent / f"{input_dir.name}{BACKUP_TAG}{stamp}"
    if backup.exists():
        raise FileExistsError(backup)
    shutil.copytree(input_dir, backup)
    return backup


def restore(input_dir: Path | str, restore_from: Path | str | None = None) -> Path:
    """Replace the input folder with the given backup, or the newest one beside it, and return the backup used."""
    destination = Path(input_dir)
    if restore_from is None:
        candidates = sorted(
            path
            for path in destination.parent.iterdir()
            if path.is_dir()
            and path.name.startswith(f"{destination.name}{BACKUP_TAG}")
        )
        if not candidates:
            raise FileNotFoundError(
                f"no {BACKUP_TAG}* backup found beside {destination}"
            )
        source = candidates[-1]
    else:
        source = Path(restore_from)
    if not source.is_dir():
        raise FileNotFoundError(source)
    if destination.exists():
        shutil.rmtree(destination)
    shutil.copytree(source, destination)
    return source


def apply_pin_rules(
    input_dir: Path | str,
    scenario: str,
    rules_csv: Path | str = RULES_CSV,
    *,
    skip_backup: bool = False,
) -> dict[str, object]:
    """Apply exact rules for one canonical root scenario.

    All structural, key, and row-wide projection checks complete before the
    first workbook assignment.  A validation failure therefore leaves the
    input workbook byte-for-byte unchanged.
    """
    if scenario not in PIN_ROOT_SCENARIOS:
        raise ValueError(f"unsupported pin scenario: {scenario!r}")
    rules_path = Path(rules_csv)
    all_rules = load_pin_rules(rules_path)
    rules = tuple(
        rule for rule in all_rules if scenario in rule.root_scenarios
    )
    if not rules:
        raise ValueError(f"no pin rules apply to scenario {scenario!r}")

    directory = Path(input_dir)
    workbook_path = directory / PARAM_FILE
    if not workbook_path.is_file():
        raise FileNotFoundError(workbook_path)
    workbook = load_workbook(workbook_path)
    temp_path = workbook_path.with_name(
        f".{workbook_path.stem}.pwr-min-pin.tmp.xlsx"
    )
    if temp_path.exists():
        workbook.close()
        raise FileExistsError(temp_path)

    rules_by_row: dict[tuple[str, str, str], list[PinRule]] = defaultdict(list)
    for rule in rules:
        rules_by_row[(rule.technology, rule.parameter, rule.timeslice)].append(rule)

    try:
        missing_sheets = [sheet for sheet in SHEETS if sheet not in workbook]
        if missing_sheets:
            raise ValueError(f"missing required sheets: {missing_sheets}")
        locations: dict[tuple[str, str, str], list[tuple[object, int]]] = defaultdict(
            list
        )
        sheet_metadata: dict[str, tuple[dict[object, int], dict[int, int]]] = {}
        for sheet_name in SHEETS:
            worksheet = workbook[sheet_name]
            headers = _headers(worksheet)
            required_headers = {
                "Tech",
                "Parameter",
                "Projection.Mode",
                "Projection.Parameter",
            }
            missing_headers = sorted(required_headers - set(headers))
            if missing_headers:
                raise ValueError(
                    f"{sheet_name!r} is missing headers {missing_headers}"
                )
            years = _year_columns(headers)
            sheet_metadata[sheet_name] = headers, years
            tech_column = headers["Tech"]
            parameter_column = headers["Parameter"]
            for row_number in range(2, worksheet.max_row + 1):
                key = (
                    worksheet.cell(row=row_number, column=tech_column).value,
                    worksheet.cell(
                        row=row_number, column=parameter_column
                    ).value,
                )
                key = (*key, str(worksheet.cell(row=row_number, column=headers["Timeslices"]).value) if sheet_name == "Capacities" else "")
                if key in rules_by_row:
                    locations[key].append((worksheet, row_number))

        missing_rows = sorted(
            key for key in rules_by_row if len(locations.get(key, ())) == 0 and any(r.present for r in rules_by_row[key])
        )
        duplicate_rows = {
            key: [(worksheet.title, row) for worksheet, row in found]
            for key, found in locations.items()
            if len(found) > 1
        }
        if missing_rows:
            raise ValueError(f"missing target workbook rows: {missing_rows[:10]}")
        if duplicate_rows:
            raise ValueError(f"duplicate target workbook rows: {duplicate_rows}")

        assignments: list[tuple[object, int, Decimal, PinRule]] = []
        projection_flips: list[tuple[object, int]] = []
        for key, row_rules in rules_by_row.items():
            if not locations.get(key):
                continue  # An absent-only operation is already satisfied.
            worksheet, row_number = locations[key][0]
            headers, year_columns = sheet_metadata[worksheet.title]
            target_years = {rule.year for rule in row_rules}
            missing_years = sorted(target_years - set(year_columns))
            if missing_years:
                raise ValueError(
                    f"{worksheet.title}/{key} is missing years {missing_years}"
                )
            mode_cell = worksheet.cell(
                row=row_number, column=headers["Projection.Mode"]
            )
            mode = mode_cell.value
            if not any(rule.present for rule in row_rules):
                pass  # Clearing inactive cells must not activate their row.
            elif mode == "User defined":
                pass
            elif mode in (None, "", "EMPTY"):
                non_target_values = [
                    (year, worksheet.cell(row=row_number, column=column).value)
                    for year, column in sorted(year_columns.items())
                    if year not in target_years
                    and worksheet.cell(row=row_number, column=column).value
                    not in (None, "")
                ]
                if non_target_values:
                    raise ValueError(
                        f"{worksheet.title}/{key} cannot activate row-wide "
                        f"Projection.Mode; populated non-target years: "
                        f"{non_target_values[:8]}"
                    )
                projection_parameter = worksheet.cell(
                    row=row_number, column=headers["Projection.Parameter"]
                ).value
                if projection_parameter not in (None, "", 0, 0.0):
                    raise ValueError(
                        f"{worksheet.title}/{key} has unsafe "
                        f"Projection.Parameter {projection_parameter!r}"
                    )
                projection_flips.append((mode_cell, row_number))
            else:
                raise ValueError(
                    f"{worksheet.title}/{key} has unsupported "
                    f"Projection.Mode {mode!r}"
                )
            for rule in row_rules:
                cell = worksheet.cell(
                    row=row_number, column=year_columns[rule.year]
                )
                if cell.value not in (None, ""):
                    _cell_decimal(
                        cell.value,
                        label=(
                            f"{worksheet.title}/{rule.technology}/"
                            f"{rule.parameter}/{rule.year}"
                        ),
                    )
                assignments.append((cell, row_number, rule.value, rule))

        backup = None if skip_backup else make_backup(directory)
        changed_value_cells = 0
        zero_cells = 0
        positive_cells = 0
        for cell, _row_number, value, rule in assignments:
            current = (
                None
                if cell.value in (None, "")
                else _cell_decimal(
                    cell.value,
                    label=f"current cell for {rule.source_rule_id}",
                )
            )
            if (current != value) if rule.present else (cell.value is not None):
                cell.value = _excel_number(value) if rule.present else None
                changed_value_cells += 1
            if rule.state == "ZERO":
                zero_cells += 1
            elif rule.present:
                positive_cells += 1
        changed_projection_modes = 0
        for mode_cell, _row_number in projection_flips:
            if mode_cell.value != "User defined":
                mode_cell.value = "User defined"
                changed_projection_modes += 1

        for cell, _row_number, value, rule in assignments:
            if not rule.present:
                if cell.value is not None:
                    raise RuntimeError(f"retirement failed: {rule.source_rule_id}")
                continue
            actual = _cell_decimal(
                cell.value,
                label=f"post-apply cell for {rule.source_rule_id}",
            )
            if actual != value:
                raise RuntimeError(
                    f"post-apply mismatch for {rule.source_rule_id}: "
                    f"{actual} != {value}"
                )

        changed = changed_value_cells + changed_projection_modes
        if changed:
            workbook.save(temp_path)
            workbook.close()
            os.replace(temp_path, workbook_path)
        else:
            workbook.close()
        return {
            "status": "PASS",
            "scenario": scenario,
            "input_dir": str(directory),
            "workbook": str(workbook_path),
            "rules_csv": str(rules_path),
            "rules_sha256": _sha256(rules_path),
            "rules_loaded": len(all_rules),
            "rules_applied": len(rules),
            "operation_counts": dict(
                sorted(Counter(rule.operation_kind for rule in rules).items())
            ),
            "workbook_rows_matched": len(rules_by_row),
            "zero_rules": zero_cells,
            "positive_rules": positive_cells,
            "changed_value_cells": changed_value_cells,
            "changed_projection_modes": changed_projection_modes,
            "saved": bool(changed),
            "backup_dir": str(backup) if backup is not None else None,
        }
    finally:
        try:
            workbook.close()
        finally:
            if temp_path.exists():
                temp_path.unlink()


def run(
    input_dir: Path | str,
    scenario: str,
    rules_csv: Path | str = RULES_CSV,
    *,
    skip_backup: bool = False,
) -> dict[str, object]:
    """Compatibility wrapper around :func:`apply_pin_rules`."""
    return apply_pin_rules(
        input_dir,
        scenario,
        rules_csv,
        skip_backup=skip_backup,
    )


def print_summary(log: dict[str, object]) -> None:
    """Print a one-line summary of an :func:`apply_pin_rules` log."""
    print(
        "apply_base_year_pin "
        f"scenario={log['scenario']} status={log['status']} "
        f"rules={log['rules_applied']} rows={log['workbook_rows_matched']} "
        f"value_changes={log['changed_value_cells']} "
        f"projection_mode_changes={log['changed_projection_modes']}"
    )


def main() -> int:
    """Parse the command line, then apply the base-year pin or restore a backup; return the exit code."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", type=Path)
    parser.add_argument("--scenario")
    parser.add_argument("--rules-csv", type=Path, default=RULES_CSV)
    parser.add_argument("--skip-backup", action="store_true")
    parser.add_argument("--restore", action="store_true")
    parser.add_argument("--restore-from", type=Path)
    args = parser.parse_args()
    if args.restore or args.restore_from is not None:
        if args.input_dir is None:
            parser.error("--input-dir is required for restore")
        try:
            source = restore(args.input_dir, args.restore_from)
        except Exception as error:
            print(f"ERROR: {error}", file=sys.stderr)
            return 1
        print(f"Restored {args.input_dir} from {source}")
        return 0
    if args.input_dir is None or args.scenario is None:
        parser.error("--input-dir and --scenario are required")
    try:
        log = apply_pin_rules(
            args.input_dir,
            args.scenario,
            args.rules_csv,
            skip_backup=args.skip_backup,
        )
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    print_summary(log)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
