"""Fail-closed loaders for migrated interconnector minimum authorities.

The production numeric values live only in ``Interconnector_Params`` in the
v18 workbook.  This module contains schema, domain and metadata guards, and
no numeric schedule fallback.

Both supported workbook shapes are accepted:

* the raw v18 sheet, whose first column is ``scenario``; and
* a scenario-materialized sheet, from which ``scenario`` has been removed.

Either shape may end with the audit-only ``IEEE Reference`` column. The loader does not consume that column and continues
to reject every other schema extension.

In the raw sheet, ``TotalAnnualMinCapacityInvestment`` rows may name a registered
root scenario instead of ``BAU``. A root's row replaces the BAU row of the same
link for that root, and so for every scenario derived from it. Every other
parameter keeps ``BAU`` rows only.
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, Mapping

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from ..registry import ROOT_SCENARIOS


AUTHORITY_SHEET_NAME = "Interconnector_Params"
AUTHORITY_YEARS = tuple(range(2023, 2051))
BAU_SCENARIO = "BAU"

MINIMUM_CONTRIBUTION_PARAMETER = "TotalAnnualMinCapacityInvestment"
MINIMUM_CONTRIBUTION_TECHS = frozenset({
    "TRNBTNXXBGDXX",
    "TRNBTNXXINDEA",
    "TRNBTNXXINDNE",
    "TRNINDEAINDWE",
    "TRNINDEANPLXX",
    "TRNINDNOINDWE",
    "TRNINDNONPLXX",
    "TRNINDSOINDWE",
    "TRNINDSOLKAXX",
    "TRNLKAXXMDVXX",
    "TRNNPLXXBGDXX",
})
# Cross-border links without a BAU row. A root scenario's row may add one.
ROOT_ONLY_CONTRIBUTION_TECHS = frozenset({
    "TRNBGDXXINDEA",
    "TRNBGDXXINDNE",
    "TRNMDVXXINDSO",
})

MINIMUM_BOUNDARY_PARAMETER = "MinimumInvestmentClampBoundary"
MINIMUM_BOUNDARY_TECHS = frozenset({
    "TRNBGDXXINDEA",
    "TRNNPLXXBGDXX",
})
MINIMUM_BOUNDARY_SOURCE = "Minimum-only LinkFreeze compatibility boundary"

_COMMON_HEADERS = (
    "Tech.ID",
    "Tech",
    "Tech.Name",
    "Parameter.ID",
    "Parameter",
    "Unit",
    "Projection.Mode",
    "Projection.Parameter",
    *AUTHORITY_YEARS,
    "Source",
)
_RAW_HEADERS = ("scenario", *_COMMON_HEADERS)
_MATERIALIZED_HEADERS = _COMMON_HEADERS
_OPTIONAL_AUDIT_HEADERS = ("IEEE Reference",)

# These guards identify the intended rows independently of their numeric
# schedules.  The values were mechanically copied from the effective target
# TotalAnnualMinCapacityInvestment metadata.
_CONTRIBUTION_TECH_METADATA = {
    "TRNBTNXXBGDXX": (
        10,
        "Transmission interconnection from Bhutan, region XX to Bangladesh, region XX",
    ),
    "TRNBTNXXINDEA": (
        10,
        "Transmission interconnection from Bhutan, region XX to India, region EA",
    ),
    "TRNBTNXXINDNE": (
        10,
        "Transmission interconnection from Bhutan, region XX to India, region NE",
    ),
    "TRNINDEAINDWE": (
        195,
        "Transmission interconnection from India, region EA to India, region WE",
    ),
    "TRNINDEANPLXX": (
        101,
        "Transmission interconnection from India, region EA to Nepal, region XX",
    ),
    "TRNINDNOINDWE": (
        93,
        "Transmission interconnection from India, region NO to India, region WE",
    ),
    "TRNINDNONPLXX": (
        123,
        "Transmission interconnection from India, region NO to Nepal, region XX",
    ),
    "TRNINDSOINDWE": (
        141,
        "Transmission interconnection from India, region SO to India, region WE",
    ),
    "TRNINDSOLKAXX": (
        22,
        "Transmission interconnection from India, region SO to Sri Lanka, region XX",
    ),
    "TRNLKAXXMDVXX": (
        110,
        "Transmission interconnection from Sri Lanka, region XX to Maldives, region XX",
    ),
    "TRNNPLXXBGDXX": (
        205,
        "Transmission interconnection from Nepal, region XX to Bangladesh, region XX",
    ),
    "TRNBGDXXINDEA": (
        205,
        "Transmission interconnection from Bangladesh, region XX to India, region EA",
    ),
    "TRNBGDXXINDNE": (
        95,
        "Transmission interconnection from Bangladesh, region XX to India, region NE",
    ),
    "TRNMDVXXINDSO": (
        141,
        "Transmission interconnection from Maldives, region XX to India, region SO",
    ),
}
_BOUNDARY_TECH_METADATA = {
    "TRNBGDXXINDEA": (
        205,
        "Transmission interconnection from Bangladesh, region XX to India, region EA",
    ),
    "TRNNPLXXBGDXX": (
        205,
        "Transmission interconnection from Nepal, region XX to Bangladesh, region XX",
    ),
}


def _numeric_decimal(value: object, context: str) -> Decimal:
    """Return an exact finite, non-negative Decimal for a numeric cell."""
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        if isinstance(value, str) and value.startswith("="):
            detail = "formula"
        elif value is None:
            detail = "blank"
        else:
            detail = f"non-numeric value {value!r}"
        raise ValueError(f"{context}: expected a numeric value, got {detail}")
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(
            f"{context}: expected a numeric value, got {value!r}"
        ) from exc
    if not number.is_finite() or number < 0:
        raise ValueError(
            f"{context}: expected a finite non-negative value, got {value!r}"
        )
    return number


def _header_columns(ws: Worksheet) -> tuple[Dict[object, int], bool]:
    headers = tuple(cell.value for cell in ws[1])
    raw_shapes = {
        _RAW_HEADERS,
        (*_RAW_HEADERS, *_OPTIONAL_AUDIT_HEADERS),
    }
    materialized_shapes = {
        _MATERIALIZED_HEADERS,
        (*_MATERIALIZED_HEADERS, *_OPTIONAL_AUDIT_HEADERS),
    }
    if headers in raw_shapes:
        raw_shape = True
    elif headers in materialized_shapes:
        raw_shape = False
    else:
        expected_raw = [list(shape) for shape in raw_shapes]
        expected_materialized = [list(shape) for shape in materialized_shapes]
        raise ValueError(
            f"{AUTHORITY_SHEET_NAME} has an unexpected header shape; "
            f"expected one of raw {expected_raw!r} or materialized "
            f"{expected_materialized!r}, got {list(headers)!r}"
        )
    return {header: index for index, header in enumerate(headers)}, raw_shape


def _validate_mapping(
    authority: Mapping[str, Mapping[int, object]],
    *,
    label: str,
    techs: frozenset[str],
    optional_techs: frozenset[str] = frozenset(),
) -> None:
    actual_techs = set(authority)
    missing = techs - actual_techs
    extra = actual_techs - techs - optional_techs
    if missing or extra:
        raise ValueError(
            f"{label} technology domain mismatch: "
            f"missing={sorted(missing)}, "
            f"extra={sorted(extra)}"
        )
    expected_years = set(AUTHORITY_YEARS)
    for tech in sorted(actual_techs):
        actual_years = set(authority[tech])
        if actual_years != expected_years:
            raise ValueError(
                f"{label} year domain mismatch for {tech}: "
                f"missing={sorted(expected_years - actual_years)}, "
                f"extra={sorted(actual_years - expected_years)}"
            )
        for year in AUTHORITY_YEARS:
            _numeric_decimal(authority[tech][year], f"{tech}/{year}")


def validate_minimum_contribution_authority(
    authority: Mapping[str, Mapping[int, object]],
) -> None:
    """Validate an already-loaded minimum/FUTURE contribution mapping."""
    _validate_mapping(
        authority,
        label="minimum contribution authority",
        techs=MINIMUM_CONTRIBUTION_TECHS,
        optional_techs=ROOT_ONLY_CONTRIBUTION_TECHS,
    )


def validate_minimum_boundary_authority(
    authority: Mapping[str, Mapping[int, object]],
) -> None:
    """Validate an already-loaded minimum-clamp boundary mapping."""
    _validate_mapping(
        authority,
        label="minimum boundary authority",
        techs=MINIMUM_BOUNDARY_TECHS,
    )


def _load_family(
    authority_path: Path,
    *,
    parameter: str,
    techs: frozenset[str],
    source: str | None,
    tech_metadata: Mapping[str, tuple[int, str]],
    label: str,
    scenario: str = BAU_SCENARIO,
    root_techs: frozenset[str] | None = None,
) -> Dict[str, Dict[int, float]]:
    """Load one dense family; ``source=None`` asks for any non-blank Source.

    ``root_techs`` is the domain of rows that may name a root scenario; when
    it is None every raw row must be a BAU row.
    """
    path = Path(authority_path)
    if not path.is_file():
        raise FileNotFoundError(f"{label} workbook not found: {path}")
    if scenario not in ROOT_SCENARIOS:
        raise ValueError(
            f"{label}: {scenario!r} is not a registered root scenario "
            f"{list(ROOT_SCENARIOS)}"
        )

    workbook = load_workbook(
        path, data_only=False, read_only=True, keep_links=False
    )
    try:
        if AUTHORITY_SHEET_NAME not in workbook.sheetnames:
            raise ValueError(
                f"authority workbook missing sheet {AUTHORITY_SHEET_NAME!r}"
            )
        worksheet = workbook[AUTHORITY_SHEET_NAME]
        columns, raw_shape = _header_columns(worksheet)

        # {row scenario: {tech: profile}}; a materialized sheet has one
        # already-merged scenario.
        by_scenario: Dict[str, Dict[str, Dict[int, float]]] = {}
        for row_index, values in enumerate(
            worksheet.iter_rows(min_row=2, values_only=True), start=2
        ):
            if values[columns["Parameter"]] != parameter:
                continue

            tech = values[columns["Tech"]]
            if not isinstance(tech, str):
                raise ValueError(
                    f"{label} row {row_index}: Tech must be non-blank text"
                )
            row_scenario = (
                values[columns["scenario"]] if raw_shape else scenario
            )
            if raw_shape and row_scenario != BAU_SCENARIO:
                if root_techs is None:
                    raise ValueError(
                        f"{label} {tech}: scenario must be {BAU_SCENARIO!r}"
                    )
                if row_scenario not in ROOT_SCENARIOS:
                    raise ValueError(
                        f"{label} {tech} row {row_index}: scenario "
                        f"{row_scenario!r} is not a registered root scenario "
                        f"{list(ROOT_SCENARIOS)}"
                    )
                if tech not in root_techs:
                    raise ValueError(
                        f"{label} {tech} row {row_index}: a "
                        f"{row_scenario!r} row is not allowed for {tech}"
                    )
            rows = by_scenario.setdefault(row_scenario, {})
            if tech in rows:
                raise ValueError(
                    f"{label} duplicate row for {tech} at row {row_index}"
                )
            if tech not in tech_metadata:
                raise ValueError(f"{label} unexpected technology {tech!r}")

            expected_id, expected_name = tech_metadata[tech]
            metadata_checks = {
                "Tech.ID": expected_id,
                "Tech.Name": expected_name,
                "Parameter.ID": 7,
                "Unit": None,
                "Projection.Mode": "User defined",
            }
            if source is not None:
                metadata_checks["Source"] = source
            for field, expected in metadata_checks.items():
                actual = values[columns[field]]
                if actual != expected:
                    raise ValueError(
                        f"{label} {tech}: {field} must be {expected!r}, "
                        f"got {actual!r}"
                    )
            row_source = values[columns["Source"]]
            if source is None and (
                not isinstance(row_source, str) or not row_source.strip()
            ):
                raise ValueError(
                    f"{label} {tech} row {row_index}: Source must name the "
                    "row's external source"
                )

            projection_parameter = _numeric_decimal(
                values[columns["Projection.Parameter"]],
                f"{label} {tech}/Projection.Parameter",
            )
            if projection_parameter != 0:
                raise ValueError(
                    f"{label} {tech}: Projection.Parameter must be 0"
                )

            profile: Dict[int, float] = {}
            for year in AUTHORITY_YEARS:
                number = _numeric_decimal(
                    values[columns[year]], f"{label} {tech}/{year}"
                )
                profile[year] = float(number)
            rows[tech] = profile
    finally:
        workbook.close()

    optional = frozenset(root_techs or ()) - techs
    if raw_shape:
        bau = by_scenario.get(BAU_SCENARIO, {})
        _validate_mapping(bau, label=label, techs=techs)
        authority = dict(bau)
        if scenario != BAU_SCENARIO:
            authority.update(by_scenario.get(scenario, {}))
    else:
        authority = by_scenario.get(scenario, {})
    _validate_mapping(
        authority, label=label, techs=techs, optional_techs=optional
    )
    return authority


def load_minimum_contribution_authority(
    authority_path: Path,
    scenario: str = BAU_SCENARIO,
) -> Dict[str, Dict[int, float]]:
    """Load the additive minimum (forced build) authority of a root scenario.

    The 11 BAU rows apply to every root; a row of ``scenario`` replaces the BAU
    row of its link or adds a cross-border link that has no BAU row.
    """
    return _load_family(
        authority_path,
        parameter=MINIMUM_CONTRIBUTION_PARAMETER,
        techs=MINIMUM_CONTRIBUTION_TECHS,
        source=None,
        tech_metadata=_CONTRIBUTION_TECH_METADATA,
        label="minimum contribution authority",
        scenario=scenario,
        root_techs=MINIMUM_CONTRIBUTION_TECHS | ROOT_ONLY_CONTRIBUTION_TECHS,
    )


def load_minimum_boundary_authority(
    authority_path: Path,
) -> Dict[str, Dict[int, float]]:
    """Load the exact two-row LinkFreeze minimum-clamp boundary authority."""
    return _load_family(
        authority_path,
        parameter=MINIMUM_BOUNDARY_PARAMETER,
        techs=MINIMUM_BOUNDARY_TECHS,
        source=MINIMUM_BOUNDARY_SOURCE,
        tech_metadata=_BOUNDARY_TECH_METADATA,
        label="minimum boundary authority",
    )
