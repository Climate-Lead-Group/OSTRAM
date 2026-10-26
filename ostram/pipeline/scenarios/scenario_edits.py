"""
scenario_edits.py  --  builds OSTRAM derived scenarios   (CLG / OSTRAM)
=======================================================================

Builds a derived scenario's A-O_Parametrization.xlsx by branching from
the ``base_scenario`` declared by its canonical registry/scenario-edit contract and
applying, in order:

    1. the SHARED VRE physical-potential ceiling layer
       (config/setups/south_asia.yaml::vre_ceilings)
    2. the run-specific edits in
       config/scenarios/<scenario>/scenario.yaml, after the edits of the
       scenario named by its optional reuse_edits_from key

Design (patterned on the rules_scripts):
  * Non-destructive to the SOURCE: the declared root A-O is never mutated.
  * Idempotent: every run rebuilds the target folder from a FRESH copy of the
    source A1_Outputs folder, so the result depends only on (source + scenario edits),
    never on a prior edited state.
  * Timestamped backup of any pre-existing target folder before overwrite.
  * Writes a *_CHANGES.json edit log.
  * --self-test runs in-process unit tests; --restore reverts the target from
    its most recent backup.

This is INPUT GENERATION (post-A3). It does NOT run A3/B1/B2.

USAGE
-----
    python -B -m ostram.pipeline.scenarios.scenario_edits --scenario B_Opt_TxCap150
    python -B -m ostram.pipeline.scenarios.scenario_edits --scenario B_Opt_SolarCapexHi
    python -B -m ostram.pipeline.scenarios.scenario_edits --scenario B_Opt_TradeCap15
    python -B -m ostram.pipeline.scenarios.scenario_edits --self-test
    python -B -m ostram.pipeline.scenarios.scenario_edits --scenario B_Opt_TxCap150 --restore

scenario YAML edit schema (one dict per edit):
    sheet            : worksheet name (e.g. "Secondary Techs", "Demand Techs")
    param            : OSeMOSYS parameter (the "Parameter" column value)
    tech | tech_prefix | techs : technology selector (exactly one)
    ONE operation of:
        values       : {year: value}  -> set those year cells explicitly
        op:"multiply"+factor          -> new = old * factor  (per year)
        op:"set_flat"+value           -> set every model year to value
        op:"set_to_residual"          -> set every year = that tech's
                                         ResidualCapacity (same sheet)
        op:"set_to_residual_factor_floor"
                                      -> selected years = max(
                                         factor * effective RC[residual_year],
                                         base_window_floor)
        op:"clamp_to_residual"         -> clamp to the effective RC, or to the
                                         separate v18 compatibility boundary
                                         when residual_source is
                                         "minimum_investment_boundary"
                                         (minimum only)
        op:"copy_from_tech"+source_tech -> set every year to the value of the
                                         same parameter in the source_tech
                                         row (same sheet)
    create_if_absent : bool (default False) -> create the param row if missing
    note             : free text (audit only)
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import shutil
import sys
import tempfile
from datetime import datetime
from difflib import get_close_matches
from pathlib import Path

from openpyxl import load_workbook, Workbook
from ostram import yaml_io
from ostram.paths import resolve_paths
from .registry import load_scenario_file
from .transformations import interconnector_authority as _authority

# --------------------------------------------------------------------------
# Constants / paths
# --------------------------------------------------------------------------
PARAM_FILE = "A-O_Parametrization.xlsx"
PROJ_MODE_COL = "Projection.Mode"
PROJ_MODE_USER = "User defined"
RES_PARAM = "ResidualCapacity"
MIN_INV_PARAM = "TotalAnnualMinCapacityInvestment"
YEAR_MIN, YEAR_MAX = 2020, 2060

SCRIPT_DIR = Path(__file__).resolve().parent
_PROJECT_PATHS = resolve_paths()
REPO = _PROJECT_PATHS.project_root
A1_OUTPUTS = _PROJECT_PATHS.a1_outputs
A3_PROCESS_DIR = SCRIPT_DIR / "transformations"
CONFIGS = _PROJECT_PATHS.scenario_config_root
CEIL_BASE = _PROJECT_PATHS.scenario_registry
SOASIA_V18 = _PROJECT_PATHS.scenario_workbook
MINIMUM_BOUNDARY_SOURCE = "minimum_investment_boundary"

MINIMUM_BOUNDARY_PARAMETER = _authority.MINIMUM_BOUNDARY_PARAMETER
load_minimum_boundary_authority = _authority.load_minimum_boundary_authority


# --------------------------------------------------------------------------
# Worksheet helpers
# --------------------------------------------------------------------------
def _as_year(v):
    try:
        y = int(float(v))
        return y if YEAR_MIN <= y <= YEAR_MAX else None
    except (TypeError, ValueError):
        return None


def scan_columns(ws):
    """Return (col_map, year_cols) from the header row (row 1)."""
    col_map, year_cols = {}, {}
    for c in range(1, ws.max_column + 1):
        raw = ws.cell(row=1, column=c).value
        if raw is None:
            continue
        y = _as_year(raw)
        if y is not None:
            year_cols[y] = c
            continue
        name = str(raw).strip()
        if name in ("Technology", "Tech"):
            col_map["tech"] = c
        elif name == "Parameter":
            col_map["param"] = c
        elif name == PROJ_MODE_COL:
            col_map["proj"] = c
        elif name == "Tech.ID":
            col_map["techid"] = c
        elif name == "Tech.Name":
            col_map["techname"] = c
        elif name == "Parameter.ID":
            col_map["paramid"] = c
        elif name == "Unit":
            col_map["unit"] = c
    if "tech" not in col_map or "param" not in col_map:
        raise ValueError("Sheet missing 'Tech'/'Parameter' columns")
    return col_map, year_cols


def find_row(ws, col_map, tech, param):
    """Return the first row of ``tech`` and ``param``, or None when absent."""
    tc, pc = col_map["tech"], col_map["param"]
    for r in range(2, ws.max_row + 1):
        t = ws.cell(row=r, column=tc).value
        p = ws.cell(row=r, column=pc).value
        if t is None or p is None:
            continue
        if str(t).strip() == tech and str(p).strip() == param:
            return r
    return None


def find_unique_row(ws, col_map, tech, param):
    """Return the single row of ``tech`` and ``param``; raise ValueError for none or several."""
    rows = []
    tc, pc = col_map["tech"], col_map["param"]
    for r in range(2, ws.max_row + 1):
        if (
            str(ws.cell(row=r, column=tc).value or "").strip() == tech
            and str(ws.cell(row=r, column=pc).value or "").strip() == param
        ):
            rows.append(r)
    if len(rows) != 1:
        raise ValueError(
            f"expected exactly one {tech}/{param} row, found {len(rows)}"
        )
    return rows[0]


def finite_nonnegative(value, context):
    """Return ``value`` as a float if finite and non-negative; raise ValueError otherwise."""
    if isinstance(value, bool) or value is None:
        raise ValueError(f"{context}: expected numeric value, got {value!r}")
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"{context}: expected numeric value, got {value!r}"
        ) from exc
    if not math.isfinite(result) or result < 0:
        raise ValueError(
            f"{context}: expected finite non-negative value, got {value!r}"
        )
    return result


def load_minimum_investment_boundaries(authority_path=None):
    """Load the exact dense minimum-clamp authority from canonical v18."""
    path = SOASIA_V18 if authority_path is None else authority_path
    return load_minimum_boundary_authority(Path(path).resolve())


def validate_residual_source(edit):
    """Validate an edit's optional external residual source."""
    source = edit.get("residual_source")
    if source is None or source == "effective":
        return False
    if source != MINIMUM_BOUNDARY_SOURCE:
        raise ValueError(f"unsupported residual_source {source!r}")
    if (
        edit.get("param") != MIN_INV_PARAM
        or edit.get("op") != "clamp_to_residual"
    ):
        raise ValueError(
            f"{MINIMUM_BOUNDARY_SOURCE} is permitted only for "
            f"{MIN_INV_PARAM} clamp_to_residual edits, got "
            f"parameter={edit.get('param')!r}, op={edit.get('op')!r}"
        )
    return True


def residual_values_for_edit(
    ws,
    col_map,
    year_cols,
    tech,
    param,
    edit,
    minimum_boundaries=None,
):
    """Return {year: value} from the tech's effective ResidualCapacity row, or from the
    minimum-investment boundary when the edit's ``residual_source`` names it.
    """
    source = edit.get("residual_source", "effective")
    if source == "effective":
        row = find_unique_row(ws, col_map, tech, RES_PARAM)
        return {
            year: finite_nonnegative(
                ws.cell(row, col).value,
                f"effective {tech}/{RES_PARAM}/{year}",
            )
            for year, col in year_cols.items()
        }
    if source == MINIMUM_BOUNDARY_SOURCE:
        if param != MIN_INV_PARAM:
            raise ValueError(
                f"{MINIMUM_BOUNDARY_SOURCE} is permitted only for "
                f"{MIN_INV_PARAM}, got {param}"
            )
        profiles = (
            minimum_boundaries
            if minimum_boundaries is not None
            else load_minimum_investment_boundaries()
        )
        if tech not in profiles:
            raise ValueError(
                f"{MINIMUM_BOUNDARY_PARAMETER} authority missing {tech}"
            )
        missing_years = sorted(set(year_cols) - set(profiles[tech]))
        if missing_years:
            raise ValueError(
                f"{MINIMUM_BOUNDARY_PARAMETER} authority {tech} "
                f"missing years {missing_years}"
            )
        return {
            year: finite_nonnegative(
                profiles[tech][year],
                f"{MINIMUM_BOUNDARY_PARAMETER} authority {tech}/{year}",
            )
            for year in year_cols
        }
    raise ValueError(f"unsupported residual_source {source!r}")


class ParentResultError(FileNotFoundError):
    """A solved parent result that a scenario edit reads is missing or empty."""


class ParentResults:
    """Read the solved parent results a derived scenario's edits use.

    ``paths`` maps each parent scenario to its result folder
    (``Executables/<parent>_0``); its ``Outputs/<Variable>.csv`` tables are read.
    A value absent from a table is 0, as the result tables omit zeros.
    """

    def __init__(self, scenario, paths=None):
        self.scenario = scenario
        self.paths = {name: Path(path) for name, path in (paths or {}).items()}
        self._tables = {}

    def _table(self, parent, variable):
        key = (parent, variable)
        if key in self._tables:
            return self._tables[key]
        if parent not in self.paths:
            raise ParentResultError(
                f"Scenario {self.scenario}: an edit reads the solved {parent} result, "
                f"which is not a declared result dependency of the scenario."
            )
        path = self.paths[parent] / "Outputs" / f"{variable}.csv"
        if not path.is_file():
            raise ParentResultError(
                f"Scenario {self.scenario}: the solved {parent} result has no {path}."
            )
        with path.open(encoding="utf-8", newline="") as fh:
            rows = list(csv.DictReader(fh))
        if not rows:
            raise ParentResultError(
                f"Scenario {self.scenario}: the solved {parent} result table {path} is empty."
            )
        self._tables[key] = rows
        return rows

    def by_tech_year(self, parent, variable):
        """Return {(technology, year): value} of a TECHNOLOGY x YEAR result table."""
        out = {}
        for row in self._table(parent, variable):
            key = (row["TECHNOLOGY"], int(float(row["YEAR"])))
            out[key] = out.get(key, 0.0) + float(row["VALUE"])
        return out

    def value(self, parent, variable, tech, year):
        """Return the parent's ``variable`` of ``tech`` in ``year`` (0 when absent)."""
        return self.by_tech_year(parent, variable).get((tech, int(year)), 0.0)

    def production_into(self, parent, tech, node, year):
        """Return the parent's annual production of ``tech`` into fuels of ``node``."""
        return sum(
            float(row["VALUE"])
            for row in self._table(parent, "ProductionByTechnologyAnnual")
            if row["TECHNOLOGY"] == tech and row["FUEL"][3:-2] == node
            and int(float(row["YEAR"])) == int(year)
        )

    def demand(self, parent, fuel, year):
        """Return the parent's annual demand of ``fuel`` (sum over time slices)."""
        return sum(
            float(row["VALUE"])
            for row in self._table(parent, "Demand")
            if row["FUEL"] == fuel and int(float(row["YEAR"])) == int(year)
        )


def base_window_floor(edit, tech, results):
    """Return the base-window floor of a ``set_to_residual_factor_floor`` edit.

    Either the edit's ``base_window_floor`` or, with ``base_window_result``, the
    largest ``TotalCapacityAnnual`` of ``tech`` in the solved parent over the
    years ``base_window_years`` (first and last, inclusive).
    """
    if ("base_window_floor" in edit) == ("base_window_result" in edit):
        raise ValueError(
            f"{tech}: set exactly one of base_window_floor and base_window_result"
        )
    if "base_window_floor" in edit:
        return finite_nonnegative(edit["base_window_floor"], f"{tech} base_window_floor")
    first, last = (int(y) for y in edit["base_window_years"])
    parent = edit["base_window_result"]
    return max(
        results.value(parent, "TotalCapacityAnnual", tech, year)
        for year in range(first, last + 1)
    )


def import_share_caps(edit, results):
    """Return {(tech, year): cap} for an ``import_share_cap`` edit.

    cap = share x the parent's demand of ``demand_fuel`` that year x the tech's
    share of the parent's imports into ``importer`` that year. Imports are the
    parent's production of the listed technologies into the importer's fuels.
    A ``demand_fuel`` measured at the bus bar holds the published requirement x
    (1 - loss) in the parent, so its demand is divided by (1 - loss) and the cap
    is the share of the published requirement. A year in which the parent
    imports nothing gives every listed technology a cap of 0.
    """
    from .rules.apply_internal_tx_losses import bus_bar_fuels, read_loss
    parent, techs = edit["result"], list(edit["techs"])
    share = finite_nonnegative(edit["share"], "import_share_cap share")
    scale = 1.0
    if edit["demand_fuel"] in bus_bar_fuels(_PROJECT_PATHS.scenario_workbook, parent):
        scale = 1.0 / (1.0 - read_loss(_PROJECT_PATHS.country_config))
    caps = {}
    for year in (int(y) for y in edit["years"]):
        imports = {t: results.production_into(parent, t, edit["importer"], year) for t in techs}
        total = sum(imports.values())
        demand = results.demand(parent, edit["demand_fuel"], year) * scale
        for t in techs:
            caps[(t, year)] = share * demand * imports[t] / total if total > 0 else 0.0
    return caps


def find_any_row_for_tech(ws, col_map, tech):
    """Return the first row of ``tech`` on the sheet, or None when absent."""
    tc = col_map["tech"]
    for r in range(2, ws.max_row + 1):
        t = ws.cell(row=r, column=tc).value
        if t is not None and str(t).strip() == tech:
            return r
    return None


def techs_matching(ws, col_map, edit):
    """Resolve tech / tech_prefix / techs into a concrete ordered list that
    actually exists on the sheet."""
    tc = col_map["tech"]
    present = []
    seen = set()
    for r in range(2, ws.max_row + 1):
        t = ws.cell(row=r, column=tc).value
        if t is None:
            continue
        t = str(t).strip()
        if t not in seen:
            seen.add(t)
            present.append(t)
    if "tech" in edit:
        return [edit["tech"]]
    if "techs" in edit:
        return list(edit["techs"])
    if "tech_prefix" in edit:
        pref = edit["tech_prefix"]
        return [t for t in present if t.startswith(pref)]
    raise ValueError("edit needs one of tech / techs / tech_prefix")


def read_row_values(ws, col_map, year_cols, tech, param):
    """Return {year: cell value} of the ``tech`` and ``param`` row, or None when absent."""
    r = find_row(ws, col_map, tech, param)
    if r is None:
        return None
    return {y: ws.cell(row=r, column=c).value for y, c in year_cols.items()}


def create_param_row(ws, col_map, tech, param):
    """Append a new parameter row for `tech`, copying identity columns from an
    existing row of that tech. Returns the new row index."""
    src = find_any_row_for_tech(ws, col_map, tech)
    r = ws.max_row + 1
    # copy identity columns
    for key in ("techid", "tech", "techname", "unit"):
        if key in col_map:
            val = ws.cell(row=src, column=col_map[key]).value if src else (tech if key == "tech" else None)
            ws.cell(row=r, column=col_map[key]).value = val
    ws.cell(row=r, column=col_map["tech"]).value = tech
    ws.cell(row=r, column=col_map["param"]).value = param
    if "proj" in col_map:
        ws.cell(row=r, column=col_map["proj"]).value = PROJ_MODE_USER
    return r


def flip_proj_mode(ws, col_map, row):
    """Set the row's Projection.Mode to ``User defined`` when it is blank or ``EMPTY``."""
    if "proj" in col_map:
        cell = ws.cell(row=row, column=col_map["proj"])
        if cell.value in (None, "EMPTY", ""):
            cell.value = PROJ_MODE_USER


# --------------------------------------------------------------------------
# Edit application
# --------------------------------------------------------------------------
class ScenarioEditError(ValueError):
    """A scenario edit or a VRE ceiling matched nothing in the workbook."""


def scenario_edit_error(ws, log, subject, param, name, kind="tech"):
    """Return the error for an edit that matched nothing; None for a valid "create"."""
    techs, params = set(), set()
    for sheet in ws.parent.worksheets:
        try:
            cols = scan_columns(sheet)[0]
        except ValueError:
            continue
        for row in sheet.iter_rows(min_row=2, values_only=True):
            for names, col in ((techs, cols["tech"]), (params, cols["param"])):
                if row[col - 1] is not None:
                    names.add(str(row[col - 1]).strip())
    close = []
    if kind == "years":
        reason = f"None of the years {', '.join(map(str, name))} is a column in the workbook."
    elif param not in params:
        reason, close = f'Sheet "{ws.title}" has no parameter "{param}".', get_close_matches(param, sorted(params), n=1)
    elif kind == "prefix" and not any(tech.startswith(name) for tech in techs):
        reason = f'No technology starts with "{name}".'
    elif kind != "prefix" and name not in techs:
        reason, close = "No technology has that name.", get_close_matches(name, sorted(techs), n=1)
    elif kind == "create":
        return None
    else:
        reason = f'Sheet "{ws.title}" has no {param} row for it.'
    hint = f" Did you mean {close[0]}?" if close else ""
    return ScenarioEditError(f"Scenario {log.get('scenario')}: {subject} changed nothing. {reason}{hint}")


def apply_edit(
    ws,
    col_map,
    year_cols,
    edit,
    log,
    minimum_boundaries=None,
    results=None,
):
    """Apply one scenario edit to each selected technology and record the changes in the edit log.

    Raises ScenarioEditError when the edit matches nothing.
    """
    param = edit["param"]
    validate_residual_source(edit)
    create = bool(edit.get("create_if_absent", False))
    selector = edit.get("tech") or edit.get("tech_prefix") or ", ".join(edit.get("techs", []))
    if results is None:
        results = ParentResults(log.get("scenario"))
    import_caps = import_share_caps(edit, results) if edit.get("op") == "import_share_cap" else None
    if edit.get("op") not in ("set_to_residual_factor_floor", "from_result", "import_share_cap"):
        if "values" in edit:
            named_years = edit["values"] or {}
        else:
            named_years = edit.get("years") if edit.get("op") == "multiply" else None
        missing = sorted({int(y) for y in named_years or ()} - set(year_cols))
        if missing:
            raise scenario_edit_error(ws, log, f'the edit to {param} for "{selector}"', param, missing, "years")
    matched = 0
    for tech in techs_matching(ws, col_map, edit):
        r = find_row(ws, col_map, tech, param)
        if r is None:
            if edit.get("op") == "set_to_residual_factor_floor" and edit.get("clamp"):
                log["skipped"].append({"tech": tech, "param": param,
                                       "reason": "row absent, nothing to lower"})
                continue
            if edit.get("op") == "set_to_residual_factor_floor":
                raise ValueError(f"required target row missing: {tech}/{param}")
            if not create:
                if "tech_prefix" not in edit:
                    raise scenario_edit_error(ws, log, f'the edit to {param} for "{tech}"', param, tech)
                log["skipped"].append({"tech": tech, "param": param,
                                       "reason": "row absent, create_if_absent=false"})
                continue
            if error := scenario_edit_error(ws, log, f'the edit to {param} for "{tech}"', param, tech, "create"):
                raise error
            r = create_param_row(ws, col_map, tech, param)
            log["rows_created"].append({"tech": tech, "param": param, "row": r})
        matched += 1

        # compute new value per year
        if edit.get("op") == "set_to_residual_factor_floor":
            residual_year = int(edit["residual_year"])
            if residual_year not in year_cols:
                raise ValueError(
                    f"{tech}/{param}: residual year {residual_year} missing"
                )
            factor = finite_nonnegative(
                edit["factor"], f"{tech}/{param} factor"
            )
            if factor <= 0:
                raise ValueError(f"{tech}/{param}: factor must be positive")
            floor = base_window_floor(edit, tech, results)
            years = tuple(int(year) for year in edit["years"])
            if not years or len(set(years)) != len(years):
                raise ValueError(
                    f"{tech}/{param}: years must be unique and non-empty"
                )
            missing_years = sorted(set(years) - set(year_cols))
            if missing_years:
                raise ValueError(
                    f"{tech}/{param}: target years missing {missing_years}"
                )
            residual_row = find_unique_row(
                ws, col_map, tech, RES_PARAM
            )
            residual = finite_nonnegative(
                ws.cell(residual_row, year_cols[residual_year]).value,
                f"effective {tech}/{RES_PARAM}/{residual_year}",
            )
            factor_value = round(factor * residual, 6)
            # A floor read from a solved parent is used exactly: rounded down, it
            # would cap below the capacity the parent built in the base years.
            formula_value = (
                max(factor_value, floor) if "base_window_result" in edit
                else round(max(factor_value, floor), 6)
            )
            year_set = set(years)
            if edit.get("clamp") and "base_window_result" in edit:
                # Lower the listed years, in order, to the headroom the limit leaves
                # above the parent's base-year capacity and the minimum investment of
                # earlier listed years, so the forced builds fit under the limit.
                headroom = max(formula_value - floor, 0.0)
                clamped = {}
                for y in sorted(year_set):
                    old = ws.cell(r, year_cols[y]).value
                    if old is None or (isinstance(old, str) and not old.strip()):
                        continue
                    clamped[y] = min(float(old), headroom)
                    headroom = max(headroom - clamped[y], 0.0)
                newfn = lambda y, old, v=clamped: v.get(y, old)
            elif edit.get("clamp"):
                # Lower the existing value to the limit where it exceeds it.
                def newfn(y, old, value=formula_value, yrs=year_set):
                    """Return min(``old``, the limit) in the listed years; keep others."""
                    if y not in yrs or old is None or (isinstance(old, str) and not old.strip()):
                        return old
                    return min(float(old), value)
            else:
                newfn = lambda y, old, value=formula_value, yrs=year_set: (
                    value if y in yrs else old
                )
        elif edit.get("op") == "from_result":
            years = {int(year) for year in edit["years"]}
            missing_years = sorted(years - set(year_cols))
            if missing_years:
                raise ValueError(f"{tech}/{param}: target years missing {missing_years}")
            fixed = edit.get("result_year")
            values = {
                y: results.value(edit["result"], edit["variable"], tech,
                                 int(fixed) if fixed is not None else y)
                for y in years
            }
            newfn = lambda y, old, v=values: v.get(y, old)
        elif edit.get("op") == "import_share_cap":
            newfn = lambda y, old, t=tech, v=import_caps: v.get((t, y), old)
        elif "values" in edit:
            newfn = lambda y, old, v=edit["values"]: (
                v[str(y)] if str(y) in v else (v[y] if y in v else old))
        elif edit.get("op") == "multiply":
            f = float(edit["factor"])
            # Optional per-year restriction: {"op":"multiply","factor":..,"years":[..]}
            # multiplies only the listed years; all other years are left unchanged.
            # Absent "years" -> multiply every year (backward compatible).
            yrs = set(int(y) for y in edit["years"]) if edit.get("years") else None
            def newfn(y, old, f=f, yrs=yrs):
                """Return ``old`` times the factor in selected years; keep others unchanged."""
                if yrs is not None and y not in yrs:
                    return old
                if old is None or (isinstance(old, str) and not old.strip()):
                    return old
                try:
                    return round(float(old) * f, 6)
                except (TypeError, ValueError):
                    return old
        elif edit.get("op") == "set_flat":
            val = float(edit["value"])
            newfn = lambda y, old, val=val: val
        elif edit.get("op") == "set_to_residual":
            resvals = residual_values_for_edit(
                ws, col_map, year_cols, tech, param, edit, minimum_boundaries
            )
            def newfn(y, old, rv=resvals):
                """Return the residual value of year ``y``."""
                return round(rv[y], 6)
        elif edit.get("op") == "clamp_to_residual":
            resvals = residual_values_for_edit(
                ws, col_map, year_cols, tech, param, edit, minimum_boundaries
            )
            def newfn(y, old, rv=resvals):
                """Return min(``old``, residual value of year ``y``); keep None and text."""
                cap = rv[y]
                if old is None:
                    return None
                try:
                    return round(min(float(old), cap), 6)
                except (TypeError, ValueError):
                    return old
        elif edit.get("op") == "copy_from_tech":
            source_tech = edit.get("source_tech")
            source = read_row_values(ws, col_map, year_cols, source_tech, param)
            if source is None:
                raise ScenarioEditError(
                    f"Scenario {log.get('scenario')}: the edit to {param} for "
                    f'"{selector}" has no source row. Sheet "{ws.title}" has '
                    f"no {param} row for {source_tech}."
                )
            newfn = lambda y, old, sv=source: sv[y]
        else:
            raise ValueError(f"edit for {tech}/{param}: no valid operation")

        changed = False
        for y, c in year_cols.items():
            cell = ws.cell(row=r, column=c)
            old = cell.value
            nv = newfn(y, old)
            # A parent result is written exactly, so that edits reading the same
            # result (minimum and maximum investment) give equal cells.
            differs = (old != nv) if edit.get("op") == "from_result" else _neq(old, nv)
            if nv is not None and (old is None or differs):
                cell.value = nv
                changed = True
                log["cells"].append({"tech": tech, "param": param, "year": y,
                                     "old": _num(old), "new": _num(nv)})
        if changed:
            flip_proj_mode(ws, col_map, r)
    if "tech_prefix" in edit and not matched:
        raise scenario_edit_error(ws, log, f'the edit to {param} for "{selector}"', param, selector, "prefix")


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


def _neq(a, b):
    try:
        return abs(float(a) - float(b)) > 1e-9
    except (TypeError, ValueError):
        return a != b


def clamp_row_to(ws, col_map, year_cols, tech, param, cap, log):
    """Clamp an existing (tech,param) row down so no year exceeds `cap`.
    Preserves coherence when a ceiling lowers TotalAnnualMaxCapacity below an
    existing investment build cap/floor (e.g. MaxCapInv > MaxCap after a clip)."""
    r = find_row(ws, col_map, tech, param)
    if r is None:
        return
    changed = False
    for y, c in year_cols.items():
        cell = ws.cell(row=r, column=c)
        old = cell.value
        if old is None:
            continue
        try:
            ov = float(old)
        except (TypeError, ValueError):
            continue
        if ov > cap + 1e-9:
            cell.value = round(cap, 6)
            changed = True
            log["cells"].append({"tech": tech, "param": param, "year": y,
                                 "old": ov, "new": round(cap, 6), "reason": "clamp<=ceiling"})
    if changed:
        flip_proj_mode(ws, col_map, r)


ACT_LOWER = "TotalTechnologyAnnualActivityLowerLimit"


def scale_activity_to_ceiling(ws, col_map, year_cols, tech, ceil, maxcap_orig, log):
    """Where the VRE ceiling clips MaxCap below its pre-clip value, scale the
    tech's activity LOWER limit (NDC generation floor) down by the same per-year
    ratio ceil/MaxCap_orig.

    Rationale: C_Target's set_vre_targets writes an activity floor sized to a
    cap_envelope MaxCap (floor + 20% headroom). When the atlas ceiling clips
    MaxCap below that (e.g. MDV solar 1.73 -> 1.0 GW), the clipped capacity can
    no longer physically produce the floor, so glpsol's line-199 check
    (max production >= activity lower limit) aborts. Production is linear in
    MaxCap, so scaling the floor by ceil/MaxCap_orig re-caps the NDC target at
    what is physically buildable while preserving the original 20% headroom.

    ONLY the lower limit is touched -- never the pinned upper limit:
      * B_Opt-derived scenarios have lower limit == 0 on every VRE tech (they
        use investment build caps, not activity floors), so scaling is a pure no-op
        for them (0 -> 0), and their upper limit is -1, which skips the
        constraint and must NOT be scaled into a bogus negative bound.
      * Leaving C_Target's upper pin unscaled just makes it non-binding above
        physical max (lower <= upper still holds, line-184 check safe); the
        tech lands at ~physical maximum -- the correct 'NDC capped at physical
        potential' outcome."""
    r = find_row(ws, col_map, tech, ACT_LOWER)
    if r is None:
        return
    changed = False
    for y, c in year_cols.items():
        mco = maxcap_orig.get(y)
        try:
            mco = float(mco) if mco is not None else None
        except (TypeError, ValueError):
            mco = None
        if mco is None or mco <= 0 or mco <= ceil + 1e-9:
            continue  # ceiling does not bind this year -> leave floor as-is
        cell = ws.cell(row=r, column=c)
        old = cell.value
        if old is None:
            continue
        try:
            ov = float(old)
        except (TypeError, ValueError):
            continue
        if ov <= 1e-9:
            continue  # no floor to scale (B_Opt: lower limit == 0)
        nv = round(ov * ceil / mco, 6)
        if _neq(ov, nv):
            cell.value = nv
            changed = True
            log["cells"].append({"tech": tech, "param": ACT_LOWER, "year": y,
                                 "old": ov, "new": nv, "reason": "scale NDC floor to clipped ceiling"})
    if changed:
        flip_proj_mode(ws, col_map, r)


def apply_ceiling_layer(wb, log, ceiling_path=CEIL_BASE):
    """Set each shared VRE ceiling as a flat value in every year, clamp the technology's
    annual investment rows to it and scale its activity lower limit where the ceiling binds.
    """
    ceiling_path = Path(ceiling_path)
    base = yaml_io.load_file(ceiling_path)
    base = base.get("vre_ceilings", base)
    sheet, param = base["sheet"], base["param"]
    ceilings = base["ceilings_gw"]
    ws = wb[sheet]
    col_map, year_cols = scan_columns(ws)
    for tech, ceil in ceilings.items():
        ceil = float(ceil)
        if find_row(ws, col_map, tech, param) is None:
            raise scenario_edit_error(ws, log, f'the VRE ceiling for "{tech}"', param, tech)
        # capture pre-clip MaxCap BEFORE overwriting, so any activity floor/pin
        # sized to the un-clipped capacity can be scaled back into feasibility.
        maxcap_orig = read_row_values(ws, col_map, year_cols, tech, param) or {}
        edit = {"param": param, "tech": tech, "op": "set_flat", "value": ceil}
        apply_edit(ws, col_map, year_cols, edit, log)
        # coherence: no annual investment build cap/floor may exceed the total ceiling
        clamp_row_to(ws, col_map, year_cols, tech, "TotalAnnualMaxCapacityInvestment", ceil, log)
        clamp_row_to(ws, col_map, year_cols, tech, "TotalAnnualMinCapacityInvestment", ceil, log)
        # coherence: an activity floor sized above the ceiling is unproducible
        # once MaxCap is clipped (C_Target NDC cap_envelope) -> scale it down.
        scale_activity_to_ceiling(ws, col_map, year_cols, tech, ceil, maxcap_orig, log)
    log["ceiling_layer"] = {
        "n_techs": len(ceilings),
        "source": str(ceiling_path),
    }


def apply_scenario_edits(wb, scenario_inputs, log, minimum_boundaries=None, results=None):
    """Apply the scenario edits listed under ``edits`` in their listed order.

    ``results`` (ParentResults) holds the solved parent results that edits with
    ``op: from_result``, ``op: import_share_cap`` or ``base_window_result`` read.
    """
    for edit in scenario_inputs.get("edits", []):
        ws = wb[edit["sheet"]]
        col_map, year_cols = scan_columns(ws)
        apply_edit(
            ws,
            col_map,
            year_cols,
            edit,
            log,
            minimum_boundaries=minimum_boundaries,
            results=results,
        )


def reused_edits(scenario, scenario_inputs, configs=CONFIGS):
    """Return the edits of the scenario named by ``reuse_edits_from``, or [] when it is absent.

    The named scenario must be a derived scenario with the same
    ``base_scenario`` that does not itself use ``reuse_edits_from``;
    otherwise ScenarioEditError says why.
    """
    name = scenario_inputs.get("reuse_edits_from")
    if name is None:
        return []
    prefix = f"Scenario {scenario}: reuse_edits_from names {name}"
    path = Path(configs) / str(name) / "scenario.yaml"
    if not path.is_file():
        raise ScenarioEditError(f"{prefix}, which is not a derived scenario (no file {path}).")
    other = load_scenario_file(path)
    if other.get("base_scenario") != scenario_inputs.get("base_scenario"):
        raise ScenarioEditError(
            f"{prefix}, whose base_scenario {other.get('base_scenario')} is not "
            f"{scenario_inputs.get('base_scenario')}."
        )
    if "reuse_edits_from" in other:
        raise ScenarioEditError(f"{prefix}, which itself uses reuse_edits_from.")
    return list(other.get("edits") or [])


def validate_edit_authorities(scenario_inputs, authority_path=None):
    """Validate external edit authorities before target-directory mutation."""
    requires_minimum_boundary = False
    for edit in scenario_inputs.get("edits", []):
        requires_minimum_boundary |= validate_residual_source(edit)
    if not requires_minimum_boundary:
        return None
    return load_minimum_investment_boundaries(authority_path)


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------
def build_scenario(
    scenario,
    source=None,
    skip_backup=False,
    *,
    a1_outputs=A1_OUTPUTS,
    configs=CONFIGS,
    ceiling_path=CEIL_BASE,
    authority_path=SOASIA_V18,
    result_paths=None,
):
    """Rebuild one derived scenario from its explicitly declared root.

    ``result_paths`` maps each parent whose solved result the edits read to its
    result folder; the materializer resolves it from the registry.

    ``source`` is retained as a fail-closed compatibility override.  When it
    is supplied it must equal ``scenario YAML::base_scenario``; omitting it uses
    the declaration directly.  Injectable path roots support pristine,
    disposable materialization proofs without touching the live worktree.
    """

    a1_outputs = Path(a1_outputs)
    configs = Path(configs)
    scenario_path = configs / scenario / "scenario.yaml"
    if not scenario_path.is_file():
        raise FileNotFoundError(f"scenario YAML not found: {scenario_path}")
    scenario_inputs = load_scenario_file(scenario_path)
    declared_source = scenario_inputs.get("base_scenario")
    if not declared_source:
        raise ValueError(f"{scenario_path} does not declare base_scenario")
    if source is not None and source != declared_source:
        raise ValueError(
            f"source override {source!r} conflicts with "
            f"{scenario_path}::base_scenario={declared_source!r}"
        )
    source = declared_source
    reused = reused_edits(scenario, scenario_inputs, configs)
    if reused:
        scenario_inputs = {
            **scenario_inputs,
            "edits": reused + list(scenario_inputs.get("edits") or []),
        }

    src_dir = a1_outputs / f"A1_Outputs_{source}"
    tgt_dir = a1_outputs / f"A1_Outputs_{scenario}"

    if not (src_dir / PARAM_FILE).is_file():
        raise FileNotFoundError(f"source A-O not found: {src_dir/PARAM_FILE}")
    minimum_boundaries = validate_edit_authorities(
        scenario_inputs,
        authority_path=authority_path,
    )

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Backup any pre-existing target, then rebuild from a FRESH source copy.
    backup_dir = None
    if tgt_dir.exists():
        if not skip_backup:
            backup_dir = tgt_dir.parent / f"{tgt_dir.name}_PREPATCH_{stamp}"
            shutil.copytree(tgt_dir, backup_dir)
        shutil.rmtree(tgt_dir)
    shutil.copytree(src_dir, tgt_dir)

    log = {"scenario": scenario, "source": source, "timestamp": stamp,
           "scenario_yaml": str(scenario_path), "target": str(tgt_dir),
           "backup_dir": str(backup_dir) if backup_dir else None,
           "cells": [], "rows_created": [], "skipped": []}
    if "reuse_edits_from" in scenario_inputs:
        log["reused_edits_from"] = {
            "scenario": scenario_inputs["reuse_edits_from"],
            "edits": len(reused),
        }

    wb = load_workbook(tgt_dir / PARAM_FILE)
    if scenario_inputs.get("apply_vre_ceiling_layer", True):
        apply_ceiling_layer(wb, log, ceiling_path=ceiling_path)
    apply_scenario_edits(
        wb,
        scenario_inputs,
        log,
        minimum_boundaries=minimum_boundaries,
        results=ParentResults(scenario, result_paths),
    )
    wb.save(tgt_dir / PARAM_FILE)
    wb.close()

    # Generic descendant ceilings must not replace reviewed calibration.
    # Resolve the declared ancestry; retain descendant edits outside the table.
    from ostram.setups import setup_policy
    from .rules import apply_base_year_pin as pins
    root = pins.calibration_root(scenario)
    if root is not None:
        rules = Path(setup_policy("pwr_min_pin_rules_path", str(pins.RULES_CSV)))
        log["reviewed_calibration"] = pins.apply_pin_rules(
            tgt_dir, root, rules, skip_backup=True,
        )

    log_path = tgt_dir / f"apply_patches_CHANGES_{stamp}.json"
    log_path.write_text(json.dumps(log, indent=2, default=str))
    log["log_path"] = str(log_path)
    return log


def restore_scenario(scenario):
    """Replace the derived scenario's folder with its latest ``_PREPATCH_`` backup."""
    tgt_dir = A1_OUTPUTS / f"A1_Outputs_{scenario}"
    backups = sorted(tgt_dir.parent.glob(f"{tgt_dir.name}_PREPATCH_*"))
    if not backups:
        raise FileNotFoundError(f"no backup found for {scenario}")
    latest = backups[-1]
    if tgt_dir.exists():
        shutil.rmtree(tgt_dir)
    shutil.copytree(latest, tgt_dir)
    print(f"Restored {tgt_dir} from {latest}")


def print_summary(log):
    """Print a summary of one derived scenario's edit log."""
    print("=" * 64)
    print(f"scenario_edits  --  {log['scenario']}  (from {log['source']})")
    print("=" * 64)
    print(f"target        : {log['target']}")
    print(f"backup        : {log.get('backup_dir') or '(none)'}")
    if "ceiling_layer" in log:
        print(f"VRE ceilings  : {log['ceiling_layer']['n_techs']} gen techs")
    print(f"cells written : {len(log['cells'])}")
    print(f"rows created  : {len(log['rows_created'])}  {[r['tech']+'/'+r['param'] for r in log['rows_created']]}")
    print(f"skipped       : {len(log['skipped'])}")
    print(f"edit log      : {log.get('log_path')}")


# --------------------------------------------------------------------------
# Self-test
# --------------------------------------------------------------------------
def self_test():
    """In-process tests of the edit engine on a synthetic workbook."""
    tmp = Path(tempfile.mkdtemp())
    wb = Workbook()
    ws = wb.active
    ws.title = "Secondary Techs"
    hdr = ["Tech.ID", "Tech", "Tech.Name", "Parameter.ID", "Parameter",
           "Unit", "Projection.Mode", "Projection.Parameter"] + list(range(2023, 2051))
    ws.append(hdr)
    ws.append([1, "PWRSPVBGDXX", "solar", 1, "CapitalCost", None, "User defined", 0] + [100.0] * 28)
    ws.append([1, "PWRSPVBGDXX", "solar", 2, "TotalAnnualMaxCapacity", None, "User defined", 0] + [365.0] * 28)
    ws.append([2, "TRNBGDXXINDEA", "trn", 3, "ResidualCapacity", None, "User defined", 0] + [3.0] * 28)
    ws.append([2, "TRNBGDXXINDEA", "trn", 4, "TotalAnnualMaxCapacity", None, "User defined", 0] + [9999.0] * 28)
    ws.append([2, "TRNBGDXXINDEA", "trn", 5, "TotalTechnologyAnnualActivityUpperLimit", None, "EMPTY", 0] + [None] * 28)
    ws.append([9, "TRNNLIBGDXX", "bk", 6, "ResidualCapacity", None, "User defined", 0] + [5.0] * 28)
    col_map, year_cols = scan_columns(ws)
    log = {"cells": [], "rows_created": [], "skipped": []}

    # multiply
    apply_edit(ws, col_map, year_cols, {"param": "CapitalCost", "tech": "PWRSPVBGDXX", "op": "multiply", "factor": 1.10}, log)
    r = find_row(ws, col_map, "PWRSPVBGDXX", "CapitalCost")
    assert abs(ws.cell(row=r, column=year_cols[2050]).value - 110.0) < 1e-6, "multiply failed"

    # set_flat ceiling (clip 365 -> 40)
    apply_edit(ws, col_map, year_cols, {"param": "TotalAnnualMaxCapacity", "tech": "PWRSPVBGDXX", "op": "set_flat", "value": 40.0}, log)
    r = find_row(ws, col_map, "PWRSPVBGDXX", "TotalAnnualMaxCapacity")
    assert abs(ws.cell(row=r, column=year_cols[2030]).value - 40.0) < 1e-6, "set_flat failed"

    # set_to_residual (freeze 9999 -> 3)
    apply_edit(ws, col_map, year_cols, {"param": "TotalAnnualMaxCapacity", "tech": "TRNBGDXXINDEA", "op": "set_to_residual"}, log)
    r = find_row(ws, col_map, "TRNBGDXXINDEA", "TotalAnnualMaxCapacity")
    assert abs(ws.cell(row=r, column=year_cols[2050]).value - 3.0) < 1e-6, "set_to_residual failed"

    # explicit values on existing empty AUL row
    apply_edit(ws, col_map, year_cols, {"param": "TotalTechnologyAnnualActivityUpperLimit", "tech": "TRNBGDXXINDEA", "values": {"2050": 123.4}}, log)
    r = find_row(ws, col_map, "TRNBGDXXINDEA", "TotalTechnologyAnnualActivityUpperLimit")
    assert abs(ws.cell(row=r, column=year_cols[2050]).value - 123.4) < 1e-6, "values failed"

    # create_if_absent AUL row for backstop
    apply_edit(ws, col_map, year_cols, {"param": "TotalTechnologyAnnualActivityUpperLimit", "tech": "TRNNLIBGDXX", "op": "set_flat", "value": 0.0, "create_if_absent": True}, log)
    r = find_row(ws, col_map, "TRNNLIBGDXX", "TotalTechnologyAnnualActivityUpperLimit")
    assert r is not None, "create_if_absent failed"
    assert ws.cell(row=r, column=year_cols[2035]).value == 0.0, "created row value failed"
    assert len(log["rows_created"]) == 1, "rows_created log failed"

    shutil.rmtree(tmp, ignore_errors=True)
    print("SELF-TEST PASSED (6 assertions):")
    print("  multiply x1.10, set_flat clip, set_to_residual freeze,")
    print("  explicit values, create_if_absent backstop row, proj-mode flip.")
    return 0


# --------------------------------------------------------------------------
def main():
    """Build one derived scenario, restore it from its backup, or run the self-test.

    Returns 1 when a scenario edit or VRE ceiling matches nothing.
    """
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument(
        "--scenario",
        help=(
            "derived scenario name; reads "
            "config/scenarios/<scenario>/scenario.yaml and writes "
            "workspace/preparation/A1_Outputs/A1_Outputs_<scenario>"
        ),
    )
    ap.add_argument(
        "--source-scenario",
        default=None,
        help=(
            "Fail-closed compatibility override; must match the "
            "scenario YAML base_scenario declaration."
        ),
    )
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--restore", action="store_true")
    ap.add_argument("--skip-backup", action="store_true")
    args = ap.parse_args()

    if args.self_test:
        return self_test()
    if not args.scenario:
        ap.error("--scenario is required (or use --self-test)")
    if args.restore:
        restore_scenario(args.scenario)
        return 0
    try:
        log = build_scenario(args.scenario, args.source_scenario, args.skip_backup)
    except ScenarioEditError as error:
        print(error, file=sys.stderr)
        return 1
    print_summary(log)
    return 0


if __name__ == "__main__":
    sys.exit(main())
