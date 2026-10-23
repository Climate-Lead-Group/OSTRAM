"""
apply_internal_transmission.py
==============================

Late A3 stage stage_internal_transmission. Calibrates the INTERNAL (intra-node)
transmission families in the
model's `A-O_Parametrization.xlsx`, making `Config_country_codes.yaml` and the
workbooks' own 2023 inputs the source of truth for:

    * ResidualCapacity   -> Demand Techs            (year-indexed rows)
    * CapitalCost         -> Demand Techs            (year-indexed rows)
    * FixedCost           -> Demand Techs            (year-indexed rows)
    * OperationalLife     -> Fixed Horizon Parameters (single 'Value' cell)

for the six internal-transmission families (10 nodes each = 60 techs):

    RE     : RNWTRN, RNWNLI, RNWRPO   (renewable-carrying lines)
    non-RE : PWRTRN, TRNNLI, TRNRPO

Background / why a late A3 stage
--------------------------------
A2 (transmission.py) injects these families flat from the YAML (CapEx 100, FOM 4, ResCap 5,
life 20) into `Demand Techs` / `Fixed Horizon Parameters`. But:
  * ResidualCapacity is a flat 5 GW placeholder for every node (Maldives ~=
    India-West), which is not physical.
  * OperationalLife is rewritten to 50/20 by the Stage-1 template merge
    (update_ao_from_extensions.py), so a YAML/snapshot life edit does NOT
    stick.
  * There is no RE-vs-non-RE cost differentiation.

The rule has two parts. A3 runs the costs part (stage_internal_transmission)
after stage 5 and the interconnector-cost stage, and the residuals part
(stage_internal_tx_residuals) after the base-year pin, so the residuals read
the 2023 values the model compiles. It leaves the flat post-A2 snapshot
untouched and writes the delivered workbook.

Values applied
--------------
  * ResidualCapacity (existing stock, base year, held flat across the horizon),
    computed in every run for each node from the workbooks in --input-dir:
        peak  = max over time slices of 2023 SpecifiedAnnualDemand(ELC<node>03)
                x SpecifiedDemandProfile / (YearSplit x 8760 h), in GW
        RE    = sum of the node's 2023 renewable generator ResidualCapacity x
                CapacityFactor in the peak slice (see compute_node_residuals)
        RNWTRN<node> = residual_peak_margin x RE
        PWRTRN<node> = residual_peak_margin x peak - RNWTRN<node>
                (RNWTRN = margin x peak and PWRTRN = 0 when that is negative)
        RNWNLI/RNWRPO/TRNNLI/TRNRPO = 0          (new-build / repower mechanisms)
  * CapitalCost / FixedCost, the same in every node, computed in every run from
    the sourced inputs of the YAML block (internal_grid_costs):
        non-RE (PWRTRN/TRNNLI/TRNRPO) = base_capital_cost / base_fixed_cost
        RE     (RNWTRN/RNWNLI/RNWRPO) = base x re_capex_multiplier
    Each input and its source or ASSUMPTION is in the YAML block.
  * OperationalLife = operational_life for all 6 families.

Scope (intentionally narrow): the four params above, on the six internal
families only. Interconnectors (13-char TRN*****), DSPTRN, generators (PWR***),
storage, losses, and the Max*/Investment caps are all left untouched.

OUTPUT
------
1. Timestamped backup of the input directory (unless --skip-backup).
2. In-place edit of A-O_Parametrization.xlsx.
3. A JSON change log next to the backup.
4. A desk-check CSV (old -> new for every cell) when --desk-check-csv is given
   or on --dry-run (which writes the CSV and makes NO workbook edits).

USAGE
-----
    python apply_internal_transmission.py --input-dir A1_Outputs/A1_Outputs_BAU [--part costs|residuals]
    python apply_internal_transmission.py --input-dir <dir> --dry-run \\
        --desk-check-csv desk_check.csv          # preview only, no mutation
    python apply_internal_transmission.py --self-test
    python apply_internal_transmission.py --input-dir <dir> --restore
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

import yaml
from openpyxl import load_workbook, Workbook
from ostram.pipeline.scenarios.rules.config_keys import INTERNAL_TRANSMISSION_KEYS, check_keys

# --------------------------------------------------------------------------- config
PARAM_FILE_NAME = "A-O_Parametrization.xlsx"
DEFAULT_TARGET_SHEETS = ["Demand Techs", "Fixed Horizon Parameters"]

# Where ResidualCapacity / CapitalCost / FixedCost live (year-indexed) vs.
# OperationalLife (scalar). Both sheets are scanned; each param is written
# wherever its row is found, with the layout auto-detected per sheet.
YEAR_PARAMS = ("ResidualCapacity", "CapitalCost", "FixedCost")
SCALAR_PARAMS = ("OperationalLife",)
APPLIED_PARAMS = YEAR_PARAMS + SCALAR_PARAMS

RE_FAMILIES = ("RNWTRN", "RNWNLI", "RNWRPO")
NONRE_FAMILIES = ("PWRTRN", "TRNNLI", "TRNRPO")
ALL_FAMILIES = RE_FAMILIES + NONRE_FAMILIES
# Only the "Existing" families carry base-year residual stock; the rest are
# repower/new-line build mechanisms and start at 0.
RESIDUAL_FAMILY = {"RNWTRN": "RNWTRN", "PWRTRN": "PWRTRN"}

PROJ_MODE_COL = "Projection.Mode"
PROJ_MODE_EMPTY = "EMPTY"
PROJ_MODE_USER = "User defined"
VALUE_COL = "Value"

# Defaults mirror the YAML block; used only if the YAML omits a key.
DEFAULTS = {
    "operational_life": 40,
}

# The sourced inputs of the internal grid costs, all required. The costs are
# computed from them in every run (internal_grid_costs).
COST_INPUTS = (
    "esmap_capital_cost_usd2010_per_kw", "muv_2010", "muv_2022", "muv_2023",
    "fixed_cost_share", "gec2_cost_rs_crore", "gec2_capacity_mw", "inr_per_usd_2022",
)
RUPEES_PER_CRORE = 1e7
KW_PER_MW = 1e3

BACKUP_TAG = "_PRE_INTERNAL_TX_"
CONFIG_NAME = "Config_country_codes.yaml"
CONFIG_BLOCK = "internal_transmission"
MARGIN_KEY = "residual_peak_margin"

# The two parts of the rule. A3 runs "costs" before the base-year pin and
# "residuals" after it, so the residuals read the values the model compiles.
PARTS = ("costs", "residuals")
COST_PARAMS = ("CapitalCost", "FixedCost", "OperationalLife")
RESIDUAL_PARAMS = ("ResidualCapacity",)

# Inputs of the residual computation, all read from the delivered workbooks.
DEMAND_FILE_NAME = "A-O_Demand.xlsx"
BASE_YEAR = 2023
HOURS_PER_YEAR = 8760
GWH_PER_PJ = 1e6 / 3600
# B1 compiles no value from a row whose Projection.Mode is one of these.
NOT_COMPILED_MODES = ("EMPTY", "Zero", "", None, "According to demand")
RESIDUAL_SHEETS = ("Primary Techs", "Secondary Techs", "Demand Techs")
# ASSUMPTION: biomass, geothermal and waste plants count at capacity factor 1 at
# the peak, because they are dispatchable and their time-slice capacity factor is
# not a measure of their availability at the peak.
RE_FIRM_CODES = frozenset({"BIO", "GEO", "WAS"})
RE_VARIABLE_CODES = frozenset({"HYD", "SPV", "WON", "WOF", "CSP"})
# ASSUMPTION: small hydro (SHP) counts as non-renewable here, as in the method the
# residuals came from; its code is not in the renewable list of that method.
RE_CODES = RE_FIRM_CODES | RE_VARIABLE_CODES


def _is_internal_tx(tech) -> bool:
    """True for the 60 internal-transmission techs (family(6) + node(5) = 11).
    Excludes interconnectors (13-char), DSPTRN, generators, storage."""
    return isinstance(tech, str) and len(tech) == 11 and tech[:6] in ALL_FAMILIES


def _family_of(tech: str) -> str:
    return tech[:6]


def _node_of(tech: str) -> str:
    return tech[6:]


# --------------------------------------------------------------------------- sources
def resolve_config(explicit: Path | None) -> Path:
    """Locate Config_country_codes.yaml (knobs). Priority: --config, env, canonical."""
    candidates = []
    if explicit:
        candidates.append(Path(explicit))
    env = os.environ.get("OSTRAM_CONFIG_PATH")
    if env:
        candidates.append(Path(env))
    # Package rule module -> project preparation configuration.
    candidates.append(Path(__file__).resolve().parent.parent.parent / CONFIG_NAME)
    for c in candidates:
        if c and c.is_file():
            return c
    raise FileNotFoundError(
        "Could not locate Config_country_codes.yaml. Tried: "
        + ", ".join(str(c) for c in candidates)
    )


def has_config_block(config_path: Path) -> bool:
    """True when the configuration file has an ``internal_transmission`` block."""
    with open(config_path, "r", encoding="utf-8") as fh:
        data = yaml.safe_load(fh) or {}
    return bool(data.get(CONFIG_BLOCK))


def internal_grid_costs(blk: dict) -> dict:
    """Return the internal grid costs computed from the sourced inputs of ``blk``.

        base_capital_cost   = ESMAP capital cost x MUV(2023) / MUV(2010)        USD2023/kW
        base_fixed_cost     = fixed_cost_share x base_capital_cost              USD2023/kW/yr
        re_add_on           = GEC-II cost in rupees / GEC-II capacity in kW
                              / rupees per USD (2022) x MUV(2023) / MUV(2022)   USD2023/kW
        re_capex_multiplier = (base_capital_cost + re_add_on) / base_capital_cost

    Raises ValueError naming every missing input.
    """
    missing = [k for k in COST_INPUTS if blk.get(k) is None]
    if missing:
        raise ValueError(f"{CONFIG_BLOCK}: missing cost inputs: {', '.join(missing)}")
    v = {k: float(blk[k]) for k in COST_INPUTS}
    base = v["esmap_capital_cost_usd2010_per_kw"] * v["muv_2023"] / v["muv_2010"]
    add_on = (v["gec2_cost_rs_crore"] * RUPEES_PER_CRORE / (v["gec2_capacity_mw"] * KW_PER_MW)
              / v["inr_per_usd_2022"] * v["muv_2023"] / v["muv_2022"])
    return {
        "base_capital_cost": base,
        "base_fixed_cost": v["fixed_cost_share"] * base,
        "re_add_on": add_on,
        "re_capex_multiplier": (base + add_on) / base,
    }


def read_config(config_path: Path) -> dict:
    """Read the internal_transmission knob block and compute the costs from it.

    The operational life falls back to DEFAULTS. The cost inputs have no
    default: the rule stops when one is absent. The residual margin has no
    default: the residual part stops when it is absent.
    """
    with open(config_path, "r", encoding="utf-8") as fh:
        data = yaml.safe_load(fh) or {}
    blk = data.get(CONFIG_BLOCK, {}) or {}
    check_keys(blk, INTERNAL_TRANSMISSION_KEYS, config_path, CONFIG_BLOCK)
    knobs = dict(DEFAULTS)
    for k in DEFAULTS:
        if k in blk and blk[k] is not None:
            knobs[k] = blk[k]
    knobs.update(internal_grid_costs(blk))
    if blk.get(MARGIN_KEY) is not None:
        knobs[MARGIN_KEY] = float(blk[MARGIN_KEY])
    return knobs


# --------------------------------------------------------------------------- residuals
class MissingInputError(ValueError):
    """An input of the residual computation is missing for a node."""


def _sheet_rows(wb, sheet: str):
    """Return (header, data rows) of a workbook sheet."""
    if sheet not in wb.sheetnames:
        raise MissingInputError(f"sheet '{sheet}' not found")
    rows = list(wb[sheet].iter_rows(values_only=True))
    if not rows:
        raise MissingInputError(f"sheet '{sheet}' is empty")
    return list(rows[0]), rows[1:]


def _year_index(header, sheet: str) -> int:
    """Return the column index of the base year in a sheet header."""
    for i, h in enumerate(header):
        if h is not None and str(h).strip() in (str(BASE_YEAR), f"{BASE_YEAR}.0"):
            return i
    raise MissingInputError(f"sheet '{sheet}' has no {BASE_YEAR} column")


def _is_node_demand(fuel) -> bool:
    """True for a node's dispatch-ready demand fuel, ELC<node>03."""
    return isinstance(fuel, str) and len(fuel) == 10 and fuel.startswith("ELC") and fuel.endswith("03")


def _re_code(tech):
    """Return the renewable code of a PWR<code><node> generator, else None."""
    if not isinstance(tech, str) or not tech.startswith("PWR") or _is_internal_tx(tech):
        return None
    code = tech[3:-5]
    return code if code in RE_CODES else None


def _compiled(mode, where: str) -> bool:
    """True when B1 compiles the row's year values as they stand.

    A row with a mode in NOT_COMPILED_MODES compiles nothing. Any other mode than
    ``User defined`` makes B1 project the values, which this rule does not repeat.
    """
    if mode == PROJ_MODE_USER:
        return True
    if mode in NOT_COMPILED_MODES:
        return False
    raise MissingInputError(f"{where}: Projection.Mode {mode!r} is not read by this rule")


def read_residual_inputs(input_dir: Path) -> dict:
    """Read the base-year inputs of the residual computation that B1 compiles.

    From ``A-O_Demand.xlsx``: SpecifiedAnnualDemand (Demand_Projection) and
    SpecifiedDemandProfile (Profiles). From ``A-O_Parametrization.xlsx``: YearSplit
    (Yearsplit), CapacityFactor (Capacities), the ResidualCapacity of the PWR
    technologies and the internal transmission nodes.
    """
    input_dir = Path(input_dir)
    out = {"demand": {}, "profile": {}, "yearsplit": {}, "cf": {}, "residual": {}, "nodes": set()}
    dem = load_workbook(input_dir / DEMAND_FILE_NAME, read_only=True, data_only=True)
    par = load_workbook(input_dir / PARAM_FILE_NAME, read_only=True, data_only=True)
    try:
        hdr, rows = _sheet_rows(dem, "Demand_Projection")
        y, f, m = _year_index(hdr, "Demand_Projection"), hdr.index("Fuel/Tech"), hdr.index(PROJ_MODE_COL)
        for r in rows:
            if (_is_node_demand(r[f]) and _compiled(r[m], f"Demand_Projection {r[f]}")
                    and r[y] is not None):
                out["demand"][r[f]] = float(r[y])
        hdr, rows = _sheet_rows(dem, "Profiles")
        y, f, m = _year_index(hdr, "Profiles"), hdr.index("Fuel/Tech"), hdr.index(PROJ_MODE_COL)
        t = hdr.index("Timeslices")
        for r in rows:
            if (_is_node_demand(r[f]) and _compiled(r[m], f"Profiles {r[f]} {r[t]}")
                    and r[y] is not None):
                out["profile"][(r[f], r[t])] = float(r[y])
        hdr, rows = _sheet_rows(par, "Yearsplit")
        y, t, m = _year_index(hdr, "Yearsplit"), hdr.index("Timeslices"), hdr.index(PROJ_MODE_COL)
        for r in rows:
            if r[t] and _compiled(r[m], f"Yearsplit {r[t]}") and r[y] is not None:
                out["yearsplit"][r[t]] = float(r[y])
        hdr, rows = _sheet_rows(par, "Capacities")
        y, t, m = _year_index(hdr, "Capacities"), hdr.index("Timeslices"), hdr.index(PROJ_MODE_COL)
        te, p = hdr.index("Tech"), hdr.index("Parameter")
        for r in rows:
            if (r[p] == "CapacityFactor" and _re_code(r[te]) in RE_VARIABLE_CODES
                    and _compiled(r[m], f"Capacities {r[te]} {r[t]}") and r[y] is not None):
                out["cf"][(r[te], r[t])] = float(r[y])
        for sheet in RESIDUAL_SHEETS:
            if sheet not in par.sheetnames:
                continue
            hdr, rows = _sheet_rows(par, sheet)
            y, te, p, m = (_year_index(hdr, sheet), hdr.index("Tech"), hdr.index("Parameter"),
                           hdr.index(PROJ_MODE_COL))
            for r in rows:
                tech = r[te]
                if r[p] != "ResidualCapacity" or not isinstance(tech, str):
                    continue
                if _is_internal_tx(tech) and _family_of(tech) == "RNWTRN":
                    out["nodes"].add(_node_of(tech))
                if _re_code(tech) is None:
                    continue
                if _compiled(r[m], f"{sheet} {tech}") and r[y] is not None:
                    if tech in out["residual"]:
                        raise MissingInputError(f"{tech}: two ResidualCapacity rows")
                    out["residual"][tech] = float(r[y])
    finally:
        dem.close()
        par.close()
    out["nodes"] = sorted(out["nodes"])
    return out


def compute_node_residuals(inputs: dict, node: str, margin: float) -> dict:
    """Return the base-year residual capacity (GW) of RNWTRN<node> and PWRTRN<node>.

    peak = the largest, over time slices, of 2023 SpecifiedAnnualDemand of
    ELC<node>03 x SpecifiedDemandProfile share / (YearSplit share x 8760 h).
    RE at peak = sum over the node's renewable generators of 2023 ResidualCapacity
    x CapacityFactor in the peak slice (1 for biomass, geothermal and waste; the
    generator's mean 2023 capacity factor when it has none in the peak slice).
    RNWTRN = margin x RE at peak; PWRTRN = margin x peak - RNWTRN. When that is
    negative, RNWTRN = margin x peak and PWRTRN = 0.
    """
    fuel = f"ELC{node}03"
    if not inputs["yearsplit"]:
        raise MissingInputError(f"{node}: no compiled {BASE_YEAR} YearSplit")
    if fuel not in inputs["demand"]:
        raise MissingInputError(f"{node}: no compiled {BASE_YEAR} SpecifiedAnnualDemand of {fuel}")
    demand = inputs["demand"][fuel]
    powers = {}
    for ts, share in inputs["yearsplit"].items():
        if (fuel, ts) not in inputs["profile"]:
            raise MissingInputError(
                f"{node}: no compiled {BASE_YEAR} SpecifiedDemandProfile of {fuel} in {ts}")
        powers[ts] = demand * inputs["profile"][(fuel, ts)] * GWH_PER_PJ / (share * HOURS_PER_YEAR)
    peak_ts = max(powers, key=powers.get)
    peak = powers[peak_ts]
    re_at_peak = 0.0
    for tech, cap in sorted(inputs["residual"].items()):
        code = _re_code(tech)
        if tech[-5:] != node or code is None:
            continue
        if code in RE_FIRM_CODES:
            cf = 1.0
        elif (tech, peak_ts) in inputs["cf"]:
            cf = inputs["cf"][(tech, peak_ts)]
        else:
            # ASSUMPTION: a renewable generator with no capacity factor in the peak
            # slice counts at the mean of its 2023 capacity factors in the other
            # slices, or 1 when it has none, as in the method the residuals came from.
            own = [v for (t, _ts), v in inputs["cf"].items() if t == tech]
            cf = sum(own) / len(own) if own else 1.0
        re_at_peak += cap * cf
    rnwtrn = re_at_peak * margin
    pwrtrn = peak * margin - re_at_peak * margin
    if pwrtrn < 0:
        rnwtrn = peak * margin
        pwrtrn = 0.0
    return {"RNWTRN": rnwtrn, "PWRTRN": pwrtrn, "peak_GW": peak,
            "peak_timeslice": peak_ts, "RE_at_peak_GW": re_at_peak}


def compute_residuals(input_dir: Path, margin: float, nodes=None) -> dict:
    """Return {node: computed residuals} for every internal transmission node."""
    inputs = read_residual_inputs(input_dir)
    node_list = list(nodes) if nodes else inputs["nodes"]
    if not node_list:
        raise MissingInputError("no RNWTRN<node> ResidualCapacity row in the workbook")
    return {node: compute_node_residuals(inputs, node, margin) for node in node_list}


def build_target_values(knobs: dict, residuals: dict, nodes: list, parts=PARTS) -> dict:
    """Return {tech: {param: value}} for the 60 internal-tx techs and the selected parts."""
    out: dict = {}
    for fam in ALL_FAMILIES:
        for node in nodes:
            out[f"{fam}{node}"] = {}
    if "costs" in parts:
        base_cap = float(knobs["base_capital_cost"])
        base_fom = float(knobs["base_fixed_cost"])
        mult = float(knobs["re_capex_multiplier"])
        life = knobs["operational_life"]
        for fam in ALL_FAMILIES:
            is_re = fam in RE_FAMILIES
            for node in nodes:
                out[f"{fam}{node}"].update({
                    "CapitalCost": base_cap * (mult if is_re else 1.0),
                    "FixedCost": base_fom * (mult if is_re else 1.0),
                    "OperationalLife": life,
                })
    if "residuals" in parts:
        for fam in ALL_FAMILIES:
            for node in nodes:
                value = residuals[node][RESIDUAL_FAMILY[fam]] if fam in RESIDUAL_FAMILY else 0.0
                out[f"{fam}{node}"]["ResidualCapacity"] = float(value)
    return out


# --------------------------------------------------------------------------- worksheet helpers
def find_year_columns(ws) -> dict:
    """Return {year: column index} for the integer year headers in row 1."""
    out = {}
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col_idx).value
        if isinstance(val, int) and 1900 <= val <= 2200:
            out[val] = col_idx
    return out


def find_named_columns(ws, names) -> dict:
    """Return {header: column index} for the row-1 headers listed in names."""
    found = {}
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col_idx).value
        if val in names:
            found[val] = col_idx
    return found


def _values_differ(a, b, tol: float = 1e-9) -> bool:
    if a is None and b is None:
        return False
    if a is None or b is None:
        return True
    try:
        return abs(float(a) - float(b)) > tol
    except (TypeError, ValueError):
        return a != b


# --------------------------------------------------------------------------- core edit
def apply_values(ws, target_values: dict, seen: set, dry_run: bool) -> dict:
    """Overwrite internal-tx cells in one sheet. Layout auto-detected: year-indexed
    (Demand Techs) writes every year cell; scalar (Fixed Horizon Parameters) writes
    the single 'Value' cell. `seen` collects (tech, param) written anywhere."""
    hdr = find_named_columns(ws, ["Tech", "Parameter", PROJ_MODE_COL, VALUE_COL])
    tech_c, par_c = hdr.get("Tech"), hdr.get("Parameter")
    pm_c, val_c = hdr.get(PROJ_MODE_COL), hdr.get(VALUE_COL)
    if tech_c is None or par_c is None:
        return {"sheet": ws.title, "skipped": "no Tech/Parameter columns"}
    year_cols = find_year_columns(ws)
    if year_cols:
        layout = "year"
    elif val_c is not None:
        layout = "scalar"
    else:
        return {"sheet": ws.title, "skipped": "no year columns and no Value column"}

    log = {"sheet": ws.title, "layout": layout, "changes": [], "rows_touched": [],
           "pm_flips": []}
    years = sorted(year_cols)

    for row_idx in range(2, ws.max_row + 1):
        tech = ws.cell(row=row_idx, column=tech_c).value
        par = ws.cell(row=row_idx, column=par_c).value
        if not _is_internal_tx(tech):
            continue
        if par not in APPLIED_PARAMS or tech not in target_values or par not in target_values[tech]:
            continue
        seen.add((tech, par))
        new_val = target_values[tech][par]
        row_changed = False
        if layout == "year":
            for y in years:
                cell = ws.cell(row=row_idx, column=year_cols[y])
                if _values_differ(cell.value, new_val):
                    log["changes"].append({"tech": tech, "param": par, "year": y,
                                            "old": cell.value, "new": new_val})
                    if not dry_run:
                        cell.value = new_val
                    row_changed = True
        else:  # scalar
            cell = ws.cell(row=row_idx, column=val_c)
            if _values_differ(cell.value, new_val):
                log["changes"].append({"tech": tech, "param": par, "year": None,
                                        "old": cell.value, "new": new_val})
                if not dry_run:
                    cell.value = new_val
                row_changed = True
        if row_changed:
            log["rows_touched"].append({"tech": tech, "param": par, "new": new_val})
            if pm_c is not None and not dry_run:
                pm = ws.cell(row=row_idx, column=pm_c)
                if pm.value in (None, "", PROJ_MODE_EMPTY):
                    pm.value = PROJ_MODE_USER
                    log["pm_flips"].append({"tech": tech, "param": par})
    return log


def edit_parametrization(filepath: Path, sheets: list, target_values: dict,
                         dry_run: bool) -> dict:
    """Write the target values into each target sheet, save the workbook unless dry_run and return the change log with the unmatched (tech, param) pairs."""
    wb = load_workbook(filepath)
    file_log = {"file": str(filepath), "sheets": []}
    seen: set = set()
    try:
        for sheet in sheets:
            if sheet not in wb.sheetnames:
                file_log["sheets"].append({"sheet": sheet, "skipped": "not present"})
                continue
            file_log["sheets"].append(apply_values(wb[sheet], target_values, seen, dry_run))
        if not dry_run:
            wb.save(filepath)
    finally:
        wb.close()
    missing = [{"tech": t, "param": p}
               for t, params in target_values.items() for p in params
               if (t, p) not in seen]
    file_log["missing_pairs"] = missing
    return file_log


def write_desk_check_csv(file_log: dict, path: Path) -> None:
    """Write every cell change of a run log to a CSV with its old and new value."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sheet", "tech", "family", "node", "parameter", "year",
                    "old_value", "new_value"])
        for s in file_log.get("sheets", []):
            for c in s.get("changes", []):
                t = c["tech"]
                w.writerow([s["sheet"], t, _family_of(t), _node_of(t), c["param"],
                            c.get("year"), c.get("old"), c.get("new")])


# --------------------------------------------------------------------------- backup / restore
def make_backup(input_dir: Path) -> Path:
    """Copy the input folder to a timestamped `_PRE_INTERNAL_TX_` sibling and return its path."""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = input_dir.parent / f"{input_dir.name}{BACKUP_TAG}{stamp}"
    if backup.exists():
        raise FileExistsError(f"Backup already exists: {backup}")
    shutil.copytree(input_dir, backup)
    return backup


def find_latest_backup(input_dir: Path) -> Path | None:
    """Return the newest `_PRE_INTERNAL_TX_` backup beside the input folder, or None."""
    parent = input_dir.parent
    cands = sorted((p for p in parent.iterdir()
                    if p.is_dir() and p.name.startswith(f"{input_dir.name}{BACKUP_TAG}")),
                   key=lambda p: p.name)
    return cands[-1] if cands else None


def restore_from_backup(input_dir: Path, backup_dir: Path | None = None) -> Path:
    """Replace the input folder with the given backup, or the newest one, and return the backup used."""
    input_dir = Path(input_dir)
    backup_dir = Path(backup_dir) if backup_dir else find_latest_backup(input_dir)
    if not backup_dir or not backup_dir.is_dir():
        raise FileNotFoundError(f"No {BACKUP_TAG}* backup found next to {input_dir}.")
    if input_dir.is_dir():
        shutil.rmtree(input_dir)
    shutil.copytree(backup_dir, input_dir)
    return backup_dir


# --------------------------------------------------------------------------- orchestration
def run(input_dir, sheets=None, skip_backup=False, config=None, dry_run=False,
        desk_check_csv=None, nodes=None, parts=PARTS) -> dict:
    """Apply the selected parts to the internal transmission families and return the change log.

    ``costs``: CapitalCost, FixedCost and OperationalLife from the configuration.
    ``residuals``: the base-year ResidualCapacity computed from the workbooks in
    ``input_dir`` (see ``compute_node_residuals``).
    """
    input_dir = Path(input_dir)
    sheets = sheets or DEFAULT_TARGET_SHEETS
    parts = tuple(parts)
    unknown = [part for part in parts if part not in PARTS]
    if unknown or not parts:
        raise ValueError(f"parts must be taken from {PARTS}, got {parts}")
    config_path = resolve_config(config)
    knobs = read_config(config_path)
    computed = {}
    if "residuals" in parts:
        if MARGIN_KEY not in knobs:
            raise MissingInputError(f"{config_path}: {CONFIG_BLOCK}.{MARGIN_KEY} is not set")
        computed = compute_residuals(input_dir, knobs[MARGIN_KEY], nodes)
        node_list = list(computed)
    else:
        node_list = list(nodes) if nodes else read_residual_inputs(input_dir)["nodes"]
    target_values = build_target_values(knobs, computed, node_list, parts)

    paramfile = input_dir / PARAM_FILE_NAME
    if not paramfile.exists():
        raise FileNotFoundError(f"{paramfile} not found")

    backup_dir = None
    if not dry_run and not skip_backup:
        backup_dir = make_backup(input_dir)

    file_log = edit_parametrization(paramfile, sheets, target_values, dry_run)
    file_log.update({
        "config": str(config_path),
        "parts": list(parts),
        "knobs": knobs,
        "nodes": node_list,
        "residuals_computed": computed,
        "dry_run": dry_run,
        "backup_dir": str(backup_dir) if backup_dir else None,
        "timestamp": datetime.now().isoformat(),
        "families": {"RE": list(RE_FAMILIES), "non_RE": list(NONRE_FAMILIES)},
        "applied_params": [p for part in parts
                           for p in (COST_PARAMS if part == "costs" else RESIDUAL_PARAMS)],
    })
    if desk_check_csv:
        write_desk_check_csv(file_log, Path(desk_check_csv))
        file_log["desk_check_csv"] = str(desk_check_csv)
    if backup_dir is not None:
        log_path = backup_dir.parent / f"{backup_dir.name}_CHANGES.json"
        log_path.write_text(json.dumps(file_log, indent=2, default=str))
        file_log["log_path"] = str(log_path)
    return file_log


def print_summary(log: dict) -> None:
    """Print the inputs, computed residuals, per-sheet changes and unmatched pairs of a run log."""
    bar = "=" * 72
    print(bar)
    print("apply_internal_transmission: " + " + ".join(log.get("parts", []))
          + ("  [DRY RUN]" if log.get("dry_run") else ""))
    print(bar)
    print(f"Config (knobs)   : {log.get('config')}")
    print(f"Knobs            : {log.get('knobs')}")
    for node, r in (log.get("residuals_computed") or {}).items():
        print(f"Residuals {node}  : peak {r['peak_GW']!r} GW in {r['peak_timeslice']}, "
              f"RE at peak {r['RE_at_peak_GW']!r} GW -> RNWTRN {r['RNWTRN']!r}, PWRTRN {r['PWRTRN']!r}")
    print(f"Backup folder    : {log.get('backup_dir', '(skipped)')}")
    for s in log.get("sheets", []):
        if "skipped" in s:
            print(f"[SKIPPED] '{s['sheet']}': {s['skipped']}")
            continue
        print(f"Sheet '{s['sheet']}' ({s['layout']}): {len(s['changes'])} cells, "
              f"{len(s['rows_touched'])} rows, {len(s.get('pm_flips', []))} proj-mode flips")
    missing = log.get("missing_pairs", [])
    if missing:
        print(f"Sourced pairs with NO row in any target sheet ({len(missing)}): "
              + ", ".join(f"{m['tech']}/{m['param']}" for m in missing[:20])
              + (" ..." if len(missing) > 20 else ""))
    if log.get("desk_check_csv"):
        print(f"Desk-check CSV   : {log['desk_check_csv']}")
    if log.get("log_path"):
        print(f"Change log       : {log['log_path']}")


# --------------------------------------------------------------------------- self-test
def run_self_test() -> int:
    """Run the costs part twice on synthetic inputs, print each check and return 0 on success."""
    import tempfile
    bar = "=" * 72
    print(bar); print("apply_internal_transmission.py: SELF-TEST"); print(bar)
    ok = True
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        # synthetic config: base 100, fixed 4, add-on 100 (10 crore for 1,000 MW
        # at 1 rupee per dollar), so the RE multiplier is 2
        cfg = {
            "internal_transmission": {
                "esmap_capital_cost_usd2010_per_kw": 100, "muv_2010": 100, "muv_2022": 100,
                "muv_2023": 100, "fixed_cost_share": 0.04, "gec2_cost_rs_crore": 10,
                "gec2_capacity_mw": 1000, "inr_per_usd_2022": 1, "operational_life": 40,
            }
        }
        cfg_path = td / CONFIG_NAME
        cfg_path.write_text(yaml.safe_dump(cfg), encoding="utf-8")

        # synthetic A-O_Parametrization: Demand Techs (year) + FHP (scalar)
        idir = td / "A1_Outputs" / "A1_Outputs_BAU"; idir.mkdir(parents=True)
        pwb = Workbook()
        dt = pwb.active; dt.title = "Demand Techs"
        for c, h in enumerate(["Tech", "Parameter", PROJ_MODE_COL, 2023, 2024, 2025], 1):
            dt.cell(row=1, column=c, value=h)
        dt_rows = [
            ("RNWTRNINDWE", "ResidualCapacity", "User defined", 5, 5, 5),   # residual part only
            ("RNWTRNINDWE", "CapitalCost", "User defined", 100, 100, 100),  # RE -> 200
            ("RNWTRNINDWE", "FixedCost", "User defined", 4, 4, 4),          # RE -> 8
            ("PWRTRNMDVXX", "ResidualCapacity", "User defined", 5, 5, 5),   # residual part only
            ("PWRTRNMDVXX", "CapitalCost", "User defined", 100, 100, 100),  # non-RE -> 100 (unchanged)
            ("DSPTRNINDWE", "ResidualCapacity", "User defined", 9999, 9999, 9999),  # must NOT change
            ("PWRSPVINDWE", "CapitalCost", "User defined", 560, 560, 560),  # non-tx, must NOT change
        ]
        for ri, row in enumerate(dt_rows, 2):
            for c, v in enumerate(row, 1):
                dt.cell(row=ri, column=c, value=v)
        fhp = pwb.create_sheet("Fixed Horizon Parameters")
        for c, h in enumerate(["Tech.Type", "Tech.ID", "Tech", "Parameter", "Unit", VALUE_COL], 1):
            fhp.cell(row=1, column=c, value=h)
        fhp_rows = [
            ("Demand", 1, "RNWTRNINDWE", "OperationalLife", "years", 50),   # -> 40
            ("Demand", 2, "PWRTRNMDVXX", "OperationalLife", "years", 20),   # -> 40
            ("Demand", 3, "RNWTRNINDWE", "CapacityToActivityUnit", "", 31.536),  # must NOT change
            ("Demand", 4, "DSPTRNINDWE", "OperationalLife", "years", 20),   # must NOT change
        ]
        for ri, row in enumerate(fhp_rows, 2):
            for c, v in enumerate(row, 1):
                fhp.cell(row=ri, column=c, value=v)
        pwb.save(idir / PARAM_FILE_NAME); pwb.close()

        run(idir, skip_backup=True, config=cfg_path, nodes=["INDWE", "MDVXX"], parts=("costs",))

        chk = load_workbook(idir / PARAM_FILE_NAME, data_only=True)
        dt = chk["Demand Techs"]; got = {}
        for r in range(2, dt.max_row + 1):
            got[(dt.cell(row=r, column=1).value, dt.cell(row=r, column=2).value)] = \
                [dt.cell(row=r, column=c).value for c in (4, 5, 6)]
        fhp = chk["Fixed Horizon Parameters"]; got_f = {}
        for r in range(2, fhp.max_row + 1):
            got_f[(fhp.cell(row=r, column=3).value, fhp.cell(row=r, column=4).value)] = \
                fhp.cell(row=r, column=6).value
        chk.close()

        def ck(cond, label):
            """Print PASS or FAIL for one check and record a failure."""
            nonlocal ok
            print(("  PASS " if cond else "  FAIL ") + label)
            ok = ok and cond
        ck(got.get(("RNWTRNINDWE", "ResidualCapacity")) == [5, 5, 5], "costs part leaves RNWTRN residual")
        ck(got.get(("PWRTRNMDVXX", "ResidualCapacity")) == [5, 5, 5], "costs part leaves PWRTRN residual")
        ck(got.get(("RNWTRNINDWE", "CapitalCost")) == [200.0, 200.0, 200.0], "RE CapitalCost 100 -> 200 (2x)")
        ck(got.get(("RNWTRNINDWE", "FixedCost")) == [8.0, 8.0, 8.0], "RE FixedCost 4 -> 8 (2x)")
        ck(got.get(("PWRTRNMDVXX", "CapitalCost")) == [100.0, 100.0, 100.0], "non-RE CapitalCost stays 100")
        ck(got.get(("DSPTRNINDWE", "ResidualCapacity")) == [9999, 9999, 9999], "DSPTRN UNCHANGED")
        ck(got.get(("PWRSPVINDWE", "CapitalCost")) == [560, 560, 560], "non-tx PWRSPV UNCHANGED")
        ck(got_f.get(("RNWTRNINDWE", "OperationalLife")) == 40, "RE life 50 -> 40")
        ck(got_f.get(("PWRTRNMDVXX", "OperationalLife")) == 40, "non-RE life 20 -> 40")
        ck(got_f.get(("RNWTRNINDWE", "CapacityToActivityUnit")) == 31.536, "CapacityToActivityUnit UNCHANGED")
        ck(got_f.get(("DSPTRNINDWE", "OperationalLife")) == 20, "DSPTRN life UNCHANGED")

        # idempotency: a second run should change nothing
        log2 = run(idir, skip_backup=True, config=cfg_path, nodes=["INDWE", "MDVXX"], parts=("costs",))
        n2 = sum(len(s.get("changes", [])) for s in log2["sheets"])
        ck(n2 == 0, f"idempotent (2nd run: {n2} changes)")
    print(bar); print("SELF-TEST", "PASSED" if ok else "FAILED"); print(bar)
    return 0 if ok else 1


# --------------------------------------------------------------------------- main
def main() -> int:
    """Parse the command line, then run the rule, the self-test or a restore; return the exit code."""
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--input-dir", type=Path, default=Path("A1_Outputs/A1_Outputs_BAU"))
    p.add_argument("--config", type=Path, default=None, help="Config_country_codes.yaml")
    p.add_argument("--part", choices=PARTS + ("all",), default="all",
                   help="costs (CapitalCost, FixedCost, OperationalLife), residuals, or all")
    p.add_argument("--sheets", nargs="+", default=DEFAULT_TARGET_SHEETS)
    p.add_argument("--skip-backup", action="store_true")
    p.add_argument("--dry-run", action="store_true",
                   help="compute + write the desk-check CSV but do NOT edit the workbook")
    p.add_argument("--desk-check-csv", type=Path, default=None,
                   help="write an old->new CSV of every intended cell change")
    p.add_argument("--self-test", action="store_true")
    p.add_argument("--restore", action="store_true")
    p.add_argument("--restore-from", type=Path, default=None)
    args = p.parse_args()

    if args.self_test:
        return run_self_test()
    if args.restore or args.restore_from is not None:
        try:
            used = restore_from_backup(args.input_dir, args.restore_from)
        except Exception as exc:
            print(f"ERROR: {exc}", file=sys.stderr); return 1
        print(f"Restored {args.input_dir} from {used}"); return 0
    try:
        parts = PARTS if args.part == "all" else (args.part,)
        log = run(args.input_dir, args.sheets, args.skip_backup, args.config,
                  args.dry_run, args.desk_check_csv, parts=parts)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr); return 1
    print_summary(log)
    return 0


if __name__ == "__main__":
    sys.exit(main())
