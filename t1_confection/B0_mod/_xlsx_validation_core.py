"""
Shared helpers for validating A-O_Parametrization.xlsx upper-bound invariants.

Used by:
    - Z_AUX_apply_parametrization_review.py (consistency sweep after review writes)
    - B1b_Pre_solver_validation.py            (pre-solver validation in B1)

Constants, indexing, and the cumulative-capacity sweep originated in
Z_AUX_apply_parametrization_review.py. The activity-vs-capacity helpers
(load_yearsplit, load_capacity_factor, load_capacity_to_activity_unit,
get_availability_factor, compute_*, validate_activity_lower_limit) are
new and support the CAb1_PlannedMaintenance vs AAC3 feasibility check.
"""
import math

# --- Parameter names in A-O_Parametrization.xlsx ---
PARAM_NAME = "TotalAnnualMinCapacityInvestment"
MAX_PARAM_NAME = "TotalAnnualMaxCapacityInvestment"
RESIDUAL_PARAM = "ResidualCapacity"
MAX_TOTAL_PARAM = "TotalAnnualMaxCapacity"
AF_PARAM = "AvailabilityFactor"
ACT_LOWER_PARAM = "TotalTechnologyAnnualActivityLowerLimit"

# --- Sheets ---
OPLIFE_SHEET = "Fixed Horizon Parameters"
OPLIFE_PARAM = "OperationalLife"
C2A_PARAM = "CapacityToActivityUnit"
YEARSPLIT_SHEET = "Yearsplit"
CAPACITIES_SHEET = "Capacities"
CAP_FACTOR_PARAM = "CapacityFactor"

# --- Tunables ---
DEFAULT_OPLIFE = 30
MAX_MULTIPLIER = 1.01           # Max = bound * 101%
ACT_LOWER_HAIRCUT = 0.99        # ActivityLowerLimit = max_activity * 99%


def norm_str(v):
    return str(v).strip() if v is not None else ""


def is_empty_number(v):
    if v is None or v == "":
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    return False


# -------------------------------------------------------------------
# Indexing of Secondary/Demand Techs sheets
# -------------------------------------------------------------------
def index_target_sheet(wb, sheet_name, extra_params=()):
    """Return (ws, cols, year_col_map, tech_to_row_by_param, dup_warnings).

    tech_to_row_by_param maps {param_name: {tech_upper: row}} for the four
    capacity parameters plus any names listed in extra_params (e.g.
    AvailabilityFactor, TotalTechnologyAnnualActivityLowerLimit).
    """
    ws = wb[sheet_name]
    header = {norm_str(c.value): i for i, c in enumerate(ws[1], 1) if c.value is not None}
    for required in ("Tech", "Parameter", "Projection.Mode"):
        if required not in header:
            raise RuntimeError(f"Column '{required}' missing in sheet '{sheet_name}'")
    cols = {
        "tech": header["Tech"],
        "parameter": header["Parameter"],
        "proj_mode": header["Projection.Mode"],
    }
    year_col_map = {}
    for name, col in header.items():
        if name.isdigit():
            year_col_map[int(name)] = col

    wanted = {
        PARAM_NAME, MAX_PARAM_NAME, RESIDUAL_PARAM, MAX_TOTAL_PARAM,
        *extra_params,
    }
    tech_to_row_by_param = {p: {} for p in wanted}
    dup_warnings = []
    for r in range(2, ws.max_row + 1):
        tech = ws.cell(r, cols["tech"]).value
        param = ws.cell(r, cols["parameter"]).value
        if not tech or not param:
            continue
        param_str = norm_str(param)
        if param_str not in tech_to_row_by_param:
            continue
        key = norm_str(tech).upper()
        bucket = tech_to_row_by_param[param_str]
        if key in bucket:
            dup_warnings.append((param_str, key, bucket[key], r))
        else:
            bucket[key] = r
    return ws, cols, year_col_map, tech_to_row_by_param, dup_warnings


# -------------------------------------------------------------------
# Workbook-wide loaders (Fixed Horizon, Yearsplit, Capacities)
# -------------------------------------------------------------------
def load_operational_life(wb):
    """Return {tech_upper: op_life_years} from 'Fixed Horizon Parameters'."""
    if OPLIFE_SHEET not in wb.sheetnames:
        return {}
    ws = wb[OPLIFE_SHEET]
    header = {norm_str(c.value): i for i, c in enumerate(ws[1], 1) if c.value is not None}
    tcol = header.get("Tech")
    pcol = header.get("Parameter")
    vcol = header.get("Value")
    if not (tcol and pcol and vcol):
        return {}
    out = {}
    for r in range(2, ws.max_row + 1):
        if norm_str(ws.cell(r, pcol).value) != OPLIFE_PARAM:
            continue
        tech = norm_str(ws.cell(r, tcol).value).upper()
        v = ws.cell(r, vcol).value
        if not tech or v is None:
            continue
        try:
            out[tech] = int(round(float(v)))
        except (TypeError, ValueError):
            pass
    return out


def load_capacity_to_activity_unit(wb):
    """Return {tech_upper: c2a_value} from 'Fixed Horizon Parameters'."""
    if OPLIFE_SHEET not in wb.sheetnames:
        return {}
    ws = wb[OPLIFE_SHEET]
    header = {norm_str(c.value): i for i, c in enumerate(ws[1], 1) if c.value is not None}
    tcol = header.get("Tech")
    pcol = header.get("Parameter")
    vcol = header.get("Value")
    if not (tcol and pcol and vcol):
        return {}
    out = {}
    for r in range(2, ws.max_row + 1):
        if norm_str(ws.cell(r, pcol).value) != C2A_PARAM:
            continue
        tech = norm_str(ws.cell(r, tcol).value).upper()
        v = ws.cell(r, vcol).value
        if not tech or v is None:
            continue
        try:
            out[tech] = float(v)
        except (TypeError, ValueError):
            pass
    return out


def load_yearsplit(wb):
    """Return {year: {timeslice: fraction}} from sheet 'Yearsplit'.

    The header layout (per repo convention): col 1 = Timeslices, then
    Parameter.ID/Parameter/Unit/Projection.Mode/Projection.Parameter,
    then year columns.
    """
    if YEARSPLIT_SHEET not in wb.sheetnames:
        return {}
    ws = wb[YEARSPLIT_SHEET]
    header = {}
    year_cols = {}
    for i, c in enumerate(ws[1], 1):
        if c.value is None:
            continue
        v = c.value
        if isinstance(v, (int, float)) and float(v).is_integer():
            year_cols[int(v)] = i
        else:
            try:
                # Some files store year headers as strings ('2023')
                year_cols[int(str(v).strip())] = i
            except (TypeError, ValueError):
                header[norm_str(v)] = i
    out = {y: {} for y in year_cols}
    for r in range(2, ws.max_row + 1):
        ts = ws.cell(r, 1).value
        if not ts:
            continue
        ts_str = norm_str(ts)
        for y, c in year_cols.items():
            v = ws.cell(r, c).value
            if is_empty_number(v):
                continue
            try:
                out[y][ts_str] = float(v)
            except (TypeError, ValueError):
                pass
    return out


def load_capacity_factor(wb, tech_upper):
    """Return {year: {timeslice: cf}} for one tech from 'Capacities' sheet."""
    if CAPACITIES_SHEET not in wb.sheetnames:
        return {}
    ws = wb[CAPACITIES_SHEET]
    header = {}
    year_cols = {}
    for i, c in enumerate(ws[1], 1):
        if c.value is None:
            continue
        v = c.value
        if isinstance(v, (int, float)) and float(v).is_integer():
            year_cols[int(v)] = i
        else:
            try:
                year_cols[int(str(v).strip())] = i
            except (TypeError, ValueError):
                header[norm_str(v)] = i
    tcol = header.get("Tech")
    pcol = header.get("Parameter")
    tscol = header.get("Timeslices") or 1
    if not (tcol and pcol):
        return {}
    out = {}
    for r in range(2, ws.max_row + 1):
        if norm_str(ws.cell(r, pcol).value) != CAP_FACTOR_PARAM:
            continue
        if norm_str(ws.cell(r, tcol).value).upper() != tech_upper:
            continue
        ts = norm_str(ws.cell(r, tscol).value)
        if not ts:
            continue
        for y, c in year_cols.items():
            v = ws.cell(r, c).value
            if is_empty_number(v):
                continue
            try:
                out.setdefault(y, {})[ts] = float(v)
            except (TypeError, ValueError):
                pass
    return out


def get_availability_factor(ws, tech_to_row_by_param, tech_upper, year, ymap):
    """AvailabilityFactor(tech, year) from Secondary/Demand Techs."""
    af_rows = tech_to_row_by_param.get(AF_PARAM, {})
    af_row = af_rows.get(tech_upper)
    if af_row is None:
        return None
    col = ymap.get(year)
    if col is None:
        return None
    v = ws.cell(af_row, col).value
    if is_empty_number(v):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# -------------------------------------------------------------------
# Math helpers
# -------------------------------------------------------------------
def compute_annual_cap_factor(yearsplit_year, capfactor_year):
    """Σ_l CapacityFactor(l) × YearSplit(l) for one year.

    yearsplit_year: {timeslice: ys}
    capfactor_year: {timeslice: cf}
    Missing timeslices contribute 0 (treated as no capacity factor data).
    """
    if not yearsplit_year or not capfactor_year:
        return 0.0
    total = 0.0
    for ts, ys in yearsplit_year.items():
        cf = capfactor_year.get(ts)
        if cf is None:
            continue
        total += cf * ys
    return total


def compute_max_capacity(ws, tech_to_row_by_param, tech_upper, year,
                         op_life, ymap, sorted_years=None):
    """Residual(y) + Σ MaxInv(y') over y' ∈ [y - op_life + 1, y].

    Returns float or None if MaxInv has at least one empty cell within the
    window (interpreted as 'unbounded' → activity check is vacuous).
    """
    res_rows = tech_to_row_by_param.get(RESIDUAL_PARAM, {})
    max_rows = tech_to_row_by_param.get(MAX_PARAM_NAME, {})
    res_row = res_rows.get(tech_upper)
    max_row = max_rows.get(tech_upper)

    residual = 0.0
    if res_row is not None:
        col = ymap.get(year)
        if col is not None:
            v = ws.cell(res_row, col).value
            if not is_empty_number(v):
                try:
                    residual = float(v)
                except (TypeError, ValueError):
                    residual = 0.0

    if max_row is None:
        return residual  # no investment possible → only residual

    if sorted_years is None:
        sorted_years = sorted(ymap.keys())
    horizon = (max(sorted_years) - min(sorted_years) + 1) if sorted_years else op_life
    window_start = year - min(op_life, horizon) + 1

    cumulative = 0.0
    for y2 in sorted_years:
        if y2 > year:
            break
        if y2 < window_start:
            continue
        col2 = ymap.get(y2)
        if col2 is None:
            continue
        v2 = ws.cell(max_row, col2).value
        if is_empty_number(v2):
            return None  # unbounded → cannot bound max_capacity
        try:
            cumulative += float(v2)
        except (TypeError, ValueError):
            return None
    return residual + cumulative


def compute_max_activity(max_cap, af, c2a, annual_cf):
    """max_capacity × AvailabilityFactor × CapacityToActivityUnit × Σ(CF·YS)."""
    if max_cap is None or af is None or c2a is None or annual_cf is None:
        return None
    return max_cap * af * c2a * annual_cf


# -------------------------------------------------------------------
# Sweep 1: max-cap consistency (V1 + V2)
# -------------------------------------------------------------------
def consistency_sweep(ws, sheet_name, cols, ymap, tech_to_row_by_param,
                      oplife_map, apply_changes):
    """V1+V2 sweep. Returns (max_adjusts, max_total_adjusts)."""
    min_rows = tech_to_row_by_param[PARAM_NAME]
    max_rows = tech_to_row_by_param[MAX_PARAM_NAME]
    res_rows = tech_to_row_by_param[RESIDUAL_PARAM]
    max_tot_rows = tech_to_row_by_param[MAX_TOTAL_PARAM]
    max_adjusts = []
    max_total_adjusts = []
    sorted_years = sorted(ymap.keys())
    horizon = (max(sorted_years) - min(sorted_years) + 1) if sorted_years else DEFAULT_OPLIFE

    for tech, min_row in min_rows.items():
        max_row = max_rows.get(tech)
        max_tot_row = max_tot_rows.get(tech)
        res_row = res_rows.get(tech)
        op_life = oplife_map.get(tech, DEFAULT_OPLIFE)
        effective_window = min(op_life, horizon)

        min_by_year = {}
        for y, c in ymap.items():
            v = ws.cell(min_row, c).value
            if is_empty_number(v):
                continue
            try:
                min_by_year[y] = float(v)
            except (TypeError, ValueError):
                pass

        for y in sorted_years:
            col = ymap[y]
            min_y = min_by_year.get(y)

            # V1: per-year
            if max_row is not None and min_y is not None and min_y > 0:
                old_max = ws.cell(max_row, col).value
                if not is_empty_number(old_max):
                    try:
                        old_max_num = float(old_max)
                    except (TypeError, ValueError):
                        old_max_num = None
                    if old_max_num is not None and min_y >= old_max_num:
                        new_max = min_y * MAX_MULTIPLIER
                        old_mode = ws.cell(max_row, cols["proj_mode"]).value
                        if apply_changes:
                            ws.cell(max_row, col).value = new_max
                            ws.cell(max_row, cols["proj_mode"]).value = "User defined"
                        max_adjusts.append({
                            "target_sheet": sheet_name,
                            "tech": tech, "year": y,
                            "row": max_row, "col": col,
                            "min": min_y, "old_max": old_max_num,
                            "new_max": new_max, "old_mode": old_mode,
                        })

            # V2: cumulative
            if max_tot_row is None:
                continue
            old_max_tot = ws.cell(max_tot_row, col).value
            if is_empty_number(old_max_tot):
                continue
            try:
                old_max_tot_num = float(old_max_tot)
            except (TypeError, ValueError):
                continue
            window_start = y - effective_window + 1
            accumulated = sum(v for y2, v in min_by_year.items() if window_start <= y2 <= y)
            if accumulated <= 0:
                continue
            residual = 0.0
            if res_row is not None:
                old_res = ws.cell(res_row, col).value
                if not is_empty_number(old_res):
                    try:
                        residual = float(old_res)
                    except (TypeError, ValueError):
                        residual = 0.0
            threshold = residual + accumulated
            if old_max_tot_num > threshold:
                continue
            new_max_tot = threshold * MAX_MULTIPLIER
            old_mode = ws.cell(max_tot_row, cols["proj_mode"]).value
            if apply_changes:
                ws.cell(max_tot_row, col).value = new_max_tot
                ws.cell(max_tot_row, cols["proj_mode"]).value = "User defined"
            max_total_adjusts.append({
                "target_sheet": sheet_name,
                "tech": tech, "year": y,
                "row": max_tot_row, "col": col,
                "op_life": op_life,
                "window_start": window_start,
                "accumulated_min": accumulated,
                "residual": residual,
                "old_max_tot": old_max_tot_num,
                "new_max_tot": new_max_tot,
                "old_mode": old_mode,
            })

    return max_adjusts, max_total_adjusts


# -------------------------------------------------------------------
# Sweep 2: activity-vs-capacity (V3)
# -------------------------------------------------------------------
def validate_activity_lower_limit(ws, sheet_name, cols, ymap,
                                   tech_to_row_by_param, oplife_map,
                                   yearsplit, c2a_map, capacities_wb,
                                   base_year, apply_changes):
    """V3 sweep: TotalTechnologyAnnualActivityLowerLimit must be ≤ max_activity.

    Returns list of dicts with diagnostic data and (if apply_changes)
    writes the haircut value back to the sheet.
    """
    al_rows = tech_to_row_by_param.get(ACT_LOWER_PARAM, {})
    issues = []
    sorted_years = sorted(ymap.keys())

    # Cache CF per tech (each call to load_capacity_factor reads the sheet,
    # which is expensive; only load techs that have ActivityLowerLimit).
    cf_cache = {}

    for tech, al_row in al_rows.items():
        op_life = oplife_map.get(tech, DEFAULT_OPLIFE)
        c2a = c2a_map.get(tech)
        if c2a is None:
            continue  # can't compute without CapToActUnit
        if tech not in cf_cache:
            cf_cache[tech] = load_capacity_factor(capacities_wb, tech)
        cf_by_year = cf_cache[tech]

        for y in sorted_years:
            col = ymap[y]
            al_val = ws.cell(al_row, col).value
            if is_empty_number(al_val):
                continue
            try:
                al_num = float(al_val)
            except (TypeError, ValueError):
                continue
            if al_num <= 0:
                continue

            af = get_availability_factor(ws, tech_to_row_by_param, tech, y, ymap)
            if af is None:
                continue
            ann_cf = compute_annual_cap_factor(yearsplit.get(y, {}), cf_by_year.get(y, {}))
            if ann_cf <= 0:
                continue
            max_cap = compute_max_capacity(ws, tech_to_row_by_param, tech, y,
                                           op_life, ymap, sorted_years)
            if max_cap is None:
                continue  # unbounded capacity → check vacuous
            max_act = compute_max_activity(max_cap, af, c2a, ann_cf)
            if max_act is None:
                continue
            if al_num <= max_act:
                continue  # consistent

            new_lower = max_act * ACT_LOWER_HAIRCUT
            old_mode = ws.cell(al_row, cols["proj_mode"]).value
            if apply_changes:
                ws.cell(al_row, col).value = new_lower
                ws.cell(al_row, cols["proj_mode"]).value = "User defined"
            issues.append({
                "target_sheet": sheet_name,
                "tech": tech, "year": y,
                "row": al_row, "col": col,
                "residual": _safe_residual(ws, tech_to_row_by_param, tech, y, ymap),
                "af": af, "c2a": c2a, "ann_cf": ann_cf,
                "max_capacity": max_cap, "max_activity": max_act,
                "old_lower": al_num, "new_lower": new_lower,
                "gap": al_num - max_act,
                "calibration_impact": (base_year is not None and y == int(base_year)),
                "old_mode": old_mode,
            })

    return issues


def _safe_residual(ws, tech_to_row_by_param, tech_upper, year, ymap):
    """Helper for diagnostics: read ResidualCapacity(tech, year) or 0."""
    res_rows = tech_to_row_by_param.get(RESIDUAL_PARAM, {})
    res_row = res_rows.get(tech_upper)
    if res_row is None:
        return 0.0
    col = ymap.get(year)
    if col is None:
        return 0.0
    v = ws.cell(res_row, col).value
    if is_empty_number(v):
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0
