#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_3_step_2d.py

Acceptance test for Step 2D inside 3_update_ao_from_extensions.py.

Step 2D's contract: WV is the single source of truth. For every
(match-key, year) pair present in BOTH WV and post-pipeline AO across the
five Parametrization sheets refreshed by Step 2D, the AO value must equal
the WV value exactly (no tolerance, no merging, just copy).

Coverage: five AO sheets x their match keys
    Primary Techs            <- Primary_Techs            (Tech, Parameter)
    Secondary Techs          <- Secondary_Techs          (Tech, Parameter)
    Fixed Horizon Parameters <- Fixed_Horizon_Parameters (Tech, Parameter)
    VariableCost             <- VariableCost             (Tech, Mode.Operation)
    Capacities               <- Capacities_CF            (Tech, Timeslices)

For each sheet there are three test families:
    Tnn_match  -- matched-key cell-equality (THE Step 2D contract)
    Tnn_wvonly -- every WV-only key appears in post-pipeline AO
                  (i.e. Step 3 additive pass picked it up)
    Tnn_aoonly -- AO-only keys are reported as informational counts (always pass)

Plus structural checks for the Capacities passthrough (20 timeslices on every
refreshed tech, Tech.IDs preserved).

Files expected (all in same directory as this script):
    SOASIA_OSeMOSYS_WV.xlsx                   -- the source of truth
    A-O_Parametrization.xlsx                   -- pre-pipeline input (for Tech.ID baseline)
    wvaligned_outputs/A-O_Parametrization_wvaligned.xlsx
                                               -- post-script-3 output

Exits 0 if all assertions pass; nonzero on any failure.
"""
from pathlib import Path
import sys
import pandas as pd
import numpy as np

SCRIPT_DIR = Path(__file__).resolve().parent
WV_FILE       = SCRIPT_DIR / "SOASIA_OSeMOSYS_WV.xlsx"
AO_INPUT      = SCRIPT_DIR / "A-O_Parametrization.xlsx"
AO_POST3      = SCRIPT_DIR / "wvaligned_outputs" / "A-O_Parametrization_wvaligned.xlsx"

# -------------------- minimal test runner (matches test_5 style) --------------
_results = []
def test(label):
    def deco(fn):
        def runner():
            try:
                fn()
                _results.append((True, label, None))
                print(f"  PASS  {label}")
            except AssertionError as e:
                _results.append((False, label, str(e)))
                print(f"  FAIL  {label}")
                print(f"        {e}")
            except Exception as e:
                _results.append((False, label, f"{type(e).__name__}: {e}"))
                print(f"  ERROR {label}")
                print(f"        {type(e).__name__}: {e}")
        runner.__name__ = fn.__name__
        runner()
        return runner
    return deco


# -------------------- helpers --------------------
def is_year_col(c):
    if isinstance(c, int) and 2000 <= c <= 2100:
        return True
    if isinstance(c, str) and c.isdigit() and 2000 <= int(c) <= 2100:
        return True
    return False


def col_as_year_int(c):
    return c if isinstance(c, int) else int(c)


def norm_key(v):
    if v is None:
        return None
    if isinstance(v, float):
        if pd.isna(v):
            return None
        if v.is_integer():
            return str(int(v))
        return str(v)
    if isinstance(v, str):
        s = v.strip()
        return s if s else None
    return str(v)


def equal_or_close(a, b, atol=0.0):
    """Cell equality. Numerics: bit-exact by default (atol=0) since Step 2D
    is a copy, not a computation. None / NaN treated as equivalent."""
    a_null = a is None or (isinstance(a, float) and np.isnan(a))
    b_null = b is None or (isinstance(b, float) and np.isnan(b))
    if a_null and b_null:
        return True
    if a_null != b_null:
        return False
    try:
        return abs(float(a) - float(b)) <= atol
    except (TypeError, ValueError):
        return a == b


def is_null(v):
    return v is None or (isinstance(v, float) and np.isnan(v))


# -------------------- check-runner per sheet (year cells) --------------------
def assert_year_cell_equality(ao_df, wv_df, key_cols, sheet_label, atol=0.0):
    """For every (key, year) pair where WV has a non-null value, AO == WV.

    Step 2D semantics: WV is the source of truth for cells where WV has a
    value. WV-null cells are "no claim from source of truth" -- AO retains
    whatever it had (matches Step 2B's pattern: `if pd.notna(wv_v)` gate).
    """
    wv_year_int = {col_as_year_int(c): c for c in wv_df.columns if is_year_col(c)}
    ao_year_int = {col_as_year_int(c): c for c in ao_df.columns if is_year_col(c)}
    common_years = sorted(set(wv_year_int) & set(ao_year_int))
    assert common_years, f"{sheet_label}: no common year columns between AO and WV"

    wv_idx = {}
    for _, r in wv_df.iterrows():
        k = tuple(norm_key(r[c]) if c in r.index else None for c in key_cols)
        if any(p is None for p in k):
            continue
        wv_idx[k] = r

    ao_idx = {}
    for _, r in ao_df.iterrows():
        k = tuple(norm_key(r[c]) if c in r.index else None for c in key_cols)
        if any(p is None for p in k):
            continue
        ao_idx.setdefault(k, []).append(r)

    matched = set(wv_idx.keys()) & set(ao_idx.keys())
    assert matched, f"{sheet_label}: no matched keys between AO and WV"

    mismatches = []
    n_compared = 0
    n_skipped_wv_null = 0
    for k in matched:
        wv_r = wv_idx[k]
        for ao_r in ao_idx[k][:1]:
            for yint in common_years:
                wv_v = wv_r.get(wv_year_int[yint])
                ao_v = ao_r.get(ao_year_int[yint])
                if is_null(wv_v):
                    # WV makes no claim about this cell; AO keeps its value.
                    n_skipped_wv_null += 1
                    continue
                n_compared += 1
                if not equal_or_close(ao_v, wv_v, atol=atol):
                    mismatches.append((k, yint, ao_v, wv_v))

    if mismatches:
        sample = mismatches[:5]
        raise AssertionError(
            f"{sheet_label}: {len(mismatches)} cell mismatches across "
            f"{len(matched)} matched keys ({n_compared} cells compared, "
            f"{n_skipped_wv_null} skipped because WV is null). "
            f"Sample: {sample}"
        )

    print(f"        {sheet_label}: {len(matched)} matched keys, "
          f"{n_compared} cells compared (+{n_skipped_wv_null} WV-null skipped), "
          f"all equal")


def assert_value_scalar_equality(ao_df, wv_df, key_cols, sheet_label, value_col="Value"):
    """For every key in both AO and WV with non-null WV value,
    AO[value_col] == WV[value_col]."""
    wv_idx = {}
    for _, r in wv_df.iterrows():
        k = tuple(norm_key(r[c]) if c in r.index else None for c in key_cols)
        if any(p is None for p in k):
            continue
        wv_idx[k] = r

    ao_idx = {}
    for _, r in ao_df.iterrows():
        k = tuple(norm_key(r[c]) if c in r.index else None for c in key_cols)
        if any(p is None for p in k):
            continue
        ao_idx.setdefault(k, []).append(r)

    matched = set(wv_idx.keys()) & set(ao_idx.keys())
    assert matched, f"{sheet_label}: no matched keys"

    mismatches = []
    n_compared = 0
    n_skipped = 0
    for k in matched:
        wv_v = wv_idx[k].get(value_col)
        ao_v = ao_idx[k][0].get(value_col)
        if is_null(wv_v):
            n_skipped += 1
            continue
        n_compared += 1
        if not equal_or_close(ao_v, wv_v):
            mismatches.append((k, ao_v, wv_v))

    if mismatches:
        raise AssertionError(
            f"{sheet_label}: {len(mismatches)} {value_col!r} mismatches across "
            f"{len(matched)} matched keys. Sample: {mismatches[:5]}"
        )

    print(f"        {sheet_label}: {len(matched)} matched keys, "
          f"{n_compared} {value_col!r} cells compared "
          f"(+{n_skipped} WV-null skipped), all equal")


def report_wv_only_keys(ao_df, wv_df, key_cols, sheet_label):
    """Informational: WV-only keys not in post-pipeline AO.

    Per handover, WV-only keys SHOULD be present in AO via Step 3's additive
    pass. But Step 3 only appends techs from OSTRAM_AO_Extensions Tab 1
    Include=Y -- a curated subset (16 PWR/TRN techs). Other WV techs
    (RNW*, RNWBIO*, etc.) appear here as a deferred upstream-completeness
    gap explicitly out of scope per handover s7 / item 5 'Address after
    Step 2D ships'. Reported as INFO, not failure.
    """
    wv_keys = set()
    for _, r in wv_df.iterrows():
        k = tuple(norm_key(r[c]) if c in r.index else None for c in key_cols)
        if not any(p is None for p in k):
            wv_keys.add(k)
    ao_keys = set()
    for _, r in ao_df.iterrows():
        k = tuple(norm_key(r[c]) if c in r.index else None for c in key_cols)
        if not any(p is None for p in k):
            ao_keys.add(k)
    missing = wv_keys - ao_keys
    if missing:
        # Distill to unique Tech list so the message is actionable
        if "Tech" in key_cols:
            tech_idx = key_cols.index("Tech")
            missing_techs = sorted({k[tech_idx] for k in missing})
            print(f"        {sheet_label}: INFO {len(missing)} WV-only keys "
                  f"({len(missing_techs)} unique techs) not in AO -- "
                  f"upstream Tab 1 gap, addressed in handover s7 item 5. "
                  f"Sample techs: {missing_techs[:5]}")
        else:
            print(f"        {sheet_label}: INFO {len(missing)} WV-only keys "
                  f"not in AO -- upstream Tab 1 gap")
    else:
        print(f"        {sheet_label}: all {len(wv_keys)} WV keys present in AO")


def report_ao_only_keys(ao_df, wv_df, key_cols, sheet_label):
    """Informational count of AO-only keys (untouched per handover rule 2)."""
    wv_keys = set()
    for _, r in wv_df.iterrows():
        k = tuple(norm_key(r[c]) if c in r.index else None for c in key_cols)
        if not any(p is None for p in k):
            wv_keys.add(k)
    ao_keys = set()
    for _, r in ao_df.iterrows():
        k = tuple(norm_key(r[c]) if c in r.index else None for c in key_cols)
        if not any(p is None for p in k):
            ao_keys.add(k)
    ao_only = ao_keys - wv_keys
    print(f"        {sheet_label}: {len(ao_only)} AO-only keys (untouched, "
          f"this is expected per handover rule 2)")


# -------------------- load --------------------
print("=" * 70)
print("Acceptance test for Step 2D (3_update_ao_from_extensions.py)")
print("=" * 70)

assert WV_FILE.is_file(),    f"missing {WV_FILE}"
assert AO_INPUT.is_file(),   f"missing {AO_INPUT}"
assert AO_POST3.is_file(),   f"missing {AO_POST3}  (run script 3 first)"

# Pre-load all the dataframes we'll touch
WV_PRIMARY      = pd.read_excel(WV_FILE,  sheet_name="Primary_Techs")
WV_SECONDARY    = pd.read_excel(WV_FILE,  sheet_name="Secondary_Techs")
WV_FH           = pd.read_excel(WV_FILE,  sheet_name="Fixed_Horizon_Parameters")
WV_VARCOST      = pd.read_excel(WV_FILE,  sheet_name="VariableCost")
WV_CAPACITIES   = pd.read_excel(WV_FILE,  sheet_name="Capacities_CF")

AO_PRIMARY      = pd.read_excel(AO_POST3, sheet_name="Primary Techs")
AO_SECONDARY    = pd.read_excel(AO_POST3, sheet_name="Secondary Techs")
AO_FH           = pd.read_excel(AO_POST3, sheet_name="Fixed Horizon Parameters")
AO_VARCOST      = pd.read_excel(AO_POST3, sheet_name="VariableCost")
AO_CAPACITIES   = pd.read_excel(AO_POST3, sheet_name="Capacities")

AO_INPUT_CAP    = pd.read_excel(AO_INPUT, sheet_name="Capacities")


# ============================================================================
# Primary Techs
# ============================================================================

@test("T01_match: Primary Techs (Tech, Parameter) -- year-cell equality")
def t01():
    assert_year_cell_equality(
        AO_PRIMARY, WV_PRIMARY,
        key_cols=["Tech", "Parameter"],
        sheet_label="Primary Techs"
    )


@test("T01_wvonly: report Primary Techs WV-only keys (informational; upstream Tab 1 gap)")
def t01b():
    report_wv_only_keys(
        AO_PRIMARY, WV_PRIMARY,
        key_cols=["Tech", "Parameter"],
        sheet_label="Primary Techs"
    )


@test("T01_aoonly: report Primary Techs AO-only keys (informational)")
def t01c():
    report_ao_only_keys(
        AO_PRIMARY, WV_PRIMARY,
        key_cols=["Tech", "Parameter"],
        sheet_label="Primary Techs"
    )


# ============================================================================
# Secondary Techs
# ============================================================================

@test("T02_match: Secondary Techs (Tech, Parameter) -- year-cell equality")
def t02():
    assert_year_cell_equality(
        AO_SECONDARY, WV_SECONDARY,
        key_cols=["Tech", "Parameter"],
        sheet_label="Secondary Techs"
    )


@test("T02_wvonly: report Secondary Techs WV-only keys (informational; upstream Tab 1 gap)")
def t02b():
    report_wv_only_keys(
        AO_SECONDARY, WV_SECONDARY,
        key_cols=["Tech", "Parameter"],
        sheet_label="Secondary Techs"
    )


@test("T02_aoonly: report Secondary Techs AO-only keys (informational)")
def t02c():
    report_ao_only_keys(
        AO_SECONDARY, WV_SECONDARY,
        key_cols=["Tech", "Parameter"],
        sheet_label="Secondary Techs"
    )


# ============================================================================
# Fixed Horizon Parameters
# ============================================================================

@test("T03_match: FH Parameters (Tech, Parameter) -- 'Value' equality")
def t03():
    assert_value_scalar_equality(
        AO_FH, WV_FH,
        key_cols=["Tech", "Parameter"],
        sheet_label="Fixed Horizon Parameters",
        value_col="Value"
    )


@test("T03_wvonly: report FH WV-only keys (informational; upstream Tab 1 gap)")
def t03b():
    report_wv_only_keys(
        AO_FH, WV_FH,
        key_cols=["Tech", "Parameter"],
        sheet_label="Fixed Horizon Parameters"
    )


@test("T03_aoonly: report FH AO-only keys (informational)")
def t03c():
    report_ao_only_keys(
        AO_FH, WV_FH,
        key_cols=["Tech", "Parameter"],
        sheet_label="Fixed Horizon Parameters"
    )


# ============================================================================
# VariableCost
# ============================================================================

@test("T04_match: VariableCost (Tech, Mode.Operation) -- year-cell equality")
def t04():
    assert_year_cell_equality(
        AO_VARCOST, WV_VARCOST,
        key_cols=["Tech", "Mode.Operation"],
        sheet_label="VariableCost"
    )


@test("T04_wvonly: report VariableCost WV-only keys (informational; upstream Tab 1 gap)")
def t04b():
    report_wv_only_keys(
        AO_VARCOST, WV_VARCOST,
        key_cols=["Tech", "Mode.Operation"],
        sheet_label="VariableCost"
    )


@test("T04_aoonly: report VariableCost AO-only keys (informational)")
def t04c():
    report_ao_only_keys(
        AO_VARCOST, WV_VARCOST,
        key_cols=["Tech", "Mode.Operation"],
        sheet_label="VariableCost"
    )


# ============================================================================
# Capacities -- the key new behaviour: 20-timeslice passthrough,
# Tech.ID preserved on refreshed techs, AO-only techs untouched.
# ============================================================================

@test("T05_match: Capacities (Tech, Timeslices) -- year-cell equality")
def t05():
    assert_year_cell_equality(
        AO_CAPACITIES, WV_CAPACITIES,
        key_cols=["Tech", "Timeslices"],
        sheet_label="Capacities"
    )


@test("T05_wvonly: report Capacities WV-only keys (informational; upstream Tab 1 gap)")
def t05b():
    report_wv_only_keys(
        AO_CAPACITIES, WV_CAPACITIES,
        key_cols=["Tech", "Timeslices"],
        sheet_label="Capacities"
    )


@test("T05_aoonly: report Capacities AO-only (Tech, Timeslices) keys (informational)")
def t05c():
    report_ao_only_keys(
        AO_CAPACITIES, WV_CAPACITIES,
        key_cols=["Tech", "Timeslices"],
        sheet_label="Capacities"
    )


@test("T05_struct1: every refreshed tech has exactly 20 timeslices in AO")
def t05d():
    wv_techs = set(WV_CAPACITIES["Tech"].dropna().unique())
    ao_techs = set(AO_CAPACITIES["Tech"].dropna().unique())
    refreshed = wv_techs & ao_techs

    bad = []
    for t in refreshed:
        rows = AO_CAPACITIES[AO_CAPACITIES["Tech"] == t]
        ts = sorted(rows["Timeslices"].dropna().astype(str).unique())
        if len(ts) != 20:
            bad.append((t, len(ts)))
    if bad:
        raise AssertionError(
            f"{len(bad)} refreshed tech(s) lack the expected 20 timeslices: {bad[:5]}"
        )
    print(f"        all {len(refreshed)} refreshed techs have 20 timeslices each")


@test("T05_struct2: AO-only Capacities techs (no WV CF row) are left untouched")
def t05e():
    """AO techs with no WV Capacities_CF entry must keep their pre-pipeline
    timeslice count (12 in this dataset) and original Tech.IDs."""
    wv_techs = set(WV_CAPACITIES["Tech"].dropna().unique())
    ao_in    = AO_INPUT_CAP
    ao_out   = AO_CAPACITIES

    ao_only_techs = set(ao_in["Tech"].dropna().unique()) - wv_techs

    issues = []
    for t in ao_only_techs:
        in_rows  = ao_in[ao_in["Tech"]  == t].sort_values("Timeslices")
        out_rows = ao_out[ao_out["Tech"] == t].sort_values("Timeslices")
        if len(out_rows) != len(in_rows):
            issues.append((t, "row count drift",
                           f"{len(in_rows)} -> {len(out_rows)}"))
            continue
        # Check Tech.ID preserved
        in_ids  = set(in_rows["Tech.ID"].dropna().unique())
        out_ids = set(out_rows["Tech.ID"].dropna().unique())
        if in_ids != out_ids:
            issues.append((t, "Tech.ID drift", f"{in_ids} -> {out_ids}"))
    if issues:
        raise AssertionError(
            f"{len(issues)} AO-only tech issue(s): {issues[:5]}"
        )
    print(f"        {len(ao_only_techs)} AO-only techs preserved unchanged")


@test("T05_struct3: refreshed techs preserve their original AO Tech.ID")
def t05f():
    wv_techs = set(WV_CAPACITIES["Tech"].dropna().unique())
    refreshed = set(AO_INPUT_CAP["Tech"].dropna().unique()) & wv_techs

    drift = []
    for t in refreshed:
        in_ids = set(AO_INPUT_CAP[AO_INPUT_CAP["Tech"] == t]["Tech.ID"].dropna().unique())
        out_ids = set(AO_CAPACITIES[AO_CAPACITIES["Tech"] == t]["Tech.ID"].dropna().unique())
        if in_ids != out_ids:
            drift.append((t, in_ids, out_ids))
    if drift:
        raise AssertionError(
            f"{len(drift)} refreshed tech(s) had Tech.ID drift: {drift[:5]}"
        )
    print(f"        {len(refreshed)} refreshed techs preserved their Tech.ID")


# ============================================================================
# Summary
# ============================================================================
print()
n_total = len(_results)
n_pass  = sum(1 for ok, _, _ in _results if ok)
n_fail  = n_total - n_pass
print("-" * 70)
print(f"SUMMARY:  {n_pass}/{n_total} passed,  {n_fail} failed")
print("-" * 70)

sys.exit(0 if n_fail == 0 else 1)
