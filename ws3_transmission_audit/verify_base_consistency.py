# -*- coding: utf-8 -*-
"""
WS-3 Phase 0 — Base-consistency verification (READ-ONLY acceptance test).

Checks, across the 3 canonical scenarios (A_Calibrated_BAU, B_Optimised_VRE,
C_Target_VRE), that the CLEANED repo's base is sound:

  1. STATIC INPUT identity  - SpecifiedAnnualDemand, ResidualCapacity, CapitalCost
                              (all techs) for base year 2023 + 2023-2027 must be
                              IDENTICAL across scenarios (reads *_Input.csv).
  2. SOLVE / objective       - read each scenario's Outputs/TotalDiscountedCost and
                              report the objective anchor per scenario.
  3. BASE-YEAR OUTPUT diff   - 2023 ProductionByTechnologyAnnual must match across
                              scenarios (within tol); flag any backstop (BCK) use.

Emits a pass/fail table. Uses existing solved artifacts (no re-solve). Reads only.
"""
from __future__ import annotations
import os, glob
import pandas as pd

REPO = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
EXE = os.path.join(REPO, "t1_confection", "Executables")
SCEN = {"A_Calibrated_BAU": "A_Calibrated_BAU_0",
        "B_Optimised_VRE": "B_Optimised_VRE_0",
        "C_Target_VRE": "C_Target_VRE_0"}
BASE_YEARS = [2023, 2024, 2025, 2026, 2027]
TOL = 1e-6
results = []   # (check, pass/fail, detail)

def rec(check, ok, detail=""):
    results.append((check, "PASS" if ok else "FAIL", detail))
    print(f"  [{'PASS' if ok else 'FAIL'}] {check}  {detail}")

# --------------------------------------------------------------- 1. static input identity
print("=" * 78, "\n1. STATIC INPUT identity across scenarios (2023 + 2023-2027)\n" + "=" * 78)
inputs = {}
for scen, d0 in SCEN.items():
    f = os.path.join(EXE, d0, f"{d0}_Input.csv")
    if os.path.isfile(f):
        inputs[scen] = pd.read_csv(f, low_memory=False)
        inputs[scen]["TECHNOLOGY"] = inputs[scen]["TECHNOLOGY"].astype(str)

ref_scen = "A_Calibrated_BAU"
for param in ["SpecifiedAnnualDemand", "ResidualCapacity", "CapitalCost"]:
    for horizon, yrs in [("2023", [2023]), ("2023-2027", BASE_YEARS)]:
        def slice_map(df):
            if param not in df.columns:
                return None
            sub = df[df["YEAR"].astype(float).isin(yrs) & df[param].notna()]
            key = ["TECHNOLOGY"] + (["FUEL"] if "FUEL" in df.columns and param == "SpecifiedAnnualDemand" else [])
            return {(tuple(str(r[k]) for k in key), float(r["YEAR"])): round(float(r[param]), 6)
                    for _, r in sub.iterrows()}
        ref = slice_map(inputs[ref_scen])
        diffs = []
        for scen, df in inputs.items():
            if scen == ref_scen:
                continue
            m = slice_map(df)
            if m is None or ref is None:
                continue
            keys = set(ref) | set(m)
            for k in keys:
                a, b = ref.get(k), m.get(k)
                if a != b and not (a is not None and b is not None and abs(a - b) < TOL):
                    diffs.append(f"{scen}:{k}={b} vs A={a}")
        rec(f"{param} identical [{horizon}]", len(diffs) == 0,
            "" if not diffs else f"{len(diffs)} diffs e.g. {diffs[:2]}")

# --------------------------------------------------------------- 2. objective anchor
print("\n" + "=" * 78, "\n2. SOLVE objective anchor (TotalDiscountedCost)\n" + "=" * 78)
def find_outputs(d0):
    for cand in [os.path.join(EXE, d0, "Outputs"), os.path.join(EXE, d0)]:
        if os.path.isdir(cand) and glob.glob(os.path.join(cand, "TotalDiscountedCost.csv")):
            return cand
    return None

obj = {}
for scen, d0 in SCEN.items():
    od = find_outputs(d0)
    f = os.path.join(od, "TotalDiscountedCost.csv") if od else None
    if f and os.path.isfile(f):
        df = pd.read_csv(f)
        vcol = "VALUE" if "VALUE" in df.columns else df.columns[-1]
        obj[scen] = float(df[vcol].sum())
        print(f"  {scen:20} sum(TotalDiscountedCost) = {obj[scen]:,.1f}")
    else:
        print(f"  {scen:20} (Outputs/TotalDiscountedCost.csv not found)")
rec("objective readable for all 3 scenarios", len(obj) == 3,
    f"{len(obj)}/3 found")

# --------------------------------------------------------------- 3. base-year output diff
print("\n" + "=" * 78, "\n3. BASE-YEAR (2023) generation identity + backstop check\n" + "=" * 78)
prod = {}
for scen, d0 in SCEN.items():
    od = find_outputs(d0)
    f = os.path.join(od, "ProductionByTechnologyAnnual.csv") if od else None
    if not (f and os.path.isfile(f)):
        f2 = os.path.join(od, "TotalTechnologyAnnualActivity.csv") if od else None
        f = f2 if (f2 and os.path.isfile(f2)) else None
    if f and os.path.isfile(f):
        df = pd.read_csv(f)
        vcol = "VALUE" if "VALUE" in df.columns else df.columns[-1]
        tcol = "TECHNOLOGY" if "TECHNOLOGY" in df.columns else df.columns[1]
        df = df[df["YEAR"].astype(float) == 2023]
        prod[scen] = df.groupby(tcol)[vcol].sum().round(3).to_dict()

if len(prod) >= 2:
    ref = prod.get(ref_scen) or next(iter(prod.values()))
    for scen, m in prod.items():
        if scen == ref_scen:
            continue
        keys = set(ref) | set(m)
        d = [k for k in keys if abs(ref.get(k, 0) - m.get(k, 0)) > max(1e-3, 1e-4 * abs(ref.get(k, 0)))]
        rec(f"2023 generation A vs {scen} identical", len(d) == 0,
            "" if not d else f"{len(d)} techs differ e.g. {d[:3]}")
    for scen, m in prod.items():
        bck = sum(v for k, v in m.items() if "BCK" in str(k).upper())
        rec(f"{scen}: no base-year backstop", bck < 1.0, f"BCK 2023 = {bck:.3f}")
else:
    rec("base-year generation comparable", False, "insufficient Outputs found")

# --------------------------------------------------------------- summary
print("\n" + "=" * 78, "\nPHASE-0 SUMMARY\n" + "=" * 78)
npass = sum(1 for _, s, _ in results if s == "PASS")
for check, status, detail in results:
    print(f"  {status:4}  {check}   {detail}")
print(f"\n  {npass}/{len(results)} checks passed")
if obj:
    print("\n  Objective anchors (WS-3 baseline for later cost-impact delta):")
    for scen, v in obj.items():
        print(f"    {scen:20} {v:,.1f}")
