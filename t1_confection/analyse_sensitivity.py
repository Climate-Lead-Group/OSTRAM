"""
analyse_sensitivity.py  --  OSTRAM Phase-B post-solve analysis (CLG / OSTRAM)
=============================================================================
Reads the post-solve combined CSV (OSTRAM_Combined_Inputs_Outputs.csv from
concat_all_scenarios_2.py) and produces the 8-scenario comparison table,
sensitivity deltas (measured against B_Opt_Clipped), per-node breakdowns and
flags. Read-only; writes sensitivity_comparison.csv + sensitivity_report.txt.

Metric definitions
------------------
  System cost (NPV)   sum(TotalDiscountedCost)                       M USD
  Capital (NPV)       sum(DiscountedCapitalInvestment[+Storage])     M USD
  Fuel+VarOpex        sum(AnnualVariableOperatingCost) (undiscounted) M USD
  Coal/Solar/Wind 2050 sum(TotalCapacityAnnual) PWRCOA/PWRSPV/PWRWON GW
  Storage 2050        sum(TotalCapacityAnnual) PWRSDS+PWRLDS          GW
  CO2 2050            sum(AnnualEmissions) EMISSION^CO2, y=2050        Mt
  BGD net imports 2050  demand - domestic gen  (energy balance)        TWh
  BGD domestic share    domestic gen / demand                          %
  BGD solar 2050        TotalCapacityAnnual PWRSPVBGDXX y=2050          GW
  Cross-border trade 2050  sum ProductionByTechnologyAnnual over
                           cross-border TRN corridors, y=2050 /3.6      TWh
  Backstop generation   sum ProductionByTechnologyAnnual over
                        PWRBCK*/TRNNLI*/TRNRPO*, all years /3.6         TWh
(1 TWh = 3.6 PJ; ProductionByTechnologyAnnual is PJ.)

Deltas are measured vs B_Opt_Clipped (isolates each lever from the clip
effect). B_Opt_Clipped - B_Opt is reported separately = cost of VRE realism.

Usage
-----
  python analyse_sensitivity.py                 # full run
  python analyse_sensitivity.py --unittest      # validate metrics vs B_Opt anchors
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path
import numpy as np
import pandas as pd

REPO = Path(__file__).resolve().parent
COMBINED = REPO / "OSTRAM_Combined_Inputs_Outputs.csv"
CEIL_JSON = REPO / "sensitivity_expansion" / "reference" / "vre_ceilings_base.json"
BASE_JSON = REPO / "sensitivity_expansion" / "reference" / "b_opt_baseline.json"
OUT_CSV = REPO / "sensitivity_comparison.csv"
OUT_TXT = REPO / "sensitivity_report.txt"

PJ_PER_TWH = 3.6
IND = ["INDEA", "INDNE", "INDNO", "INDSO", "INDWE"]
COUNTRIES = {"BGD", "IND", "LKA", "NPL", "BTN", "MDV"}
# display order
SCEN_ORDER = ["A_Calibrated_BAU", "A_Calibrated_BAU_Clipped",
              "B_Optimised_VRE", "B_Opt_Clipped",
              "C_Target_VRE", "C_Target_VRE_Clipped",
              "B_Opt_TradeCap15", "B_Opt_SolarCapexHi", "B_Opt_TxCap150",
              "B_Opt_IndiaCosts", "B_Opt_IndiaCostsFuel",
              "B_Opt_DirBidir", "B_Opt_DirContractual"]
SENSITIVITIES = ["B_Opt_TradeCap15", "B_Opt_SolarCapexHi", "B_Opt_TxCap150",
                 "B_Opt_IndiaCosts", "B_Opt_IndiaCostsFuel",
                 "B_Opt_DirBidir", "B_Opt_DirContractual"]
REF = "B_Opt_Clipped"          # sensitivities measured against this
CALBAU = "A_Calibrated_BAU"
BOPT = "B_Optimised_VRE"
# cost of physical VRE realism = each pathway's clipped run minus its unclipped run
CLIP_PAIRS = [("A_Calibrated_BAU_Clipped", "A_Calibrated_BAU"),
              ("B_Opt_Clipped", "B_Optimised_VRE"),
              ("C_Target_VRE_Clipped", "C_Target_VRE")]

USECOLS = ["Scenario", "REGION", "YEAR", "TECHNOLOGY", "EMISSION",
           "TotalDiscountedCost", "DiscountedCapitalInvestment",
           "DiscountedCapitalInvestmentStorage", "AnnualVariableOperatingCost",
           "TotalCapacityAnnual", "AnnualEmissions", "ProductionByTechnologyAnnual"]


def _is_cross_border_trn(t: str) -> bool:
    return len(t) == 13 and t.startswith("TRN") and t[3:6] in COUNTRIES and t[8:11] in COUNTRIES and t[3:6] != t[8:11]


def load(path=COMBINED):
    df = pd.read_csv(path, usecols=lambda c: c in USECOLS, low_memory=False)
    for c in USECOLS:
        if c not in df.columns:
            df[c] = np.nan
    df["YEAR"] = pd.to_numeric(df["YEAR"], errors="coerce")
    df["TECHNOLOGY"] = df["TECHNOLOGY"].astype(str)
    return df


def _sum(df, col):
    return float(pd.to_numeric(df[col], errors="coerce").sum())


def _cap2050(df, prefixes):
    m = df["YEAR"].eq(2050) & df["TECHNOLOGY"].str.startswith(tuple(prefixes))
    return float(pd.to_numeric(df.loc[m, "TotalCapacityAnnual"], errors="coerce").sum())


def metrics_for(df_s, bgd_demand_2050_twh):
    """df_s = combined rows for one scenario."""
    m = {}
    m["System cost (NPV) [M USD]"] = _sum(df_s, "TotalDiscountedCost")
    m["Capital (NPV) [M USD]"] = _sum(df_s, "DiscountedCapitalInvestment") + _sum(df_s, "DiscountedCapitalInvestmentStorage")
    m["Fuel+VarOpex [M USD]"] = _sum(df_s, "AnnualVariableOperatingCost")
    m["Coal 2050 [GW]"] = _cap2050(df_s, ["PWRCOA"])
    m["Solar 2050 [GW]"] = _cap2050(df_s, ["PWRSPV"])
    m["Wind 2050 [GW]"] = _cap2050(df_s, ["PWRWON"])
    m["Storage 2050 [GW]"] = _cap2050(df_s, ["PWRSDS", "PWRLDS"])
    co2 = df_s[df_s["YEAR"].eq(2050) & df_s["EMISSION"].astype(str).str.startswith("CO2")]
    m["CO2 2050 [Mt]"] = _sum(co2, "AnnualEmissions")
    # BGD energy balance
    # domestic generation = PWR generators at BGD, EXCLUDING storage (SDS/LDS),
    # backstop (BCK) and the grid/transmission converter (TRN, which re-outputs
    # generated+imported energy and would double-count).
    gen = df_s[df_s["YEAR"].eq(2050) & df_s["TECHNOLOGY"].str.match(r"PWR(?!BCK|SDS|LDS|TRN).{3}BGDXX$")]
    dom_gen_twh = _sum(gen, "ProductionByTechnologyAnnual") / PJ_PER_TWH
    m["BGD domestic gen 2050 [TWh]"] = dom_gen_twh
    m["BGD net imports 2050 [TWh]"] = bgd_demand_2050_twh - dom_gen_twh
    m["BGD domestic share 2050 [%]"] = 100.0 * dom_gen_twh / bgd_demand_2050_twh
    m["BGD solar 2050 [GW]"] = _cap2050(df_s[df_s["TECHNOLOGY"].eq("PWRSPVBGDXX")], ["PWRSPV"])
    xb = df_s[df_s["YEAR"].eq(2050) & df_s["TECHNOLOGY"].map(_is_cross_border_trn)]
    m["Cross-border trade 2050 [TWh]"] = _sum(xb, "ProductionByTechnologyAnnual") / PJ_PER_TWH
    # Backstop = PWRBCK (unserved-energy penalty tech) only. NOTE: TRNNLI*/TRNRPO*
    # are modeled imports-from-outside-region that B_Opt legitimately uses, NOT
    # backstop; TradeCap15/TxCap150 disallow the non-India ones as a lever.
    bk = df_s[df_s["TECHNOLOGY"].str.startswith("PWRBCK")]
    m["Backstop generation [TWh]"] = _sum(bk, "ProductionByTechnologyAnnual") / PJ_PER_TWH
    return m


def build(df, bgd_demand_2050_twh):
    scen_present = [s for s in SCEN_ORDER if s in set(df["Scenario"].unique())]
    table = {s: metrics_for(df[df["Scenario"].eq(s)], bgd_demand_2050_twh) for s in scen_present}
    return table, scen_present


def ceiling_flags(df, ceilings):
    """Return {scenario: [(tech, cap2050, ceil, pct)]} for node-techs >=95% of ceiling."""
    out = {}
    for s in df["Scenario"].unique():
        ds = df[df["Scenario"].eq(s) & df["YEAR"].eq(2050)]
        cap = pd.to_numeric(ds.set_index("TECHNOLOGY")["TotalCapacityAnnual"], errors="coerce")
        hits = []
        for t, ceil in ceilings.items():
            if ceil <= 0:
                continue
            v = float(cap.get(t, 0) if not isinstance(cap.get(t), pd.Series) else cap.get(t, pd.Series([0])).sum())
            if v >= 0.95 * ceil:
                hits.append((t, round(v, 2), ceil, round(100 * v / ceil, 1)))
        if hits:
            out[s] = sorted(hits, key=lambda x: -x[3])
    return out


def fmt(v, nd=1):
    if isinstance(v, float):
        return f"{v:,.{nd}f}"
    return str(v)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--unittest", action="store_true")
    ap.add_argument("--combined", default=str(COMBINED))
    args = ap.parse_args()

    ceilings = json.loads(CEIL_JSON.read_text())["ceilings_gw"]
    base = json.loads(BASE_JSON.read_text())
    bgd_demand_2050_twh = base["bgd_demand_PJ"]["2050"] / PJ_PER_TWH

    df = load(Path(args.combined))
    table, present = build(df, bgd_demand_2050_twh)

    if args.unittest:
        return unittest(table, present, ceilings)

    metrics = list(next(iter(table.values())).keys())
    # comparison CSV
    rows = []
    for mtr in metrics:
        rows.append({"metric": mtr, **{s: table[s][mtr] for s in present}})
    comp = pd.DataFrame(rows)
    comp.to_csv(OUT_CSV, index=False)

    # report
    L = []
    L.append("=" * 100)
    L.append("OSTRAM PHASE-B SENSITIVITY ANALYSIS")
    L.append("=" * 100)
    L.append(f"scenarios present: {present}")
    L.append(f"BGD 2050 demand: {bgd_demand_2050_twh:.1f} TWh (constant across scenarios)")
    L.append("")
    w = 26
    hdr = f"{'metric':34}" + "".join(f"{s.replace('B_Opt_','').replace('_VRE','')[:14]:>15}" for s in present)
    L.append(hdr); L.append("-" * len(hdr))
    for mtr in metrics:
        L.append(f"{mtr:34}" + "".join(f"{fmt(table[s][mtr]):>15}" for s in present))

    # cost of physical VRE realism — per pathway (clipped - unclipped)
    hdr_done = False
    for clip, base in CLIP_PAIRS:
        if clip in table and base in table:
            cost = "System cost (NPV) [M USD]"
            d = table[clip][cost] - table[base][cost]
            if not hdr_done:
                L.append("")
                L.append("Cost of physical VRE realism  (clipped - unclipped, per pathway):")
                hdr_done = True
            L.append(f"  {base:20} {d:>12,.0f} M USD  ({100*d/table[base][cost]:+.2f}%)")

    # deltas vs B_Opt_Clipped
    if REF in table:
        L.append("")
        L.append("=" * 100)
        L.append(f"SENSITIVITY DELTAS vs {REF}  (absolute / %)")
        L.append("=" * 100)
        for s in [x for x in SENSITIVITIES if x in table]:
            L.append(f"\n--- {s} ---")
            for mtr in metrics:
                a, b = table[s][mtr], table[REF][mtr]
                da = a - b
                pct = (100 * da / b) if abs(b) > 1e-9 else float("nan")
                L.append(f"  {mtr:34} {fmt(a):>14}   d={da:>+12,.1f}   {pct:>+7.1f}%")

    # flags
    L.append("")
    L.append("=" * 100)
    L.append("FLAGS")
    L.append("=" * 100)
    cf = ceiling_flags(df, ceilings)
    L.append("\n[>=95% of VRE ceiling  (physical potential binding)]")
    if cf:
        for s in present:
            if s in cf:
                for t, v, c, p in cf[s]:
                    L.append(f"  {s:20} {t:14} {v:>8} / {c:>6} GW  ({p}%)")
    else:
        L.append("  (none)")
    L.append("\n[backstop generation > 0  (model stress)]")
    anyb = False
    for s in present:
        b = table[s]["Backstop generation [TWh]"]
        if b > 1e-6:
            L.append(f"  {s:20} {b:,.3f} TWh"); anyb = True
    if not anyb:
        L.append("  (none)")
    L.append("\n[sensitivity objective > CalBAU  (constraint economically unreasonable)]")
    anyc = False
    if CALBAU in table:
        cb = table[CALBAU]["System cost (NPV) [M USD]"]
        for s in [x for x in SENSITIVITIES if x in table]:
            if table[s]["System cost (NPV) [M USD]"] > cb:
                L.append(f"  {s:20} {table[s]['System cost (NPV) [M USD]']:,.0f} > CalBAU {cb:,.0f}"); anyc = True
    if not anyc:
        L.append("  (none)")

    # per-node breakdown BGD/INDNO/INDSO
    L.append("")
    L.append("=" * 100)
    L.append("PER-NODE 2050 CAPACITY (GW) — BGDXX / INDNO / INDSO")
    L.append("=" * 100)
    fams = ["PWRCOA", "PWRNGS", "PWROIL", "PWRSPV", "PWRWON", "PWRHYD", "PWRURN", "PWRSDS", "PWRLDS"]
    for node in ["BGDXX", "INDNO", "INDSO"]:
        L.append(f"\n[{node}]")
        L.append(f"  {'family':10}" + "".join(f"{s.replace('B_Opt_','')[:12]:>13}" for s in present))
        for fam in fams:
            vals = []
            for s in present:
                ds = df[df["Scenario"].eq(s) & df["YEAR"].eq(2050) & df["TECHNOLOGY"].eq(fam + node)]
                vals.append(_sum(ds, "TotalCapacityAnnual"))
            if any(abs(v) > 1e-6 for v in vals):
                L.append(f"  {fam:10}" + "".join(f"{v:>13.2f}" for v in vals))

    txt = "\n".join(L)
    OUT_TXT.write_text(txt, encoding="utf-8")
    print(txt)
    print(f"\nWrote {OUT_CSV.name} and {OUT_TXT.name}")
    return 0


def unittest(table, present, ceilings):
    """Validate metric extraction against known Phase-A anchors (B_Opt) and
    the trivial B_Opt-vs-B_Opt=0 identity. Runs on the existing combined CSV
    (baselines only; sensitivities not yet solved)."""
    print("UNIT TEST — metric validation on existing combined CSV")
    print(f"  scenarios present: {present}")
    fails = []
    def check(name, got, exp, tol):
        ok = abs(got - exp) <= tol
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}: got {got:,.2f}  exp {exp:,.2f} (+-{tol})")
        if not ok: fails.append(name)

    if BOPT not in table:
        print("  B_Optimised_VRE not in combined CSV — cannot validate."); return 2
    b = table[BOPT]
    check("B_Opt System cost (NPV)", b["System cost (NPV) [M USD]"], 2_113_984.5, 50)
    check("B_Opt Coal 2050", b["Coal 2050 [GW]"], 298.79, 1.0)
    check("B_Opt Solar 2050", b["Solar 2050 [GW]"], 886.61, 1.0)
    check("B_Opt Wind 2050", b["Wind 2050 [GW]"], 506.54, 1.0)
    check("B_Opt BGD net imports 2050 (TWh)", b["BGD net imports 2050 [TWh]"], 301.0, 15.0)
    check("B_Opt BGD domestic share (%)", b["BGD domestic share 2050 [%]"], 38.0, 3.0)
    check("B_Opt Backstop generation (TWh)", b["Backstop generation [TWh]"], 0.0, 0.1)

    # B_Opt vs B_Opt = 0
    zero_ok = True
    for mtr in b:
        if abs(b[mtr] - table[BOPT][mtr]) > 1e-9:
            zero_ok = False
    print(f"  [{'PASS' if zero_ok else 'FAIL'}] B_Opt - B_Opt == 0 for all metrics")
    if not zero_ok: fails.append("self-delta")

    # CalBAU vs B_Opt known-direction deltas (from Phase A: CalBAU has MORE coal, LESS solar)
    if CALBAU in table:
        c = table[CALBAU]
        d_coal = c["Coal 2050 [GW]"] - b["Coal 2050 [GW]"]
        d_solar = c["Solar 2050 [GW]"] - b["Solar 2050 [GW]"]
        ok = d_coal > 100 and d_solar < -400
        print(f"  [{'PASS' if ok else 'FAIL'}] CalBAU-B_Opt direction: dCoal={d_coal:+.1f} (>100), dSolar={d_solar:+.1f} (<-400)")
        if not ok: fails.append("calbau-direction")

    print(f"\n{'ALL UNIT TESTS PASSED' if not fails else 'FAILURES: ' + str(fails)}")
    return 0 if not fails else 2


if __name__ == "__main__":
    sys.exit(main())
