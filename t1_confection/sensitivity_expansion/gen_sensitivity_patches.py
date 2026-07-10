"""
gen_sensitivity_patches.py  --  OSTRAM sensitivity patch generator (CLG / OSTRAM)
================================================================================
Deterministically builds patches.json for the three *computed* sensitivity
runs from the validated B_Optimised_VRE A-O + reference baseline. Nothing is
hardcoded: corridor lists, residuals and India reference costs are all read
from the A-O. Prints verification tables. Run once; commit the outputs.

  B_Opt_TradeCap30  : BGD imports <= 30% of demand (per-year, split by B_Opt
                      import share) + 1.5x export allowance; backstops -> 0.
  B_Opt_TxCap150    : every cross-border TRN corridor MaxCap = 1.5 x Residual_2023
                      (India-internal kept; non-India backstops -> MaxCap 0).
  B_Opt_IndiaCosts  : CapitalCost + FixedCost of every non-India gen/storage tech
                      set to the India reference (INDNO anchor where India nodes
                      disagree). Transmission excluded; fuel off by default.
"""
import json
from pathlib import Path
import pandas as pd

REPO = Path(r"C:/Users/luisfernando/Desktop/OSeMOSYS/OSTRAM_clean/t1_confection")
AO = REPO / "A1_Outputs" / "A1_Outputs_B_Optimised_VRE" / "A-O_Parametrization.xlsx"
CONFIGS = REPO / "A3_process" / "rules_scripts" / "configs"
BASE = json.loads((REPO / "sensitivity_expansion" / "reference" / "b_opt_baseline.json").read_text())
YEARS = list(range(2023, 2051))
IND = ["INDEA", "INDNE", "INDNO", "INDSO", "INDWE"]
NONIND = ["BGDXX", "BTNXX", "LKAXX", "MDVXX", "NPLXX"]


def load_sheet(sheet):
    df = pd.read_excel(AO, sheet_name=sheet, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    tcol = "Tech" if "Tech" in df.columns else "Technology"
    ycols = {}
    for c in df.columns:
        try:
            y = int(float(c))
            if 2020 <= y <= 2060:
                ycols[y] = c
        except (ValueError, TypeError):
            pass
    out = {}
    for _, r in df.iterrows():
        t, p = r.get(tcol), r.get("Parameter")
        if pd.isna(t) or pd.isna(p):
            continue
        out[(str(t).strip(), str(p).strip())] = {
            y: (None if pd.isna(r[ycols[y]]) else float(r[ycols[y]])) for y in ycols}
    return out


def write_patches(scen, obj):
    p = CONFIGS / scen / "patches.json"
    p.write_text(json.dumps(obj, indent=1))
    print(f"  wrote {p.relative_to(REPO)}  ({len(obj['edits'])} edits)")


# ---------------------------------------------------------------- TradeCap30
def gen_tradecap(frac, export_factor=1.5):
    """BGD cross-border import cap at `frac` of demand. Bangladesh is the region's
    ONLY net importer (verified: LKA/NPL/BTN net exporters, IND/MDV self-sufficient
    in B_Opt 2050), so a region-wide import cap binds on BGD alone -- capping the
    exporters would be a no-op at best and could throttle their exports.

    export_factor: multiplier on each corridor's baseline export added to its AUL.
    Default 1.5 preserves historical behaviour (TradeCap30/50). Pass 0.0 for a STRICT
    cap: total corridor throughput <= frac*demand, so net imports actually land at
    <=frac (no export headroom the optimiser could otherwise spend on imports -- which
    is why TradeCap15 with the 1.5x allowance landed at ~18%, not 15%)."""
    pct = int(round(frac * 100))
    scen = f"B_Opt_TradeCap{pct}"
    print("\n" + "=" * 70 + f"\n{scen}\n" + "=" * 70)
    dem = BASE["bgd_demand_PJ"]; imp = BASE["corridor_import_PJ"]; exp = BASE["corridor_export_PJ"]
    REAL = ["TRNBGDXXINDEA", "TRNBGDXXINDNE", "TRNBTNXXBGDXX", "TRNNPLXXBGDXX"]

    def auls(f, ef):
        out = {c: {} for c in REAL}
        for y in map(str, YEARS):
            tot = sum(imp[c][y] for c in REAL)
            for c in REAL:
                share = imp[c][y] / tot if tot > 0 else 0.0
                out[c][y] = round(f * float(dem[y]) * share + ef * exp[c][y], 4)
        return out

    alw = f"+ {export_factor:g}x export allowance" if export_factor else "(strict cap, no export allowance)"
    # SANITY: recompute at 0.50 with the 1.5x allowance TradeCap50 was built with
    a50 = auls(0.50, 1.5)
    old = json.loads((CONFIGS / "B_Opt_TradeCap50" / "patches.json").read_text())
    old_vals = {e["tech"]: e["values"] for e in old["edits"] if "values" in e}
    maxerr = 0.0
    for c in REAL:
        for y in map(str, YEARS):
            if c in old_vals:
                maxerr = max(maxerr, abs(a50[c][y] - old_vals[c][y]))
    print(f"  formula sanity (recomputed 0.50 vs built TradeCap50): max abs err = {maxerr:.4f}  {'OK' if maxerr < 0.01 else 'MISMATCH!'}")

    a = auls(frac, export_factor)
    edits = []
    for c in REAL:
        edits.append({"sheet": "Secondary Techs", "tech": c,
                      "param": "TotalTechnologyAnnualActivityUpperLimit",
                      "values": a[c], "create_if_absent": False,
                      "note": f"import<={pct}% budget x per-year B_Opt share {alw}"})
    for c in ["TRNNLIBGDXX", "TRNRPOBGDXX"]:
        edits.append({"sheet": "Demand Techs", "tech": c,
                      "param": "TotalTechnologyAnnualActivityUpperLimit",
                      "op": "set_flat", "value": 0.0, "create_if_absent": True,
                      "note": "backstop import disallowed"})
    write_patches(scen, {
        "scenario": scen, "base_scenario": "B_Optimised_VRE",
        "cap_fraction": frac, "export_factor": export_factor,
        "description": f"BGD cross-border imports capped to {pct}% of demand (per-year, split by B_Opt import share) {alw}; backstop imports (TRNNLIBGDXX/TRNRPOBGDXX) zeroed. Region-wide {pct}% import cap binds on Bangladesh only (sole net importer; LKA/NPL/BTN export, IND/MDV self-sufficient). Over shared VRE-ceiling layer. Generated by gen_sensitivity_patches.py.",
        "apply_vre_ceiling_layer": True, "edits": edits})
    print(f"  2050 AUL @{pct}%: " + ", ".join(f"{c[3:]}={a[c]['2050']}" for c in REAL) + f"   [{alw}]")


# ---------------------------------------------------------------- TxCap150
def gen_txcap150():
    print("\n" + "=" * 70 + "\nB_Opt_TxCap150\n" + "=" * 70)
    st = load_sheet("Secondary Techs")
    resid = {t: v.get(2023) for (t, p), v in st.items() if p == "ResidualCapacity" and t.startswith("TRN")}
    cross = sorted(t for t in resid if len(t) == 13 and t[3:6] != t[8:11])
    internal = sorted(t for t in resid if len(t) == 13 and t[3:6] == t[8:11])
    build = {}
    capcsv = REPO / "Executables" / "B_Optimised_VRE_0" / "Outputs" / "TotalCapacityAnnual.csv"
    import csv as _csv
    with open(capcsv, newline="") as f:
        for row in _csv.DictReader(f):
            if row["TECHNOLOGY"].startswith("TRN") and int(float(row["YEAR"])) == 2050:
                build[row["TECHNOLOGY"]] = float(row["VALUE"])

    edits = []
    print(f"  {'corridor':16}{'resid_2023':>11}{'1.5x_cap':>10}{'B_Opt_2050':>12}  effect")
    for c in cross:
        cap = round(1.5 * resid[c], 6)
        vals = {str(y): cap for y in YEARS}
        edits.append({"sheet": "Secondary Techs", "tech": c, "param": "TotalAnnualMaxCapacity",
                      "values": vals, "note": "cross-border cap = 1.5 x Residual_2023"})
        edits.append({"sheet": "Secondary Techs", "tech": c, "param": "TotalAnnualMaxCapacityInvestment",
                      "values": vals, "note": "coherence: MaxCapInv <= MaxCap"})
        headroom = round(cap - resid[c], 6)  # MaxCap - Residual
        mincinv = st.get((c, "TotalAnnualMinCapacityInvestment"), {})
        mci_vals = {}
        for y in YEARS:
            orig = mincinv.get(y)
            if orig is not None:
                mci_vals[str(y)] = round(min(orig, headroom), 6)
        if mci_vals:
            edits.append({"sheet": "Secondary Techs", "tech": c, "param": "TotalAnnualMinCapacityInvestment",
                          "values": mci_vals, "note": f"coherence: MinCapInv <= MaxCap-Residual ({headroom})"})
        b = build.get(c, 0.0)
        eff = "blocked (0 residual)" if cap < 1e-9 else ("no bind" if b <= cap + 1e-6 else f"cut {b:.1f}->{cap:.2f}")
        print(f"  {c:16}{resid[c]:>11.3f}{cap:>10.3f}{b:>12.3f}  {eff}")
    # non-India backstops -> AUL 0 (activity-based block; MaxCap=0 breaks Residual<=MaxCap)
    for pre in ("TRNNLI", "TRNRPO"):
        for n in NONIND:
            t = pre + n
            edits.append({"sheet": "Demand Techs", "tech": t,
                          "param": "TotalTechnologyAnnualActivityUpperLimit",
                          "op": "set_flat", "value": 0.0, "create_if_absent": True,
                          "note": "non-India backstop import disallowed (AUL=0; MaxCap=0 breaks Residual<=MaxCap)"})
    print(f"  india-internal (kept, untouched): {[c[3:] for c in internal]}")
    write_patches("B_Opt_TxCap150", {
        "scenario": "B_Opt_TxCap150", "base_scenario": "B_Optimised_VRE",
        "rule": "TotalAnnualMaxCapacity = 1.5 x ResidualCapacity_2023 for every cross-border TRN corridor (flat, all years). India-internal corridors kept at B_Opt. Non-India TRNNLI*/TRNRPO* backstops -> AUL 0 (activity-blocked; MaxCap=0 violates Residual<=MaxCap). Zero-residual corridors are blocked (pure rule, no floor).",
        "apply_vre_ceiling_layer": True, "edits": edits})


# ---------------------------------------------------------------- IndiaCosts
NONIND_C = ["BGD", "BTN", "LKA", "MDV", "NPL"]   # 3-char fuel-supply country codes
FUELS = ["COA", "COG", "GAS", "OIL", "OTH", "PET", "URN"]


def gen_indiacosts(scen="B_Opt_IndiaCosts", include_fuel=False):
    print("\n" + "=" * 70 + f"\n{scen}  (include_fuel={include_fuel})\n" + "=" * 70)
    st = load_sheet("Secondary Techs")
    fams = sorted({t[:6] for (t, p) in st if t.startswith("PWR") and t[:6] != "PWRBCK"})
    edits = []; spread = []; n_over = 0
    for param in ("CapitalCost", "FixedCost"):
        for fam in fams:
            ind_traj = {n: st[(fam + n, param)] for n in IND if (fam + n, param) in st}
            if not ind_traj:
                continue
            equal = all(
                (lambda vs: not vs or max(vs) - min(vs) <= 1e-6)([ind_traj[n][y] for n in ind_traj if ind_traj[n].get(y) is not None])
                for y in YEARS)
            ref_node = "INDNO" if "INDNO" in ind_traj else sorted(ind_traj)[0]
            ref = ind_traj[ref_node]
            if not equal:
                yr = 2030
                spr = {n: ind_traj[n].get(yr) for n in ind_traj}
                spread.append((param, fam, ref_node, spr))
            for n in NONIND:
                key = (fam + n, param)
                if key not in st:
                    continue
                vals = {str(y): round(ref[y], 6) for y in YEARS if ref.get(y) is not None}
                if not vals:
                    continue
                edits.append({"sheet": "Secondary Techs", "tech": fam + n, "param": param,
                              "values": vals, "note": f"India reference ({ref_node}) {param}"})
                n_over += 1
    print(f"  overwrote {n_over} CapitalCost/FixedCost cells across non-India nodes")
    for param, fam, rn, spr in spread:
        print(f"    UNEQUAL {param:12} {fam}: anchor {rn}={spr.get(rn)}  spread={spr}")
    n_fuel = 0
    if include_fuel:
        vc = load_sheet("VariableCost")
        for F in FUELS:
            ref = vc.get((f"MIN{F}IND", "VariableCost"))
            if not ref:
                continue
            for C in NONIND_C:
                t = f"MIN{F}{C}"
                if (t, "VariableCost") not in vc:
                    continue
                vals = {str(y): round(ref[y], 6) for y in YEARS if ref.get(y) is not None}
                edits.append({"sheet": "VariableCost", "tech": t, "param": "VariableCost",
                              "values": vals, "note": f"India ({F}) fuel price (MIN{F}IND)"})
                n_fuel += 1
        print(f"  + {n_fuel} fuel-price (MIN* VariableCost) overwrites -> India (fuel ON)")
    desc = ("CapitalCost + FixedCost of every non-India generation/storage tech set to the India "
            "reference (per family/year; INDNO anchor where India nodes disagree: PWRCOA/PWRHYD/PWRWOF). "
            "Transmission (TRN corridors) excluded. Over shared VRE-ceiling layer.")
    desc += (" FUEL ON: non-India MIN* fuel-supply VariableCost also set to the India (MIN*IND) price -- "
             "note India has cheap coal / dear gas, so this is a MIXED shift."
             if include_fuel else " Fuel/VariableCost NOT changed (include_fuel=false).")
    write_patches(scen, {
        "scenario": scen, "base_scenario": "B_Optimised_VRE",
        "description": desc, "include_fuel": include_fuel,
        "apply_vre_ceiling_layer": True, "edits": edits})


if __name__ == "__main__":
    gen_tradecap(0.30)
    gen_tradecap(0.15, export_factor=0.0)
    gen_txcap150()
    gen_indiacosts("B_Opt_IndiaCosts", include_fuel=False)
    gen_indiacosts("B_Opt_IndiaCostsFuel", include_fuel=True)
    print("\nDONE. Four patches.json generated (2 IndiaCosts variants).")
