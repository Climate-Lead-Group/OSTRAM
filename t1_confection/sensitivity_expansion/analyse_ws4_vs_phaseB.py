"""
analyse_ws4_vs_phaseB.py  --  OSTRAM behavioural cross-check (CLG / OSTRAM)
===========================================================================
Compares the 15 solved scenarios on the WS-3/WS-4 foundation against the
pre-WS-3 Phase-B oracle (sensitivity_expansion/PHASE_B_METHODOLOGY_AND_RESULTS.md,
Appendix A). The WS-3/WS-4 foundation shifted every absolute number, so we check
the *behaviour* -- same SIGNS and same coarse RANKING -- NOT identical magnitudes.

Reads each scenario's solved Executables/<s>_0/Outputs/:
  syscost  = Sum(TotalDiscountedCost)              [M USD]
  co2_50   = Sum(AnnualEmissions, YEAR=2050)        [Mt]
  coal/solar/wind 2050 GW from TotalCapacityAnnual
  bgd_dom  = Sum(ProductionByTechnologyAnnual PWR generators at BGDXX, 2050) [PJ] (self-sufficiency proxy)
  backstop = Sum(ProductionByTechnologyAnnual PWRBCK*)                        [PJ] (must be 0)

Emits: the metric table, the system-cost ranking, and per-lever BGD-self-sufficiency
+ CO2 direction vs B_Opt_Clipped, each flagged OK/MISMATCH against the oracle sign.
Run from t1_confection after all 15 have solved (del_files:False keeps the Outputs).
"""
import csv
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent          # t1_confection
EXEC = REPO / "Executables"
SCEN = ["A_Calibrated_BAU", "A_Calibrated_BAU_Clipped", "B_Optimised_VRE", "B_Opt_Clipped",
        "C_Target_VRE", "C_Target_VRE_Clipped", "B_Opt_TradeCap15", "B_Opt_TxCap150",
        "B_Opt_DirContractual", "B_Opt_SolarCapexHi", "B_Opt_SolarCapex130",
        "B_Opt_SolarCapexSpike", "B_Opt_IndiaCosts", "B_Opt_IndiaCostsFuel", "B_Opt_DirBidir"]

# Pre-WS-3 oracle (Appendix A) -- for SIGN/RANK direction only, never magnitude.
# BGD domestic-gen (TWh) direction vs B_Opt_Clipped, and system-cost order.
ORACLE_BGDDOM = {"B_Opt_Clipped": 182.6, "B_Opt_TradeCap15": 421.1, "B_Opt_TxCap150": 459.1,
                 "B_Opt_DirContractual": 155.9, "B_Opt_IndiaCosts": 249.2, "B_Opt_DirBidir": 182.6}
ORACLE_COST_ORDER = ["B_Opt_IndiaCosts", "B_Opt_IndiaCostsFuel", "B_Optimised_VRE",
                     "B_Opt_Clipped", "B_Opt_DirBidir", "B_Opt_DirContractual",
                     "B_Opt_SolarCapexHi", "B_Opt_TradeCap15", "B_Opt_TxCap150",
                     "C_Target_VRE_Clipped", "C_Target_VRE", "A_Calibrated_BAU",
                     "A_Calibrated_BAU_Clipped"]  # SolarCapex130/Spike are new (post-oracle)

FAM = {"coal": ["PWRCOA"], "solar": ["PWRSPV"], "wind": ["PWRWON", "PWRWOF"]}


def _sum(path, keep=None):
    if not path.is_file():
        return None
    t = 0.0
    for r in csv.DictReader(open(path, newline="")):
        if keep and not keep(r):
            continue
        try:
            t += float(r["VALUE"])
        except (ValueError, KeyError):
            pass
    return t


def _cap2050(path, prefixes):
    return _sum(path, lambda r: r["YEAR"] == "2050" and any(r["TECHNOLOGY"].startswith(p) for p in prefixes))


def _is_bgd_gen(r):
    t = r["TECHNOLOGY"]
    return t.startswith("PWR") and t.endswith("BGDXX") and not t.startswith(("PWRTRN", "PWRBCK"))


def collect():
    rows = {}
    for s in SCEN:
        od = EXEC / f"{s}_0" / "Outputs"
        tdc = od / "TotalDiscountedCost.csv"
        if not tdc.is_file():
            rows[s] = None
            continue
        rows[s] = {
            "syscost": _sum(tdc),
            "co2": _sum(od / "AnnualEmissions.csv", lambda r: r["YEAR"] == "2050"),
            "coal": _cap2050(od / "TotalCapacityAnnual.csv", FAM["coal"]),
            "solar": _cap2050(od / "TotalCapacityAnnual.csv", FAM["solar"]),
            "wind": _cap2050(od / "TotalCapacityAnnual.csv", FAM["wind"]),
            "bgd": _sum(od / "ProductionByTechnologyAnnual.csv", lambda r: _is_bgd_gen(r) and r["YEAR"] == "2050"),
            "bck": _sum(od / "ProductionByTechnologyAnnual.csv", lambda r: r["TECHNOLOGY"].startswith("PWRBCK")),
        }
    return rows


def main():
    rows = collect()
    bc = rows.get("B_Opt_Clipped")
    bc_cost = bc["syscost"] if bc else None
    print(f"{'scenario':26} {'syscost':>13} {'d_vs_Clip':>11} {'CO2_50':>9} {'coalGW':>8} {'solarGW':>8} {'windGW':>8} {'BGDdomPJ':>9} {'bckstop':>9}")
    for s in SCEN:
        m = rows[s]
        if m is None:
            print(f"{s:26} {'-- NO OUTPUTS --':>13}")
            continue
        d = m["syscost"] - bc_cost if bc_cost else float("nan")
        print(f"{s:26} {m['syscost']:13,.0f} {d:>+11,.0f} {m['co2'] or 0:9.1f} {m['coal'] or 0:8.1f} {m['solar'] or 0:8.1f} {m['wind'] or 0:8.1f} {m['bgd'] or 0:9.1f} {m['bck'] or 0:9.4f}")

    print("\n--- SYSTEM-COST RANKING (cheapest->dearest) vs oracle order ---")
    mine = sorted([s for s in SCEN if rows[s]], key=lambda s: rows[s]["syscost"])
    ora_pos = {s: i for i, s in enumerate([s for s in ORACLE_COST_ORDER if s in set(mine)])}
    for i, s in enumerate(mine):
        op = ora_pos.get(s)
        note = "" if op is None else ("" if op == len([x for x in mine[:i] if x in ora_pos]) else f"  (oracle pos {op+1})")
        newf = "  [new post-oracle]" if s not in ora_pos else ""
        print(f"  {i+1:2}. {s:26} {rows[s]['syscost']:13,.0f}{newf}{note}")

    print("\n--- BGD domestic-gen + CO2 DIRECTION vs B_Opt_Clipped (sign vs oracle) ---")
    bc_bgd = bc["bgd"] if bc else None
    for s in ["B_Opt_TradeCap15", "B_Opt_TxCap150", "B_Opt_DirContractual", "B_Opt_IndiaCosts", "B_Opt_DirBidir"]:
        if rows.get(s) and rows[s]["bgd"] is not None and bc_bgd is not None:
            d = rows[s]["bgd"] - bc_bgd
            mysign = "UP" if d > 1 else ("DOWN" if d < -1 else "~SAME")
            od_ = ORACLE_BGDDOM.get(s, 0) - ORACLE_BGDDOM.get("B_Opt_Clipped", 0)
            osign = "UP" if od_ > 1 else ("DOWN" if od_ < -1 else "~SAME")
            print(f"  {s:26} mine dBGD={d:+9.1f} ({mysign:5})  oracle {osign:5}  {'OK' if mysign == osign else 'MISMATCH'}")

    # Neutrality identities
    def eq(a, b):
        return rows.get(a) and rows.get(b) and abs(rows[a]["syscost"] - rows[b]["syscost"]) < 1.0
    print("\n--- Neutralities ---")
    print(f"  IndiaCosts == IndiaCostsFuel : {'OK' if eq('B_Opt_IndiaCosts','B_Opt_IndiaCostsFuel') else 'DIFFER'}")
    print(f"  DirBidir   == B_Opt_Clipped  : {'OK' if eq('B_Opt_DirBidir','B_Opt_Clipped') else 'DIFFER'}")


if __name__ == "__main__":
    main()
