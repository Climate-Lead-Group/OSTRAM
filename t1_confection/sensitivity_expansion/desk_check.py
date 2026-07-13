"""
desk_check.py  --  OSTRAM sensitivity pre-CPLEX desk-check  (CLG)
=================================================================

A PARTIAL ENERGY/COST BALANCE of what the constraints force, computed from the
validated B_Optimised_VRE baseline. This is NOT a dispatch prediction; it bounds
what CPLEX must do, to sanity-check the runs before solving.

Run 1 (B_Opt_TradeCap50) -- the important one:
   For BGD, with imports capped at 50% of demand and domestic VRE bounded by the
   physical ceilings:
       max_imports[y]      = 0.50 * demand[y]
       min_domestic_gen[y] = demand[y] - max_imports[y]
       max_VRE_gen[y]      = (SPV_ceil*CF_spv + WON_ceil*CF_won) * 31.536   [PJ]
       forced_firm[y]      = min_domestic_gen[y] - max_VRE_gen[y]
   forced_firm is the generation BGD MUST get from coal/gas/existing-firm once
   imports are capped and VRE is ceiling-limited -- the physical core of the
   energy-security premium.

Run 2 (B_Opt_SolarHi10): solar LCOE at B_Opt CapEx vs +10%, against coal/gas
   LCOE, 2035 & 2050. Shows the marginal shift and that ceilings stay idle.

Run 3 (B_Opt_LinkFreeze): B_Opt throughput on the frozen corridors vs the
   residual-capped throughput -> import capacity removed; affected node headroom.

Outputs: reports/desk_check_report.txt and reports/desk_check_matrix.csv
"""

from __future__ import annotations
import json, csv
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
REF = SCRIPT_DIR / "reference"
REPORTS = SCRIPT_DIR / "reports"
C2A = 31.536  # PJ per GW-yr
YEARS = list(range(2023, 2051))
HILITE = [2030, 2040, 2050]


def crf(r, n):
    return r * (1 + r) ** n / ((1 + r) ** n - 1)


def lcoe_usd_mwh(capex, fixed, cf, life, r, fuel_var=0.0, iar=0.0):
    cap_fixed = (crf(r, life) * capex + fixed) / (cf * 8760.0) * 1000.0
    fuel = fuel_var * iar * 3.6 if fuel_var else 0.0
    return cap_fixed + fuel


def main():
    REPORTS.mkdir(exist_ok=True)
    bl = json.loads((REF / "b_opt_baseline.json").read_text(encoding="utf-8"))
    dem = {int(y): v for y, v in bl["bgd_demand_PJ"].items()}
    ceil = bl["vre_ceiling_GW"]
    cf_spv = bl["bgd_cf_realized_2050"]["PWRSPVBGDXX"]
    cf_won = bl["bgd_cf_realized_2050"]["PWRWONBGDXX"]
    spv_ceil = ceil["PWRSPVBGDXX"]   # 40
    won_ceil = ceil["PWRWONBGDXX"]   # 3 (atlas-enforced clip)

    L = []
    def p(s=""): L.append(s)
    p("=" * 82); p("OSTRAM SENSITIVITY DESK-CHECK (pre-CPLEX, partial balance)"); p("=" * 82)
    p(f"BGD VRE ceilings: solar {spv_ceil} GW (CF {cf_spv}), wind {won_ceil} GW (CF {cf_won})")
    p(f"max domestic VRE @ ceiling = ({spv_ceil}x{cf_spv}+{won_ceil}x{cf_won})x{C2A} "
      f"= {round((spv_ceil*cf_spv+won_ceil*cf_won)*C2A,1)} PJ (flat; capacity-bound)")

    matrix = []

    # ---------------- RUN 1 ----------------
    p(""); p("-" * 82); p("RUN 1  B_Opt_TradeCap50  --  BGD import cap 50%, backstops zeroed"); p("-" * 82)
    max_vre = (spv_ceil * cf_spv + won_ceil * cf_won) * C2A
    p(f"{'year':>5} {'demand':>9} {'maxImp50%':>10} {'minDom':>9} {'maxVRE':>8} {'FORCED_FIRM':>12}")
    ff = {}
    for y in YEARS:
        mi = 0.50 * dem[y]; md = dem[y] - mi; ffy = md - max_vre
        ff[y] = ffy
        if y in HILITE or y == 2023:
            p(f"{y:>5} {dem[y]:>9.1f} {mi:>10.1f} {md:>9.1f} {max_vre:>8.1f} {ffy:>12.1f}")
    bgd_firm_2050 = sum(bl["run2_lcoe_inputs"]["bgd_firm_prod_2050_PJ"].values())
    p("")
    p(f">> BGD FORCED-FIRM 2050 = {ff[2050]:.1f} PJ  ({ff[2050]/3.6:.1f} TWh) that MUST come from")
    p(f"   coal/gas/existing-firm once imports are capped at 50% and VRE hits its ceilings.")
    p(f"   B_Opt 2050 BGD firm generation = {bgd_firm_2050:.1f} PJ  (coal 160.6 + gas 201.6 + oil 66.8 + hydro 10.2)")
    p(f"   => additional firm required vs B_Opt ~= {ff[2050]-bgd_firm_2050:.1f} PJ (new coal/gas; MaxCap uncapped => feasible).")
    p(f"   Domestic VRE is CAPPED: solar ceiling {spv_ceil} GW is the active limit (B_Opt built 10.5 GW).")
    matrix.append(["B_Opt_TradeCap50", "forced_firm_2030_PJ", round(ff[2030], 1)])
    matrix.append(["B_Opt_TradeCap50", "forced_firm_2040_PJ", round(ff[2040], 1)])
    matrix.append(["B_Opt_TradeCap50", "forced_firm_2050_PJ", round(ff[2050], 1)])
    matrix.append(["B_Opt_TradeCap50", "extra_firm_vs_bopt_2050_PJ", round(ff[2050] - bgd_firm_2050, 1)])

    # ---------------- RUN 2 ----------------
    p(""); p("-" * 82); p("RUN 2  B_Opt_SolarHi10  --  solar CapitalCost x1.10"); p("-" * 82)
    r2 = bl["run2_lcoe_inputs"]; r = r2["discount_rate_assumed"]
    p(f"(illustrative LCOE, discount rate {r} assumed; USD/MWh)")
    p(f"{'tech/node':>18} {'2035':>9} {'2050':>9}")
    def solar_lcoe(node, yr, mult=1.0):
        s = r2[node]["solar"]; capex = s["capex_usd_kw"][str(yr)] * mult
        return lcoe_usd_mwh(capex, s["fixed_usd_kw_yr"], s["cf"], s["life"], r)
    for node in ("BGD", "INDNO"):
        p(f"{'solar '+node+' B_Opt':>18} {solar_lcoe(node,2035):>9.1f} {solar_lcoe(node,2050):>9.1f}")
        p(f"{'solar '+node+' +10%':>18} {solar_lcoe(node,2035,1.10):>9.1f} {solar_lcoe(node,2050,1.10):>9.1f}")
    cb = r2["BGD"]["coal"]; gb = r2["BGD"]["gas"]
    coal_l = lcoe_usd_mwh(cb["capex_usd_kw"], cb["fixed_usd_kw_yr"], cb["cf"], cb["life"], r, cb["fuel_var"]["2050"], cb["iar"])
    gas_l = lcoe_usd_mwh(gb["capex_usd_kw"], gb["fixed_usd_kw_yr"], gb["cf"], gb["life"], r, gb["fuel_var"]["2050"], gb["iar"])
    p(f"{'coal BGD 2050':>18} {'':>9} {coal_l:>9.1f}")
    p(f"{'gas  BGD 2050':>18} {'':>9} {gas_l:>9.1f}")
    d_bgd = solar_lcoe("BGD", 2050, 1.10) - solar_lcoe("BGD", 2050)
    p("")
    p(f">> Solar +10% CapEx raises BGD solar LCOE by {d_bgd:.1f} USD/MWh "
      f"(+{100*d_bgd/solar_lcoe('BGD',2050):.1f}%), i.e. ~{d_bgd*277778/1e6:.2f} M USD per PJ solar.")
    p(f"   Solar (even +10%: {solar_lcoe('BGD',2050,1.10):.0f}) stays FAR below BGD coal ({coal_l:.0f}) / gas ({gas_l:.0f}).")
    p(f"   => small system-cost rise, minimal substitution; VRE ceilings stay IDLE for Run 2.")
    matrix.append(["B_Opt_SolarHi10", "solarBGD_LCOE_2050", round(solar_lcoe("BGD", 2050), 1)])
    matrix.append(["B_Opt_SolarHi10", "solarBGD_LCOE_2050_+10%", round(solar_lcoe("BGD", 2050, 1.10), 1)])
    matrix.append(["B_Opt_SolarHi10", "coalBGD_LCOE_2050", round(coal_l, 1)])
    matrix.append(["B_Opt_SolarHi10", "delta_solar_LCOE_pct", round(100 * d_bgd / solar_lcoe("BGD", 2050), 1)])

    # ---------------- RUN 3 ----------------
    p(""); p("-" * 82); p("RUN 3  B_Opt_LinkFreeze  --  freeze TRNBGDXXINDEA & TRNNPLXXBGDXX at residual"); p("-" * 82)
    imp = bl["corridor_import_PJ"]; res = bl["trn_residual_GW"]
    p(f"{'corridor':>16} {'B_Opt imp 2050':>15} {'residual GW':>12} {'frozen cap PJ':>14} {'removed PJ':>11}")
    tot_removed = 0.0
    for c in ("TRNBGDXXINDEA", "TRNNPLXXBGDXX"):
        thr = imp[c]["2050"]; cap = res[c] * C2A; rem = max(thr - cap, 0.0); tot_removed += rem
        p(f"{c:>16} {thr:>15.1f} {res[c]:>12.2f} {cap:>14.1f} {rem:>11.1f}")
        matrix.append(["B_Opt_LinkFreeze", f"removed_{c}_PJ", round(rem, 1)])
    p("")
    p(f">> Import capacity removed from BGD ~= {tot_removed:.0f} PJ/yr (2050); affected node = BGDXX.")
    p(f"   Backstops are NOT frozen in Run 3 -> BGD can reroute via TRNBGDXXINDNE/TRNBTNXXBGDXX + backstops,")
    p(f"   so forced-firm pressure is milder than Run 1; BGD VRE headroom still capped (solar {spv_ceil}, wind {won_ceil} GW).")
    matrix.append(["B_Opt_LinkFreeze", "total_removed_2050_PJ", round(tot_removed, 1)])

    # ---------------- ALL: ceiling contact ----------------
    p(""); p("-" * 82); p("VRE CEILINGS the constraints could drive against"); p("-" * 82)
    p("  Already clipping baseline (atlas-enforced): PWRSPVLKAXX 16, PWRWONBGDXX 3, PWRWONMDVXX 0")
    p("  Active domestic-VRE limit under Run 1: PWRSPVBGDXX 40 GW (BGD solar)")
    p("  Pure-guard (preserve B_Opt): PWRWON INDEA/INDNE/NPLXX = 0; PWRSPVBTNXX 1.81")
    p("  India big nodes: non-binding headroom (buildout << ceiling)")

    # expected directions
    p(""); p("-" * 82); p("EXPECTED DIRECTIONS (sanity, not truth)"); p("-" * 82)
    p("  Run              | Sys cost | Coal GW | VRE GW    | Net imports | Binding")
    p("  TradeCap50       | UP sm-mod| UP(BGD) | UP->CAP   | DOWN(BGD)   | imports<=50%; BGD SPV 40")
    p("  SolarHi10        | UP small | UP small| DOWN small| shift/flat  | solar+10% vs firm; ceilings idle")
    p("  LinkFreeze       | UP mod   | UP      | redistrib | DOWN(BGD)   | frozen corridors; BGD ceilings")

    txt = "\n".join(L)
    (REPORTS / "desk_check_report.txt").write_text(txt, encoding="utf-8")
    with open(REPORTS / "desk_check_matrix.csv", "w", newline="") as f:
        w = csv.writer(f); w.writerow(["scenario", "metric", "value"]); w.writerows(matrix)
    print(txt)
    print(f"\nWrote {REPORTS/'desk_check_report.txt'} and desk_check_matrix.csv")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
