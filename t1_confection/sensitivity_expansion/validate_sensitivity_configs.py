"""
validate_sensitivity_configs.py  --  OSTRAM sensitivity pre-run validator (CLG)
==============================================================================

Validates the three sensitivity scenarios (configs + patched A-O) against the
validated B_Optimised_VRE baseline and the run specifications. Read-only.

Checks per run:
  1. COMPLETENESS   - the 4 required YAMLs exist and parse; patches.json parses.
  2. SCHEMA         - lid_rule has relaxation_schedule + family_relaxation_ceilings;
                      relax_interconnectors has mode + headroom_factor + overrides.
  3. DIFF_vs_BOPT   - every (sheet,tech,param) that changed vs B_Opt is expected
                      for that run; no unexpected edits.
  4. COHERENCE      - MaxCapInv<=MaxCap; MinCapInv<=MaxCapInv; Residual<=MaxCap;
                      AUL>=ALL, for every (tech,year) where both are set.
  5. VRE_CEILINGS   - each PWRSPV*/PWRWON* MaxCap present, constant across years,
                      >= ResidualCapacity, and == min(atlas, B_Opt MaxCap).
                      Buildout coherence reported; the 3 approved clips
                      (LKAXX SPV, BGDXX WON, MDVXX WON) are expected, any other
                      clip is a FAIL.
  6. RUN1_CAPS      - real-corridor AUL sums to 0.5*demand + export allowance;
                      backstop AUL == 0.
  7. RUN3_FREEZE    - each frozen corridor MaxCap == ResidualCapacity (not 9999).

Outputs: reports/validation_report.txt and reports/validation_matrix.csv
"""

from __future__ import annotations
import json, sys
from pathlib import Path
import pandas as pd
import yaml

SCRIPT_DIR = Path(__file__).resolve().parent
REPO = SCRIPT_DIR.parent
A1 = REPO / "A1_Outputs"
CONFIGS = REPO / "A3_process" / "rules_scripts" / "configs"
REF = SCRIPT_DIR / "reference"
REPORTS = SCRIPT_DIR / "reports"
PARAM_FILE = "A-O_Parametrization.xlsx"
BOPT = "B_Optimised_VRE"

SCEN = ["B_Opt_Clipped", "B_Opt_TradeCap30", "B_Opt_TradeCap15", "B_Opt_SolarCapexHi", "B_Opt_TxCap150", "B_Opt_IndiaCosts", "B_Opt_IndiaCostsFuel"]
YAMLS = ["lid_rule.yaml", "relax_interconnectors.yaml", "retirement_schedule.yaml", "storage_floors.yaml"]
REAL_CORR = ["TRNBGDXXINDEA", "TRNBGDXXINDNE", "TRNBTNXXBGDXX", "TRNNPLXXBGDXX"]
BACKSTOPS = ["TRNNLIBGDXX", "TRNRPOBGDXX"]
NONIND = ["BGDXX", "BTNXX", "LKAXX", "MDVXX", "NPLXX"]
COUNTRIES = {"BGD", "IND", "LKA", "NPL", "BTN", "MDV"}
NONIND_BACKSTOPS = [p + n for p in ("TRNNLI", "TRNRPO") for n in NONIND]
APPROVED_CLIPS = {"PWRSPVLKAXX", "PWRWONBGDXX", "PWRWONMDVXX"}


def _is_cross_border(t):
    return (len(t) == 13 and t.startswith("TRN") and t[3:6] in COUNTRIES
            and t[8:11] in COUNTRIES and t[3:6] != t[8:11])
YEARS = list(range(2023, 2051))

ATLAS = {"PWRSPV": {"INDEA":150,"INDNE":25,"INDNO":600,"INDSO":500,"INDWE":650,"BGDXX":40,"LKAXX":16,"NPLXX":40,"BTNXX":5,"MDVXX":1},
         "PWRWON": {"INDEA":8,"INDNE":0.5,"INDNO":285,"INDSO":445,"INDWE":410,"BGDXX":3,"LKAXX":15,"NPLXX":2,"BTNXX":0.5,"MDVXX":0}}


def load_sheet(path, sheet):
    """Return {(tech,param): {year: float|None}} for a sheet."""
    df = pd.read_excel(path, sheet_name=sheet, header=0)
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
        key = (str(t).strip(), str(p).strip())
        out[key] = {y: (None if pd.isna(r[ycols[y]]) else float(r[ycols[y]])) for y in ycols}
    return out


def is_set(v):
    return v is not None


class Report:
    def __init__(self):
        self.lines = []
        self.matrix = {}  # scenario -> {check: PASS/FAIL}

    def head(self, s):
        self.lines += ["", "=" * 78, s, "=" * 78]

    def p(self, s=""):
        self.lines.append(s)

    def result(self, scen, check, ok, detail=""):
        tag = "PASS" if ok else "FAIL"
        self.matrix.setdefault(scen, {})[check] = tag
        self.p(f"  [{tag}] {check}: {detail}")


def chk_completeness(rep, scen):
    ok = True
    det = []
    for y in YAMLS:
        f = CONFIGS / scen / y
        if not f.is_file():
            ok = False; det.append(f"MISSING {y}"); continue
        try:
            yaml.safe_load(f.read_text(encoding="utf-8"))
        except Exception as e:
            ok = False; det.append(f"UNPARSEABLE {y}: {e}")
    pj = CONFIGS / scen / "patches.json"
    if not pj.is_file():
        ok = False; det.append("MISSING patches.json")
    else:
        try:
            json.loads(pj.read_text(encoding="utf-8"))
        except Exception as e:
            ok = False; det.append(f"UNPARSEABLE patches.json: {e}")
    rep.result(scen, "1_COMPLETENESS", ok, "; ".join(det) or "4 YAMLs + patches.json present & parse")


def chk_schema(rep, scen):
    ok = True; det = []
    lid = yaml.safe_load((CONFIGS / scen / "lid_rule.yaml").read_text(encoding="utf-8"))
    for k in ("relaxation_schedule", "family_relaxation_ceilings"):
        if k not in lid:
            ok = False; det.append(f"lid_rule missing {k}")
    rlx = yaml.safe_load((CONFIGS / scen / "relax_interconnectors.yaml").read_text(encoding="utf-8"))
    for k in ("mode", "headroom_factor", "overrides"):
        if k not in rlx:
            ok = False; det.append(f"relax_interconnectors missing {k}")
    rep.result(scen, "2_SCHEMA", ok, "; ".join(det) or "lid & relax schema keys present")


def chk_diff(rep, scen, bopt_sheets, scen_sheets):
    """Every changed (sheet,tech,param) must be expected for this run."""
    unexpected = []
    changed = []
    for sheet in scen_sheets:
        b = bopt_sheets.get(sheet, {})
        s = scen_sheets[sheet]
        keys = set(b) | set(s)
        for (t, p) in keys:
            bv = b.get((t, p)); sv = s.get((t, p))
            if bv is None and sv is not None:
                # new row (e.g. backstop AUL)
                changed.append((sheet, t, p, "NEW"))
                if not _expected(scen, sheet, t, p):
                    unexpected.append((sheet, t, p, "NEW"))
                continue
            if bv is None or sv is None:
                continue
            diff = any(_neq(bv.get(y), sv.get(y)) for y in set(bv) | set(sv))
            if diff:
                changed.append((sheet, t, p, "MOD"))
                if not _expected(scen, sheet, t, p):
                    unexpected.append((sheet, t, p, "MOD"))
    ok = not unexpected
    detail = f"{len(changed)} (tech,param) changed; {len(unexpected)} unexpected"
    if unexpected:
        detail += " -> " + ", ".join(f"{s}:{t}/{p}({k})" for s, t, p, k in unexpected[:8])
    rep.result(scen, "3_DIFF_vs_BOPT", ok, detail)
    return changed


def _expected(scen, sheet, tech, param):
    gen = tech.startswith("PWRSPV") or tech.startswith("PWRWON")
    # shared VRE ceiling layer (all runs): sets MaxCap and clamps the
    # investment lids/floors down to the ceiling for coherence.
    if sheet == "Secondary Techs" and gen and param in (
            "TotalAnnualMaxCapacity", "TotalAnnualMaxCapacityInvestment",
            "TotalAnnualMinCapacityInvestment"):
        return True
    if scen == "B_Opt_Clipped":
        return False  # ceiling-only reference; no run-specific edits expected
    if scen in ("B_Opt_TradeCap30", "B_Opt_TradeCap15"):
        if param == "TotalTechnologyAnnualActivityUpperLimit" and tech in REAL_CORR + BACKSTOPS:
            return True
    if scen == "B_Opt_SolarCapexHi":
        if param == "CapitalCost" and tech.startswith("PWRSPV"):
            return True
    if scen == "B_Opt_TxCap150":
        if param in ("TotalAnnualMaxCapacity", "TotalAnnualMaxCapacityInvestment",
                     "TotalAnnualMinCapacityInvestment") and _is_cross_border(tech):
            return True
        if param == "TotalTechnologyAnnualActivityUpperLimit" and tech in NONIND_BACKSTOPS:
            return True
    if scen in ("B_Opt_IndiaCosts", "B_Opt_IndiaCostsFuel"):
        if param in ("CapitalCost", "FixedCost") and tech.startswith("PWR") and len(tech) >= 11 and tech[6:11] in NONIND:
            return True
        # IndiaCostsFuel also overwrites MIN* fuel VariableCost, but that lives in the
        # 'VariableCost' sheet which this validator does not load, so it is not diffed here.
    return False


def _neq(a, b):
    if a is None and b is None:
        return False
    if a is None or b is None:
        return True
    return abs(a - b) > 1e-6


def coherence_violations(sheets):
    """Return set of (tech, year, kind) coherence violations on Secondary Techs AND Demand Techs."""
    viol = set()
    for sheet_name in ("Secondary Techs", "Demand Techs"):
        st = sheets.get(sheet_name, {})
        for t in sorted({t for (t, p) in st}):
            res = st.get((t, "ResidualCapacity"), {})
            mc = st.get((t, "TotalAnnualMaxCapacity"), {})
            mci = st.get((t, "TotalAnnualMaxCapacityInvestment"), {})
            mni = st.get((t, "TotalAnnualMinCapacityInvestment"), {})
            aul = st.get((t, "TotalTechnologyAnnualActivityUpperLimit"), {})
            low = st.get((t, "TotalTechnologyAnnualActivityLowerLimit"), {})
            for y in YEARS:
                mcv, mciv, mniv = mc.get(y), mci.get(y), mni.get(y)
                rv, av, lv = res.get(y), aul.get(y), low.get(y)
                if is_set(mciv) and is_set(mcv) and mciv > mcv + 1e-6:
                    viol.add((t, y, "MaxCapInv>MaxCap"))
                if is_set(mniv) and is_set(mciv) and mniv > mciv + 1e-6:
                    viol.add((t, y, "MinCapInv>MaxCapInv"))
                if is_set(rv) and is_set(mcv) and rv > mcv + 1e-6:
                    viol.add((t, y, "Residual>MaxCap"))
                if is_set(av) and is_set(lv) and lv > av + 1e-6:
                    viol.add((t, y, "ALL>AUL"))
    return viol


def chk_coherence(rep, scen, sheets, bopt_viol):
    viol = coherence_violations(sheets)
    introduced = sorted(viol - bopt_viol)
    inherited = len(viol & bopt_viol)
    ok = not introduced
    detail = f"{len(introduced)} patch-introduced; {inherited} inherited from B_Opt (benign, MaxCap dominates)"
    if introduced:
        detail += " -> " + "; ".join(f"{t}@{y}:{k}" for t, y, k in introduced[:8])
    rep.result(scen, "4_COHERENCE", ok, detail)


def chk_ceilings(rep, scen, sheets, baseline):
    st = sheets["Secondary Techs"]
    problems = []; clips = []
    present = 0
    for pre in ("PWRSPV", "PWRWON"):
        for n in ATLAS[pre]:
            t = pre + n
            mc = st.get((t, "TotalAnnualMaxCapacity"))
            if mc is None:
                problems.append(f"{t}: MaxCap ABSENT"); continue
            present += 1
            vals = [mc.get(y) for y in YEARS if is_set(mc.get(y))]
            if len(vals) != len(YEARS):
                problems.append(f"{t}: MaxCap not set all years")
            if vals and (max(vals) - min(vals) > 1e-6):
                problems.append(f"{t}: MaxCap NOT constant ({min(vals)}..{max(vals)})")
            ceil = vals[0] if vals else None
            # expected == min(atlas, B_Opt maxcap)
            bmc = (baseline["spv_maxcap_2050_GW"] if pre == "PWRSPV" else baseline["won_maxcap_2050_GW"]).get(t)
            exp = min(ATLAS[pre][n], bmc if bmc is not None else 0.0)
            if ceil is not None and abs(ceil - exp) > 1e-3:
                problems.append(f"{t}: ceiling {ceil} != min(atlas,maxcap) {exp}")
            # residual coherence
            resv = [v for v in (st.get((t, "ResidualCapacity")) or {}).values() if is_set(v)]
            if resv and ceil is not None and max(resv) > ceil + 1e-6:
                problems.append(f"{t}: Residual {max(resv)}>ceiling {ceil}")
            # buildout coherence
            b = (baseline["spv_buildout_2050_GW"] if pre == "PWRSPV" else baseline["won_buildout_2050_GW"])[n]
            if ceil is not None and b > ceil + 1e-6:
                if t in APPROVED_CLIPS:
                    clips.append(f"{t}(build {round(b,1)}>ceil {ceil}, approved)")
                else:
                    problems.append(f"{t}: UNAPPROVED CLIP build {round(b,1)}>ceil {ceil}")
    ok = not problems
    detail = f"{present}/20 present; clips: {', '.join(clips) or 'none'}"
    if problems:
        detail += " | PROBLEMS: " + "; ".join(problems[:6])
    rep.result(scen, "5_VRE_CEILINGS", ok, detail)


def chk_run1(rep, scen, sheets, baseline):
    if not scen.startswith("B_Opt_TradeCap"):
        return
    try:
        frac = float(json.loads((CONFIGS / scen / "patches.json").read_text(encoding="utf-8")).get("cap_fraction"))
    except Exception:
        return
    st = sheets["Secondary Techs"]; dt = sheets.get("Demand Techs", {})
    dem = baseline["bgd_demand_PJ"]; exp = baseline["corridor_export_PJ"]
    ok = True; det = []
    for y in [2030, 2040, 2050]:
        aul_sum = sum((st.get((c, "TotalTechnologyAnnualActivityUpperLimit"), {}) or {}).get(y, 0.0) for c in REAL_CORR)
        exp_alw = sum(exp[c][str(y)] * 1.5 for c in REAL_CORR)
        target = frac * dem[str(y)] + exp_alw
        if abs(aul_sum - target) > 0.5:
            ok = False; det.append(f"{y}: AUL sum {round(aul_sum,1)} != {frac:.2f}*dem+exp {round(target,1)}")
    for c in BACKSTOPS:
        row = dt.get((c, "TotalTechnologyAnnualActivityUpperLimit"))
        if row is None:
            ok = False; det.append(f"{c}: backstop AUL row absent")
        elif any(is_set(row.get(y)) and abs(row[y]) > 1e-9 for y in YEARS):
            ok = False; det.append(f"{c}: backstop AUL != 0")
    rep.result(scen, "6_RUN1_CAPS", ok, "; ".join(det) or f"real-corridor AUL == {frac:.2f}*demand+export; backstops==0")


def chk_run3(rep, scen, sheets):
    if scen != "B_Opt_TxCap150":
        return
    st = sheets["Secondary Techs"]; dt = sheets.get("Demand Techs", {})
    ok = True; det = []
    cross = sorted({t for (t, p) in st if _is_cross_border(t) and p == "ResidualCapacity"})
    for c in cross:
        res = st.get((c, "ResidualCapacity"), {})
        mc = st.get((c, "TotalAnnualMaxCapacity"), {})
        for y in YEARS:
            exp = 1.5 * (res.get(y) or 0.0)
            if abs((mc.get(y) if mc.get(y) is not None else -99) - exp) > 1e-3:
                ok = False; det.append(f"{c}@{y}: MaxCap {mc.get(y)} != 1.5xRes {exp:.3f}"); break
        if any((mc.get(y) or 0) >= 9998 for y in YEARS):
            ok = False; det.append(f"{c}: MaxCap still 9999")
    for c in NONIND_BACKSTOPS:
        aul = dt.get((c, "TotalTechnologyAnnualActivityUpperLimit"), {})
        if not aul or not any(is_set(aul.get(y)) for y in YEARS):
            ok = False; det.append(f"{c}: non-India backstop AUL row absent")
        elif any(is_set(aul.get(y)) and abs(aul[y]) > 1e-9 for y in YEARS):
            ok = False; det.append(f"{c}: non-India backstop AUL != 0")
    rep.result(scen, "7_RUN3_TXCAP", ok,
               "; ".join(det) or f"{len(cross)} cross-border MaxCap==1.5xRes; non-India backstops AUL==0")


def main():
    REPORTS.mkdir(exist_ok=True)
    baseline = json.loads((REF / "b_opt_baseline.json").read_text(encoding="utf-8"))
    bopt_path = A1 / f"A1_Outputs_{BOPT}" / PARAM_FILE
    print("Loading B_Opt A-O ...")
    bopt_sheets = {sh: load_sheet(bopt_path, sh) for sh in ("Secondary Techs", "Demand Techs")}
    bopt_viol = coherence_violations(bopt_sheets)

    rep = Report()
    rep.head("OSTRAM SENSITIVITY CONFIG VALIDATION")
    rep.p(f"baseline objective: {baseline['objective_MUSD']} M USD   (validated B_Optimised_VRE)")
    rep.p(f"generated: {pd.Timestamp.now().isoformat(timespec='seconds')}")

    for scen in SCEN:
        rep.head(f"SCENARIO: {scen}")
        ao = A1 / f"A1_Outputs_{scen}" / PARAM_FILE
        chk_completeness(rep, scen)
        chk_schema(rep, scen)
        if not ao.is_file():
            rep.result(scen, "PATCHED_AO", False, f"missing {ao} (run apply_patches.py first)")
            continue
        sheets = {sh: load_sheet(ao, sh) for sh in ("Secondary Techs", "Demand Techs")}
        chk_diff(rep, scen, bopt_sheets, sheets)
        chk_coherence(rep, scen, sheets, bopt_viol)
        chk_ceilings(rep, scen, sheets, baseline)
        chk_run1(rep, scen, sheets, baseline)
        chk_run3(rep, scen, sheets)

    # matrix
    checks = ["1_COMPLETENESS", "2_SCHEMA", "3_DIFF_vs_BOPT", "4_COHERENCE",
              "5_VRE_CEILINGS", "6_RUN1_CAPS", "7_RUN3_TXCAP"]
    rep.head("VALIDATION MATRIX (run x check)")
    rep.p(f"{'scenario':22} " + " ".join(f"{c.split('_')[0]:>3}" for c in checks))
    rows = []
    for scen in SCEN:
        cells = [rep.matrix.get(scen, {}).get(c, "-") for c in checks]
        rep.p(f"{scen:22} " + " ".join(f"{('P' if v=='PASS' else 'F' if v=='FAIL' else '-'):>3}" for v in cells))
        rows.append([scen] + cells)

    txt = "\n".join(rep.lines)
    (REPORTS / "validation_report.txt").write_text(txt, encoding="utf-8")
    pd.DataFrame(rows, columns=["scenario"] + checks).to_csv(REPORTS / "validation_matrix.csv", index=False)
    print(txt)
    print(f"\nWrote {REPORTS/'validation_report.txt'} and validation_matrix.csv")
    allpass = all(v in ("PASS",) for scen in SCEN for k, v in rep.matrix.get(scen, {}).items())
    return 0 if allpass else 2


if __name__ == "__main__":
    sys.exit(main())
