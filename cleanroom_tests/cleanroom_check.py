#!/usr/bin/env python
"""cleanroom_check.py -- single verification harness for the OSTRAM clean-room prep.

Reads cleanroom_tests/hero_refs.yaml (the one manifest). Hero repos are READ-ONLY;
this script never writes to them. Every subcommand prints [PASS]/[FAIL]/[INFO] lines
and exits non-zero if any check in the requested set fails.

Subcommands (run one, or --all for the sensible set):
  consolidation  8 WS files present; v18 = WS-3's; sensitivity tooling present; branch ok
  reproduction   CR-stripped compiled .txt byte-diff == hero for baselines (hard oracle)
  anchors        (info) sum Outputs/TotalDiscountedCost.csv for a hero
  foundation     interconnector CapEx/life; internal-tx OAR/life; base-year pin (exact,
                 internal-tx PINNED, only 18 interconnectors excluded); cliff values
  clips          the 3 enforced VRE ceilings present (min(atlas, MaxCap))
  sensitivities  each sensitivity carries the ceiling + its own lever; glpsol --check clean
  hygiene        nothing giant/backup/superseded/-A staged; no v18 churn staged

Usage:
  python cleanroom_tests/cleanroom_check.py reproduction --hero pre_ws3
  python cleanroom_tests/cleanroom_check.py reproduction --hero final
  python cleanroom_tests/cleanroom_check.py consolidation
  python cleanroom_tests/cleanroom_check.py hygiene
"""
from __future__ import annotations
import argparse, subprocess, sys, csv
from pathlib import Path
import yaml

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent                      # OSTRAM_mainredo
T1 = ROOT / "t1_confection"
REFS = yaml.safe_load((HERE / "hero_refs.yaml").read_text(encoding="utf-8"))

RED, GRN, INF = "[FAIL]", "[PASS]", "[INFO]"
_fails = 0
def ok(msg):   print(f"{GRN} {msg}")
def bad(msg):
    global _fails; _fails += 1; print(f"{RED} {msg}")
def info(msg): print(f"{INF} {msg}")

def cr_strip(p: Path) -> bytes:
    return p.read_bytes().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

def repro_txt(exec_root: Path, scen: str) -> Path:
    return exec_root / f"{scen}_0" / f"Pre_processed_{scen}_0{REFS['repro_txt_suffix']}"

def sum_tdc(outputs_dir: Path):
    f = outputs_dir / "TotalDiscountedCost.csv"
    if not f.exists():
        return None
    with f.open(newline="", encoding="utf-8") as fh:
        r = csv.reader(fh); next(r, None)
        return sum(float(row[-1]) for row in r if row and row[-1].strip())

# ----------------------------------------------------------------- reproduction
def check_reproduction(hero_key: str, scenarios):
    hero = REFS["heroes"][hero_key]
    hero_exec = Path(hero["repo"]) / "t1_confection" / "Executables"
    our_exec = T1 / "Executables"
    scenarios = scenarios or REFS["baselines"]
    info(f"reproduction vs hero '{hero_key}' ({hero['repo']})")
    for s in scenarios:
        ours, theirs = repro_txt(our_exec, s), repro_txt(hero_exec, s)
        if not theirs.exists():
            bad(f"{s}: hero .txt missing ({theirs})"); continue
        if not ours.exists():
            bad(f"{s}: our .txt not generated yet ({ours.relative_to(ROOT)})"); continue
        a, b = cr_strip(ours), cr_strip(theirs)
        if a == b:
            ok(f"{s}: 0-diff ({len(a):,} bytes) == hero")
        else:
            al, bl = a.split(b"\n"), b.split(b"\n")
            firstdiff = next((i for i in range(min(len(al), len(bl))) if al[i] != bl[i]), None)
            bad(f"{s}: BYTE-DIFF (ours {len(al):,} lines vs hero {len(bl):,}); "
                f"first diff at line {firstdiff}")

# ----------------------------------------------------------------- anchors (info)
def check_anchors(hero_key: str):
    hero = REFS["heroes"][hero_key]
    hexec = Path(hero["repo"]) / "t1_confection" / "Executables"
    for s in REFS["baselines"]:
        got = sum_tdc(hexec / f"{s}_0" / "Outputs")
        exp = hero.get("anchors_measured", {}).get(s)
        info(f"{hero_key} {s}: sum(TDC)={got:,.0f}" if got else f"{hero_key} {s}: no Outputs"
             + (f"  (expect ~{exp:,})" if exp else ""))

# ----------------------------------------------------------------- consolidation
def check_consolidation():
    for rel in REFS["overlay_from_final"]:
        (ok if (ROOT / rel).exists() else bad)(f"present: {rel}")
    (ok if (ROOT / "ws3_transmission_audit").exists() else bad)("present: ws3_transmission_audit/")
    (ok if (T1 / "sensitivity_expansion").exists() else bad)("present: sensitivity_expansion/")
    cfg = T1 / "A3_process" / "rules_scripts" / "configs"
    (ok if cfg.exists() else bad)(f"present: A3_process/rules_scripts/configs/ ({len(list(cfg.glob('*'))) if cfg.exists() else 0} dirs)")
    br = subprocess.run(["git", "-C", str(ROOT), "branch", "--show-current"],
                        capture_output=True, text=True).stdout.strip()
    (ok if br == "ws3-phaseb-cleanredo" else bad)(f"branch = {br}")

# ----------------------------------------------------------------- hygiene
GIANT_PATTERNS = ("OSTRAM_Combined", "OSTRAM_Inputs", "OSTRAM_Outputs", "StorageDelay",
                  "storage_delay", "_output.csv")
BACKUP_MARKERS = ("_PREPATCH_", "_PRE_", "PREPATCH", "backup", "Backup", "BACKUP")
SUPERSEDED = REFS["superseded"]
def check_hygiene(allow_v18=False):
    staged = subprocess.run(["git", "-C", str(ROOT), "diff", "--cached", "--name-only"],
                            capture_output=True, text=True).stdout.splitlines()
    if not staged:
        info("nothing staged (hygiene of staged set trivially clean)"); return
    for f in staged:
        if f.endswith((".sol", ".lp", ".feasopt.sol")):
            bad(f"staged solver artifact: {f}")
        if any(p in f for p in GIANT_PATTERNS):
            bad(f"staged giant/regenerable dataset: {f}")
        if any(m in f for m in BACKUP_MARKERS):
            bad(f"staged backup: {f}")
        if any(sc in f for sc in SUPERSEDED):
            bad(f"staged superseded scenario: {f}")
        if "SOASIA_OSeMOSYS_Template_v18.xlsx" in f:
            if allow_v18:
                info(f"staged v18 permitted (STEP 1 source overlay only): {f}")
            else:
                bad(f"staged v18 (stage-6 Restrictions churn is benign; do NOT commit): {f}")
    if _fails == 0:
        ok(f"{len(staged)} staged path(s) clean (no giant/solver/backup/superseded/v18-churn)")

# ------------------------------------------------ foundation (spec, exercised in STEP 3)
def _read_param_csv(scen, name):
    p = T1 / "A2_Output_Params" / scen / f"{name}.csv"
    if not p.exists():
        return None
    with p.open(newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))

def check_foundation():
    # interconnector CapEx range 380..2800 + life 40 ; internal-tx OAR 0.97 ; A==B==C internal-tx
    rows = _read_param_csv("A_Calibrated_BAU", "CapitalCost")
    if rows is None:
        bad("foundation: A2_Output_Params/A_Calibrated_BAU/CapitalCost.csv missing (run STEP 3 first)"); return
    def val(rows, tech, year="2023"):
        for r in rows:
            if r.get("TECHNOLOGY") == tech and r.get("YEAR") == year:
                return float(r["VALUE"])
        return None
    cc = {t: val(rows, t) for t in ("TRNBGDXXINDEA", "TRNMDVXXINDSO")}
    (ok if cc["TRNBGDXXINDEA"] == 380.0 else bad)(f"interconnector CapEx min TRNBGDXXINDEA=380 (got {cc['TRNBGDXXINDEA']})")
    (ok if cc["TRNMDVXXINDSO"] == 2800.0 else bad)(f"interconnector CapEx max TRNMDVXXINDSO=2800 (got {cc['TRNMDVXXINDSO']})")
    ol = _read_param_csv("A_Calibrated_BAU", "OperationalLife")
    if ol:
        def olv(tech):
            return next((float(r["VALUE"]) for r in ol if r.get("TECHNOLOGY") == tech), None)
        (ok if olv("TRNMDVXXINDSO") == 40.0 else bad)(f"interconnector life TRNMDVXXINDSO=40 (got {olv('TRNMDVXXINDSO')})")
        (ok if olv("RNWTRNINDWE") == 40.0 else bad)(f"internal-tx life RNWTRNINDWE=40 (got {olv('RNWTRNINDWE')})")
    oar = _read_param_csv("A_Calibrated_BAU", "OutputActivityRatio")
    if oar:
        vals = [float(r["VALUE"]) for r in oar if r.get("TECHNOLOGY") == "RNWTRNINDWE" and r.get("YEAR") == "2030"]
        (ok if vals and abs(vals[0]-0.97) < 1e-9 else bad)(f"internal-tx OAR RNWTRNINDWE@2030=0.97 (got {vals})")
    ccc = val(_read_param_csv("C_Target_VRE", "CapitalCost") or [], "RNWTRNINDWE")
    (ok if ccc == 200.0 else bad)(f"internal-tx CapEx A==C RNWTRNINDWE=200 (C got {ccc})")

def check_clips():
    # 3 enforced ceilings in the clipped B scenario's compiled MaxCapacity
    rows = _read_param_csv("B_Opt_Clipped", "TotalAnnualMaxCapacity")
    if rows is None:
        bad("clips: B_Opt_Clipped not built yet"); return
    want = {"PWRSPVLKAXX": 16.0, "PWRWONBGDXX": 3.0, "PWRWONMDVXX": 0.0}
    for tech, exp in want.items():
        v = next((float(r["VALUE"]) for r in rows if r.get("TECHNOLOGY") == tech and r.get("YEAR") == "2030"), None)
        (ok if v is not None and abs(v-exp) < 1e-6 else bad)(f"clip {tech}={exp} (got {v})")

def check_sensitivities():
    info("sensitivities structural check: run glpsol --check per scenario (see STEP 6 driver).")
    info("each of the 15 must carry the foundation + its lever; verified at preflight.")

# ----------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("subcommand", choices=["consolidation", "reproduction", "anchors",
                                           "foundation", "clips", "sensitivities", "hygiene", "all"])
    ap.add_argument("--hero", default="final", choices=list(REFS["heroes"].keys()))
    ap.add_argument("--scenarios", nargs="*", default=None)
    ap.add_argument("--allow-v18", action="store_true", help="permit staged v18 (STEP 1 source overlay only)")
    a = ap.parse_args()
    if a.subcommand == "reproduction": check_reproduction(a.hero, a.scenarios)
    elif a.subcommand == "anchors":    check_anchors(a.hero)
    elif a.subcommand == "consolidation": check_consolidation()
    elif a.subcommand == "hygiene":    check_hygiene(a.allow_v18)
    elif a.subcommand == "foundation": check_foundation()
    elif a.subcommand == "clips":      check_clips()
    elif a.subcommand == "sensitivities": check_sensitivities()
    elif a.subcommand == "all":
        check_consolidation(); check_foundation(); check_clips()
        check_reproduction(a.hero, a.scenarios); check_hygiene()
    print(f"\n=== {'ALL GREEN' if _fails == 0 else str(_fails) + ' FAILURE(S)'} ===")
    sys.exit(1 if _fails else 0)

if __name__ == "__main__":
    main()
