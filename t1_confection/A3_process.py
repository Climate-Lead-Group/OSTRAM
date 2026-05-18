"""A3_process.py
==============
Orchestrator for the A3 modification workflow, multi-scenario aware.

Transforms the 4 fresh A1 outputs in `A1_Outputs/A1_Outputs_<scenario>/` into
their final form by chaining the sequential operations bundled in
`t1_confection/A3_process/`. The source snapshot is ALWAYS the BAU one
(`_post_a2_snapshot_BAU`); other scenarios diverge from BAU only through
SOASIA_OSeMOSYS_Template_v18.xlsx overrides, optional inherited restrictions
from previous runs, and the rules_script assigned to the scenario.

Pipeline:
  Stage 0    materialize scenario template  v18 -> flat v17-shaped per scenario
  Stage 0.5  fix_rnwbio_restore             input fix (RNWBIO rows)
  Stage 1    scripts 1-5                    AO/WV alignment pipeline
  Stage 1b   A0_insert_reserve_margin       adds System Parameters sheet
             add_max OLD (8ee8056)          fills 9999 / zeroes
             add_max NEW (2be1616)          flips Projection.Mode
             fix_elc_pmode_revert           reverts 20 ELC*01 PM cells
             B1b_Pre_solver_validation      V2 fix (PWRHYDLKAXX)
  Stage 2    patch_ao_c2a                   adds CapacityToActivityUnit
  Stage 2.5  fix_pwrpet_clear               clears 336 PWRPETBGDXX cells
  Stage 3    fix_trn_residuals              + clear_stale_unbinding_caps
                                            + cap_trn_to_residual
  Stage 4    consolidate                    move 4 files to stage5/
  Stage 4.5  apply_inherited_restrictions   write Restrictions rows from
                                            inherit_restrictions_from scenarios
  Stage 5    <rules_script>                 applies the scenario's rules
  Stage 6    persist_run_restrictions       write CHANGES.json into v18

Usage:
    python A3_process.py                         # runs BAU (default)
    python A3_process.py --scenario NDC          # runs NDC scenario
    python A3_process.py --keep-workdir          # debug: preserve intermediates

Without SOASIA_OSeMOSYS_Template_v18.xlsx present, the script falls back to
the legacy single-scenario BAU behavior.
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

T1_CONFECTION = Path(__file__).resolve().parent
A3_PROCESS_DIR = T1_CONFECTION / "A3_process"
RULES_SCRIPTS_DIR = A3_PROCESS_DIR / "rules_scripts"
SOASIA_V18 = A3_PROCESS_DIR / "SOASIA_OSeMOSYS_Template_v18.xlsx"

# =============================================================================
# USER CONFIGURATION — defaults; can be overridden via CLI
# =============================================================================
DEFAULT_SCENARIO = "BAU"

# When SOASIA v18 is absent and we fall back to legacy single-scenario mode,
# this is the rules_script invoked at stage 5.
LEGACY_RULES_SCRIPT = "add_max_cap_investment_lid_rule.py"

# Don't auto-clean the runtime workdir (A3_process/_run_<ts>/) — useful for debugging.
KEEP_WORKDIR_DEFAULT = False
# =============================================================================


def _resolve(p):
    """Resolve a config path: absolute as-is, relative to T1_CONFECTION."""
    p = Path(p)
    return p if p.is_absolute() else (T1_CONFECTION / p)


INPUT_FILES = (
    "A-O_AR_Model_Base_Year.xlsx",
    "A-O_AR_Projections.xlsx",
    "A-O_Demand.xlsx",
    "A-O_Parametrization.xlsx",
)


def parse_cli_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="A3_process.py",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument(
        "--scenario", default=DEFAULT_SCENARIO,
        help=f"Scenario name (default: {DEFAULT_SCENARIO}). Must be present in "
             f"SOASIA v18 Control sheet when v18 exists.",
    )
    p.add_argument(
        "--soasia", default=None, type=Path,
        help=f"Path to SOASIA v18 (default: {SOASIA_V18}). Pass an alternative "
             f"to test scenarios without touching the canonical file.",
    )
    p.add_argument(
        "--rules-script", default=None,
        help="Override the rules_script declared in Control. Empty string skips stage 5.",
    )
    p.add_argument(
        "--inherit-from", default=None,
        help="Override inherit_restrictions_from from Control. CSV list, e.g. 'BAU,NDC'.",
    )
    p.add_argument(
        "--input-dir", default=None, type=Path,
        help="Override the A1_Outputs/A1_Outputs_<scenario> input dir.",
    )
    p.add_argument(
        "--output-dir", default=None, type=Path,
        help="Override the delivery dir (default: in-place over input-dir).",
    )
    p.add_argument(
        "--keep-workdir", action="store_true", default=KEEP_WORKDIR_DEFAULT,
        help="Preserve the runtime _run_<ts>/ folder for debugging.",
    )
    return p.parse_args()

PYTHON = sys.executable


# ---------------------------------------------------------------------------
# CLI helpers
# ---------------------------------------------------------------------------
def banner(msg: str) -> None:
    bar = "=" * 78
    print(f"\n{bar}\n{msg}\n{bar}")


def step(label: str) -> None:
    print(f"\n>>> {label}")


def run_subproc(cmd: list, cwd: Path | None = None, label: str | None = None) -> str:
    """Run a subprocess; abort on non-zero. Returns stdout."""
    if label:
        step(label)
    cmd_str = " ".join(str(c) for c in cmd)
    print(f"    $ {cmd_str}" + (f"  (cwd={cwd.name})" if cwd else ""))
    res = subprocess.run(
        [str(c) for c in cmd], cwd=str(cwd) if cwd else None,
        capture_output=True, text=True,
    )
    if res.returncode != 0:
        if res.stdout:
            print("--- stdout ---")
            print(res.stdout[-3000:])
        if res.stderr:
            print("--- stderr ---")
            print(res.stderr[-3000:])
        sys.exit(f"FAILED: {label or cmd_str}")
    # Surface a few last lines of stdout for visibility
    tail = [l for l in (res.stdout or "").strip().splitlines() if l.strip()]
    for line in tail[-3:]:
        print(f"    {line}")
    return res.stdout or ""


# ---------------------------------------------------------------------------
# Workdir setup
# ---------------------------------------------------------------------------
def build_workdir(parent: Path, ts: str, rules_script: str | None) -> dict:
    """Create runtime workdir with stage subfolders + copies of all scripts/assets.

    If rules_script is given, copy it from rules_scripts/ into wd so stage 5 can
    run it. None or empty -> stage 5 is skipped.

    Returns a dict of {label: Path} for downstream use.
    """
    wd = parent / f"_run_{ts}"
    if wd.exists():
        shutil.rmtree(wd)
    wd.mkdir(parents=True)
    s1 = wd / "stage1";   s1.mkdir()
    s1b = wd / "stage1b"; s1b.mkdir()
    s2 = wd / "stage2";   s2.mkdir()
    s3 = wd / "stage3";   s3.mkdir()
    s5 = wd / "stage5";   s5.mkdir()

    # Stage 1: scripts + asset templates. SOASIA_OSeMOSYS_Template_v17.xlsx is
    # kept as a fallback for standalone Spyder runs of script 1; the orchestrator
    # overrides it via OSTRAM_TEMPLATE_PATH env var pointing at the materialized
    # per-scenario template.
    for f in ("1_merge_timeslices_into_WV.py", "2_extract_ao_extensions.py",
              "3_update_ao_from_extensions.py", "4_apply_manual_fixes.py",
              "5_propagate_timeslice_fabric.py",
              "SOASIA_OSeMOSYS_Template_v17.xlsx",
              "OSTRAM_Timeslice_Outputs.xlsx",
              "OSTRAM_AO_Extensions_FILLED.xlsx"):
        shutil.copy(A3_PROCESS_DIR / f, s1 / f)

    # Stage 2: patch_ao_c2a + TECH_TYPES
    shutil.copy(A3_PROCESS_DIR / "patch_ao_c2a.py", s2)
    shutil.copy(A3_PROCESS_DIR / "TECH_TYPES.csv", s2)

    # Stage 3: FIX_2 scripts + NATY reference
    for f in ("fix_trn_residuals.py", "clear_stale_unbinding_caps.py",
              "cap_trn_to_residual.py", "A-O_Parametrization_NATY.xlsx"):
        shutil.copy(A3_PROCESS_DIR / f, s3)

    # Workdir-level scripts (run from `wd`, operate on subdirs via --input args)
    for f in (
        "A0_insert_reserve_margin.py",
        "add_max_capacity_investment_rule_OLD_8ee8056.py",
        "add_max_capacity_investment_rule_NEW_2be1616.py",
        "B1b_Pre_solver_validation.py", "_xlsx_validation_core.py",
        "Config_MOMF_T1_A.yaml", "TECH_TYPES.csv",
        "fix_rnwbio_restore.py", "fix_pwrpet_clear.py", "fix_elc_pmode_revert.py",
        "6_sync_og_to_ts20.py",
        "A-O_Parametrization_REFERENCE_with_RNWBIO.xlsx",
    ):
        shutil.copy(A3_PROCESS_DIR / f, wd / f)

    # Per-scenario rules_script (lives under rules_scripts/). Copied to wd so
    # stage 5 can invoke it directly. TECH_TYPES.csv is already copied above
    # and the script resolves it via ../TECH_TYPES.csv -> wd/TECH_TYPES.csv.
    # We replicate the rules_scripts/ folder relationship by placing the
    # script inside wd/rules_scripts/, so its `script_dir.parent` lookup
    # finds wd/TECH_TYPES.csv just like in the real layout.
    if rules_script:
        src = RULES_SCRIPTS_DIR / rules_script
        if not src.is_file():
            raise FileNotFoundError(
                f"rules_script '{rules_script}' not found at {src}. "
                f"Available: {[p.name for p in RULES_SCRIPTS_DIR.glob('*.py')]}"
            )
        rs_wd = wd / "rules_scripts"
        rs_wd.mkdir()
        shutil.copy(src, rs_wd / rules_script)

    return {
        "wd": wd, "s1": s1, "s1b": s1b, "s2": s2, "s3": s3, "s5": s5,
    }


# ---------------------------------------------------------------------------
# Pipeline stages
# ---------------------------------------------------------------------------
def stage_0_5_rnwbio(wd: Path, s1: Path) -> None:
    banner("Stage 0.5 — fix_rnwbio_restore")
    run_subproc([
        PYTHON, wd / "fix_rnwbio_restore.py",
        "--input", s1 / "A-O_Parametrization.xlsx",
        "--source", wd / "A-O_Parametrization_REFERENCE_with_RNWBIO.xlsx",
    ], label="fix_rnwbio_restore.py")


def stage_1_scripts_1_to_5(s1: Path) -> None:
    banner("Stage 1 — scripts 1-5 (AO/WV alignment pipeline)")
    run_subproc([PYTHON, s1 / "1_merge_timeslices_into_WV.py"], cwd=s1, label="1_merge_timeslices_into_WV.py")
    run_subproc([PYTHON, s1 / "2_extract_ao_extensions.py"], cwd=s1, label="2_extract_ao_extensions.py")
    # Script 3 reads OSTRAM_AO_Extensions.xlsx — we have to overwrite it with the FILLED version
    shutil.copy(s1 / "OSTRAM_AO_Extensions_FILLED.xlsx", s1 / "OSTRAM_AO_Extensions.xlsx")
    print("    (OSTRAM_AO_Extensions.xlsx <- OSTRAM_AO_Extensions_FILLED.xlsx)")
    run_subproc([PYTHON, s1 / "3_update_ao_from_extensions.py"], cwd=s1, label="3_update_ao_from_extensions.py")
    run_subproc([PYTHON, s1 / "4_apply_manual_fixes.py"], cwd=s1, label="4_apply_manual_fixes.py")
    run_subproc([PYTHON, s1 / "5_propagate_timeslice_fabric.py"], cwd=s1, label="5_propagate_timeslice_fabric.py")


def stage_1b(wd: Path, s1: Path, s1b: Path) -> None:
    banner("Stage 1b — A0 + add_max OLD/NEW + ELC revert + B1b")

    # Move stage1 outputs into stage1b/
    src_dir = s1 / "wvaligned_outputs_v2"
    shutil.copy(src_dir / "A-O_Parametrization_wvaligned_v2_ts20.xlsx",
                s1b / "A-O_Parametrization.xlsx")
    shutil.copy(src_dir / "A-O_AR_Model_Base_Year_wvaligned_v2.xlsx",
                s1b / "A-O_AR_Model_Base_Year.xlsx")
    shutil.copy(src_dir / "A-O_AR_Projections_wvaligned_v2.xlsx",
                s1b / "A-O_AR_Projections.xlsx")
    shutil.copy(src_dir / "A-O_Demand_wvaligned_v2.xlsx",
                s1b / "A-O_Demand.xlsx")
    print("    (Stage 1 outputs copied into stage1b/)")

    # 1) A0
    run_subproc([
        PYTHON, wd / "A0_insert_reserve_margin.py",
        "--input", s1b / "A-O_Parametrization.xlsx",
    ], label="A0_insert_reserve_margin.py")

    # 2) add_max OLD (commit 8ee8056) — cell value changes
    run_subproc([
        PYTHON, wd / "add_max_capacity_investment_rule_OLD_8ee8056.py",
        "--input-dir", s1b,
    ], cwd=wd, label="add_max_capacity_investment_rule (OLD 8ee8056)")

    # 3) add_max NEW (commit 2be1616) — Projection.Mode flips
    run_subproc([
        PYTHON, wd / "add_max_capacity_investment_rule_NEW_2be1616.py",
        "--input-dir", s1b,
    ], cwd=wd, label="add_max_capacity_investment_rule (NEW 2be1616)")

    # 4) fix_elc_pmode_revert — manual ELC*01 revert
    run_subproc([
        PYTHON, wd / "fix_elc_pmode_revert.py",
        "--input", s1b / "A-O_Parametrization.xlsx",
    ], label="fix_elc_pmode_revert.py")

    # 5) B1b validation (V2 fix on PWRHYDLKAXX)
    run_subproc([
        PYTHON, wd / "B1b_Pre_solver_validation.py",
        "--xlsx", s1b / "A-O_Parametrization.xlsx",
        "--auto-fix-all",
    ], cwd=wd, label="B1b_Pre_solver_validation.py")


def stage_2_and_2_5(wd: Path, s1b: Path, s2: Path) -> None:
    banner("Stage 2 + 2.5 — patch_ao_c2a + fix_pwrpet_clear")

    # Move Stage 1b output into stage2/ as the input to patch_ao_c2a
    shutil.copy(s1b / "A-O_Parametrization.xlsx", s2 / "A-O_Parametrization_ORIGINAL.xlsx")

    # patch_ao_c2a (output: A-O_Parametrization_c2a_patched.xlsx in s2)
    run_subproc([
        PYTHON, s2 / "patch_ao_c2a.py",
        "--src", "A-O_Parametrization_ORIGINAL.xlsx",
        "--tax", "TECH_TYPES.csv",
        "--out", "A-O_Parametrization_c2a_patched.xlsx",
    ], cwd=s2, label="patch_ao_c2a.py")

    # Fallback: when there's nothing to patch (target techs aren't present),
    # patch_ao_c2a doesn't write the output file. Copy the original so
    # downstream stages have something to consume.
    patched_xlsx = s2 / "A-O_Parametrization_c2a_patched.xlsx"
    if not patched_xlsx.exists():
        shutil.copy(s2 / "A-O_Parametrization_ORIGINAL.xlsx", patched_xlsx)
        print("    [FALLBACK] patch_ao_c2a produced no output; copied ORIGINAL "
              "to *_c2a_patched.xlsx so pipeline can continue.")

    # fix_pwrpet_clear — the manual edit Luis did
    run_subproc([
        PYTHON, wd / "fix_pwrpet_clear.py",
        "--input", s2 / "A-O_Parametrization_c2a_patched.xlsx",
    ], label="fix_pwrpet_clear.py")


def stage_3_fix_2(s2: Path, s3: Path) -> Path:
    banner("Stage 3 — FIX_2 pipeline (fix_trn + clear_stale + cap_trn)")

    # Move Stage 2.5 output into stage3/
    shutil.copy(s2 / "A-O_Parametrization_c2a_patched.xlsx",
                s3 / "A-O_Parametrization_c2a_patched.xlsx")

    # fix_trn_residuals
    run_subproc([
        PYTHON, s3 / "fix_trn_residuals.py",
        "--input", "A-O_Parametrization_c2a_patched.xlsx",
        "--output", "A-O_Parametrization_c2a_patched_FIXED.xlsx",
        "--reference", "A-O_Parametrization_NATY.xlsx",
        "--diff-csv", "diff_log.csv",
        "--diff-md", "diff_log.md",
        "--mode", "min",
        "--cutoff-year", "2023",
    ], cwd=s3, label="fix_trn_residuals.py")

    # clear_stale_unbinding_caps
    run_subproc([
        PYTHON, s3 / "clear_stale_unbinding_caps.py",
        "--input", "A-O_Parametrization_c2a_patched_FIXED.xlsx",
    ], cwd=s3, label="clear_stale_unbinding_caps.py")

    # Find the auto-timestamped POST_CAP_RESET file
    post_cap_reset = sorted(
        s3.glob("A-O_Parametrization_c2a_patched_FIXED_POST_CAP_RESET_*.xlsx"),
        key=lambda p: p.stat().st_mtime, reverse=True,
    )[0]
    print(f"    POST_CAP_RESET: {post_cap_reset.name}")

    # cap_trn_to_residual
    run_subproc([
        PYTHON, s3 / "cap_trn_to_residual.py",
        "--input", post_cap_reset.name,
    ], cwd=s3, label="cap_trn_to_residual.py")

    post_trn_cap = sorted(
        s3.glob("A-O_Parametrization_c2a_patched_FIXED_POST_CAP_RESET_*_POST_TRN_CAP_*.xlsx"),
        key=lambda p: p.stat().st_mtime, reverse=True,
    )[0]
    print(f"    POST_TRN_CAP:   {post_trn_cap.name}")
    return post_trn_cap


def stage_4_consolidate(s1: Path, s3: Path, s5: Path, post_trn_cap: Path) -> None:
    """Stage 4: consolidate the 4 final files into stage5/. Separate from stage 5
    so stage 4.5 (inherit restrictions) can run between them."""
    banner("Stage 4 — consolidate 4 final files into stage5/")
    shutil.copy(post_trn_cap, s5 / "A-O_Parametrization.xlsx")
    src_dir = s1 / "wvaligned_outputs_v2"
    shutil.copy(src_dir / "A-O_AR_Model_Base_Year_wvaligned_v2.xlsx",
                s5 / "A-O_AR_Model_Base_Year.xlsx")
    shutil.copy(src_dir / "A-O_AR_Projections_wvaligned_v2.xlsx",
                s5 / "A-O_AR_Projections.xlsx")
    shutil.copy(src_dir / "A-O_Demand_wvaligned_v2.xlsx",
                s5 / "A-O_Demand.xlsx")
    print(f"    Consolidated 4 files into {s5.name}/")


def stage_4_5_apply_inherited_restrictions(
    s5: Path,
    soasia: Path,
    inherit_from: list[str],
) -> None:
    """Stage 4.5: write inherited Restrictions rows from `inherit_from` scenarios
    into stage5/A-O_Parametrization.xlsx. Skipped when inherit_from is empty.

    Applied AFTER stage 4 (where the file is consolidated into stage5/) and
    BEFORE stage 5 (where this scenario's rules_script runs on top). Stages 1b
    earlier in the pipeline may have written 9999 placeholders into the same
    MaxCapInv cells — overwriting them here is the intended behavior.
    """
    if not inherit_from:
        return
    banner(
        f"Stage 4.5 — apply inherited restrictions from {', '.join(inherit_from)}"
    )
    sys.path.insert(0, str(A3_PROCESS_DIR))
    try:
        import _scenarios
    finally:
        sys.path.pop(0)

    restrictions = _scenarios.read_restrictions(soasia, inherit_from)
    if not restrictions:
        print("    (no rows matched; nothing applied)")
        return
    target = s5 / "A-O_Parametrization.xlsx"
    written = _scenarios.apply_restrictions(target, restrictions)
    print(f"    Wrote {written} inherited restriction cell(s) into {target.name}")


def stage_5_rules_script(
    wd: Path,
    s5: Path,
    rules_script: str | None,
) -> None:
    """Stage 5: invoke the scenario's rules_script against stage5/. Skipped if
    rules_script is None or empty (allowed: a scenario can choose to only
    inherit restrictions and apply no rule of its own)."""
    if not rules_script:
        banner("Stage 5 — SKIPPED (no rules_script for this scenario)")
        return
    banner(f"Stage 5 — {rules_script}")
    rs_path = wd / "rules_scripts" / rules_script
    if not rs_path.is_file():
        sys.exit(f"ERROR: rules_script not staged at {rs_path}")
    run_subproc(
        [PYTHON, rs_path, "--input-dir", s5],
        cwd=wd, label=rules_script,
    )


def stage_6_persist_restrictions(
    s5: Path,
    soasia: Path,
    scenario: str,
    rules_script: str | None,
) -> None:
    """Stage 6: persist the rules_script's CHANGES.json into v18.Restrictions.

    The rules_script writes its change log as `<input-dir>_PRE_LID_<ts>_CHANGES.json`
    next to (sibling of) the input dir, not inside it. So we look in s5.parent
    for `<s5.name>_PRE_*_CHANGES.json`, falling back to any *_CHANGES.json in
    that folder if the naming convention changes. Existing rows for `scenario`
    are replaced (clear-and-write).
    """
    if not rules_script:
        return
    if not soasia.is_file():
        return
    banner("Stage 6 — persist run restrictions to SOASIA v18 Restrictions sheet")
    search_dir = s5.parent
    candidates = sorted(
        search_dir.glob(f"{s5.name}_PRE_*_CHANGES.json"),
        key=lambda p: p.stat().st_mtime, reverse=True,
    )
    if not candidates:
        # Fallback: any *_CHANGES.json in the same dir (in case future
        # rules_scripts adopt a different naming convention).
        candidates = sorted(
            search_dir.glob("*_CHANGES.json"),
            key=lambda p: p.stat().st_mtime, reverse=True,
        )
    if not candidates:
        print(f"    [WARN] no *_CHANGES.json found in {search_dir}; nothing persisted")
        return
    changes = candidates
    changes_json = changes[0]
    sys.path.insert(0, str(A3_PROCESS_DIR))
    try:
        import _scenarios
    finally:
        sys.path.pop(0)
    n = _scenarios.persist_run_restrictions(soasia, scenario, changes_json)
    print(f"    Persisted {n} restriction row(s) to {soasia.name}::Restrictions"
          f" (scenario={scenario})")


def stage_6_sync_og_to_ts20(wd: Path, s1: Path) -> None:
    """Propagate the 20-ts fabric from WV down to OG_csvs_inputs and the YAML,
    so the next A1 run produces consistent 20-ts CSVs and B1_Compiler stops
    aborting on the YAML-vs-xlsx timeslice mismatch."""
    banner("Stage 6 — sync OG_csvs_inputs + YAML to 20-ts fabric")
    wv_file = s1 / "SOASIA_OSeMOSYS_WV.xlsx"
    og_csvs_dir = T1_CONFECTION / "OG_csvs_inputs"
    yaml_file = T1_CONFECTION / "Config_MOMF_T1_A.yaml"
    if not wv_file.is_file():
        print(f"    [SKIP] WV file not found at {wv_file}; sync stage skipped.")
        return
    run_subproc([
        PYTHON, wd / "6_sync_og_to_ts20.py",
        "--wv", wv_file,
        "--og-csvs-dir", og_csvs_dir,
        "--yaml", yaml_file,
    ], cwd=wd, label="6_sync_og_to_ts20.py")


def deliver_outputs(s5: Path, output_dir: Path) -> None:
    banner(f"Delivering 4 final files to {output_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)
    for f in INPUT_FILES:
        src = s5 / f
        dst = output_dir / f
        shutil.copy(src, dst)
        print(f"    {f}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def _resolve_scenario_config(
    args: argparse.Namespace,
    soasia: Path,
) -> tuple[str, str | None, list[str]]:
    """Resolve (scenario, rules_script, inherit_from) from CLI + Control.

    CLI flags take precedence. When SOASIA v18 is absent, fall back to legacy
    BAU-only behavior with LEGACY_RULES_SCRIPT and no inheritance.
    """
    scenario = args.scenario

    if not soasia.is_file():
        if scenario != DEFAULT_SCENARIO:
            sys.exit(
                f"ERROR: SOASIA v18 not found at {soasia} and scenario "
                f"'{scenario}' != '{DEFAULT_SCENARIO}'. Build v18 first via "
                f"_build_v18_from_v17.py."
            )
        rules_script = (
            args.rules_script if args.rules_script is not None
            else LEGACY_RULES_SCRIPT
        )
        return scenario, rules_script or None, []

    sys.path.insert(0, str(A3_PROCESS_DIR))
    try:
        import _scenarios
    finally:
        sys.path.pop(0)
    configs = _scenarios.read_control_sheet(soasia)
    cfg = next((c for c in configs if c.scenario == scenario), None)
    if cfg is None:
        names = [c.scenario for c in configs]
        sys.exit(
            f"ERROR: scenario '{scenario}' not in Control sheet of "
            f"{soasia.name}. Available: {names}"
        )
    rules_script = (
        args.rules_script if args.rules_script is not None
        else (cfg.rules_script or None)
    )
    # Empty string explicitly means "skip rules_script"; treat "" -> None
    if rules_script == "":
        rules_script = None
    if args.inherit_from is not None:
        inherit_from = [s.strip() for s in args.inherit_from.split(",") if s.strip()]
    else:
        inherit_from = list(cfg.inherit_restrictions_from)
    return scenario, rules_script, inherit_from


def main() -> int:
    args = parse_cli_args()

    if not A3_PROCESS_DIR.is_dir():
        sys.exit(f"ERROR: A3_process folder missing: {A3_PROCESS_DIR}")

    soasia = args.soasia if args.soasia is not None else SOASIA_V18
    scenario, rules_script, inherit_from = _resolve_scenario_config(args, soasia)

    # Resolve input/output dirs. Default: A1_Outputs/A1_Outputs_<scenario>.
    if args.input_dir is not None:
        input_dir = _resolve(args.input_dir)
    else:
        input_dir = T1_CONFECTION / "A1_Outputs" / f"A1_Outputs_{scenario}"
    output_dir = _resolve(args.output_dir) if args.output_dir is not None else input_dir
    workdir_base = A3_PROCESS_DIR

    # Snapshot fuente: SIEMPRE BAU. All scenarios diverge from BAU only via
    # SOASIA v18 overrides + inherited restrictions + scenario rules_script.
    snapshot_dir = T1_CONFECTION / "A1_Outputs" / "_post_a2_snapshot_BAU"
    if not snapshot_dir.is_dir():
        sys.exit(
            f"ERROR: snapshot post-A2 not found: {snapshot_dir}\n"
            f"       Run A1 + A2 (for BAU) first; A2 creates the snapshot."
        )

    t_start = time.time()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    banner(f"A3 workflow run @ {ts}")
    print(f"  scenario          : {scenario}")
    print(f"  input-dir         : {input_dir}")
    print(f"  output-dir        : {output_dir}")
    print(f"  snapshot (source) : {snapshot_dir}")
    print(f"  SOASIA v18        : {soasia if soasia.is_file() else '(legacy mode, v18 absent)'}")
    print(f"  rules_script      : {rules_script or '(none)'}")
    print(f"  inherit_from      : {inherit_from or '(none)'}")

    # 0. Restore input_dir from snapshot (clean canonical post-A2 state)
    if input_dir.exists():
        shutil.rmtree(input_dir)
    shutil.copytree(snapshot_dir, input_dir)
    print(f"  -> {input_dir.name} restored from {snapshot_dir.name}")

    # 1. Build workdir (stages incl. the chosen rules_script staged into wd/rules_scripts/)
    paths = build_workdir(workdir_base, ts, rules_script)
    wd = paths["wd"]
    s1 = paths["s1"]; s1b = paths["s1b"]; s2 = paths["s2"]; s3 = paths["s3"]; s5 = paths["s5"]
    print(f"  workdir           : {wd}")

    # 2. Stage 0 — materialize per-scenario template + expose it via env var so
    #    1_merge_timeslices_into_WV.py picks it up instead of the raw v17.
    materialized_template = None
    if soasia.is_file():
        materialized_template = wd / f"_materialized_{scenario}.xlsx"
        banner(f"Stage 0 — materialize scenario template for '{scenario}'")
        sys.path.insert(0, str(A3_PROCESS_DIR))
        try:
            import _scenarios
        finally:
            sys.path.pop(0)
        _scenarios.materialize_scenario_template(soasia, scenario, materialized_template)
        os.environ["OSTRAM_TEMPLATE_PATH"] = str(materialized_template)
        print(f"    materialized -> {materialized_template.name}")
        print(f"    OSTRAM_TEMPLATE_PATH set; stage 1 will read it instead of v17")

    # 3. Copy inputs into stage1
    for f in INPUT_FILES:
        src = input_dir / f
        if not src.exists():
            sys.exit(f"ERROR: input file missing: {src}")
        shutil.copy(src, s1 / f)

    # 4. Pipeline stages
    stage_0_5_rnwbio(wd, s1)
    stage_1_scripts_1_to_5(s1)
    stage_1b(wd, s1, s1b)
    stage_2_and_2_5(wd, s1b, s2)
    param_for_stage4 = stage_3_fix_2(s2, s3)
    stage_4_consolidate(s1, s3, s5, param_for_stage4)
    stage_4_5_apply_inherited_restrictions(s5, soasia, inherit_from)
    stage_5_rules_script(wd, s5, rules_script)
    stage_6_sync_og_to_ts20(wd, s1)
    stage_6_persist_restrictions(s5, soasia, scenario, rules_script)

    # 5. Deliver
    deliver_outputs(s5, output_dir)

    # 6. Cleanup workdir
    if not args.keep_workdir:
        shutil.rmtree(wd, ignore_errors=True)
        print(f"\n  Cleaned up workdir: {wd.name}")
    else:
        print(f"\n  Workdir preserved: {wd}")
    # Clear env var so other code in the same process is not affected
    os.environ.pop("OSTRAM_TEMPLATE_PATH", None)

    elapsed = time.time() - t_start
    banner(f"DONE in {elapsed:.1f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
