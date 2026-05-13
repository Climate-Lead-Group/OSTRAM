"""A3_process.py
==============
Orchestrator for the A3 modification workflow.

Transforms the 4 fresh A1 outputs in `A1_Outputs/A1_Outputs_BAU/`
into the final form by chaining the sequential operations bundled in
`t1_confection/A3_process/`.

Pipeline:
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
  Stage 5    add_max_cap_investment_lid_rule  applies lid + untie

Usage:
    python A3_process.py

All behavior is controlled by the USER CONFIGURATION block below.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

T1_CONFECTION = Path(__file__).resolve().parent
A3_PROCESS_DIR = T1_CONFECTION / "A3_process"

# =============================================================================
# USER CONFIGURATION — edit these to control the run
# =============================================================================
# Folder with the 4 fresh A-O_*.xlsx (inputs).
# Path is relative to this script's folder (t1_confection/), or absolute.
INPUT_DIR = "A1_Outputs/A1_Outputs_BAU"

# Folder to write the 4 final files. None = same as INPUT_DIR (in-place overwrite).
OUTPUT_DIR = None  # or e.g. "A1_Outputs/A1_Outputs_BAU_post_A3"

# Don't auto-clean the runtime workdir (A3_process/_run_<ts>/) — useful for debugging.
KEEP_WORKDIR = False
# =============================================================================


def _resolve(p):
    """Resolve a config path: absolute as-is, relative to T1_CONFECTION."""
    p = Path(p)
    return p if p.is_absolute() else (T1_CONFECTION / p)


DEFAULT_IO = _resolve(INPUT_DIR)
DEFAULT_OUTPUT = _resolve(OUTPUT_DIR) if OUTPUT_DIR else None

INPUT_FILES = (
    "A-O_AR_Model_Base_Year.xlsx",
    "A-O_AR_Projections.xlsx",
    "A-O_Demand.xlsx",
    "A-O_Parametrization.xlsx",
)

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
def build_workdir(parent: Path, ts: str) -> dict:
    """Create runtime workdir with stage subfolders + copies of all scripts/assets.

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

    # Stage 1: scripts + asset templates
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
        "add_max_cap_investment_lid_rule.py",
        "B1b_Pre_solver_validation.py", "_xlsx_validation_core.py",
        "Config_MOMF_T1_A.yaml", "TECH_TYPES.csv",
        "fix_rnwbio_restore.py", "fix_pwrpet_clear.py", "fix_elc_pmode_revert.py",
        "6_sync_og_to_ts20.py",
        "A-O_Parametrization_REFERENCE_with_RNWBIO.xlsx",
    ):
        shutil.copy(A3_PROCESS_DIR / f, wd / f)

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


def stage_4_and_5(wd: Path, s1: Path, s3: Path, s5: Path, post_trn_cap: Path) -> None:
    banner("Stage 4 + 5 — consolidate + add_max_cap_investment_lid_rule")

    # Stage 4: consolidate the 4 final files into stage5
    shutil.copy(post_trn_cap, s5 / "A-O_Parametrization.xlsx")
    src_dir = s1 / "wvaligned_outputs_v2"
    shutil.copy(src_dir / "A-O_AR_Model_Base_Year_wvaligned_v2.xlsx",
                s5 / "A-O_AR_Model_Base_Year.xlsx")
    shutil.copy(src_dir / "A-O_AR_Projections_wvaligned_v2.xlsx",
                s5 / "A-O_AR_Projections.xlsx")
    shutil.copy(src_dir / "A-O_Demand_wvaligned_v2.xlsx",
                s5 / "A-O_Demand.xlsx")
    print(f"    Consolidated 4 files into {s5.name}/")

    # Stage 5: lid script
    run_subproc([
        PYTHON, wd / "add_max_cap_investment_lid_rule.py",
        "--input-dir", s5,
    ], cwd=wd, label="add_max_cap_investment_lid_rule.py")


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
def main() -> int:
    input_dir = DEFAULT_IO
    output_dir = DEFAULT_OUTPUT or input_dir
    workdir_base = A3_PROCESS_DIR

    if not A3_PROCESS_DIR.is_dir():
        sys.exit(f"ERROR: A3_process folder missing: {A3_PROCESS_DIR}")
    if not input_dir.is_dir():
        sys.exit(f"ERROR: input dir missing: {input_dir}")

    t_start = time.time()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    banner(f"A3 workflow run @ {ts}")
    print(f"  input-dir : {input_dir}")
    print(f"  output-dir: {output_dir}")

    # 1. Build workdir
    paths = build_workdir(workdir_base, ts)
    wd = paths["wd"]
    s1 = paths["s1"]; s1b = paths["s1b"]; s2 = paths["s2"]; s3 = paths["s3"]; s5 = paths["s5"]
    print(f"  workdir   : {wd}")

    # 2. Copy inputs into stage1
    for f in INPUT_FILES:
        src = input_dir / f
        if not src.exists():
            sys.exit(f"ERROR: input file missing: {src}")
        shutil.copy(src, s1 / f)

    # 3. Run pipeline
    stage_0_5_rnwbio(wd, s1)
    stage_1_scripts_1_to_5(s1)
    stage_1b(wd, s1, s1b)
    stage_2_and_2_5(wd, s1b, s2)
    param_for_stage4 = stage_3_fix_2(s2, s3)
    stage_4_and_5(wd, s1, s3, s5, param_for_stage4)
    stage_6_sync_og_to_ts20(wd, s1)

    # 4. Deliver
    deliver_outputs(s5, output_dir)

    # 5. Cleanup workdir
    if not KEEP_WORKDIR:
        shutil.rmtree(wd, ignore_errors=True)
        print(f"\n  Cleaned up workdir: {wd.name}")
    else:
        print(f"\n  Workdir preserved: {wd}")

    elapsed = time.time() - t_start
    banner(f"DONE in {elapsed:.1f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
