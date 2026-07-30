"""A3_process.py
==============
Orchestrator for the A3 modification workflow, multi-scenario aware.

Transforms the 4 fresh A1 outputs in `A1_Outputs/A1_Outputs_<scenario>/` into
their final form by chaining the sequential operations bundled in
`t1_confection/A3_process/`. The source snapshot is ALWAYS the BAU one
(`_post_a2_snapshot_BAU`); other scenarios diverge from BAU only through
  OSTRAM_Scenario_Inputs.xlsx overrides, optional inherited restrictions
from previous runs, and the rules_script assigned to the scenario.

Pipeline:
  Stage 0    materialize scenario template  v18 -> flat v17-shaped per scenario
  Stage 1    scripts 1-5                    AO/WV alignment pipeline, with
                                            maintained AO decision overlay
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
  Late WS-4  apply_base_year_pin            exact audited 2023-2026 PWR/MIN
                                            keys on the three canonical roots
  Stage 6    persist_run_restrictions       write CHANGES.json into a
                                            disposable scenario-state copy

Usage:
    python A3_process.py                         # runs BAU (default)
    python A3_process.py --scenario NDC          # runs NDC scenario
    python A3_process.py --keep-workdir          # debug: preserve intermediates

OSTRAM_Scenario_Inputs.xlsx is required. Each run materializes a
scenario-specific working template from that maintained authority.
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

try:
    from t1_confection import a3_orchestrator as _orchestrator
except ModuleNotFoundError as error:
    if error.name != "t1_confection":
        raise
    import a3_orchestrator as _orchestrator

T1_CONFECTION = Path(__file__).resolve().parent
A3_PROCESS_DIR = T1_CONFECTION / "A3_process"
RULES_SCRIPTS_DIR = A3_PROCESS_DIR / "rules_scripts"
SOASIA_V18 = A3_PROCESS_DIR / "OSTRAM_Scenario_Inputs.xlsx"
PIN_ROOT_SCENARIOS = _orchestrator.PWR_MIN_PIN_ROOT_SCENARIOS

# =============================================================================
# USER CONFIGURATION — defaults; can be overridden via CLI
# =============================================================================
DEFAULT_SCENARIO = "BAU"

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
             f"OSTRAM scenario-input Control sheet.",
    )
    p.add_argument(
        "--soasia", default=None, type=Path,
        help=f"Path to scenario inputs (default: {SOASIA_V18}). Pass an alternative "
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
def _read_script_yaml_name(script_path: Path) -> str | None:
    """Inspect a rules_script source for its `YAML_FILE_NAME` constant.

    Returns the YAML filename if declared at module level, else None
    (scripts like the lid_rule use in-script config and have no YAML).
    """
    import re
    try:
        text = script_path.read_text(encoding="utf-8")
    except Exception:
        return None
    m = re.search(r'^YAML_FILE_NAME\s*=\s*["\']([^"\']+)["\']', text, re.MULTILINE)
    return m.group(1) if m else None


def _resolve_script_yaml(script_name: str, scenario: str) -> Path | None:
    """Locate the YAML config to use for `script_name` under `scenario`.

    Resolution order:
      1. rules_scripts/configs/<scenario>/<YAML_FILE_NAME>   (per-scenario override)
      2. rules_scripts/<YAML_FILE_NAME>                       (default next-to-script)
      3. None — script has no YAML or no file present.
    """
    src = RULES_SCRIPTS_DIR / script_name
    yaml_name = _read_script_yaml_name(src)
    if not yaml_name:
        return None
    scenario_yaml = RULES_SCRIPTS_DIR / "configs" / scenario / yaml_name
    if scenario_yaml.is_file():
        return scenario_yaml
    default_yaml = RULES_SCRIPTS_DIR / yaml_name
    if default_yaml.is_file():
        return default_yaml
    return None


def build_workdir(
    parent: Path,
    ts: str,
    rules_scripts: list[str],
    scenario: str,
) -> dict:
    """Create runtime workdir with stage subfolders + copies of all scripts/assets.

    Each name in `rules_scripts` is copied from `rules_scripts/` into
    `wd/rules_scripts/` so stage 5 can invoke it. Each script's YAML config
    (when declared via `YAML_FILE_NAME` at module level) is also staged into
    `wd/rules_scripts/`, resolved from `configs/<scenario>/` when present,
    otherwise from the default location next to the script.

    Empty list -> stage 5 is skipped.

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

    # Stage 1: scripts + asset templates. Script 1 and the decision overlay
    # read the materialized per-scenario template through OSTRAM_TEMPLATE_PATH.
    for f in ("1_merge_timeslices_into_WV.py", "2_extract_ao_extensions.py",
              "apply_ao_extension_decisions.py",
              "3_update_ao_from_extensions.py", "4_apply_manual_fixes.py",
              "5_propagate_timeslice_fabric.py",
              "OSTRAM_Timeslice_Inputs.xlsx"):
        shutil.copy(A3_PROCESS_DIR / f, s1 / f)

    # Stage 2: patch_ao_c2a + TECH_TYPES
    shutil.copy(A3_PROCESS_DIR / "patch_ao_c2a.py", s2)
    shutil.copy(A3_PROCESS_DIR / "TECH_TYPES.csv", s2)

    # Stage 3: FIX_2 scripts + shared v18 authority loader
    for f in ("fix_trn_residuals.py", "clear_stale_unbinding_caps.py",
              "cap_trn_to_residual.py", "interconnector_authority.py"):
        shutil.copy(A3_PROCESS_DIR / f, s3)

    # Workdir-level scripts (run from `wd`, operate on subdirs via --input args)
    for f in (
        "A0_insert_reserve_margin.py",
        "add_max_capacity_investment_rule_OLD_8ee8056.py",
        "add_max_capacity_investment_rule_NEW_2be1616.py",
        "B1b_Pre_solver_validation.py", "_xlsx_validation_core.py",
        "Config_MOMF_T1_A.yaml", "TECH_TYPES.csv",
        "fix_pwrpet_clear.py", "fix_elc_pmode_revert.py",
        "6_sync_og_to_ts20.py",
    ):
        shutil.copy(A3_PROCESS_DIR / f, wd / f)

    # Per-scenario rules_scripts chain (each lives under rules_scripts/).
    # Each script + its YAML are copied into wd/rules_scripts/ so stage 5 can
    # invoke them with cwd=wd. TECH_TYPES.csv is already copied above and the
    # scripts resolve it via ../TECH_TYPES.csv -> wd/TECH_TYPES.csv. YAML is
    # resolved from rules_scripts/configs/<scenario>/ when present, otherwise
    # from the default next to the source script.
    if rules_scripts:
        rs_wd = wd / "rules_scripts"
        rs_wd.mkdir()
        available = {p.name for p in RULES_SCRIPTS_DIR.glob("*.py")}
        for script in rules_scripts:
            src = RULES_SCRIPTS_DIR / script
            if not src.is_file():
                raise FileNotFoundError(
                    f"rules_script '{script}' not found at {src}. "
                    f"Available: {sorted(available)}"
                )
            shutil.copy(src, rs_wd / script)
            yaml_path = _resolve_script_yaml(script, scenario)
            if yaml_path is not None:
                yaml_name = yaml_path.name
                shutil.copy(yaml_path, rs_wd / yaml_name)

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
    run_subproc(
        [
            PYTHON,
            s1 / "apply_ao_extension_decisions.py",
            "--extensions",
            s1 / "OSTRAM_AO_Extensions.xlsx",
        ],
        cwd=s1,
        label="apply_ao_extension_decisions.py",
    )
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

    authority_path_raw = os.environ.get("OSTRAM_TEMPLATE_PATH")
    if not authority_path_raw:
        raise RuntimeError(
            "OSTRAM_TEMPLATE_PATH is required for RC Authority V1"
        )
    authority_path = Path(authority_path_raw).resolve()
    if not authority_path.is_file():
        raise FileNotFoundError(
            f"RC Authority V1 materialized workbook not found: {authority_path}"
        )

    # Move Stage 2.5 output into stage3/
    shutil.copy(s2 / "A-O_Parametrization_c2a_patched.xlsx",
                s3 / "A-O_Parametrization_c2a_patched.xlsx")

    # fix_trn_residuals
    run_subproc([
        PYTHON, s3 / "fix_trn_residuals.py",
        "--input", "A-O_Parametrization_c2a_patched.xlsx",
        "--output", "A-O_Parametrization_c2a_patched_FIXED.xlsx",
        "--authority", authority_path,
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


def stage_5_rules_scripts(
    wd: Path,
    s5: Path,
    rules_scripts: list[str],
) -> None:
    """Stage 5: invoke each rules_script in the scenario's chain against stage5/.

    The chain order is whatever the Control sheet declares (left-to-right CSV).
    Each script edits A-O_Parametrization.xlsx in place; subsequent scripts see
    prior edits. Skipped when the chain is empty (a scenario can choose to only
    inherit restrictions and apply no rule of its own).
    """
    if not rules_scripts:
        banner("Stage 5 — SKIPPED (no rules_scripts for this scenario)")
        return
    banner(f"Stage 5 — rules_scripts chain ({len(rules_scripts)} script(s))")
    for i, script in enumerate(rules_scripts, start=1):
        rs_path = wd / "rules_scripts" / script
        if not rs_path.is_file():
            sys.exit(f"ERROR: rules_script not staged at {rs_path}")
        step(f"[{i}/{len(rules_scripts)}] {script}")
        run_subproc(
            [PYTHON, rs_path, "--input-dir", s5],
            cwd=wd, label=script,
        )


def stage_ws3_interconnector_costs(
    s5: Path,
    soasia: Path,
    materialized_template: Path | None,
) -> None:
    """WS-3: wire v18 Interconnector_Params cost columns into the model.

    Interconnector CapitalCost/FixedCost historically came from the OG_csvs base
    (distance-computed legacy values) and the sourced Interconnector_Params sheet
    was never consumed. This stage makes that sheet the source of truth: it
    applies Interconnector_Params -> Secondary Techs (CapitalCost, FixedCost) and
    Fixed Horizon Parameters (OperationalLife), reading the per-scenario
    materialized template so any scenario override flows through. Residuals/caps
    (owned by fix_trn_residuals / relax_interconnectors) and losses are untouched.
    Skipped in legacy mode (no v18 template).
    """
    if not soasia.is_file():
        return
    script = RULES_SCRIPTS_DIR / "apply_interconnector_costs.py"
    if not script.is_file():
        sys.exit(f"ERROR: apply_interconnector_costs.py not found at {script}")
    banner("WS-3 — apply v18 Interconnector_Params costs (CapitalCost / FixedCost / OperationalLife)")
    cmd = [PYTHON, script, "--input-dir", s5, "--skip-backup"]
    if materialized_template is not None:
        cmd += ["--template", materialized_template]
    run_subproc(cmd, label="apply_interconnector_costs.py")


def stage_ws3_internal_transmission(s5: Path) -> None:
    """WS-3 D5: calibrate the six INTERNAL (intra-node) transmission families.

    A2_AddTx injects RNWTRN/RNWNLI/RNWRPO (RE) and PWRTRN/TRNNLI/TRNRPO (non-RE)
    flat (CapEx 100, FOM 4, ResidualCapacity 5, life 20); the Stage-1 template
    merge then rewrites their OperationalLife to 50/20. This late stage makes
    Config_country_codes.yaml + a desk-checked per-node residuals file the source
    of truth, writing on the final A-O_Parametrization.xlsx:
      * Demand Techs -> per-node ResidualCapacity (RNWTRN/PWRTRN; NLI/RPO = 0),
        RE CapitalCost/FixedCost = base x re_capex_multiplier, non-RE = base;
      * Fixed Horizon Parameters -> OperationalLife = 40 (all six families).
    Runs after stage 5 and the interconnector-cost stage, so nothing downstream
    clobbers it. Interconnectors (13-char TRN*****), DSPTRN, generators, storage
    and losses are all left untouched. Uniform across nodes (intra-node tx is
    accounting; the study is about interties).
    """
    script = RULES_SCRIPTS_DIR / "apply_internal_transmission.py"
    config = T1_CONFECTION / "Config_country_codes.yaml"
    residuals = RULES_SCRIPTS_DIR / "internal_tx_residuals.csv"
    if not (script.is_file() and config.is_file() and residuals.is_file()):
        print("    [SKIP] internal-transmission stage: script/config/residuals missing")
        return
    banner("WS-3 — calibrate internal transmission (per-node ResidualCapacity / RE CapEx / OperationalLife=40)")
    run_subproc([
        PYTHON, script, "--input-dir", s5, "--skip-backup",
        "--config", config, "--residuals", residuals,
    ], label="apply_internal_transmission.py")


def stage_ws3_internal_tx_losses(s5: Path) -> None:
    """WS-4: give the six internal transmission families a non-zero loss.

    A2 injects internal-tx output activity = 1.0 (0% loss). This stage sets
    OutputActivityRatio = 1 - internal_transmission.transmission_loss (default
    0.03 -> 0.97) on the internal families' Output rows in the A-O_AR files'
    'Demand Techs' sheet, matching how the interconnectors carry per-corridor
    losses. Runs after the internal-tx cost/residual stage; interconnectors,
    DSPTRN, generators and storage untouched.
    """
    script = RULES_SCRIPTS_DIR / "apply_internal_tx_losses.py"
    config = T1_CONFECTION / "Config_country_codes.yaml"
    if not (script.is_file() and config.is_file()):
        print("    [SKIP] internal-tx losses stage: script/config missing")
        return
    banner("WS-4 — internal transmission losses (OutputActivityRatio = 1 - loss)")
    run_subproc([
        PYTHON, script, "--input-dir", s5, "--skip-backup", "--config", config,
    ], label="apply_internal_tx_losses.py")


def stage_ws4_pwr_min_pin(s5: Path, scenario: str) -> None:
    """Restore only the audited 2023-2026 PWR/MIN keys on canonical roots.

    This late A3 stage consumes a version-controlled static allowlist.  It
    neither reads solver output nor emits a generic ``*_CHANGES.json`` that
    Stage 6 could persist as a competing Restrictions authority.
    """
    if scenario not in PIN_ROOT_SCENARIOS:
        raise ValueError(f"unsupported PWR/MIN pin scenario: {scenario!r}")
    script = RULES_SCRIPTS_DIR / "apply_base_year_pin.py"
    rules_csv = RULES_SCRIPTS_DIR / "pwr_min_2023_2026_pin.csv"
    missing = [path for path in (script, rules_csv) if not path.is_file()]
    if missing:
        raise FileNotFoundError(
            "PWR/MIN pin production asset missing: "
            + ", ".join(str(path) for path in missing)
        )
    banner(
        "WS-4 — restore audited non-Maldives 2023-2026 PWR/MIN calibration"
    )
    run_subproc(
        [
            PYTHON,
            script,
            "--input-dir",
            s5,
            "--scenario",
            scenario,
            "--rules-csv",
            rules_csv,
            "--skip-backup",
        ],
        label="apply_base_year_pin.py",
    )


def stage_6_persist_restrictions(
    s5: Path,
    soasia: Path,
    scenario: str,
    rules_scripts: list[str],
) -> None:
    """Persist rule change logs into the supplied disposable state workbook.

    Each script in the chain writes its change log as a sibling of the input
    dir (e.g. `<s5.name>_PRE_<TAG>_<ts>_CHANGES.json` in s5.parent). We collect
    ALL of them and persist as a single clear-and-write so the Restrictions
    rows for this scenario reflect the full chain output. Rows belonging to
    other scenarios stay untouched. The orchestrator always supplies a
    workdir-local copy; this helper must never receive the maintained input.
    """
    if not rules_scripts:
        return
    if not soasia.is_file():
        return
    banner("Stage 6 — persist run restrictions to disposable scenario state")
    search_dir = s5.parent
    candidates = sorted(
        search_dir.glob(f"{s5.name}_PRE_*_CHANGES.json"),
        key=lambda p: p.stat().st_mtime,
    )
    if not candidates:
        # Fallback: any *_CHANGES.json in the same dir (in case scripts adopt
        # a different naming convention).
        candidates = sorted(
            search_dir.glob("*_CHANGES.json"),
            key=lambda p: p.stat().st_mtime,
        )
    if not candidates:
        print(f"    [WARN] no *_CHANGES.json found in {search_dir}; nothing persisted")
        return
    # When the chain has N scripts, we expect ~N CHANGES.json files (one per
    # script that touched cells). They're sorted by mtime ascending so the
    # chain order is preserved in the Restrictions audit trail.
    print(f"    Found {len(candidates)} change-log file(s):")
    for p in candidates:
        print(f"      - {p.name}")
    sys.path.insert(0, str(A3_PROCESS_DIR))
    try:
        import _scenarios
    finally:
        sys.path.pop(0)
    n = _scenarios.persist_run_restrictions(soasia, scenario, candidates)
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
) -> tuple[str, list[str], list[str]]:
    """Resolve (scenario, rules_scripts, inherit_from) from CLI + Control.

    CLI flags take precedence. The `--rules-script` arg is parsed as a CSV
    list so a chain can be passed from the command line too (e.g.
    `--rules-script "set_retirement_schedule.py, set_min_capacity_floors.py"`).
    An explicit empty string means "skip stage 5".

    OSTRAM_Scenario_Inputs.xlsx is the required scenario authority.
    """
    scenario = args.scenario

    def _parse_csv(val: str) -> list[str]:
        return [s.strip() for s in val.replace("\n", ",").split(",") if s.strip()]

    if not soasia.is_file():
        sys.exit(f"ERROR: required scenario-input workbook not found at {soasia}")

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
    if args.rules_script is not None:
        rules_scripts = _parse_csv(args.rules_script)
    else:
        rules_scripts = list(cfg.rules_scripts)
    if args.inherit_from is not None:
        inherit_from = [s.strip() for s in args.inherit_from.split(",") if s.strip()]
    else:
        inherit_from = list(cfg.inherit_restrictions_from)
    return scenario, rules_scripts, inherit_from


def _materialize_scenario_template(
    soasia: Path,
    scenario: str,
    output_path: Path,
) -> None:
    """Load the existing scenario helper only when materialization is needed."""
    sys.path.insert(0, str(A3_PROCESS_DIR))
    try:
        import _scenarios
    finally:
        sys.path.pop(0)
    _scenarios.materialize_scenario_template(soasia, scenario, output_path)


def _orchestration_paths() -> _orchestrator.A3Paths:
    """Expose script-anchored A3 paths as one immutable plan input."""
    return _orchestrator.A3Paths(
        t1_confection=T1_CONFECTION,
        process_dir=A3_PROCESS_DIR,
        default_soasia=SOASIA_V18,
    )


def _orchestration_dependencies() -> _orchestrator.A3Dependencies:
    """Bind existing helpers to the isolated orchestration effect seams."""
    return _orchestrator.A3Dependencies(
        resolve_scenario_config=_resolve_scenario_config,
        resolve_path=_resolve,
        build_workdir=build_workdir,
        materialize_scenario_template=_materialize_scenario_template,
        stage_1_scripts_1_to_5=stage_1_scripts_1_to_5,
        stage_1b=stage_1b,
        stage_2_and_2_5=stage_2_and_2_5,
        stage_3_fix_2=stage_3_fix_2,
        stage_4_consolidate=stage_4_consolidate,
        stage_4_5_apply_inherited_restrictions=(
            stage_4_5_apply_inherited_restrictions
        ),
        stage_5_rules_scripts=stage_5_rules_scripts,
        stage_ws3_interconnector_costs=stage_ws3_interconnector_costs,
        stage_ws3_internal_transmission=stage_ws3_internal_transmission,
        stage_ws3_internal_tx_losses=stage_ws3_internal_tx_losses,
        stage_ws4_pwr_min_pin=stage_ws4_pwr_min_pin,
        stage_6_sync_og_to_ts20=stage_6_sync_og_to_ts20,
        stage_6_persist_restrictions=stage_6_persist_restrictions,
        deliver_outputs=deliver_outputs,
        remove_tree=shutil.rmtree,
        copy_tree=shutil.copytree,
        copy_file=shutil.copy,
        environment=os.environ,
        clock=time.time,
        timestamp_now=datetime.now,
        banner=banner,
        emit=print,
    )


def main() -> int:
    return _orchestrator.orchestrate_a3(
        parse_cli_args(),
        _orchestration_paths(),
        _orchestration_dependencies(),
        INPUT_FILES,
    )


if __name__ == "__main__":
    sys.exit(main())
