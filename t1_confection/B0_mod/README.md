# B0_mod — Self-contained B0 modification workflow

This folder bundles every script and asset needed to reproduce Luis's A3
modification workflow that transforms 4 fresh A1 outputs into their final
form. The orchestrator at `t1_confection/B0_mod.py` runs the full pipeline.

## Quick start

```bash
# From t1_confection/
python B0_mod.py                              # in-place on A1_Outputs/A1_Outputs_BAU/
python B0_mod.py --verify                     # also compare against A1_Outputs_Luis/
python B0_mod.py --keep-workdir               # keep intermediates for inspection
python B0_mod.py --input-dir <dir> --output-dir <dir>
```

## What gets transformed

The pipeline reads 4 files from `--input-dir`:
- `A-O_AR_Model_Base_Year.xlsx`
- `A-O_AR_Projections.xlsx`
- `A-O_Demand.xlsx`
- `A-O_Parametrization.xlsx`

…and writes the post-pipeline versions of the same 4 files to `--output-dir`
(default: same as input — in-place).

## Pipeline stages

| Stage | Action | Script(s) |
|---|---|---|
| 0.5 | Restore 2 RNWBIO rows in VariableCost (lost in `f1aad25` pull) | `fix_rnwbio_restore.py` + `A-O_Parametrization_REFERENCE_with_RNWBIO.xlsx` |
| 1 | `test_a3_mod_v2` pipeline (timeslice merge, AO extensions, manual fixes, ts20 fabric) | `1_*.py` … `5_*.py` |
| 1b | Add `System Parameters` sheet (ReserveMargin=1.15) | `A0_insert_reserve_margin.py` |
| 1b | Fill EMPTY/0 with 9999 in MaxCapInv (commit `8ee8056` behavior) | `add_max_capacity_investment_rule_OLD_8ee8056.py` |
| 1b | Flip Projection.Mode EMPTY → User defined (commit `2be1616` behavior) | `add_max_capacity_investment_rule_NEW_2be1616.py` |
| 1b | Revert Projection.Mode of 10 ELC*01 lockout rows back to EMPTY | `fix_elc_pmode_revert.py` |
| 1b | V2 fix on PWRHYDLKAXX TotalAnnualMaxCapacity (residual+ΣMin constraint) | `B1b_Pre_solver_validation.py` |
| 2 | Add CapacityToActivityUnit=31.536 rows for GENERATION+STORAGE techs | `patch_ao_c2a.py` |
| 2.5 | Manual edit: clear PWRPETBGDXX CapacityFactor (336 cells) | `fix_pwrpet_clear.py` |
| 3 | TRN ResidualCapacity split with NATY reference | `fix_trn_residuals.py` |
| 3 | Reset stale unbinding caps (PWRSPV/WON/HYD/SHP) | `clear_stale_unbinding_caps.py` |
| 3 | Cap TRN MaxCapacity to ResidualCapacity | `cap_trn_to_residual.py` |
| 4 | Consolidate the 4 files into the output folder | (orchestrator) |
| 5 | Apply MaxCapacityInvestment lid + V1 untie rule | `add_max_cap_investment_lid_rule.py` |

## Files in this folder

### Pipeline scripts (executed by the orchestrator)
- `1_merge_timeslices_into_WV.py` … `5_propagate_timeslice_fabric.py`
  (stage1 scripts; WORK_DIR auto-detects to script folder)
- `patch_ao_c2a.py`
- `fix_trn_residuals.py`, `clear_stale_unbinding_caps.py`, `cap_trn_to_residual.py`
- `add_max_capacity_investment_rule_OLD_8ee8056.py`
  (snapshot of the pre-modification version Luis used at 07:23 on 2026-04-30)
- `add_max_capacity_investment_rule_NEW_2be1616.py`
  (snapshot of the version after commit `2be1616`)
- `add_max_cap_investment_lid_rule.py`
- `B1b_Pre_solver_validation.py` + `_xlsx_validation_core.py`
- `A0_insert_reserve_margin.py`
- `fix_rnwbio_restore.py`, `fix_pwrpet_clear.py`, `fix_elc_pmode_revert.py`
- `compare_xlsx.py` (utility)

### Asset templates / fixtures
- `SOASIA_OSeMOSYS_Template_v17.xlsx` (script 1 input)
- `OSTRAM_Timeslice_Outputs.xlsx` (script 1 input)
- `OSTRAM_AO_Extensions_FILLED.xlsx` (human-curated input to script 3)
- `A-O_Parametrization_NATY.xlsx` (`fix_trn_residuals.py --reference`)
- `A-O_Parametrization_REFERENCE_with_RNWBIO.xlsx` (`fix_rnwbio_restore.py --source`)
- `TECH_TYPES.csv` (used by `patch_ao_c2a.py` and the lid script)
- `Config_MOMF_T1_A.yaml` (used by B1b for `base_year`)

### Runtime artifacts (auto-cleaned unless `--keep-workdir`)
- `_run_<timestamp>/` — created at runtime, contains `stage1/`, `stage1b/`,
  `stage2/`, `stage3/`, `stage5/` subfolders with intermediates.

## Why the two add_max_capacity_investment_rule versions?

Luis ran the script twice on 2026-04-30:
1. **07:23**: commit `8ee8056` — treated explicit `0` as EMPTY → filled with 9999. Wrote 1708 ALLOWED + 4192 ZEROED + 0 Projection.Mode flips.
2. **10:55**: commit `2be1616` — preserves explicit `0`s, but adds the Projection.Mode flip behavior. Idempotent on values (0 changes), 335 PM flips.

Running ONLY the current version on a fresh input misses the value changes from the OLD version (the new version preserves what the OLD would have replaced). Both must run sequentially to reproduce Luis's intermediate state.

## Why `fix_rnwbio_restore.py`?

Commit `f1aad25` ("Rename A1_Outputs to A1_Outputs_Luis and pull A1_Outputs from first_asia_model") removed two `RNWBIOBGDXX` and `RNWBIONPLXX` rows from `VariableCost` in the fresh A1 input. Luis's workflow used a version that still had them. The fix re-inserts them at their canonical positions (before `RNWBIOINDEA` and after `RNWWASINDWE`).

## Why `fix_pwrpet_clear.py` and `fix_elc_pmode_revert.py`?

Both are explicit Luis manual edits that we automated:
- **PWRPETBGDXX clear**: Luis blanked the 12×28 = 336 CapacityFactor cells in the `Capacities` sheet between `patch_ao_c2a` output and the FIX_2 input. The auto-generated CFs (≈0.01–0.22) likely seemed unrealistically low for petroleum power generation in Bangladesh; clearing them makes OSeMOSYS use `AvailabilityFactor=0.8` instead.
- **ELC*01 PM revert**: After `add_max_capacity_investment_rule.py NEW` flipped Projection.Mode for ZEROED ELC*01 rows to "User defined", Luis reverted them to "EMPTY" (since the year cells are 0 — the rows are functionally lockout placeholders, "User defined" is misleading).

## Verifying reproducibility

`python B0_mod.py --verify` compares the 4 final files against
`A1_Outputs_Luis/A1_Outputs_BAU/` cell-by-cell.

Two independent runs in fresh workdirs produce identical outputs (verified
2026-05-05).
