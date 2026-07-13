# WS-3 — Promotion Handoff

**Purpose:** self-contained spec of *everything* WS-3 changed, so a promotion sequence can bring it into `OSTRAM_clean` cleanly. Branch-only, no merge, specific files.
**Date:** 2026-07-09 · **Status of the work:** COMPLETE + verified (all 3 scenarios re-run and solved; feasible; base-consistent).

---

## 0. Source & baseline

| Role | Location |
|---|---|
| **Source of truth (all WS-3 changes, verified)** | `C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_ws3_workcopy_D5` |
| Frozen milestone (interconnector-only; rollback point) | `…\OSTRAM_ws3_workcopy` |
| **Promotion target / baseline** | `…\OSTRAM_clean`, branch `validated-baseline-3scenario` @ commit `11d53d4` (= main tip at fork) |

**Golden rule:** `git add` **only** the specific files in §2 — **never `git add -A`**. The copy shows the 8.5 GB `t1_confection/Executables` as deleted, plus regenerated `A1_Outputs/` and `A2_Output_Params/` noise.

---

## 1. What WS-3 did (three workstreams, all in `_D5`)

1. **Interconnectors** — made the v18 `Interconnector_Params` sheet the *source of truth* for corridor `CapitalCost`/`FixedCost`/`OperationalLife` (previously legacy distance-computed values from `OG_csvs`, never consumed). 18 corridors sourced/cited; submarine raised (LK↔IN 1250, MV↔IN 2800); LK↔MV repriced subsea (1250); 3 corridors added; `OperationalLife` 40; FOM = 1.5% × CapEx.
2. **D5 — internal (intra-node) transmission** — per-node `ResidualCapacity` (existing grid at peak×1.2; replaces flat 5 GW), a **2× RE CapEx premium** exposed as a live YAML slider, and `OperationalLife` = 40 for the six internal families (RNWTRN/RNWNLI/RNWRPO, PWRTRN/TRNNLI/TRNRPO). Uniform across nodes.
3. **D7 — citations** — backfilled the two `[Pending]` interconnector citations (NP↔BD, BT↔BD) in the REFS workbook. (+ optional polish, see §7.)

---

## 2. Promotion set — exact files (git-verified vs `11d53d4`)

### 2A. Model files — CHANGE THE MODEL, must promote

| File | git | Change |
|---|:--:|---|
| `t1_confection/A3_process.py` | M | +2 core late stages (`stage_ws3_interconnector_costs`, `stage_ws3_internal_transmission`) + their two calls after stage 5 |
| `t1_confection/Config_country_codes.yaml` | M | new `internal_transmission` knob block (RE multiplier=2.0, base 100/4, life 40) + the 6 family `OperationalLife` 20→40 |
| `t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx` | M | `Interconnector_Params` sourced values (18 corridors) + a documentation `Source` column (18 tags; inert to A3) |
| `t1_confection/A3_process/rules_scripts/apply_interconnector_costs.py` | NEW | interconnector-cost wiring stage (self-tested) |
| `t1_confection/A3_process/rules_scripts/apply_internal_transmission.py` | NEW | internal-tx calibration stage (self-tested) |
| `t1_confection/A3_process/rules_scripts/internal_tx_residuals.csv` | NEW | frozen desk-checked per-node residuals |

### 2B. Deliverables / audit trail — promote if audit docs live in-repo

| Path | Contents |
|---|---|
| `ws3_transmission_audit/` (folder) | `WS3_calibration_report.md`, `WS3_value_audit.md`, `WS3_TASK_LEDGER.md`, this file; audit scripts (`compute_internal_tx_residuals.py`, `audit_transmission_values.py`, `verify_base_consistency.py`, `set_final_v18_interconnector_values.py`); `inputs/` (REFS w/ D7 citations + cost DB); `outputs/` |

---

## 3. Mechanism (what the promoted code does)

Both fixes are **late core A3 stages** invoked in `A3_process.py:main()` **after** stage 5 (the scenario rules chain), before delivery, in this order:

1. `stage_ws3_interconnector_costs(...)` → `apply_interconnector_costs.py`: writes v18 `Interconnector_Params` CapEx/FOM into `Secondary Techs`, `OperationalLife` into `Fixed Horizon Parameters` (reads the per-scenario materialized template).
2. `stage_ws3_internal_transmission(s5)` → `apply_internal_transmission.py`: writes per-node `ResidualCapacity` + RE/non-RE `CapitalCost`/`FixedCost` into `Demand Techs`, `OperationalLife`=40 into `Fixed Horizon Parameters`; knobs from `Config_country_codes.yaml` (`internal_transmission`) + residuals from `internal_tx_residuals.csv`.

**Why late stages (not snapshot/A2/YAML edits):** Stage-1 (`3_update_ao_from_extensions.py`) stamps transmission `OperationalLife` to 50/20 from its template merge; only a post-stage-5 write survives. Stage-3 residual scripts don't touch these (interconnector: `cap_trn_to_residual` allowlist; internal-tx live in `Demand Techs`, not the `Secondary Techs` `fix_trn_residuals` reads). The post-A2 snapshot is left pristine.

---

## 4. Promotion sequence (recommended)

1. Branch from `11d53d4`.
2. Bring the six §2A files into `OSTRAM_clean`.
3. Re-run per scenario (changes take effect at A3): `A3_process.py --scenario <s>` → `B1_Run_Compiler.py --scenarios "<s>"`, for `A_Calibrated_BAU`, `B_Optimised_VRE`; then B2-solve A + B; then `A3→B1` for `C_Target_VRE` (its `set_vre_targets` reads the **solved** A); then B2-solve C.
4. Verify against §5 anchors.
5. `git add` the §2 files explicitly; commit on the branch.

Env: `C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe` (prepend env dirs to PATH; `PYTHONUTF8=1`). B2 in an Anaconda Prompt (`cplex_threads=4`).

---

## 5. Verification anchors (StorageDelayN5 solve; metric = sum `TotalDiscountedCost`; all feasible, base-year backstop 0)

| Scenario | Pre-WS-3 | Post-interconnector | **Post-D5 (final target)** |
|---|--:|--:|--:|
| A_Calibrated_BAU | 2,224,447 | 2,229,145 | **2,222,829** |
| B_Optimised_VRE | 2,113,985 | 2,117,860 | **2,118,643** |
| C_Target_VRE | 2,158,340 | 2,163,127 | **2,164,880** |

Input checks: internal-tx compiled **A == B == C**; interconnector CapEx 380…2800 + life 40 intact; Phase-0 base-consistency **10/12** (the 2 "fails" are the known-benign 2023-generation-differs-by-policy items). Verifier: `ws3_transmission_audit/verify_base_consistency.py`.

---

## 6. Caveats

- **Never `git add -A`** — add the §2 files explicitly.
- `A3_process.py` stage-6 writes the v18 template's `Restrictions` sheet on every run (pre-existing behavior) — the template will show as modified after any A3 run; that's expected, not a WS-3 change.
- Solve config (`Config_MOMF_T1_AB.yaml`): `storage_delay_active: True` (StorageDelayN5 variant; outputs prefixed `OSTRAM_StorageDelay_`), `cplex_threads: 4`, `only_main_scenario: False`, `solver: cplex`. The §5 anchors were produced with this config.

---

## 7. Optional polish (documentation-only, no model effect)

- **v18 `Interconnector_Params` `Source` column — ✔ DONE (2026-07-09).** A compact source tag added on the `CapitalCost` row of all 18 corridors (new column `Source`, index 38). Verified: CapEx values + all 20 sheets intact; A3 reads by header name so the column is inert. Backup: `SOASIA_OSeMOSYS_Template_v18_PRE_SOURCE_COL_*.xlsx`. The live template now self-documents all 18 corridors.
- **3 added-corridor rows in REFS — SKIPPED (superseded).** The v18 `Source` column above now documents the 3 added corridors (`TRNINDEAINDWE`, `TRNINDNEINDNO`, `TRNLKAXXMDVXX`) in the live artifact, and their basis is in calibration report §3 — so a redundant structured insert into REFS was not worth the risk. Full IEEE citations remain in REFS for the original 15 (incl. the 2 D7 backfills).
