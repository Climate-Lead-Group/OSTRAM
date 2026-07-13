# OSTRAM — CLEAN-ROOM FINAL PROMPT (overnight autonomous PREP; no CPLEX)

**How to use:** open a fresh Claude Code session with cwd = `C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_mainredo`, and say *"Read and execute CLEANROOM_FINALPROMPT.md."* This is the single authoritative prompt (supersedes CLEANROOM_SUPERPROMPT.md).

---

## GOAL
Rebuild a clean, canonical OSTRAM on a working branch: consolidate Phase-B + the WS-3/WS-4 transmission
& calibration foundation, apply EVERY change, regenerate ALL inputs + compiled `.txt`, run `glpsol
--check` preflight on all, prove reproduction by byte-diff vs the hero repos, and stage the CPLEX batch
— running **NO B2 solve**. Commit the prepped tree to the working branch. **Do NOT push.**

Absolute values shift vs pre-WS-3 numbers (WS-3/WS-4 change the foundation): verify **baselines EXACTLY**
(byte-diff), **sensitivities BEHAVIOURALLY** (after the batch solves, later).

## HARD RULES
- **No push.** **No CPLEX** in this run (prep only; stage a batch). Explicit `git add <path>` — **never `git add -A`**.
- **v18 = WS-3's version** is the source of truth (interconnector values). Do NOT hand-edit it or restore an older one. A3 stage-6 rewrites v18's `Restrictions` sheet every run — **benign, do not commit that churn**.
- **Reference repos are READ-ONLY oracles — never modify:** `OSTRAM_latest` (pre-WS-3 hero), `OSTRAM_ws3_workcopy_D5` (WS-3 hero), `OSTRAM_ws4_workcopy` (WS-4 hero + cumulative source), `OSTRAM_clean` (Phase-B).
- Env: `C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe` (conda NOT on PATH; call the exe directly). `PYTHONUTF8=1`, `PYTHONIOENCODING=utf-8`, `chcp 65001`. CPLEX 22.1.2 + license, otoole, glpsol installed.

## CONFIG
`t1_confection/Config_MOMF_T1_AB.yaml`: `storage_delay_active: True`, `strip_storage_active: False`,
`cplex_threads: 4`, `only_main_scenario: False`, `solver: cplex`, `reuse_existing_sol: False`. For the
no-solve preflight/diff set `execute_model: False` + `create_matrix: False`, then restore both to `True` before committing.
`t1_confection/Config_country_codes.yaml` → `internal_transmission`: `transmission_loss: 0.03`,
`re_capex_multiplier: 2.0`, base costs 100/4, `operational_life: 40`; the 6 internal families `OperationalLife` 20→40.

---

## AUTONOMY, CHECKPOINTS & SELF-CORRECTION (run unattended, overnight, before any CPLEX)
Steps run STRICTLY SEQUENTIALLY. Per step: do work → run that step's checks → if GREEN, `git add <paths>`
+ commit `"cleanroom: <step> — green"` (working branch, NO push) + log → proceed; if RED, self-correct.

**Self-correction (bounded, ≤3 attempts/step):** diagnose → fix the PIPELINE/INPUTS/ENV → re-check. Log
each attempt (hypothesis → fix → result). **NEVER weaken/skip/loosen a check, anchor, or tolerance to go
green — fix what the check measures, not the measurement.**
- **MAY auto-fix (mechanical):** encoding/PATH/env; glob-scoping artifacts (confirm via `ls`/abs path
  before "missing"); a missing seed/mirror copy; the per-year `multiply` op for the solar spike;
  stale-output collisions (delete stale `*NoStorage*_output.csv`, keep StorageDelayN5); a documented-benign item.
- **MUST NOT auto-fix → STOP the step, mark RED, log, continue INDEPENDENT steps:** a baseline byte-diff
  mismatch vs a hero (foundation didn't reproduce — don't fudge inputs); a non-benign `glpsol --check`
  fail; an unmatchable anchor; a source file locatable nowhere. Escalate to a human.

**Checkpoint/resume:** `CLEANROOM_RUNLOG.md` — one line per step (status, checks+results, fixes, commit
hash). Resume from the last green commit (steps idempotent). END: all self-correctable steps green,
RED-for-human items listed, CPLEX batch + `RUN_ORDER.md` staged → final runlog summary, commit, STOP. No CPLEX.

## HERO REFS — one connection point for VERIFY + MIRROR (few files, simple names)
Create ONE manifest `cleanroom_tests/hero_refs.yaml` naming the solved folders + anchors — the single place
the whole run reads, for both verification and mirroring/seeding:
- `pre_ws3: OSTRAM_latest` (A 2,224,447 / B 2,113,985 / C 2,158,340)
- `ws3_d5: OSTRAM_ws3_workcopy_D5` (A 2,222,829 / B 2,118,643 / C 2,164,880)
- `final: OSTRAM_ws4_workcopy` (A 2,314,332 / B 2,215,073 / C 2,257,995)  ← primary hero
- `seeds:` the solved outputs the CPLEX-free prep must MIRROR in (A-with-loss `Outputs/` for the base-year pin + C's `set_vre_targets`).

**VERIFY:** checks read the hero paths from the manifest and byte-diff new-vs-hero (READ-ONLY; never modify a hero folder).
**MIRROR:** copy the manifest `seeds` (solved baseline `Outputs/`) from the hero into the new repo so the
CPLEX-free chain completes; AND mirror the hero's solved A/B/C `Outputs/` into `Executables/<s>_0/Outputs/`
(labelled "mirrored from ws4 hero — pending local re-solve") so the repo is a complete, runnable mirror —
but keep the mirrored outputs **GITIGNORED / UNCOMMITTED** (committed footprint = source + regenerated inputs
only; the byte-diff of the regenerated INPUTS remains the actual reproduction proof).

## TEST HARNESS — ONE script + the manifest (simplicity first; ~3 new files total)
One file `cleanroom_tests/cleanroom_check.py` with subcommands (run one, or `--all`), reading `hero_refs.yaml`:
`consolidation` (8 WS files + Phase-B tooling present; v18=WS-3's; branch ok) · `reproduction` (CR-stripped
`.txt` byte-diff == hero for A/B/C — **hard oracle: 0-diff or RED**) · `foundation` (interconnector CapEx
380…2800 + life 40; internal-tx OAR 0.97; D5 values; pin bands ±0.2% both activity+build, gen/mining/storage
only, transmission excluded; C cliff 227.414/255.078; base-years A==B==C within 0.2%) · `clips`
(min(atlas,MaxCap); 16/3/0) · `sensitivities` (each of 15 carries foundation + its lever; `glpsol --check`
clean) · `hygiene` (no giant CSV/.sol/.lp, backups, superseded, or `-A` staged; no v18 churn).
Whole run adds only: `cleanroom_check.py`, `hero_refs.yaml`, `CLEANROOM_RUNLOG.md` (+ the staged `.bat`s + `RUN_ORDER.md`). Don't create more; clear names; no per-check sprawl.

---

## STEP 1 — CONSOLIDATE (Git Bash; clone already exists at OSTRAM_mainredo on `main`)
```
cd /c/Users/luisfernando/Desktop/OSeMOSYS/OSTRAM_mainredo
git checkout phaseB-sensitivity-repro        # Phase-B source (pushed to origin); pre-WS-3 pipeline
ls t1_confection/sensitivity_expansion/      # sanity: Phase-B tooling present
git checkout -b ws3-phaseb-cleanredo         # working branch
```
Then overlay the **WS-3+WS-4 foundation** by copying these from `..\OSTRAM_ws4_workcopy\` (cumulative — one folder):
- `t1_confection/A3_process.py` (WS-3 stages + WS-4 `stage_ws3_internal_tx_losses`, all after stage 5)
- `t1_confection/Config_country_codes.yaml`
- `t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx` (WS-3 Interconnector_Params + Source col)
- `t1_confection/A3_process/rules_scripts/apply_interconnector_costs.py`
- `t1_confection/A3_process/rules_scripts/apply_internal_transmission.py`
- `t1_confection/A3_process/rules_scripts/internal_tx_residuals.csv`
- `t1_confection/A3_process/rules_scripts/apply_internal_tx_losses.py`
- `t1_confection/A3_process/rules_scripts/apply_base_year_pin.py`
- `ws3_transmission_audit/` (audit trail + `verify_base_consistency.py`)
**Checkpoint A prerequisite:** run STEP 2 Checkpoint A (pre-WS-3 reproduction) BEFORE this overlay, or on a stash — the overlay changes the pipeline.

## STEP 2 — REPRODUCIBILITY GATE (staged; byte-diff, NO CPLEX)
Anchors = Σ `Outputs/TotalDiscountedCost.csv` (NOT the raw `.sol` objective — offset ~157k pre-loss, ~165,245 post-loss).
Reproduction proven WITHOUT solving via the no-solve `.txt` diff: generate the compiled
`Pre_processed_*_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt` with `execute_model:False`, CR-strip, diff vs the hero → **0-diff guarantees the solve reproduces**.
- **Checkpoint A (pre-WS-3):** on the phaseB base (BEFORE the WS-4 overlay), `A3→B1` the 3 baselines, byte-diff vs `OSTRAM_latest` → proves clone + pipeline sound. Expected: A 2,224,447 / B 2,113,985 / C 2,158,340.
- **Checkpoint B (post-WS-4):** after Steps 1+3, byte-diff the 3 baselines vs `OSTRAM_ws4_workcopy` → proves the foundation reproduces. Expected FINAL anchors below.

## STEP 3 — FOUNDATION (baselines with WS-3+WS-4; A3 re-runs the late stages)
Per scenario: `A3_process.py --scenario <s>` (interconnector → internal-tx → 3% loss stages) → the two post-A3 A-O edits below → `B1_Run_Compiler.py --scenarios "<s>"`.
1. **Seed/MIRROR solved baselines** from `OSTRAM_ws4_workcopy`: C's `set_vre_targets` and the base-year pin both read the **solved A-with-loss** → copy `Executables/A_Calibrated_BAU_0/Outputs/` from ws4_workcopy before A3 on C and before the pin. (Confirm exact dir: `BAU_0` vs `A_Calibrated_BAU_0`.)
2. **C_Target VRE cliff relax** (C only), `A1_Outputs/A1_Outputs_C_Target_VRE/A-O_Parametrization.xlsx`, 'Secondary Techs' · `TotalTechnologyAnnualActivityLowerLimit` · `PWRWONINDWE` (back up first): 2027 `237.221→227.414`, 2028 `260.745→255.078` (0.98×maxReach). Guardrails: < that year's UpperLimit; monotone. Re-scan LP → 0 cliffs.
3. **Base-year pin** (all 3), `apply_base_year_pin.py`, referencing the seeded A-with-loss solve:
   - Pin **both** activity (`Activity Lower/UpperLimit`) **and** build (`Min/MaxCapacityInvestment`) as **±0.2% BANDS**, for **2023–2026**.
   - **Generation + mining + storage ONLY. EXCLUDE all transmission** (18 interconnectors; RNWTRN/RNWNLI/RNWRPO/PWRTRN/TRNNLI/TRNRPO; DSPTRN). `--tech-csv` gate to the ~384 real techs.
   - These edit the delivered A-O — do NOT re-run A3 after (pin/cliff would be lost); go straight to B1.
4. **Verify (Checkpoint B):** byte-diff 3 compiled `.txt` vs `OSTRAM_ws4_workcopy`; `verify_base_consistency.py` (expect 10/12; 2 benign); interconnector CapEx 380…2800 + life 40; internal-tx OAR 0.97 (6 families), A==B==C; base-years A==B==C within 0.2%.

**FINAL foundation anchors (Σ TotalDiscountedCost; loss + pin + cliff):** A_Calibrated_BAU **2,314,332** · B_Optimised_VRE **2,215,073** · C_Target_VRE **2,257,995**.

## STEP 4 — CLIPS (VRE ceiling layer → clipped baselines; A3-FREE, apply_patches)
Ceiling = `min(atlas, B_Opt MaxCap)`, flat, from `sensitivity_expansion/reference/vre_ceilings_base.json` (3 enforced clips: PWRSPVLKAXX 16, PWRWONBGDXX 3, PWRWONMDVXX 0). Build `A_Calibrated_BAU_Clipped`, `B_Opt_Clipped`, `C_Target_VRE_Clipped` via `apply_patches --scenario S --source-scenario <base>` (ceiling-only static patches). For C, keep the NDC-floor↔ceiling coherence scaling (only PWRSPVMDVXX affected).

## STEP 5 — SENSITIVITIES (A3-FREE; apply_patches off the WS-4 **pinned** B_Optimised_VRE A-O)
`python sensitivity_expansion/gen_sensitivity_patches.py`, then `apply_patches --scenario S --source-scenario B_Optimised_VRE` per scenario; direction scenarios also get the `set_interconnector_direction.py` overlay on the AR files. All inherit the WS-3+WS-4 foundation + ceiling. **15 scenarios = 3 baselines + 3 clipped + 9 below:**
- `B_Opt_TradeCap15`  — imports ≤15% demand, STRICT (export_factor=0); backstops AUL=0
- `B_Opt_TxCap150`    — cross-border TRN MaxCap=1.5×Residual_2023; India-internal kept; non-India TRNNLI*/TRNRPO* backstops **AUL=0 NOT MaxCap=0**
- `B_Opt_SolarCapexHi`    — PWRSPV CapitalCost ×1.10 (mild reference; known immaterial)
- `B_Opt_SolarCapex130`   — **NEW** PWRSPV CapitalCost ×1.30, all nodes/years (PRIMARY sustained solar stress; literature-anchored: NREL ATB Conservative-vs-Moderate spread ~30%, IEA diversification premium, WACC-stress equiv). Same mechanism as SolarCapexHi, factor 1.30.
- `B_Opt_SolarCapexSpike` — **NEW** PWRSPV CapitalCost ×1.50 for **2028–2030 only**, revert to 1.0 elsewhere (transient severe spike). Needs a **per-year** multiply (factor 1.50 on the 2028/2029/2030 columns only); if the patcher's `multiply` op is all-years, extend it to a year list or write the 3 columns explicitly.
- `B_Opt_IndiaCosts`     — non-India gen/storage CapitalCost+FixedCost → India ref (INDNO anchor); fuel OFF
- `B_Opt_IndiaCostsFuel` — as IndiaCosts + non-India MIN* fuel VariableCost → India
- `B_Opt_DirBidir`       — empty direction map (neutrality check → == B_Opt_Clipped)
- `B_Opt_DirContractual` — 9 corridors one-way, 2 (Nepal↔India) bidirectional
Solar band note: ×1.20 and ×1.40 are trivial factor-swaps of SolarCapex130 — run them to trace the robustness threshold. WACC/financing stress is a SEPARATE discount-rate sensitivity — out of scope for the CapEx knob. The multiplier rides on the model's existing South-Asia-calibrated per-node CapEx — do NOT change the base.

## STEP 6 — COMPILE + PREFLIGHT (all 15; NO CPLEX)
`B1_Run_Compiler.py` all; generate compiled `.txt` with execute_model:False; `glpsol … --check` each (expect "successfully generated", NO `check[...] failed`); `validate_sensitivity_configs.py` (all PASS — already scans Secondary AND Demand Techs, do NOT re-add), `desk_check.py`. Then restore execute_model:True + create_matrix:True.

## STEP 7 — STAGE THE CPLEX BATCH (do NOT run)
Leave `run_sensitivities.bat` + `run_directions.bat` + a baseline-solve batch ready; write `RUN_ORDER.md`: solve A → C after A (its set_vre_targets + the pin reference A's solve — already seeded/mirrored); then clipped + TradeCap15/TxCap150/SolarCapex{Hi,130,Spike}/IndiaCosts{,Fuel}; then DirBidir/DirContractual.

## VERIFICATION SUMMARY
- **Baselines:** byte-diff `.txt` vs `OSTRAM_latest` (Checkpoint A) and `OSTRAM_ws4_workcopy` (Checkpoint B) → exact.
- **Sensitivities:** structural (WS-3+WS-4 foundation inherited + lever applied) + `glpsol --check`. Behavioural anchor comparison vs the pre-WS-3 Phase-B story happens AFTER the batch solves — expect same SIGNS/RANKING, not identical magnitudes.

## DON'T-CHASE-GHOSTS (expected, not bugs)
- Both `_output.sol` and `_output.feasopt.sol` written each solve; real status ("Dual simplex - Optimal") is in `_output.cplex.log`. NOT infeasibility.
- `IndiaCostsFuel` == `IndiaCosts` (fuel non-binding); `DirBidir` == `B_Opt_Clipped` (neutrality check).
- Unclipped B_Opt Capital(NPV) reads ~6.89M (concat glitch); affects no Σ-TDC anchor; don't "fix" it.
- v18 shows modified after any A3 run (stage-6 Restrictions rewrite) — benign, don't commit that churn.

## CLEANUP + COMMIT (working branch; NO push)
- **SOURCE (commit):** WS-3/WS-4 files (Step 1); Phase-B tooling + all scenario configs; `reference/*`; `Config_*.yaml`; the docs; `ws3_transmission_audit/`; `cleanroom_tests/`; `RUN_ORDER.md` + the new SolarCapex130/Spike configs. v18 = WS-3's.
- **GENERATED (commit as snapshot):** `A1_Outputs/<15>`, `A2_Output_Params/<15>`, `A2_Outputs_Params_otoole/<15>`, compiled `.txt`. (Executables Outputs stay empty/mirrored until the batch runs.)
- **GITIGNORE (never commit):** `OSTRAM_*Combined*_*.csv`, `*StorageDelay*` (catches giant `.sol/.feasopt.sol/.lp` + mirrored solve outputs), `Pre_processed_*_output.csv`. Confirm coverage.
- **SKIP (live in other repos):** backups (`*_PREPATCH_*`, `*_PRE_*`), superseded scenarios (TradeCap50/30, SolarHi10, LinkFreeze).
- Before EACH commit: `git status` — no `-A`, no v18 churn, no giant CSV/.sol/.lp, no backups/superseded/mirrored-outputs. Commit source, then prepped inputs. **STOP — no push.**

## REFERENCES
Phase-B: `sensitivity_expansion/PHASE_B_IMPLEMENTATION_LOG.md`, `PHASE_B_METHODOLOGY_AND_RESULTS.md`. WS-3/WS-4: `ws3_transmission_audit/WS3_PROMOTION_HANDOFF.md`, `WS4_HANDOVER_PROMPT.md`, `WS4_PREFLIGHT.md`, `WS3_calibration_report.md`.

## PASS (overnight)
All 15 compiled + `glpsol --check` clean; baselines 0-diff vs `OSTRAM_latest` (pre-WS-3) and `OSTRAM_ws4_workcopy` (final); WS-3+WS-4 values present; levers structurally correct; full `cleanroom_tests/` suite green (or RED items logged in `CLEANROOM_RUNLOG.md`); each step checkpoint-committed; CPLEX batch + `RUN_ORDER.md` staged; committed to `ws3-phaseb-cleanredo`; **not pushed**.
