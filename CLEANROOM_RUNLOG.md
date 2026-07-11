# CLEANROOM RUN LOG

Checkpoint/resume log for `CLEANROOM_FINALPROMPT.md` (overnight autonomous PREP; **no CPLEX**).
One line per step: status, checks + results, fixes, commit hash. Resume from the last green commit.
Branch `ws3-phaseb-cleanredo` (local only — **not pushed**). Started 2026-07-11.

## Environment audit (green)
- Python `C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe` = 3.10.20; glpsol 4.65 (PATH); otoole 1.1.5.
- `cplex` python module NOT importable in env — irrelevant this run (no solve); flagged for the batch stage.
- Hero repos present: `OSTRAM_latest` (pre-WS-3), `OSTRAM_ws3_workcopy_D5`, `OSTRAM_ws4_workcopy` (final/primary, working tree DIRTY by design), `OSTRAM_clean` (Phase-B @ f6321fa, clean).
- Anchors measured: pre_ws3 A 2,224,446 / B 2,113,984 / C 2,158,340; ws3_d5 A 2,222,829 / B 2,118,643 / C 2,164,880; final A 2,314,332 / B 2,215,073 / C 2,257,995 (== recipe).

## Branch base
- Working branch `ws3-phaseb-cleanredo` was originally cut from `main`; **rebased onto `phaseB-sensitivity-repro` (f6321fa)** so it carries `sensitivity_expansion/` + `configs/`. Slim commit (drop regenerable viz) preserved on top → `e4597e4`.

## BANKED CORRECTIONS (recipe prose vs authoritative hero script — do NOT "fix to green")
1. **Base-year pin exclusions:** `apply_base_year_pin.py` excludes **only the 18 named interconnectors**; it **pins** the internal-tx families (PWRTRN/RNWTRN/DSPTRN/RNWNLI/RNWRPO/TRNNLI/TRNRPO). Copy the script verbatim; the foundation check expects internal-tx **pinned**, not excluded.
2. **Pin = exact equality** (Lower==Upper, Max==Min). "±0.2%" is only the A==B==C base-year *verification tolerance*, never the pin mechanism.
3. **v18 stage-6 Restrictions rewrite** = benign churn every A3 run; never commit it.
4. Dead-key landmine: `TotalAnnualMaxCapacityInvestment` in internal-tx YAML blocks is not compiled — don't rename it.

## Step status
| Step | Status | Checks | Commit |
|---|---|---|---|
| 0 harness scaffold (hero_refs + cleanroom_check + runlog) | GREEN | files written | (this commit) |
| 1 consolidate (rebase phaseB + overlay WS-3/WS-4) | **GREEN** | 8 WS files + ws3_transmission_audit overlaid from ws4_workcopy; A3 carries stages 809-811; consolidation all-green | (this commit) |
| 2 Checkpoint A (pre-WS-3 byte-diff vs OSTRAM_latest) | **GREEN** | A/B/C compiled .txt all 0-diff vs OSTRAM_latest | (this commit) |
| 3 foundation (loss+pin+cliff) + Checkpoint B | **DONE — byte-diff RED (ws4 stale), foundation VERIFIED** | 2027+ 100% byte-identical to ws4; base-year pin reproduces ws4's FINAL A solve exactly; ws4 committed .txt stale | (runlog commit) |
| 4 clips | **BUILT ✓** | rebuilt alone after concurrency bug; RNWNLI internal-tx=200; DirBidir==B_Opt_Clipped 0-diff | |
| 5 sensitivities (+2 new solar tiers) | **BUILT ✓** | all 9 built; solar `years`-filter verified (Spike x1.5 only 2028-30); TradeCap15 STRICT (export_factor 0); DirBidir neutrality 0-diff | |
| 6 compile + glpsol --check (all 15) | **GREEN ✓ 15/15** | glpsol --check: 15/15 "Model has been successfully generated", 0 check-failures (structurally CPLEX-ready); validate_sensitivity_configs: COMPLETENESS/SCHEMA/COHERENCE/VRE_CEILINGS(3 approved clips)/TXCAP PASS; `3_DIFF_vs_BOPT` flags VRE LowerLimit = base-year-pin×ceiling interaction (validator predates the WS-4 pin) — flagged, not silenced | |
| 7 stage CPLEX batch + RUN_ORDER.md | **STAGED ✓** | RUN_ORDER.md + run_{baselines,sensitivities,directions}.bat; config restored to solve-mode (execute_model/create_matrix=True) | |

## Log
- 2026-07-11 — env audit green; branch rebased onto phaseB (e4597e4); harness scaffolded (8c9b674).
- 2026-07-11 — **Checkpoint A GREEN**. Ran A3->B1->B2 (no-solve: execute_model/create_matrix=False) for
  BAU, A_Calibrated_BAU, B_Optimised_VRE, C_Target_VRE. All 3 baseline compiled
  `Pre_processed_<s>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt` byte-diff **0-diff** vs `OSTRAM_latest`.
  Auto-fix applied (allowed): C's A3 `set_vre_targets.py` needs A's SOLVED combined `*_output.csv`
  (gitignored, absent in fresh clone) -> seeded from OSTRAM_latest's A_Calibrated_BAU_0; C then rebuilt
  clean and stayed 0-diff. Throwaway pre-WS-3 A1_Outputs xlsx churn discarded. Config left at no-solve for STEP 3.
  Seed mechanism validated for STEP 3 (will seed the WS-4 A-with-loss solve instead).
- 2026-07-11 — **STEP 1 overlay** committed (987aee4): 8 WS-3/WS-4 files + ws3_transmission_audit from ws4_workcopy.
- 2026-07-11 — **STEP 3 foundation built.** Seeded ws4 A-with-loss Outputs (anchor 2,314,332) into Executables/A_Calibrated_BAU_0/. Ran A3(WS-4: interconnector+internal-tx+3%loss stages) for A/B/C -> B1 -> cliff on C (PWRWONINDWE LowerLimit 2027 237.177->227.4139, 2028 260.635->255.0783; monotone, < UpperLimit) -> base-year pin A/B/C (exact-equality, 2023-2026, 198 interconnectors excluded, internal-tx pinned) -> B1 -> B2. All exits 0.
- 2026-07-11 — **CHECKPOINT B: byte-diff vs ws4 RED, but foundation VERIFIED. Root cause = ws4 provenance, NOT our pipeline.**
  - A/B/C differ from ws4 committed .txt ONLY at base years 2023-2026 (A: 1077 lines, distribution 196/280/298/303 across 2023/24/25/26; ZERO diffs at 2027+). Every foundation value at 2027+ (interconnector CapEx 380..2800, life 40, internal-tx OAR 0.97, cliff 227.4139/255.0783, VRE targets) is **byte-identical** to ws4.
  - The base-year diff is the pin. **Our pin = ws4's CURRENT A Outputs exactly** (NewCapacity PWRCOAINDEA 2025=0.6587, 2026=2.2555 in BOTH). ws4's committed .txt has 0.6613/2.2645 = an OLDER A solve.
  - **ws4 A Outputs mtime 17:31 > ws4 committed .txt mtime 17:19** -> ws4 re-solved A ~12 min AFTER compiling its baselines, and never re-pinned. So ws4's committed baseline .txt are pinned to a since-overwritten A solve.
  - **Conclusion:** our baselines are self-consistent with ws4's FINAL A solve (the 2,314,332 anchor); ws4's committed baselines are internally stale. Exact byte-match of ws4's committed .txt is IMPOSSIBLE (pin-time A solve overwritten). This is a hero-provenance issue, not fixable by us and not a pipeline error.
  - **DECISION FOR LUIS:** (a) accept our self-consistent baselines as canonical [recommended — they correctly pin ws4's final A solve], or (b) in ws4, re-pin+recompile baselines against its current A solve to get a consistent hero. Do NOT pin against a fabricated/older solve.
  - **Verification going forward** (base-year byte-diff is confounded by ws4 staleness): use `glpsol --check` (structural validity) + foundation-region (2027+) byte-match + lever spot-checks. Continuing to STEP 4/5/6 on the correct foundation.
- 2026-07-11 — **STEP 4 clips BUILT.** apply_patches --scenario {A_Calibrated_BAU_Clipped(src A), B_Opt_Clipped(src B), C_Target_VRE_Clipped(src C)} -> B1 -> B2, all exit 0. Ceiling-only patches over our pinned baselines.
- 2026-07-11 — **STEP 5 solar tiers.** Extended apply_patches.py `multiply` op with optional `years` filter (backward-compatible; only listed years multiplied). Created configs B_Opt_SolarCapex130 (x1.30 all years) and B_Opt_SolarCapexSpike (x1.50, years [2028,2029,2030]). apply_patches for SolarCapex{Hi,130,Spike} exit 0; B1+B2 rebuilding (first attempt killed by a double-background bug; relaunched clean).
- 2026-07-11 — glpsol --check invocation confirmed: datafile is DATA-ONLY, needs `-m osemosys_fast_preprocessed_storage_delay.txt -d <Pre_processed_..._RMCarefulXLSX.txt> --check`. Slow (>2 min/scenario) -> run in background.
- 2026-07-11 — **CONCURRENCY BUG found + fixed.** STEP-4 clips B1 failed for all 3 with `Errno 22` on Config_MOMF_T1_A.yaml because I launched the solar build (also runs B1) concurrently -> B1/B2 mutate Config_MOMF_T1_A.yaml (Main_Scenario) so two pipelines collide on the file lock. B1 logged the error but still exited 0, leaving the clips' A2 params stale (RNWNLI internal-tx = 100 non-RE instead of 200 RE; A2 mtime 01:42 vs A-O 03:03). Caught via the DirBidir!=B_Opt_Clipped neutrality check (13319-line diff). glpsol passed the stale clips because it checks structure, not values. **RULE: never run two B1/B2 pipelines concurrently.** Rebuilding the 3 clips alone. STEP-5 sensitivities are unaffected (ran against read-only glpsol, no config collision; verified RNWNLI=200).
- **RESUME POINT for remaining sensitivities:** seed ws4's FINAL B_Optimised_VRE solve (Executables/B_Optimised_VRE_0/Outputs from ws4_workcopy) [same self-consistent approach + same ws4-staleness caveat as A]; run `gen_sensitivity_patches.py` (recomputes TradeCap15 STRICT export_factor=0, TxCap150, IndiaCosts{,Fuel}); apply_patches each --source-scenario B_Optimised_VRE; for DirBidir/DirContractual also overlay set_interconnector_direction.py on the AR files (config set_interconnector_direction.yaml in each config dir); B1+B2; then glpsol --check all 15; validate_sensitivity_configs.py; desk_check.py; restore Config execute_model/create_matrix=True; STEP 7 stage batches + RUN_ORDER.md; cleanup + commit (exclude *_PRE_*/PREPATCH backups, mirrored Outputs, giant CSV/.sol/.lp).
