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
| 1 consolidate (rebase phaseB + overlay WS-3/WS-4) | rebase done; overlay pending | | e4597e4 |
| 2 Checkpoint A (pre-WS-3 byte-diff vs OSTRAM_latest) | **GREEN** | A/B/C compiled .txt all 0-diff vs OSTRAM_latest | (this commit) |
| 3 foundation (loss+pin+cliff) + Checkpoint B | PENDING | | |
| 4 clips | PENDING | | |
| 5 sensitivities (+2 new solar tiers) | PENDING | | |
| 6 compile + glpsol --check (all 15) | PENDING | | |
| 7 stage CPLEX batch + RUN_ORDER.md | PENDING | | |

## Log
- 2026-07-11 — env audit green; branch rebased onto phaseB (e4597e4); harness scaffolded (8c9b674).
- 2026-07-11 — **Checkpoint A GREEN**. Ran A3->B1->B2 (no-solve: execute_model/create_matrix=False) for
  BAU, A_Calibrated_BAU, B_Optimised_VRE, C_Target_VRE. All 3 baseline compiled
  `Pre_processed_<s>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt` byte-diff **0-diff** vs `OSTRAM_latest`.
  Auto-fix applied (allowed): C's A3 `set_vre_targets.py` needs A's SOLVED combined `*_output.csv`
  (gitignored, absent in fresh clone) -> seeded from OSTRAM_latest's A_Calibrated_BAU_0; C then rebuilt
  clean and stayed 0-diff. Throwaway pre-WS-3 A1_Outputs xlsx churn discarded. Config left at no-solve for STEP 3.
  Seed mechanism validated for STEP 3 (will seed the WS-4 A-with-loss solve instead).
