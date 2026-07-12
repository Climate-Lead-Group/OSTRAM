# OSTRAM — CLEAN-ROOM SOLVE PROMPT (fresh session; CPLEX, little-by-little)

**How to use:** open a NEW Claude Code session, cwd = `C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_mainredo`,
and say *"Read and execute CLEANROOM_SOLVE_PROMPT.md."* Prep is DONE; this session only SOLVES + debugs.

## STATE: FULL SOLVE DONE (2026-07-12) — 12/15 OPTIMAL, 3 infeasible. FIX the 3 FIRST.
All 15 re-solved with the banded pin (`apply_base_year_pin.py --band 0.002`, now committed) + dual simplex +
feasopt-off. **12 Optimal**: baselines A/B/C 2,314,128/2,214,920/2,257,930 (within 0.01% of ws4); 3 clips;
Solar Hi/130/Spike; IndiaCosts==IndiaCostsFuel 2,177,458; DirBidir==B_Opt_Clipped 2,215,995 (neutrality ok).
The earlier A knife-edge is RESOLVED by the band.
**⛔ 3 INFEASIBLE — one modeling decision (see CLEANROOM_RUNLOG.md solve-tracker for full detail):**
`B_Opt_TradeCap15`, `B_Opt_TxCap150`, `B_Opt_DirContractual` — each infeasible on a BASE-YEAR backstop
ActivityUpperLimit (PWRBCK{BGDXX,BGDXX,BTNXX} @ 2024/2025/2023). These are exactly the 3 sensitivities that
restrict BASE-YEAR network flows; the WS-4 base-year pin freezes the calibrated import-reliant base-year mix, so
cutting base-year imports leaves a node's base-year demand unmeetable (genuine demand-balance infeasibility, not
a pin-vs-lever conflict; band won't help). They solved in pre-WS-3 Phase-B because there was no base-year pin.
**FIX (modeling call):** apply the trade/tx/direction levers to the STUDY PERIOD 2027+ only (leave 2023-2026 =
pinned calibrated mix), OR keep the base-year backstop available (zero it from 2027). Then rebuild+resolve those 3.
The WACC test target `B_Opt_Clipped` solved fine (2,215,995) -> WACC_TEST_PROMPT.md is unblocked.

## STATE (already done — do NOT redo)
- Branch `ws3-phaseb-cleanredo`. All **15 scenario datafiles are prepared, committed (`092dbb5`), and
  glpsol-`--check` clean (15/15)**. Foundation byte-exact vs `ws4_workcopy` at 2027+.
- Config already solve-mode (`Config_MOMF_T1_AB.yaml`: `execute_model: True`, `create_matrix: True`,
  `solver: cplex`, `cplex_threads: 4`). Full history: `CLEANROOM_RUNLOG.md`. Order: `RUN_ORDER.md`.
- **Baselines pin ws4's FINAL A-solve** (self-consistent); ws4's committed baselines are stale — see runlog.

## HARD RULES
- **Never run two B1/B2 pipelines at once.** They mutate `t1_confection/Config_MOMF_T1_A.yaml`
  (Main_Scenario) and collide on the file lock (`Errno 22`), silently corrupting a compile. Solve
  **one scenario (or one batch) at a time, sequentially.** Do NOT push.
- Env: `conda activate OSTRAM-env`; `set PYTHONUTF8=1` / `PYTHONIOENCODING=utf-8` / `chcp 65001`.

## SOLVE — one at a time (recommended for debugging)
From `t1_confection`:  `python B2_Executing_OG_Model.py --scenarios "<ONE_SCENARIO>"`
After each, verify:
1. **Optimal** — `Executables/<s>_0/Pre_processed_<s>_0_..._output.cplex.log` says `Dual simplex - Optimal`
   (NOT the `.feasopt.sol` — see ghosts). 2. **Outputs** produced under `Executables/<s>_0/Outputs/`.
3. **Anchor** — Sum of `Outputs/TotalDiscountedCost.csv` in the expected range below.

Order: `A_Calibrated_BAU` -> `C_Target_VRE` -> `B_Optimised_VRE`, then the clips + `SolarCapex{Hi,130,Spike}`
+ `TradeCap15` + `TxCap150` + `IndiaCosts{,Fuel}`, then `DirBidir` + `DirContractual`.
(Or use the staged batches `run_baselines.bat` / `run_sensitivities.bat` / `run_directions.bat`, one at a time.)

## EXPECTED baseline anchors (Sum TotalDiscountedCost, M USD)
A_Calibrated_BAU ~ **2,314,332** · B_Optimised_VRE ~ **2,215,073** · C_Target_VRE ~ **2,257,995**.
(Our .txt pin ws4's final A-solve, so solving should reproduce these WS-4 anchors. Tiny base-year
deviations are expected/benign given ws4's A-solve provenance.)

## DON'T-CHASE-GHOSTS (expected, not bugs)
- Both `_output.sol` and `_output.feasopt.sol` written each solve; real status is in `_output.cplex.log`. NOT infeasibility.
- `IndiaCostsFuel` == `IndiaCosts` (fuel non-binding); `DirBidir` == `B_Opt_Clipped` (neutrality; already 0-diff in the .txt).
- v18 shows modified after any A3 run (stage-6 Restrictions rewrite) — benign; do NOT commit; do NOT re-run A3 (would drop the pin/cliff).

## VERIFY sensitivities BEHAVIOURALLY (not by number-match)
Compare each solved sensitivity to `OSTRAM_clean`'s Phase-B solve of the same scenario
(`../OSTRAM_clean/t1_confection/Executables/<s>_0/`): expect the **same signs and ranking**, NOT identical
magnitudes (WS-3/WS-4 shifted the foundation). `OSTRAM_clean` (read-only) is the behavioural oracle;
`sensitivity_expansion/analyse_*` + `PHASE_B_METHODOLOGY_AND_RESULTS.md` describe the expected story.

## VALIDATORS — DONE (already green)
`validate_sensitivity_configs.py` = 32 PASS / 0 FAIL (taught DIFF_vs_BOPT about the pin x ceiling VRE
lower-limit; 6_RUN1_CAPS now reads export_factor; dropped superseded TradeCap30). glpsol --check = 15/15.
OPTIONAL later: its SCEN sanity subset is 6 scenarios — SolarCapex130/Spike, the Dir scenarios, and the
A/C clips are glpsol-covered but not in that subset; add them to SCEN + `_expected` if you want fuller config coverage.

## COMMIT (working branch; NO push)
- Solve Outputs (`Executables/*/Outputs/`, `.sol`, `.txt`, combined CSVs) are GITIGNORED/regenerable — do
  NOT commit them. Commit only analysis artifacts you deliberately produce. STOP — no push until Luis approves the PR.
