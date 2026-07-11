# OSTRAM — CLEAN-ROOM SOLVE PROMPT (fresh session; CPLEX, little-by-little)

**How to use:** open a NEW Claude Code session, cwd = `C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_mainredo`,
and say *"Read and execute CLEANROOM_SOLVE_PROMPT.md."* Prep is DONE; this session only SOLVES + debugs.

## ⛔ KNOWN ISSUE — FIX THIS FIRST: A_Calibrated_BAU is INFEASIBLE
First CPLEX run (2026-07-11): **B_Optimised_VRE OPTIMAL (2,212,351 ✓), C_Target_VRE OPTIMAL (2,253,104 ✓)** —
foundation solves. **A_Calibrated_BAU INFEASIBLE**: CPLEX presolve `Row TotalAnnualTechnologyActivityLowerLimit
(PWRSPVINDNO,2023) infeasible` -> feasopt garbage (Sum TDC 19.3e9). Cause: base-year pin set that solar tech's
2023 activity Lower=Upper=**136.0577 PJ EXACT**, but 2023 capacity (residual 22.76 + build 0.20 = 22.96 GW) at
solar CF maxes ~136 PJ -> knife-edge. B/C have no such pin, so they're fine.
**FIX:** the recipe TEXT said pin activity as **+/-0.2% BANDS**, but `apply_base_year_pin.py` does EXACT equality.
Re-pin A (+ A_Calibrated_BAU_Clipped, which inherits A's pin) with a band/tolerance on the activity lower-limit
(Lower = ~0.998x solved), recompile, re-solve; expect A ~ 2,314,332. If 0.2% is not enough or other VRE techs
also go infeasible, widen the band or investigate the capacity pin. B/C need NO rework. A's current Outputs are
the infeasible feasopt — discard on re-solve.

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
