# RUN_ORDER — CPLEX solve batch (staged; NOT run in the cleanroom prep)

All 15 scenario datafiles are prepared (no-solve). To solve with CPLEX, run the batches
below **from the repo root**, one at a time. Config is already restored to solve-mode
(`Config_MOMF_T1_AB.yaml`: `execute_model: True`, `create_matrix: True`, `solver: cplex`,
`cplex_threads: 4`).

## ⚠️ HARD RULE: never run two batches (or two B1/B2 pipelines) at once
B1/B2 mutate `t1_confection/Config_MOMF_T1_A.yaml` (Main_Scenario). Two concurrent
pipelines collide on that file lock (`Errno 22`) and silently corrupt a scenario's compile.
Each batch runs its scenarios **sequentially** in one B2 process — that is safe. Run
`run_baselines.bat` → wait → `run_sensitivities.bat` → wait → `run_directions.bat`.

## Order
1. **`run_baselines.bat`** — `A_Calibrated_BAU`, then `C_Target_VRE`, then `B_Optimised_VRE`.
   (A first: C's set_vre_targets + the base-year pin reference A's solve — already seeded/mirrored
   from OSTRAM_ws4_workcopy for the prep; the prepared .txt are self-contained, so solve order is
   for convention/robustness.)
2. **`run_sensitivities.bat`** — 3 clipped + `SolarCapex{Hi,130,Spike}` + `TradeCap15` + `TxCap150`
   + `IndiaCosts` + `IndiaCostsFuel`.
3. **`run_directions.bat`** — `DirBidir` (neutrality: must == B_Opt_Clipped) + `DirContractual`.

## Expected anchors (Sum TotalDiscountedCost)
Baselines: A 2,314,332 / B 2,215,073 / C 2,257,995 — **but see the base-year-pin caveat in
CLEANROOM_RUNLOG.md**: our baselines pin ws4's FINAL A-solve (self-consistent); ws4's committed
baselines are pinned to a since-overwritten A-solve, so a byte/anchor match at base years is not
expected. Foundation (2027+) is byte-exact to ws4.

## Don't-chase-ghosts (expected, not bugs)
- Both `_output.sol` and `_output.feasopt.sol` written each solve; real status ("Dual simplex -
  Optimal") is in `_output.cplex.log`. NOT infeasibility.
- `IndiaCostsFuel` == `IndiaCosts` (fuel non-binding); `DirBidir` == `B_Opt_Clipped` (neutrality — the
  prepared .txt are already 0-diff).
- v18 shows modified after any A3 run (stage-6 Restrictions rewrite) — benign, do not commit.

## Verify after solving
- Behavioural (not byte): expect same SIGNS/RANKING vs the pre-WS-3 Phase-B story, not identical
  magnitudes (WS-3/WS-4 shifted the foundation). `sensitivity_expansion/analyse_*`.
