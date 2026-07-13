# OSTRAM — WACC sensitivity TEST (one cell). Prove the mechanism, report loudly, PERSIST until it works.

**How to use:** fresh Claude Code session, cwd = the RUN PATH below, say *"Read and execute WACC_TEST_PROMPT.md."*

RUN PATH:  `C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_mainredo`
TEST CELL: `B_Opt_Clipped` @ DiscountRate **0.13** (base is 0.10)

## ⏱ PRECONDITION (hard): the 15-scenario solve must be FINISHED first
Do NOT start until the `phase4b` solve is done (a running B1/B2 collides on `Config_MOMF_T1_A.yaml` -> Errno 22,
silent corruption). Check: `tasklist | findstr /I "cplex python"` shows none, and CLEANROOM_RUNLOG.md solve-tracker = 15/15.
**Run only ONE B1/B2 at a time.** Do NOT push; do NOT `git add -A`; back up any file before editing.

## GROUNDWORK ALREADY DONE (verified read-only this session — saves you the dig)
- **Base B_Opt_Clipped @ 10% = Sum TotalDiscountedCost 2,215,995 (Optimal)** — this is your comparison baseline.
- The rate is NOT explicit: the solved `.txt` has `param default 0.1 : DiscountRate :=` (empty list) and
  `param default 0.1 : DiscountRateStorage :=`. The `0.1` is **otoole's default**, injected because
  `A2_Outputs_Params_otoole/B_Opt_Clipped/DiscountRate.csv` is **header-only** (just `REGION,VALUE`... actually
  `VALUE` col). It is NOT set in B1_Compiler.py or the Config_*.yaml.
- **REGION set = `GLOBAL`** (single region). Injection config: `t1_confection/Miscellaneous/conversion_format.yaml`
  section `DiscountRate:` (~line 82, `indices:[REGION]`) and `DiscountRateStorage:` (~line 87).

## STEP 1 — SET THE RATE (per-scenario; do NOT touch conversion_format.yaml's global default)
Preferred route (per-scenario, reversible): write the otoole CSV explicitly, then re-run B2 ONLY (B2 uses the
otoole CSVs and does NOT regenerate them; do NOT re-run B1 or it overwrites your edit):
  - `A2_Outputs_Params_otoole/B_Opt_Clipped/DiscountRate.csv`  ->  header + `GLOBAL,0.13`
  - `A2_Outputs_Params_otoole/B_Opt_Clipped/DiscountRateStorage.csv` -> `GLOBAL,<each STORAGE>,0.13`
  (back up both first). If otoole still emits default 0.1 (ignores the CSV), THEN escalate the injection point:
  set the value in the A-O / A2_Output_Params, or as a last resort change conversion_format.yaml's DiscountRate
  `default` to 0.13 (global — only if the per-scenario CSV route provably cannot work; note it in the report).

## STEP 2 — RECOMPILE (B2 only, one at a time)
`python B2_Executing_OG_Model.py --scenarios "B_Opt_Clipped"` (env below). This regenerates the compiled
`Pre_processed_B_Opt_Clipped_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt` and solves it.

## STEP 3 — VERIFY PROPAGATION (the persistence step — do not proceed until it passes)
`grep DiscountRate` the compiled `.txt`. It MUST show 0.13 (explicit `GLOBAL 0.13` or `default 0.13`). If it
still shows 0.10, the edit didn't take -> try the next injection point (STEP 1 escalation) and repeat until 0.13.

## STEP 4-6 — CHECK / SOLVE / COMPARE
4. glpsol --check the `.txt` (`glpsol -m osemosys_fast_preprocessed_storage_delay.txt -d <txt> --check`) -> "successfully generated", no `check[...] failed`.
5. CPLEX solve (B2 already does it; solver is dual simplex + feasopt-off in B2 -> Optimal, base-year backstop 0).
6. Sum `Outputs/TotalDiscountedCost.csv` -> compare to base **2,215,995**. MUST DIFFER (that proves the knob is live).
   DIRECTION: a higher discount rate usually LOWERS total *discounted* cost (future discounted harder) and shifts
   the build mix VRE->firm/fossil. Success = different Sum-TDC + a visible 2050 solar/wind vs coal/gas shift, NOT a sign.

## ON PASS — emit loudly + persist the result
Banner `✅ WACC TEST PASS — B_Opt_Clipped @ 13%`; show compiled rate=0.13, glpsol OK, CPLEX Optimal, backstop 0,
Sum-TDC 10%=2,215,995 -> 13%=Y (Δ, %), 2050 solar/wind vs coal/gas shift. Write `WACC_TEST_RESULT.md` with those
numbers AND the exact edit made (so the fan-out is reproducible). Then RESTORE B_Opt_Clipped to 0.10 (restore the
CSV backups + re-run B2) so the base set stays intact — unless keeping a separate 13% copy is preferred.

## PERSISTENCE / HARD-STOP
On any failure: log (failed -> hypothesis -> fix -> result), fix ROOT cause, retry. Never weaken a check or fake a
pass. ONE hard-stop (escalate, don't fake): if the rate cannot be made to propagate after exhausting injection
points, or the solve is infeasible for a non-obvious reason -> STOP, write everything tried to WACC_TEST_RESULT.md.

## ON PASS -> the full matrix (same mechanism, no new prompt)
7/13% x {B_Opt_Clipped, TradeCap15, TxCap150, DirContractual}; document method+results in OSTRAM_METHODOLOGY.md §8-C.

## ENV
`conda activate OSTRAM-env` (or call `C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe` directly);
`set PYTHONUTF8=1` / `PYTHONIOENCODING=utf-8` / `chcp 65001`. Config_MOMF_T1_AB.yaml: storage_delay_active True,
cplex_threads 4, execute_model True, create_matrix True.
