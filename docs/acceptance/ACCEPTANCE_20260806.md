# Acceptance record — Track 2 re-anchor, 2026-08-06

This record adopts the Track 2 re-anchor campaign results as the accepted solver
baseline for the 15 governed decision scenarios. It closes evidence gap **I2**:
before this campaign the 12 derived scenarios had never been solved on the
governed 2026-08-04 compiled inputs.

## Why the objectives are recorded here

`docs/regression.md` establishes that the external
`STAGE_2_GOVERNED_COMPARATOR_MANIFEST.csv` is the only *compiled-input*
acceptance authority, and that housekeeping never reruns a solver. That policy
is unchanged by this record. It leaves *solver-level* results without a tracked
home: `tests/regression/reports/accepted_compiled_solver_baseline_15.json` was
deliberately removed in `b5845ff` ("docs: finalize lean runtime guidance"), and
no manifest generator survives at HEAD.

This document is therefore the tracked record of the accepted objectives. It is
narrative evidence in the same spirit as
`examples/unescap/references/provenance.md`, not a test input — no code reads
it, and it introduces no new validator binding.

## Provenance

| item | value |
|---|---|
| repository | `OSTRAM_main` (Climate-Lead-Group/OSTRAM), branch `main` |
| git HEAD | `eaf3ae4046ecfc3595970c9b5a12e58a5bf2fae5` |
| HEAD subject | Merge pull request #30 from Climate-Lead-Group/fix/unescap-training-foundation |
| results capsule | `OSTRAM_track2_capsule_20260806T223122Z` |
| capsule manifest | `CAPSULE_MANIFEST.csv`, 107,476 bytes, 651 rows |
| capsule manifest SHA-256 | `09f9b1b9177ee9bc0731a41495318241eb7adfb1ef7aecf6eff2c99c2137fa1e` |
| capsule verdict | `VERIFIED` — 612 scenario files re-verified against source, no mismatches |
| solver | IBM CPLEX 22.1.2.0 |
| determinism | `cplex_random_seed: 12345`, `parallel: False`, `reuse_existing_sol: False`, `PYTHONHASHSEED=0` |
| entrypoint | `python -m ostram run` with `config/execution/Config_MOMF_T1_AB.yaml` |

The capsule is a results capsule, not an authenticated governance capsule. It
carries no `CAPSULE_AUTHENTICATION.json` and does not supersede
`OSTRAM_final_validation_capsule`.

## Byte gate — compiled inputs

Compiled inputs were gated against the governed Stage 2 comparator manifest
(SHA-256 `9c5c01526049d38cdfe9cedb0505c10ead1b09a83514a7877f98495620617aab`)
**twice**: before any solve and again after all 15 solves. Comparison is exact
SHA-256, byte size and line count, with no normalization, tolerance or waiver.

| gate | scenarios | matched | verdict |
|---|---|---|---|
| pre-solve (`stage0b_byte_gate.json`) | 15 | 15 | **PASS** |
| post-solve (`post_solve_byte_gate.json`) | 15 | 15 | **PASS** |

Running the gate at this commit was mandatory: HEAD is 13 commits past the
commit the 2026-08-04 diagnostic verified (`8636ccc`), and those commits modify
production code on the A3/B1/B2 paths.

## Accepted objectives

The three roots reproduce their previously accepted objectives **exactly** (zero
absolute difference). Their compiled inputs are byte-identical between the
accepted 2026-07-28 record and the governed 2026-08-04 vintage, so with a fixed
solver seed the objectives are identities, not a tolerance band. The 12 derived
scenarios are accepted here for the first time on governed bytes.

| scenario | accepted objective | syscost | LP constant term | wobble |
|---|---|---|---|---|
| `A_Calibrated_BAU` | 2148526.8396765944 | 2312838.471174838 | 164648.790588061 | 337.1590898174 |
| `A_Calibrated_BAU_Clipped` | 2148528.1032923907 | 2312839.7348966096 | 164648.790588061 | 337.1589838421 |
| `B_Optimised_VRE` | 2049837.7893340178 | 2214139.6638268325 | 164648.790588061 | 346.9160952463 |
| `B_Opt_Clipped` | 2050917.9308471703 | 2215219.9058978604 | 164648.790588061 | 346.8155373709 |
| `B_Opt_DirBidir` | 2050917.9308471703 | 2215219.9058978604 | 164648.790588061 | 346.8155373709 |
| `B_Opt_DirContractual` | 2059418.8199527734 | 2223721.072800284 | 164648.790588061 | 346.5377405504 |
| `B_Opt_IndiaCosts` | 2016600.8454936626 | 2176523.261959734 | 160268.875109329 | 346.4586432576 |
| `B_Opt_IndiaCostsFuel` | 2016600.8454936626 | 2176523.261959734 | 160268.875109329 | 346.4586432576 |
| `B_Opt_SolarCapex130` | 2072581.2848773515 | 2236883.333992209 | 164648.790588061 | 346.7414732035 |
| `B_Opt_SolarCapexHi` | 2058144.0963817907 | 2222446.078165668 | 164648.790588061 | 346.8088041837 |
| `B_Opt_SolarCapexSpike` | 2055804.580399432 | 2220106.564716033 | 164648.790588061 | 346.806271460 |
| `B_Opt_TradeCap15` | 2059013.5452385298 | 2223318.2308652657 | 164648.790588061 | 344.1049613251 |
| `B_Opt_TxCap150` | 2074302.7072198503 | 2238607.733166456 | 164648.790588061 | 343.7646414553 |
| `C_Target_VRE` | 2093472.4534030464 | 2257774.634617115 | 164648.790588061 | 346.6093739924 |
| `C_Target_VRE_Clipped` | 2081516.5132870669 | 2245818.9978743093 | 164648.790588061 | 346.3060008186 |

`syscost` is the sum of `TotalDiscountedCost`. The offset between `syscost` and
the LP objective is the LP constant term — the residual-fleet fixed O&M the LP
drops as a constant — so `syscost` is the correct system-cost metric.

## Verification checks

All checks pass. Values are from the campaign's `VERIFICATION.md`.

| # | check | verdict | detail |
|---|---|---|---|
| 0 | compiled inputs byte-identical to the governed manifest | **PASS** | 15/15 |
| 1 | `obj(B_Opt_IndiaCosts) == obj(B_Opt_IndiaCostsFuel)` | **PASS** | both 2016600.8454936626, difference 0, from *different* inputs |
| 2 | `obj(B_Opt_DirBidir) == obj(B_Opt_Clipped)` | **PASS** | both 2050917.9308471703, difference 0 |
| 3 | `PWRBCK*` activity is zero in every scenario | **PASS** | 15/15 zero; backstop present but unused in all inputs |
| 4 | `syscost == LP objective + LP constant term` less documented wobble | **PASS** | wobble range 337.1589838421–346.9160952463, inside the documented 330–350 |
| 5 | India delta on the new vintage vs July | **PASS** | new −38696.6439 (−1.7469 %), July −38537.3340 (−1.7391 %), drift 0.0078 pp |
| 6 | the 3 roots reproduce their accepted objectives exactly | **PASS** | zero absolute difference at 1e-6 relative tolerance |

Additional structural check: `DiscountRate` and `DiscountRateStorage` are `0.1`
by default with zero overrides across all 15 solver inputs.

## Numerical quality

Three of 15 scenarios returned `optimal with unscaled infeasibilities`
(`A_Calibrated_BAU_Clipped`, `B_Opt_TradeCap15`, `B_Opt_TxCap150`). The residual
is dual-only; all 15 are primal feasible within 1e-06.

| vintage | scenarios flagged |
|---|---|
| **Track 2 (this record)** | **3 / 15** |
| accepted 2026-07-28 | 5 / 15 |
| July `OSTRAM_mainredo` | 8 / 15 |

This is an improvement on both comparators. Tracked as issue **I10**.

## Provenance disclosure — interim configuration edit

One deviation from "verify, don't edit" occurred during the campaign, under
explicit user authorisation, and is disclosed here in full.

`concat_scenarios_csv` was set to `False` mid-campaign. Because the campaign
driver invoked `python -m ostram run` once per scenario, each invocation
re-concatenated every scenario already on disk; the measured cross-scenario
concat phase grew from 183.5 s at one scenario to 754.5 s at fourteen.

| variant | SHA-256 | `concat_scenarios_csv` | scenarios run under it |
|---|---|---|---|
| A (original, 6,495 B) | `a61382e3ad6c2840aae6eccbc22df4dc9efdf5aa6f0182422ee92934108982a6` | `True` | 1–7 |
| B | `cc31c3bc6d207a8479b0221bd2bb86125dfd0db5e024be45b8de8866bec8dc38` | `False` | 8–13 |
| A (restored) | `a61382e3…`, byte-identical to the original | `True` | 14–15 |

Scenario 14 ran under the **restored** variant A, proven by its 754.5 s concat
time. `concat_otoole_csv` remained `True` throughout — it produces the
per-scenario `TotalDiscountedCost.csv` on which `syscost` and checks 3, 4 and 5
depend; disabling it would have silently broken the verification.

Concatenation is pure post-processing strictly downstream of the `.sol`: it does
not touch the compiled inputs, the LP, the solve, the solver status, the
objective, or the per-scenario result CSVs. The claim is falsifiable and was
checked — the post-solve byte re-gate above re-hashed every compiled input
against the governed manifest after solving, so a config-induced input change
could not have passed unnoticed. The working tree was clean at the start and end
of the campaign; nothing was committed, staged or pushed from it.

## Scope of this acceptance

Adopted: the 15 objectives above as the accepted solver baseline on governed
2026-08-04 compiled inputs, at HEAD `eaf3ae4`.

Not adopted, and unchanged by this record:

- `OSTRAM_final_validation_capsule` remains the authenticated governance capsule.
- The external governed comparator manifest remains the only compiled-input
  acceptance authority used by `tests/regression/accepted_baseline.py`.
- No protected manifest is published; no such mechanism exists at HEAD.
