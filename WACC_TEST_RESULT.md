# ✅ WACC TEST PASS — B_Opt_Clipped @ 13%

**Date:** 2026-07-12 (cleanroom session 2) · Branch `ws3-phaseb-cleanredo` (local only).
**Cell:** `B_Opt_Clipped` DiscountRate + DiscountRateStorage **0.10 → 0.13** (single knob).

## Result banner
| Check | Value |
|---|---|
| Compiled rate | `param DiscountRate := GLOBAL 0.13` + all 20 storages `GLOBAL <STORAGE> 0.13` (verified in the `.txt`) |
| glpsol `--check` | successfully generated, no `check[...] failed` |
| CPLEX | **Dual simplex — Optimal** (feasopt-off), CPLEX objective `1.6256547237e6` |
| Base-year backstop gen | **0** (clean) |
| **Sum TotalDiscountedCost** | **10% = 2,215,995 → 13% = 1,761,993  (Δ −454,002, −20.49%)** |

**The knob is live:** a −20.5% swing in Σ-TotalDiscountedCost proves the discount rate propagates end-to-end.
Direction is correct — a higher discount rate weights future (2027–2050) costs less, so the NPV of total system
cost falls (a 2050 cost is valued ~half as much at 13% vs 10%: `1/1.13^27 = 0.037` vs `1/1.10^27 = 0.076`).

## 2050 build-mix shift (10% → 13%)
| Family (GW) | 10% | 13% | Δ |
|---|--:|--:|--:|
| solar (PWRSPV) | 878.26 | 873.99 | **−4.27** |
| wind (PWRWON/WOF) | 498.03 | 498.03 | 0.00 |
| coal (PWRCOA) | 310.04 | 310.03 | −0.01 |
| gas (PWRNGS) | 90.81 | 90.71 | −0.10 |
| oil (PWROIL) | 57.61 | 58.08 | **+0.47** |
| nuclear (PWRURN) | 32.01 | 31.71 | −0.30 |
| hydro (PWRHYD) | 243.52 | 244.09 | +0.57 |
| **VRE (solar+wind)** | **1376.3** | **1372.0** | **−4.3** |
| CO₂ 2050 (Mt) | 1899.2 | 1906.8 | **+7.6** |

The shift is **directionally correct** (VRE down, oil/hydro up, CO₂ up) but **small**: at 13% the ceiling-clipped
VRE is still far below firm alternatives, so it stays built to nearly the same level. This is the *robustness*
result of Phase-B §5.4 restated through the cost of capital — the least-cost pathway's reliance on cheap VRE is
not fragile to a +3-pt WACC shock; the discount rate mostly re-values (not re-mixes) the system.

## Exact edit made (reproducible)
The rate is not set in B1_Compiler or the Config; the solved `.txt` gets otoole's `default 0.1` because the otoole
`DiscountRate.csv`/`DiscountRateStorage.csv` are header-only, and B1 regenerates them empty. `DiscountRate.csv` is
also a header-only *template*, and with `A2_otoole_outputs: True` B2's `process_scenario_folder` overwrites the
otoole CSV with that empty template every run — so editing only the otoole CSV is clobbered. Injection that works:

1. `A2_Outputs_Params_otoole/B_Opt_Clipped/DiscountRate.csv`  →  `REGION,VALUE` / `GLOBAL,0.13`.
2. `A2_Outputs_Params_otoole/B_Opt_Clipped/DiscountRateStorage.csv` → `REGION,STORAGE,VALUE` + `GLOBAL,<STORAGE>,0.13`
   for all 20 storages (LDS/SDS × 10 nodes).
3. `Config_MOMF_T1_AB.yaml`: `A2_otoole_outputs: True → False` (so B2 uses the edited otoole CSVs and does NOT
   regenerate them from the empty template). Do **not** run B1 (it would regenerate them empty).
4. `python B2_Executing_OG_Model.py --scenarios "B_Opt_Clipped"` (env: OSTRAM-env, PYTHONUTF8=1, chcp 65001).
5. Verify `.txt` shows `GLOBAL 0.13`; glpsol `--check`; CPLEX Optimal; Σ-TDC.

Backups of both CSVs + the config are kept; **B_Opt_Clipped is then restored to 0.10** (config + CSVs reverted,
B2 re-run) so the base 15-scenario set stays intact.

## Fan-out (7%/13% × {B_Opt_Clipped, TradeCap15, TxCap150, DirContractual})
Same mechanism, per scenario. Documented in `OSTRAM_METHODOLOGY.md` §8-C. (Not run in this session — the required
deliverable was the single B_Opt_Clipped @ 13% mechanism proof, now green.)
