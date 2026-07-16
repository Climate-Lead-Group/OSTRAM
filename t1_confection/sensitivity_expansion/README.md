<!-- CLG / OSTRAM -->
# OSTRAM Sensitivity Expansion — Task 1 (Design & Validate)

> **Historical scope:** This document records the original three-sensitivity design and
> its former checkout names. The scenarios remain regression-protected, but these are not
> current repository setup or execution instructions. Use explicit `--source-scenario`
> arguments and the 20-scenario inventory in `tests/regression/scenarios.yaml` for current
> audit work.

**CLG · OSTRAM** — three one-at-a-time sensitivities branched from the validated
**B_Optimised_VRE** baseline, over a shared VRE physical-potential ceiling layer.
Built and validated pre-CPLEX. **Task 2 (separate session) runs the solver.**

## The three runs
| Scenario | Lever (one thing changes) | Discharges |
|---|---|---|
| `B_Opt_TradeCap50` | BGD cross-border **imports ≤ 50%** of demand (per-year, split by B_Opt import shares); backstop imports zeroed | high import-share risk (model 62% vs historical 10–16%) |
| `B_Opt_SolarHi10`  | Solar **CapitalCost × 1.10** (all `PWRSPV*`, all years) | "VRE is cheap" fragility |
| `B_Opt_LinkFreeze` | Freeze `TRNBGDXXINDEA` & `TRNNPLXXBGDXX` at 2023 residual | Zixuan's interconnector-disallow ask |

All three also inherit the **shared VRE-ceiling layer**: per-node `TotalAnnualMaxCapacity`
on solar/onshore-wind = `min(atlas, B_Opt MaxCap)` (NISE-2025 solar; NIWE-150 m wind).

## Layout
```
sensitivity_expansion/
  apply_patches.py                 # post-A3 patcher: ceiling layer + run patches -> A1_Outputs_<scen>/A-O
  validate_sensitivity_configs.py  # 7-check pre-run validator
  desk_check.py                    # partial energy/cost balance (pre-CPLEX)
  reference/
    b_opt_baseline.json            # validated B_Opt values (self-contained source of truth)
    vre_ceilings_base.json         # shared ceiling patch (single source)
    vre_ceilings.csv               # documented ceilings + provenance + clip/confound flags
    historical_trade_bounds.csv    # cross-border realism anchor
  reports/
    validation_report.txt / .csv
    desk_check_report.txt / .csv

A3_process/rules_scripts/configs/{B_Opt_TradeCap50,B_Opt_SolarHi10,B_Opt_LinkFreeze}/
    lid_rule.yaml, relax_interconnectors.yaml, retirement_schedule.yaml,
    storage_floors.yaml   (all copied UNCHANGED from B_Optimised_VRE)
    patches.json          (run-specific post-A3 edits)

A1_Outputs/A1_Outputs_<scenario>/A-O_Parametrization.xlsx   (patched, ready for B1)
```

## Design decisions (Luis, Step-0 gate)
1. **Historical build used `OSTRAM_clean` and read validated B_Opt outputs from `OSTRAM_latest`.** That checkout's
   config + A-O inputs are byte-identical to the validated run; only its cached `Executables`
   solve is stale (solar 415 vs validated 887). Task 2 should re-solve B_Opt here first to
   confirm it reproduces the anchors (obj 2,113,984; solar 887; wind 507; coal 299).
2. **Ceiling = `min(atlas, B_Opt MaxCap)`** — pure guard (never relaxes B_Opt; preserves
   `PWRWON INDEA/INDNE/NPLXX = 0`).
3. **Atlas enforced on the 3 clips** (LKAXX SPV 16, BGDXX WON 3, MDVXX WON 0) even though
   B_Opt overbuilds them — a deliberate physical-realism correction to the shared base.

## Two locks to set before Task 2
- **Trade-cap level:** default **50%** (built). 30% is the stricter variant — regenerate Run 1
  `patches.json` with `cap_fraction=0.30`.
- **Run-3 link list:** default `TRNBGDXXINDEA` + `TRNNPLXXBGDXX`. If Zixuan named other links
  or a directional (single-mode) restriction, edit Run 3 `patches.json`.

## How to (re)build & validate  (conda env `OSTRAM-env`)
```
set PYTHONIOENCODING=utf-8   &&  chcp 65001
python apply_patches.py --self-test
python apply_patches.py --scenario B_Opt_TradeCap50
python apply_patches.py --scenario B_Opt_SolarHi10
python apply_patches.py --scenario B_Opt_LinkFreeze
python validate_sensitivity_configs.py     # -> reports/
python desk_check.py                       # -> reports/
```
`apply_patches` is idempotent (rebuilds each target from a fresh B_Opt copy), non-destructive
to the source B_Opt A-O, and supports `--restore` and `--self-test`.

## Historical Task 2 recipe (runs CPLEX; not a current all-scenario command)
1. (Recommended) re-solve `B_Optimised_VRE` and confirm it reproduces the anchors.
2. `python B1_Run_Compiler.py --scenarios B_Opt_TradeCap50,B_Opt_SolarHi10,B_Opt_LinkFreeze`
3. `python B2_Executing_OG_Model.py --scenarios ...`  (settings: `strip_storage: False`, `delay: True`)
4. `python ..\tools\analysis\concat_all_scenarios.py`
   - **Note:** new scenarios are not in the v18 Control sheet (which we must not edit). The
     configs here are built for the **patch-B_Opt's-A-O** path (no A3 re-run needed). If Task 2
     instead re-runs A3 per scenario, the v18-Control gating must be resolved first.
