# ostram/pipeline/scenarios/rules — scenario rule scripts driven by `config/scenarios/`

Independent patches applied to a root scenario's workbooks late in A3. The Stage 5 chain is the
`rules_script` list of the scenario workbook's `Control` sheet, run left to right by
`../transform.py` as `python -B -m ostram.pipeline.scenarios.rules.<module> --input-dir <stage5>
[--yaml <config/scenarios/<scenario>/<YAML_FILE_NAME>>]`; the WS-3/WS-4 rules run for every root
after that chain, and `set_interconnector_direction.py` runs only as a derived-scenario overlay
from `../materializer.py`. YAML anatomy: `docs/configuration.md`, "A3 Scenario Rule YAMLs".
Each rule edits `A-O_Parametrization.xlsx` (or the AR workbooks) in place and, unless
`--skip-backup`, writes a `<dir>_PRE_<TAG>_<stamp>` backup folder and a `_CHANGES.json` log.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `set_retirement_schedule.py` | Draws down `ResidualCapacity` of thermal generation over time (age-based lifetimes, then scheduled overrides), never touching caps or activity limits | `A-O_Parametrization.xlsx` (`Secondary Techs`), `retirement_schedule.yaml` (`base_year`, `age_based`, `scheduled`, `exempt`), the taxonomy | Workbook in place; `_PRE_RETIRE_` backup and change log | Stage 5, first for A, B and C; `--input-dir D [--yaml F] [--sheets ...] [--skip-backup] [--self-test] [--restore]` |
| `add_max_cap_investment_lid_rule.py` | Applies the per-year lid and untie rule to `TotalAnnualMaxCapacityInvestment` of generation technologies with residual capacity or minimum investment, filling only empty, 0 or 9999 cells | `A-O_Parametrization.xlsx`, `A-O_Demand.xlsx` (`Demand_Projection`), `lid_rule.yaml` (`rule_mode`, `percentage_*`, `security_factor`, `exempt_prefixes`, `relaxation_schedule`, `lid_floor_gw`, ...), the taxonomy, policy `lid_rule_new_semantics` | Workbook in place; `_PRE_LID_` backup and change log | Stage 5 for BAU, A, B and C; same flags as above plus `--restore-from` |
| `set_min_capacity_floors.py` | Writes minimum capacity or generation floors and optional ceilings from exogenous policy schedules (`min_capacity` → `TotalAnnualMinCapacityInvestment` GW, `min_activity` → `TotalTechnologyAnnualActivityLowerLimit` PJ, `max_capacity`, `max_activity`) | `A-O_Parametrization.xlsx`, `bau_calibration.yaml` (`floors[]`, `ceilings[]`), the taxonomy | Workbook in place; `_PRE_BAU_CAL_` backup and change log | Stage 5 for A (third) and C (second); same flags |
| `set_vre_targets.py` | Converts percentage renewable targets into absolute activity and/or capacity floors from a solved `A_Calibrated_BAU` production, with envelopes and pins | `OSTRAM_A_CALIBRATED_BAU_RESULT` or `bau_results_path` (a result CSV, `ProductionByTechnology.csv` or the combined B2 output), `set_vre_targets.yaml` (`targets[]`, `constraint_type`, `max_floor_share`, `pin_generation_to_target`, `activity_lower_caps[]`, ...), `A-O_Parametrization.xlsx`, the taxonomy | Workbook in place; `_PRE_VRE_TARGETS_` backup and change log | Stage 5 for C only; same flags |
| `relax_interconnectors.py` | Relaxes `TotalAnnualMaxCapacityInvestment` of the interconnectors above their residual baseline (`multiplicative`, `additive` or `unconstrained`, with overrides and an optional `capacity_ceiling`), zeroes stale minima and caps `ResidualCapacity` to the ceiling | `A-O_Parametrization.xlsx`, `relax_interconnectors.yaml`, `TECH_TYPES.csv` one level above the script | Workbook in place; `_PRE_TRN_RELAX_` backup and change log | Stage 5 for A, B and C, run from the copy `wd/rules_scripts/relax_interconnectors.py` as `python -B -m rules_scripts.relax_interconnectors` with cwd = run dir; same flags |
| `add_storage_min_investment.py` | Writes small `TotalAnnualMinCapacityInvestment` values for storage technologies (linear interpolation between yearly anchors, per technology, GW) into empty or zero cells | `A-O_Parametrization.xlsx`, `storage_floors.yaml` (`storage_floors: {family: {year: GW}}`, required) | Workbook in place; `_PRE_STOFLOOR_` backup and change log | Stage 5, last for B and C; `--input-dir D [--yaml F] [--skip-backup]` |
| `apply_interconnector_costs.py` | Makes the template's `Interconnector_Params` sheet the source of `CapitalCost`, `FixedCost` and `OperationalLife` for the 13-character `TRN` interconnectors | `--template` / `OSTRAM_TEMPLATE_PATH` workbook (`Interconnector_Params`), `A-O_Parametrization.xlsx` (`Secondary Techs`, `Fixed Horizon Parameters`) | Workbook in place; `_PRE_TRN_COSTS_` backup and change log unless skipped | WS-3 for every root (`--skip-backup`); `--input-dir D [--template F] [--sheets ...] [--skip-backup] [--self-test] [--restore]` |
| `apply_internal_transmission.py` | Calibrates the six internal transmission families per node (costs from `Config_country_codes.yaml::internal_transmission`, flat residuals from the residuals CSV) | `--config` country YAML, `--residuals` `internal_tx_residuals.csv`, `A-O_Parametrization.xlsx` (`Demand Techs`, `Fixed Horizon Parameters`) | Workbook in place; `_PRE_INTERNAL_TX_` backup and change log; optional desk-check CSV | WS-3 for every root; `--input-dir D --config F --residuals F [--dry-run] [--desk-check-csv F] [--skip-backup] [--self-test] [--restore]` |
| `apply_internal_tx_losses.py` | Gives the six internal transmission families a uniform loss through the output activity ratio (`1 − transmission_loss`, default 0.97) in both AR workbooks | `--config` country YAML (`internal_transmission.transmission_loss`), `A-O_AR_Projections.xlsx` (required), `A-O_AR_Model_Base_Year.xlsx` (optional) | Both AR workbooks in place; `_PRE_INTERNAL_LOSS_` backup and change log | WS-4 for every root; `--input-dir D --config F [--skip-backup] [--self-test] [--restore]` |
| `apply_base_year_pin.py` | Materializes the audited base-year generation calibration and retained controls from the frozen, SHA-256-guarded rules table; only complete explicit keys may change a cell | `config/scenarios/rules/pwr_min_2023_2026_pin.csv` (3429 rules; hash from policy `pwr_min_pin_rules_sha256` or the built-in constant), `A-O_Parametrization.xlsx` (`Primary Techs`, `Secondary Techs`, `Capacities`), policy `apply_pwr_min_pin`, the registry | Workbook in place via a `.pwr-min-pin.tmp.xlsx` swap; `_PRE_PWR_MIN_PIN_` backup unless skipped; no change log | WS-4 for `A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE` (`--skip-backup`) and re-applied by `../apply_patches.py`; `--input-dir D --scenario S [--rules-csv F] [--skip-backup] [--restore]` |
| `set_interconnector_direction.py` | Enforces one-way corridors by zeroing one flow mode's activity ratios in the AR workbooks, from a study start year onward | `set_interconnector_direction.yaml` (`directions: {tech: forward|reverse|bidirectional}`), `config/scenarios/technology_types.csv`, `A-O_AR_Projections.xlsx`, `A-O_AR_Model_Base_Year.xlsx` | Both AR workbooks in place; `_PRE_IC_DIR_` backup and change log unless skipped | Overlay from `../materializer.py` for `B_Opt_DirBidir` (no-op) and `B_Opt_DirContractual` (start year 2027); CLI `--input-dir D [--yaml F] [--study-start-year N] [--skip-backup] [--restore]` |

## How to run

Do not run these against maintained inputs. `python -m ostram transform --scenario <root>` runs
the chain inside a disposable run directory; `python -m ostram run` does the same for every
selected scenario. `--self-test` on the rules that offer it exercises them on fixtures.

## Order / dependencies

Stage 5 chains as declared in the scenario workbook's `Control` sheet (verified for the four roots):

- `BAU`: `add_max_cap_investment_lid_rule`
- `A_Calibrated_BAU`: `set_retirement_schedule` → `add_max_cap_investment_lid_rule` →
  `set_min_capacity_floors` → `relax_interconnectors`
- `B_Optimised_VRE`: `set_retirement_schedule` → `add_max_cap_investment_lid_rule` →
  `relax_interconnectors` → `add_storage_min_investment`
- `C_Target_VRE`: `set_retirement_schedule` → `set_min_capacity_floors` →
  `add_max_cap_investment_lid_rule` → `set_vre_targets` → `relax_interconnectors` →
  `add_storage_min_investment`

Then, for every root: `apply_interconnector_costs` → `apply_internal_transmission` →
`apply_internal_tx_losses` → `apply_base_year_pin` (A, B, C only). Later scripts see the earlier
scripts' in-place edits. `--yaml` is passed only when the module declares `YAML_FILE_NAME` and
`config/scenarios/<scenario>/<that file>` exists.

## Gotchas

- `set_vre_targets.py` hard-depends on a completed `A_Calibrated_BAU` solve (its docstring says
  "BAU"); a relative `bau_results_path` resolves against the staged working directory, not the
  repo root (`docs/configuration.md`). `C2AU = 31.536` PJ/GW-year; `max_floor_share` is 0.95 in
  the C configuration.
- `add_max_cap_investment_lid_rule.py` preserves manual values unless policy
  `lid_rule_new_semantics` is true (false in `full`); its `LID_PERCENTAGE_BY_YEAR` comment does not
  match the table values; `A_Calibrated_BAU/lid_rule.yaml` sets `lid_floor_gw: 0.0`.
- `add_storage_min_investment.py`'s docstring says it writes a sibling file; the code saves in
  place.
- `relax_interconnectors.py` must run from the run-directory copy because it looks for
  `TECH_TYPES.csv` one level above the script; MaxCapInv is never set below `ResidualCapacity`, and
  `ResidualCapacity` is capped to the ceiling because OSeMOSYS requires it.
- `apply_base_year_pin.py` accepts only the three decision roots, refuses to run if its
  `.tmp.xlsx` already exists, and validates everything before writing; the rules CSV is CRLF and
  byte-hash guarded, do not normalise it.
- `apply_interconnector_costs.py` requires values constant across years and falls back to a
  template path (`<rules dir>/../OSTRAM_Scenario_Inputs.xlsx`) that does not exist in the current
  layout; the pipeline always passes `--template`.
- `apply_internal_transmission.py` and `apply_internal_tx_losses.py` fall back to a
  `Config_country_codes.yaml` next to the pipeline package that does not exist; the pipeline passes
  `--config` explicitly, and skips the stage with a printed `[SKIP]` if inputs are missing.
- `set_interconnector_direction.py`'s docstring says it runs after `relax_interconnectors` in the
  chain; in the current code it runs only as the materializer's overlay after `apply_patches`.
  With an empty `directions` map it returns `{"skipped": True}`.
- Several docstrings still mention `rules_scripts/configs/<scenario>/`; the code uses
  `config/scenarios/<scenario>/`.
