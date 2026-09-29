# ostram/pipeline/scenarios/transformations — ordered A3 workbook transformations

The steps that `ostram/pipeline/scenarios/transform.py` runs, in a fixed order, to turn the
materialized scenario template and the post-A2 `A-O_*` workbooks into a root scenario's final
workbooks. Most steps are module CLIs launched as `python -B -m <module>` subprocesses inside a
disposable run directory (`<workspace>/scenarios/.../_run_<timestamp>/`); a few are libraries.
Stage contract: `docs/pipeline.md`, "Build selected scenarios (A3)". Scenario rules that run
after these steps live in `../rules/`.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `merge_timeslices.py` | Builds `SOASIA_OSeMOSYS_WV.xlsx`, a working copy of the materialized template with the four timeslice sheets rebuilt from the timeslice workbook and two audit sheets added | `OSTRAM_TEMPLATE_PATH` (materialized template), `<workdir>/OSTRAM_Timeslice_Inputs.xlsx`, `Config_country_codes.yaml::countries` | `<workdir>/SOASIA_OSeMOSYS_WV.xlsx` | `python -B -m ...merge_timeslices` with `OSTRAM_STAGE_WORKDIR` and `OSTRAM_TEMPLATE_PATH` set; top-level script, no flags |
| `extract_ao_extensions.py` | Compares the WV workbook with the four `A-O_*` workbooks and writes the review workbook of technologies to add and parameter rows to replicate | `<workdir>/SOASIA_OSeMOSYS_WV.xlsx`, `A-O_Parametrization.xlsx`, `A-O_AR_Model_Base_Year.xlsx`, `A-O_AR_Projections.xlsx`, `A-O_Demand.xlsx` | `<workdir>/OSTRAM_AO_Extensions.xlsx` | `python -B -m ...extract_ao_extensions` with `OSTRAM_STAGE_WORKDIR`; top-level script |
| `ao_extension_decisions.py` | Overlays the maintained human decisions (`AO_Extension_Decisions` sheet) onto the freshly generated extensions and drops `Interconnector_Params` rows owned by later stages | `ProjectPaths.ao_decisions` (the scenario workbook), `--extensions` workbook, the four `A-O_*` workbooks | The `--extensions` workbook, in place | `python -B -m ...ao_extension_decisions [--extensions PATH] [--authority PATH]` |
| `update_ao_from_extensions.py` | Applies the decided extensions: appends new technology rows, refreshes demand projection and profiles from WV, writes `_wvaligned` copies of the four workbooks | `OSTRAM_AO_Extensions.xlsx`, `SOASIA_OSeMOSYS_WV.xlsx`, the four `A-O_*` workbooks, `ProjectPaths.interconnector_taxonomy` | `<workdir>/wvaligned_outputs/<stem>_wvaligned.xlsx` and `AO_alignment_log.txt` | `python -B -m ...update_ao_from_extensions` with `OSTRAM_STAGE_WORKDIR`; top-level script |
| `apply_manual_fixes.py` | Applies hard-coded, human-curated cell corrections (fuel-code substitutions, `OutputActivityRatio` fixes, technology renames) to the `_wvaligned` workbooks | `<workdir>/wvaligned_outputs/*_wvaligned.xlsx`, `ProjectPaths.interconnector_taxonomy` | `<workdir>/wvaligned_outputs_v2/*_wvaligned_v2.xlsx` and `manual_fixes_log.txt` | `python -B -m ...apply_manual_fixes` with `OSTRAM_STAGE_WORKDIR`; top-level script |
| `propagate_timeslice_fabric.py` | Rewrites the `Yearsplit` and `DaySplit` sheets of the v2 Parametrization from the WV fabric (DaySplit scaled by 365) | `SOASIA_OSeMOSYS_WV.xlsx`, `wvaligned_outputs_v2/A-O_Parametrization_wvaligned_v2.xlsx` | `wvaligned_outputs_v2/A-O_Parametrization_wvaligned_v2_ts20.xlsx` | `python -B -m ...propagate_timeslice_fabric` with `OSTRAM_STAGE_WORKDIR`; no flags |
| `insert_reserve_margin.py` | Adds a `System Parameters` sheet with `ReserveMargin` = 1.15 for 2023-2050 | `--input` workbook | The `--input` workbook, in place (sheet replaced) | `python -B -m ...insert_reserve_margin --input X` |
| `add_max_capacity_investment_rule.py` | MaxCapacityInvestment guard: fills blank (and, with `--fill-zeros`, zero) caps of allowed technologies with 9999 and zeroes the caps of technologies without residual capacity or minimum investment, on `Secondary Techs` | `<input-dir>/A-O_Parametrization.xlsx` | The workbook in place, a `<dir>_PRE_MAXCAP_<ts>` backup folder and `_CHANGES.json` | `python -B -m ...add_max_capacity_investment_rule --input-dir DIR [--sheets ...] [--fill-zeros] [--skip-backup]` |
| `fix_elc_pmode_revert.py` | Reverts `Projection.Mode` to `EMPTY` for the `ELC<region>01` lockout rows that the previous step set to `User defined` | `--input` workbook, `Config_country_codes.yaml::countries` | The workbook in place, or `--output` | `python -B -m ...fix_elc_pmode_revert --input X [--output Y]` |
| `pre_solver_validation.py` | Detects and auto-fixes three infeasibility classes (min-investment above max, max-capacity below residual plus commitments, activity lower limit above capacity) | `--xlsx` workbook (`Secondary Techs`, `Demand Techs`, `Fixed Horizon Parameters`, `Yearsplit`, `Capacities`), `compilation_config::base_year`, profile policy `validation_no_mutation` | The workbook in place, a `.backup_pre_validate_<ts>.xlsx`, `_validation_report.csv` under `executables/<scenario>_0/` | `python -B -m ...pre_solver_validation [--scenario S] [--xlsx X] [--auto-fix-all | --non-interactive | --report-only] [--fail-on-issues]`; also imported by the B1 compiler as a report-only gate |
| `patch_ao_c2a.py` | Sets `CapacityToActivityUnit` = 31.536 for generation and storage technologies that lack one | `--src` workbook (`Fixed Horizon Parameters`), `--tax` taxonomy CSV | `--out` workbook (never the source) | `python -B -m ...patch_ao_c2a [--src X] [--tax CSV] [--out Y] [--dry-run]` |
| `fix_pwrpet_clear.py` | Clears the 2023-2050 `Capacities` cells of `PWRPETBGDXX` so OSeMOSYS falls back to the availability factor | `--input` workbook (`Capacities`) | The workbook in place, or `--output` | `python -B -m ...fix_pwrpet_clear --input X [--output Y]` |
| `fix_trn_residuals.py` | Maintains the 18 interconnectors' commissioning constraints: merges the minimum-investment authority into the minimum rows and applies the RC authority to `ResidualCapacity` (mode `min`, cutoff 2023) | `--input` workbook, `--authority` v18 workbook (`Interconnector_Params`, SHA-256 guarded), the loaders in `interconnector_authority.py`, optional `--override` CSV | `--output` workbook, `--diff-csv`, `--diff-md` | `python -B -m ...fix_trn_residuals --input X --output Y --authority Z [--mode min|max] [--cutoff-year N] [--diff-csv F] [--diff-md F] [--override CSV]` |
| `clear_stale_unbinding_caps.py` | Resets stale flat 2024-2050 `TotalAnnualMaxCapacityInvestment` values of 27 allow-listed solar, wind and hydro technologies to 9999 so the lid rule can write its schedule | `--input` workbook (`Secondary Techs`) | A `_POST_CAP_RESET` sibling workbook | `python -B -m ...clear_stale_unbinding_caps [--input X] [--output Y] [--dry-run]` |
| `cap_trn_to_residual.py` | For the 18 `TRN*` interconnectors sets `TotalAnnualMaxCapacity` = `ResidualCapacity` per year and `TotalAnnualMaxCapacityInvestment` = 9999 for 2023-2050 | `--input` workbook (`Secondary Techs`) | A `_POST_TRN_CAP` sibling workbook | `python -B -m ...cap_trn_to_residual [--input X] [--output Y] [--dry-run]` |
| `sync_og_to_ts20.py` | Propagates the WV timeslice fabric into the A1 CSV inputs and the compilation YAML so B1 sees 20 timeslices | `--wv` workbook (`Yearsplit_Template`, `DaySplit`, `Demand_Profiles`), `--yaml` compilation config | `TIMESLICE`, `DAILYTIMEBRACKET`, `YearSplit`, `DaySplit`, `SpecifiedDemandProfile`, `Conversionls/ld/lh` CSVs under `--og-csvs-dir`; `xtra_scen` keys of `--yaml` | `python -B -m ...sync_og_to_ts20 --wv X --og-csvs-dir D --yaml Y` |
| `scenario_workbooks.py` | Reader/writer for the multi-scenario `OSTRAM_Scenario_Inputs.xlsx`: materializes a flat scenario template, applies inherited `Restrictions`, persists run restrictions | `ProjectPaths.scenario_workbook` (`Control`, `Restrictions` and the parametric sheets), `*_CHANGES.json` logs | Materialized template, `A-O_Parametrization.xlsx` restriction cells, `Restrictions` rows of the disposable run-state workbook | Imported by `transform.py` and `materializer.py`; `python -B -m ...scenario_workbooks list-active [--soasia PATH]` prints active scenarios |
| `support_bau_transmission.py` | Final support-`BAU` ceiling: `TotalAnnualMaxCapacity` of each `TRN` technology = residual capacity plus surviving committed additions | Delivered `A-O_Parametrization.xlsx` (`Secondary Techs`, `Fixed Horizon Parameters`), `TRN_TECHS` from `cap_trn_to_residual.py` | The workbook in place and `support_bau_transmission.json` beside it | Imported only (`transform.finalize_support_bau`, when scenario is `BAU` and profile policy `support_bau_committed_transmission` is true) |
| `interconnector_authority.py` | Fail-closed loaders for the interconnector minimum-investment and clamp-boundary authorities (schema, domain and SHA-256 guards, no numeric fallback) | The v18 workbook sheet `Interconnector_Params` | Nothing | Imported only (`fix_trn_residuals.py`, `../apply_patches.py`) |
| `xlsx_validation.py` | Shared helpers for the upper-bound invariants checked by `pre_solver_validation.py` (indexing, cumulative-capacity sweep, activity lower limit) | An open workbook passed in | Nothing | Imported only (`pre_solver_validation.py`) |

`...` in the Run column stands for `ostram.pipeline.scenarios.transformations`.

## How to run

Never run these steps by hand against maintained inputs. The supported entry points are
`python -m ostram transform --scenario <root>` (one root scenario) and `python -m ostram run`
(stage A3), which create the run directory, set `OSTRAM_STAGE_WORKDIR` / `OSTRAM_TEMPLATE_PATH`
and call the steps below in order. `--keep-workdir` keeps the run directory for inspection.

## Order / dependencies

As executed by `transform.py` (stage functions) and `../orchestrator.py` (`execute_plan`):

0. `scenario_workbooks.materialize_scenario_template` writes the flat template; the four
   post-A2 `A-O_*` workbooks are copied into `stage1/`.
1. Stage 1, cwd `stage1/`: `merge_timeslices` → `extract_ao_extensions` →
   `ao_extension_decisions` → `update_ao_from_extensions` → `apply_manual_fixes` →
   `propagate_timeslice_fabric`.
2. Stage 1b, on copies in `stage1b/`: `insert_reserve_margin` → `add_max_capacity_investment_rule`
   (`--fill-zeros`) → `fix_elc_pmode_revert` → `pre_solver_validation` (`--auto-fix-all`).
3. Stage 2 / 2.5: `patch_ao_c2a` → `fix_pwrpet_clear`.
4. Stage 3: `fix_trn_residuals` (`--mode min --cutoff-year 2023`) → `clear_stale_unbinding_caps`
   → `cap_trn_to_residual`.
5. Stage 4 / 4.5: consolidation into `stage5/`, inherited `Restrictions` applied through
   `scenario_workbooks`.
6. Stage 5: the scenario's rule scripts (`../rules/`), then WS-3 (interconnector costs, internal
   transmission, internal losses) and WS-4 (base-year pin) from `../rules/`.
7. `sync_og_to_ts20`, then `scenario_workbooks.persist_run_restrictions` (Stage 6), delivery, and
   `support_bau_transmission` last, for support `BAU` only.

## Gotchas

- `merge_timeslices`, `extract_ao_extensions`, `update_ao_from_extensions` and
  `apply_manual_fixes` have no `__main__` guard: importing them runs the whole step and requires
  `OSTRAM_STAGE_WORKDIR` (and `OSTRAM_TEMPLATE_PATH` for `merge_timeslices`). Do not import them
  casually.
- `sync_og_to_ts20` writes outside the disposable run directory: it rewrites CSVs under
  `<workspace>/preparation/og_csvs_inputs/` and `xtra_scen` in
  `<workspace>/compilation/Config_MOMF_T1_A.yaml` on every scenario run. `CapacityFactor.csv` is
  deliberately not touched.
- The documented Stage 3 order in `cap_trn_to_residual.py` (FIXED → POST_TRN_CAP →
  POST_CAP_RESET) differs from the executed order (FIXED → POST_CAP_RESET → POST_TRN_CAP); the two
  steps touch disjoint technologies, so the result is the same.
- `add_max_capacity_investment_rule` must stay scoped to `Secondary Techs`; applying it to
  `Primary Techs` would zero every fuel source. Its `--fill-zeros` behaviour reproduces the two
  frozen copies that ran back to back until September 2026 (module docstring, HISTORY).
- `fix_elc_pmode_revert` must run after `add_max_capacity_investment_rule`, which is what sets the
  rows it reverts.
- `pre_solver_validation` derives the scenario name from the workbook's parent folder, so inside
  the chain its `_validation_report.csv` lands under `executables/stage1b_0/`.
- `fix_trn_residuals` validates the authority workbook against a semantic SHA-256 and fails closed;
  reduced profiles (not `full`) skip the contribution and boundary authorities.
- `insert_reserve_margin` hard-codes 1.15 and 2023-2050; `patch_ao_c2a` hard-codes 31.536 and
  never overwrites existing values.
- `persist_run_restrictions` clears all existing `Restrictions` rows of the scenario before
  writing and must only receive the disposable run-state copy, never the maintained workbook
  (`transform.py`).
- Several docstrings still say "Run in Spyder with F5" or carry pre-refactor file names
  (`A0_insert_reserve_margin.py`, `3_update_ao_from_extensions.py`, `_scenarios.py`,
  `B1b_Pre_solver_validation.py`); the module CLIs above are the current form. The `transform.py`
  docstring still lists two `add_max` copies that are now one merged call.
