# ostram/pipeline/preparation — stages A1 and A2 plus country-management tools

A1 (`base_inputs.py`) stages the OSeMOSYS-Global CSV authorities into the preparation workspace,
filters and cleans them, and builds the four `A-O_*` workbooks per root scenario. A2
(`transmission.py`) adds the transmission and dispatch technologies to those workbooks and takes
the post-A2 snapshot that A3 restores. The other modules are operator tools for adding a country
(`python -m ostram country ...`, `python -m ostram scenario ...`) and standalone helpers that are
not part of `python -m ostram run`. Stage contract: `docs/pipeline.md`, "Prepare the base model
(A1) and add the transmission network (A2)"; country workflow: `docs/country-management.md`;
editor workflow: `docs/secondary-techs-editor.md`. All outputs go under
`<workspace>/preparation/`.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `base_inputs.py` | Stage A1: stages `inputs/osemosys_global/*.csv` into `og_csvs_inputs/`, sorts and normalises them, drops years before `first_year`, applies the tech-country matrix, NGS unification, region consolidation and PWR cleanup, then writes the base-model workbooks for every registry root | `inputs/osemosys_global/`, `inputs/preparation/workbook_templates/`, `secondary_technologies/Tech_Country_Matrix.xlsx` and the OLADE capacity workbook, `Config_country_codes.yaml`, `Config_region_consolidation.yaml`, the scenario registry, `config/compilation/Config_MOMF_T1_A.yaml` | `<workspace>/preparation/og_csvs_inputs/*.csv`, `A1_Outputs/A1_Outputs_<root>/A-O_{Demand,Parametrization,AR_Model_Base_Year,AR_Projections}.xlsx`, `extra_inputs/A-Xtra_{Emissions,Storage}.xlsx`, `<workspace>/compilation/Config_MOMF_T1_A.yaml` (runtime copy) | `python -m ostram run` (stage A1); directly `python -B -m ostram.pipeline.preparation.base_inputs`, no flags |
| `transmission.py` | Stage A2: adds `RNWTRN`, `RNWRPO`, `RNWNLI`, `TRNRPO`, `TRNNLI`, `PWRTRN` (and `DSPTRN` when `enable_dsptrn`) per country-region, rewrites PWR output and demand fuels to the tiered `ELC*` scheme, then snapshots each root | `Config_country_codes.yaml` (`countries`, per-prefix parameters, `enable_dsptrn`), the A1 workbooks of every root | The A1 workbooks in place; `A1_Outputs/_post_a2_snapshot_<root>/` (replaced if present) | `python -m ostram run` (stage A2); directly `python -B -m ostram.pipeline.preparation.transmission [--yaml F] [--base F] [--proj F] [--param F] [--demand F]` |
| `configuration.py` | Cached loader for `Config_country_codes.yaml`: countries, names, technology mappings, renewable fuels, flags | `ProjectPaths.country_config` | Nothing | Imported only (A1, A2, validation, templates, `secondary_techs/`) |
| `country_commands.py` | Routes `python -m ostram country {template,merge,validate}` to the three country modules | Nothing | Nothing | `python -m ostram [--profile P] country template|merge|validate ...` |
| `country_templates.py` | Clones a reference country's CSV rows into a new country's template folder, remapping interconnections, and copies its timeslice sheets | `inputs/osemosys_global/*.csv`, `Config_country_codes.yaml::template_generation` and `countries`, `OSTRAM_Timeslice_Inputs.xlsx` | `<workspace>/preparation/country_templates/<CODE>/*.csv`, `centerpoint.csv`, a helper `merge_into_inputs.py`; new `<CODE>_CF` / `<CODE>_Dem` sheets in the timeslice workbook (backup written first) | `python -m ostram country template [--new CC --ref REF] [--region RR] [-i NEIGHBOUR ...] [--lat F --lon F] [--demand-scale F] [--output DIR]` |
| `merge_country_template.py` | Merges a reviewed template into the active CSV inputs and `centerpoints.csv`, with `.bak.csv` backups, and renames every post-A2 snapshot to `_stale_<ts>` so the next run rebuilds A1/A2 | `country_templates/<CODE>/`, the destination CSVs (`og_csvs_inputs/` for `full`, the profile's `osemosys_inputs` otherwise), `centerpoints.csv` | The destination CSVs, `centerpoints.csv`, renamed snapshots | `python -m ostram country merge COUNTRY | --template PATH [--input-dir DIR] [--centerpoints FILE]` |
| `country_validation.py` | Checks that a country has the required sets, parameters, prefixes, fuel patterns and value ranges in the CSV inputs | `ProjectPaths.osemosys_inputs/*.csv`, `Config_country_codes.yaml::country_data` | Nothing (stdout; exit 1 on failure) | `python -m ostram country validate [COUNTRY] [--report] [--quiet]` |
| `scenario_country_sync.py` | Copies a country's technology rows from the post-A2 `BAU` workbook into the prepared scenario workbook (scenario `BAU` only), atomically | `A1_Outputs/_post_a2_snapshot_BAU/A-O_Parametrization.xlsx` (or `--ao`), the prepared scenario workbook | The target workbook in place (via a `.syncing` temp file) | `python -m ostram --profile P scenario sync-country --country CC [--ao F] [--target F] [--dry-run]` |
| `tech_country_matrix.py` | Generates the `Tech_Country_Matrix.xlsx` configuration workbook (YES/NO grid, NGS unification, aggregation rules) with implausible combinations pre-marked | `Config_country_codes.yaml` (`country_data`, `implausible_combinations`) | `<workspace>/preparation/Tech_Country_Matrix.xlsx` | `python -B -m ostram.pipeline.preparation.tech_country_matrix`, no flags; not called by the pipeline |
| `sort_csv.py` | Sorts every CSV in a folder by all columns and overwrites it | A folder path typed at the prompt | The CSVs in place, no backup | `python -B -m ostram.pipeline.preparation.sort_csv` (interactive); not called by the pipeline, A1 has its own copy of the routine |
| `secondary_techs/editor_template.py` | Builds the `Secondary_Techs_Editor.xlsx` workbook from every `A1_Outputs_*` Parametrization | `A1_Outputs/*/A-O_Parametrization.xlsx`, the base scenario's `A-O_AR_Model_Base_Year.xlsx`, `Config_MOMF_T1_AB.yaml::base_scenario`, `Config_country_codes.yaml` | `<workspace>/preparation/Secondary_Techs_Editor.xlsx` | `python -B -m ostram.pipeline.preparation.secondary_techs.editor_template`, no flags; not called by the pipeline |
| `secondary_techs/transmission_limits.py` | Writes `TRN` activity lower and upper limits (flow × 0.95 / × 1.05, GWh to PJ) into the editor from a bilateral-flow workbook | `<workspace>/preparation/secondary_techs/flujos_energia_estimados_optimizacion.xlsx`, `Secondary_Techs_Editor.xlsx` | `Secondary_Techs_Editor.xlsx` in place (years 2023-2050) | `python -B -m ostram.pipeline.preparation.secondary_techs.transmission_limits`, no flags; not called by the pipeline |
| `secondary_techs/update.py` | Applies the editor's instructions (residual capacities, petroleum split, demand, activity limits, trade balance, interconnections) to the A1 workbooks of its scenario list, with backups | `Secondary_Techs_Editor.xlsx`, the four `inputs/preparation/secondary_technologies/` source workbooks, the bilateral-flow workbook, `A1_Outputs/A1_Outputs_<scenario>/A-O_*.xlsx`, `conversion_format.yaml`, `Config_MOMF_T1_AB.yaml::base_scenario` | The A1 workbooks in place (`<stem>_backup_<ts>.xlsx` first), `A1_Outputs/secondary_techs_update_log_<ts>.txt`, optional `DEBUG_*.csv` | `python -B -m ostram.pipeline.preparation.secondary_techs.update`, no flags; not called by the pipeline |

## How to run

From the repo root in `OSTRAM-env`:

```bash
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU   # A1 + A2 (+ A3, B1) into the workspace
python -m ostram country template --new MMR --ref BGD -i BGD IND               # then merge, validate
python -m ostram country merge MMR
python -m ostram country validate MMR
python -m ostram --profile unescap scenario sync-country --country MMR         # after A1/A2 have re-run
```

The runner skips A1 and A2 when every selected root already has its `_post_a2_snapshot_<root>`
(`ostram/pipeline/orchestration.py`); `country merge` invalidates those snapshots on purpose.

## Order / dependencies

`base_inputs.py` (A1) → `transmission.py` (A2, requires every root's `A1_Outputs_<root>/`) →
snapshot → A3. Country workflow: `country template` → review the CSVs and add the country to
`Config_country_codes.yaml` → `country merge` → `country validate` → `python -m ostram run`
(rebuilds A1/A2) → `scenario sync-country`. Editor workflow: `editor_template` →
optional `transmission_limits` → `update`, all after A1/A2.

## Gotchas

- A1 converts OLADE data with 1 GWh = 0.0036 PJ, uses `LAST_YEAR = 2050`, drops rows before
  `first_year`, and replaces the country code `JAM` with `BRB` (`base_inputs.py`). Errors inside a
  single workbook builder are printed and do not stop the run.
- `base_inputs.py` and `transmission.py` have authorship-only module docstrings; their behaviour
  above is taken from the code.
- A2 edits the A1 workbooks in place and then snapshots them; there is no explicit guard against
  running A2 twice on already-processed workbooks. Rely on the runner's snapshot check.
- `tech_country_matrix.py` writes to `<workspace>/preparation/Tech_Country_Matrix.xlsx`, but A1
  reads `inputs/preparation/secondary_technologies/Tech_Country_Matrix.xlsx`
  (`base_inputs.py` line 51). Nothing copies one to the other; the tracked copy is the input.
- The three `secondary_techs/` docstrings say they are "invoked as part of `python -m ostram run`",
  but nothing in `ostram/` calls them. `update.py` hard-codes
  `["BAU", "NDC", "NDC+ELC", "NDC_NoRPO"]` as its scenario list, which does not match the current
  registry, and it edits A1 outputs directly, so they diverge from the post-A2 snapshot.
- `transmission_limits.py` and `update.py` look for the bilateral-flow workbook under
  `<workspace>/preparation/secondary_techs/`, which is not a tracked input.
- `country_templates.py` also modifies `OSTRAM_Timeslice_Inputs.xlsx` (adding `<CODE>_CF` /
  `<CODE>_Dem` sheets after a timestamped backup); that is a write under `inputs/` when the full
  profile is active.
- `country validate` checks `ProjectPaths.osemosys_inputs`, while `country merge` under the full
  profile writes to `<workspace>/preparation/og_csvs_inputs/`; under the full profile the two may
  not see the same data.
- `scenario sync-country` refuses to run without an activated profile workspace and only ever
  writes rows tagged `BAU`.
- `configuration.py` resolves its YAML path at import time and caches it; profile activation must
  precede the import, and a YAML edit is not seen within one process.
- `sort_csv.py` overwrites files in place with no backup.
