# ostram/pipeline/preparation/secondary_techs — Secondary Techs editor workflow

Three standalone modules that generate the `Secondary_Techs_Editor.xlsx` workbook from the A1
outputs, optionally fill its transmission activity limits from bilateral flow data, and apply the
edited workbook back to the A1 `A-O_*` workbooks. They are package internals, not part of
`python -m ostram run` (`docs/secondary-techs-editor.md`, `docs/auxiliary-tools.md`); run them as
modules, never by file path. The editor workbook is generated state and must not become a
competing input authority.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `editor_template.py` | Builds the editor workbook (`Editor`, `OSTRAM_Config`, `Interconnections`, growth-rate, renewability, weights, scenario-demand and documentation sheets) from every `A1_Outputs_*` Parametrization it finds | `<workspace>/preparation/A1_Outputs/A1_Outputs_*/A-O_Parametrization.xlsx` (`Secondary Techs`), the base scenario's `A-O_AR_Model_Base_Year.xlsx` (`Secondary`), `Config_MOMF_T1_AB.yaml::base_scenario`, `Config_country_codes.yaml` | `<workspace>/preparation/Secondary_Techs_Editor.xlsx` | `python -B -m ostram.pipeline.preparation.secondary_techs.editor_template`, no flags |
| `transmission_limits.py` | Writes `TotalTechnologyAnnualActivityLowerLimit` / `UpperLimit` for `TRN` rows of the editor from estimated bilateral flows (GWh × 0.0036 → PJ; lower = × 0.95, upper = × 1.05; 2023-2025 from data, later years flat from 2025) | `<workspace>/preparation/secondary_techs/flujos_energia_estimados_optimizacion.xlsx` (sheet `Flujos por Interconexion` / `Flows by Interconnection`), `Secondary_Techs_Editor.xlsx` (`Editor`) | `Secondary_Techs_Editor.xlsx` in place, years 2023-2050 | `python -B -m ostram.pipeline.preparation.secondary_techs.transmission_limits`, no flags |
| `update.py` | Applies the editor's `OSTRAM_Config` flags and instructions to the A1 workbooks: residual capacities, PET/OIL split, demand growth, activity limits, trade-balance demand adjustment (last), interconnections | `Secondary_Techs_Editor.xlsx`, `inputs/preparation/secondary_technologies/` (installed capacity, generation, `Shares_PET_OIL_Split.xlsx`, `Shares_Power_Generation_Technologies.xlsx`), the bilateral-flow workbook, `A1_Outputs/A1_Outputs_<scenario>/A-O_{Parametrization,Demand,AR_Model_Base_Year,AR_Projections}.xlsx`, `ostram/resources/compilation/conversion_format.yaml`, `Config_MOMF_T1_AB.yaml::base_scenario`, `Config_country_codes.yaml` | The A1 workbooks in place after a `<stem>_backup_<ts>.xlsx` copy; `A1_Outputs/secondary_techs_update_log_<ts>.txt`; optional `A1_Outputs/DEBUG_*_<ts>.csv` | `python -B -m ostram.pipeline.preparation.secondary_techs.update`, no flags; exit 1 if any editor row fails |

## How to run

After A1/A2 have populated `<workspace>/preparation/A1_Outputs/`, from the repo root in
`OSTRAM-env`:

```bash
python -B -m ostram.pipeline.preparation.secondary_techs.editor_template
# edit <workspace>/preparation/Secondary_Techs_Editor.xlsx
python -B -m ostram.pipeline.preparation.secondary_techs.transmission_limits   # optional
python -B -m ostram.pipeline.preparation.secondary_techs.update
```

## Order / dependencies

`editor_template` → (`transmission_limits`) → `update`. `update.py` edits the A1 outputs, so any
later `python -m ostram run` that rebuilds A1/A2 overwrites its changes, and the post-A2 snapshot
that A3 restores does not include them unless A2 is re-run afterwards.

## Gotchas

- The module docstrings say the scripts are "invoked as part of `python -m ostram run`"; nothing
  in `ostram/` calls them. They are manual tools.
- `update.py` hard-codes its scenario list (`BAU`, `NDC`, `NDC+ELC`, `NDC_NoRPO`), which does not
  match the current registry roots; scenarios without an `A1_Outputs_<scenario>` folder are logged
  as failed rows. `editor_template.py` instead discovers scenarios from the folders present.
- Reference year is 2023 with linear growth `Demand(year) = Demand(2023) × (1 + rate × (year − 2023))`;
  non-base scenarios reuse the base scenario's shares for 2023-2025, so policy effects start in
  2026 (`editor_template.py` documentation strings, `update.py`).
- The bilateral-flow workbook is expected under `<workspace>/preparation/secondary_techs/`, which
  is not a tracked input; both consumers print an error or disable the feature when it is absent.
- `transmission_limits.py` matches interconnectors at country level (`tech_code[3:6]` and
  `[8:11]`), ignoring the region letters of multi-region codes, and uses a hard-coded 20-country
  name map instead of the YAML.
- `update.py` is not idempotent: every run creates new backups and re-edits already edited
  workbooks. Missing OSTRAM source workbooks disable the corresponding feature with a warning
  rather than an error.
- The lower activity limit is capped by what `TotalAnnualMaxCapacity` can produce, not
  `ResidualCapacity` (`update.py`).
