# OSTRAM Energy Modeling Pipeline Technical Inventory

Generated from the local repository at `C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_latest`.

This inventory is based on reading the repository tree, every Python file, every YAML file, and workbook sheet names/headers from the main Excel inputs and generated Excel workbooks. It describes what the code currently does, including observed hardcoded paths, defaults, and generated outputs.

## 1. Repository Structure And File Classes

### 1.1 Top-Level Tree

Legend: `[script]` Python executable or helper; `[config]` YAML or process control; `[data]` source/model data; `[template]` reusable model/workbook template; `[output]` generated model output; `[doc]` documentation or guide.

```text
.
|-- .dvcignore [config]
|-- .gitignore [config]
|-- .readthedocs.yaml [config]
|-- dvc.lock [config/output]
|-- dvc.yaml [config]
|-- environment.yaml [config]
|-- LICENSE [doc]
|-- OSTRAM_data.txt [output/model datafile]
|-- OSTRAM_Git_Setup_Guide.html [doc]
|-- README.md [doc]
|-- run.py [script/orchestrator]
|-- concatenate_files/
|   `-- concatenate_ostram.py [script/output concatenation]
`-- t1_confection/
    |-- A0_generate_tech_country_matrix.py [script/preprocessing utility]
    |-- A1_Pre_processing_OG_csvs.py [script/preprocessing]
    |-- A2_AddTx.py [script/preprocessing]
    |-- A3_process.py [script/scenario pipeline]
    |-- B1_Compiler.py [script/compiler]
    |-- B1_Run_Compiler.py [script/scenario compiler runner]
    |-- B2_Executing_OG_Model.py [script/otoole+solver runner]
    |-- D1_generate_editor_template.py [script/editor generator]
    |-- D2_update_secondary_techs.py [script/editor applier]
    |-- inject_DaysInDayType.py [script/datafile patch]
    |-- open_pwrbck_caps.py [script/datafile patch]
    |-- patch_reserve_margin_repair_careful.py [script/datafile patch]
    |-- patch_reserve_margin_repair_careful_xlsx.py [script/datafile patch]
    |-- patch_storage_delay.py [script/datafile patch]
    |-- strip_storage.py [script/datafile patch]
    |-- test_strip_storage.py [script/test]
    |-- Z_AUX_*.py [script/auxiliary analytics, config, repair, visualization]
    |-- Z_generate_country_template.py [script/template generator]
    |-- Z_validate_country_data.py [script/validator]
    |-- Config_country_codes.yaml [config]
    |-- Config_MOMF_T1_A.yaml [config/compiler]
    |-- Config_MOMF_T1_AB.yaml [config/solver]
    |-- Config_region_consolidation.yaml [config]
    |-- Config_tech_equivalences.yaml [config/reference mapping]
    |-- CapacityAndDistances.xlsx [data]
    |-- firm_capacity_fallbacks_by_cr.xlsx [data/config]
    |-- OSTRAM - Electric Generation by Source - Annual.xlsx [data]
    |-- OSTRAM - Installed Capacity by Source - Annual.xlsx [data]
    |-- RateGrowthDemand_RenovabilityGoals.xlsx [data/editor input]
    |-- Shares_PET_OIL_Split.xlsx [data/editor input]
    |-- Shares_Power_Generation_Technologies.xlsx [data/editor input]
    |-- Tech_Country_Matrix.xlsx [data/config]
    |-- A1_Outputs/ [data/output workbooks]
    |-- A2_Extra_Inputs/ [data/templates]
    |-- A2_Output_Params/ [output/B1 CSV parameters]
    |-- A2_Outputs_Params_otoole/ [output/otoole-ready CSV parameters]
    |-- A3_process/ [scripts/config/templates for A3]
    |-- Executables/ [output/datafiles, solver artifacts, scenario CSVs]
    |-- Figures/ [output/visualizations]
    |-- Miscellaneous/ [templates, model, conversion config, preprocess script]
    |-- OG_csvs_inputs/ [data/raw OSeMOSYS CSVs]
    `-- templates/MDV/ [template country additions]
```

### 1.2 Python Scripts Present

```text
concatenate_files/concatenate_ostram.py
run.py
t1_confection/A0_generate_tech_country_matrix.py
t1_confection/A1_Pre_processing_OG_csvs.py
t1_confection/A2_AddTx.py
t1_confection/A3_process.py
t1_confection/A3_process/_build_v18_from_v17.py
t1_confection/A3_process/_scenarios.py
t1_confection/A3_process/_test_scenarios_lite.py
t1_confection/A3_process/_xlsx_validation_core.py
t1_confection/A3_process/1_merge_timeslices_into_WV.py
t1_confection/A3_process/2_extract_ao_extensions.py
t1_confection/A3_process/3_update_ao_from_extensions.py
t1_confection/A3_process/4_apply_manual_fixes.py
t1_confection/A3_process/5_propagate_timeslice_fabric.py
t1_confection/A3_process/6_sync_og_to_ts20.py
t1_confection/A3_process/A0_insert_reserve_margin.py
t1_confection/A3_process/add_max_capacity_investment_rule_NEW_2be1616.py
t1_confection/A3_process/add_max_capacity_investment_rule_OLD_8ee8056.py
t1_confection/A3_process/B1b_Pre_solver_validation.py
t1_confection/A3_process/cap_trn_to_residual.py
t1_confection/A3_process/clear_stale_unbinding_caps.py
t1_confection/A3_process/docs/build_user_guide.py
t1_confection/A3_process/fix_elc_pmode_revert.py
t1_confection/A3_process/fix_pwrpet_clear.py
t1_confection/A3_process/fix_rnwbio_restore.py
t1_confection/A3_process/fix_trn_residuals.py
t1_confection/A3_process/patch_ao_c2a.py
t1_confection/A3_process/rules_scripts/add_max_cap_investment_lid_rule.py
t1_confection/A3_process/rules_scripts/add_max_cap_investment_lid_rule_BACKUP_20260525.py
t1_confection/A3_process/rules_scripts/relax_interconnectors.py
t1_confection/A3_process/rules_scripts/set_min_capacity_floors.py
t1_confection/A3_process/rules_scripts/set_retirement_schedule.py
t1_confection/A3_process/rules_scripts/set_vre_targets.py
t1_confection/B1_Compiler.py
t1_confection/B1_Run_Compiler.py
t1_confection/B2_Executing_OG_Model.py
t1_confection/D1_generate_editor_template.py
t1_confection/D2_update_secondary_techs.py
t1_confection/inject_DaysInDayType.py
t1_confection/Miscellaneous/preprocess_data.py
t1_confection/open_pwrbck_caps.py
t1_confection/patch_reserve_margin_repair_careful.py
t1_confection/patch_reserve_margin_repair_careful_xlsx.py
t1_confection/patch_storage_delay.py
t1_confection/strip_storage.py
t1_confection/templates/MDV/merge_into_inputs.py
t1_confection/test_strip_storage.py
t1_confection/Z_AUX_capital_annualization_script.py
t1_confection/Z_AUX_config_loader.py
t1_confection/Z_AUX_D1b_set_trn_limits_from_flows.py
t1_confection/Z_AUX_fix_excel_profiles.py
t1_confection/Z_AUX_generate_interactive_dashboards_aggregated.py
t1_confection/Z_AUX_generate_RES_diagram.py
t1_confection/Z_AUX_generate_transmission_maps.py
t1_confection/Z_AUX_interconnections_dashboard.py
t1_confection/Z_AUX_sort_csv.py
t1_confection/Z_AUX_united_regions.py
t1_confection/Z_generate_country_template.py
t1_confection/Z_validate_country_data.py
```

### 1.3 YAML Files Present

```text
.readthedocs.yaml
dvc.yaml
environment.yaml
t1_confection/Config_country_codes.yaml
t1_confection/Config_MOMF_T1_A.yaml
t1_confection/Config_MOMF_T1_AB.yaml
t1_confection/Config_region_consolidation.yaml
t1_confection/Config_tech_equivalences.yaml
t1_confection/Miscellaneous/conversion_format.yaml
t1_confection/A3_process/Config_MOMF_T1_A.yaml
t1_confection/A3_process/rules_scripts/configs/A_Calibrated_BAU/bau_calibration.yaml
t1_confection/A3_process/rules_scripts/configs/A_Calibrated_BAU/deprecate/retirement_schedule.yaml
t1_confection/A3_process/rules_scripts/configs/A_Calibrated_BAU/retirement_schedule.yaml
t1_confection/A3_process/rules_scripts/configs/A_Calibrated_BAU/retirement_schedule_-_bau_v2.yaml
t1_confection/A3_process/rules_scripts/configs/B_Optimised_VRE/deprecate/retirement_schedule.yaml
t1_confection/A3_process/rules_scripts/configs/B_Optimised_VRE/lid_rule.yaml
t1_confection/A3_process/rules_scripts/configs/B_Optimised_VRE/lid_rule_B_Optimised_VRE.yaml
t1_confection/A3_process/rules_scripts/configs/B_Optimised_VRE/relax_interconnectors.yaml
t1_confection/A3_process/rules_scripts/configs/B_Optimised_VRE/retirement_schedule.yaml
t1_confection/A3_process/rules_scripts/configs/B_Optimised_VRE/retirement_schedule_-_opti_v2.yaml
t1_confection/A3_process/rules_scripts/configs/C_Target_VRE/lid_rule.yaml
t1_confection/A3_process/rules_scripts/configs/C_Target_VRE/relax_interconnectors.yaml
t1_confection/A3_process/rules_scripts/configs/C_Target_VRE/retirement_schedule.yaml
t1_confection/A3_process/rules_scripts/configs/C_Target_VRE/set_vre_targets.yaml
```

### 1.4 Main Excel Workbooks Present

```text
t1_confection/Miscellaneous/A-O_AR_Model_Base_Year.xlsx [template]
t1_confection/Miscellaneous/A-O_AR_Projections.xlsx [template]
t1_confection/Miscellaneous/A-O_Demand.xlsx [template]
t1_confection/Miscellaneous/A-O_Parametrization.xlsx [template]
t1_confection/Miscellaneous/A-Xtra_Emissions.xlsx [template]
t1_confection/Miscellaneous/A-Xtra_Storage.xlsx [template]
t1_confection/A2_Extra_Inputs/A-Xtra_Battery_Replacement.xlsx [data/template]
t1_confection/A2_Extra_Inputs/A-Xtra_Emissions.xlsx [data/template]
t1_confection/A2_Extra_Inputs/A-Xtra_Projections.xlsx [data/template]
t1_confection/A2_Extra_Inputs/A-Xtra_Storage.xlsx [data/template]
t1_confection/A3_process/A-O_Parametrization_NATY.xlsx [data/reference]
t1_confection/A3_process/A-O_Parametrization_REFERENCE_with_RNWBIO.xlsx [data/reference]
t1_confection/A3_process/OSTRAM_AO_Extensions_FILLED.xlsx [data/template]
t1_confection/A3_process/OSTRAM_Timeslice_Outputs.xlsx [data]
t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v17.xlsx [template]
t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx [template]
t1_confection/A1_Outputs/_post_a2_snapshot_BAU/*.xlsx [snapshot output]
t1_confection/A1_Outputs/A1_Outputs_BAU/*.xlsx [scenario output]
t1_confection/A1_Outputs/A1_Outputs_A_Calibrated_BAU/*.xlsx [scenario output]
t1_confection/A1_Outputs/A1_Outputs_B_Optimised_VRE/*.xlsx [scenario output]
t1_confection/A1_Outputs/A1_Outputs_C_Target_VRE/*.xlsx [scenario output]
```

Each `A1_Outputs/A1_Outputs_<scenario>/` directory contains:

```text
A-O_AR_Model_Base_Year.xlsx
A-O_AR_Projections.xlsx
A-O_AR_Projections_COMPLETED.xlsx
A-O_Demand.xlsx
A-O_Demand_COMPLETED.xlsx
A-O_Parametrization.xlsx
A-O_Parametrization_COMPLETED.xlsx
A-O_Parametrization_Natural_COMPLETED.xlsx
```

Current generated scenarios discovered in B1/B2 output folders are:

```text
A_Calibrated_BAU
B_Optimised_VRE
BAU
C_Target_VRE
```

### 1.5 Raw CSV And Template CSV Parameter Set

The raw `OG_csvs_inputs/` and `Miscellaneous/templates/` folders both contain one CSV per OSeMOSYS set/parameter listed here:

```text
AccumulatedAnnualDemand, AnnualEmissionLimit, AnnualExogenousEmission,
AvailabilityFactor, CapacityFactor, CapacityOfOneTechnologyUnit,
CapacityToActivityUnit, CapitalCost, CapitalCostStorage, Conversionld,
Conversionlh, Conversionls, DAILYTIMEBRACKET, DaysInDayType, DaySplit,
DAYTYPE, DepreciationMethod, DiscountRate, DiscountRateStorage, EMISSION,
EmissionActivityRatio, EmissionsPenalty, FixedCost, FUEL, InputActivityRatio,
MinStorageCharge, MODE_OF_OPERATION, ModelPeriodEmissionLimit,
ModelPeriodExogenousEmission, OperationalLife, OperationalLifeStorage,
OutputActivityRatio, REGION, REMinProductionTarget, ReserveMargin,
ReserveMarginTagFuel, ReserveMarginTagTechnology, ResidualCapacity,
ResidualStorageCapacity, RETagFuel, RETagTechnology, SEASON,
SpecifiedAnnualDemand, SpecifiedDemandProfile, STORAGE, StorageLevelStart,
StorageMaxChargeRate, StorageMaxDischargeRate, TECHNOLOGY,
TechnologyFromStorage, TechnologyToStorage, TIMESLICE,
TotalAnnualMaxCapacity, TotalAnnualMaxCapacityInvestment,
TotalAnnualMinCapacity, TotalAnnualMinCapacityInvestment,
TotalTechnologyAnnualActivityLowerLimit,
TotalTechnologyAnnualActivityUpperLimit,
TotalTechnologyModelPeriodActivityLowerLimit,
TotalTechnologyModelPeriodActivityUpperLimit, TradeRoute, VariableCost,
YEAR, YearSplit
```

## 2. Pipeline Sequence: Raw Inputs To CPLEX Solve

### 2.1 Primary Orchestrator: `run.py`

CLI parameters:

| Argument | Default | Effect |
|---|---:|---|
| `--env-name` | value from `environment.yaml`, else `OSTRAM-env` | Conda environment name. |
| `--env-file` | `environment.yaml` | Conda spec file. |
| `--dvc-file` | `dvc.yaml` | DVC pipeline file used for dependency checks. |
| `--skip-pull` | false | Skips `dvc pull`. |
| `--skip-a3` | false | Skips A3 scenario workbook processing. |
| `--skip-b1` | false | Skips B1 CSV compilation. |
| `--skip-b2` | false | Skips B2 otoole/solver execution. |
| `--scenarios` | all active scenarios | Comma-separated scenario subset. |

Observed logic:

```text
load environment.yaml
ensure conda is on PATH
create/update conda environment if needed
initialize DVC if needed
if DVC remote exists and --skip-pull is false: run dvc pull
if no A1_Outputs/_post_a2_snapshot_* exists:
    run A1_Pre_processing_OG_csvs.py
    run A2_AddTx.py
else:
    skip A1/A2 and use existing snapshot
discover active scenarios via A3_process/_scenarios.py list-active
if --scenarios was supplied: filter scenario list
for each scenario:
    run A3_process.py --scenario <scenario>
run B1_Run_Compiler.py, passing --scenarios when filtered
run B2_Executing_OG_Model.py, passing --scenarios when filtered
```

Important behavior: when a post-A2 snapshot exists, A1 and A2 are not rerun. In the current tree `_post_a2_snapshot_BAU` exists.

### 2.2 DVC Pipeline

`dvc.yaml` defines two stages:

| Stage | Command | Dependencies | Outputs |
|---|---|---|---|
| `preprocess` | `python -u t1_confection/B1_Run_Compiler.py` | `B1_Compiler.py`, `Config_MOMF_T1_A.yaml`, `A1_Outputs/`, `A2_Extra_Inputs/` | `A2_Structure_Lists.xlsx`, `A2_Output_Params/` |
| `executing` | `python -u t1_confection/B2_Executing_OG_Model.py` | `Config_MOMF_T1_AB.yaml`, `osemosys_fast_preprocessed.txt`, `concatenate_files/concatenate_ostram.py`, `Z_AUX_capital_annualization_script.py`, `A2_Output_Params/`, `Miscellaneous/` | `A2_Outputs_Params_otoole/`, `Executables/`, final root `OSTRAM_*.csv` |

The README describes `python run.py` as the intended user command.

### 2.3 Stage A1: `A1_Pre_processing_OG_csvs.py`

Reads:

- `t1_confection/OG_csvs_inputs/*.csv`
- `t1_confection/Miscellaneous/A-O_Demand.xlsx`
- `t1_confection/Miscellaneous/A-O_Parametrization.xlsx`
- `t1_confection/Miscellaneous/A-O_AR_Model_Base_Year.xlsx`
- `t1_confection/Miscellaneous/A-O_AR_Projections.xlsx`
- `t1_confection/Miscellaneous/A-Xtra_Emissions.xlsx`
- `t1_confection/Miscellaneous/A-Xtra_Storage.xlsx`
- `t1_confection/Config_country_codes.yaml`
- `t1_confection/Config_region_consolidation.yaml`
- `t1_confection/Tech_Country_Matrix.xlsx`
- optional `t1_confection/OSTRAM - Installed Capacity by Source - Annual.xlsx`

Writes:

- Updated scenario workbooks in `A1_Outputs/A1_Outputs_<scenario>/`
- Updated `OG_csvs_inputs/*.csv` after normalization/filtering
- Updated `Config_MOMF_T1_A.yaml` pieces for years, conversions, and scenario structure

Parameters and hardcoded values:

- `INPUT_FOLDER = OG_csvs_inputs`
- `OUTPUT_FOLDER = A1_Outputs`
- `MISCELLANEOUS_FOLDER = Miscellaneous`
- `A2_EXTRA_INPUTS_FOLDER = A2_Extra_Inputs`
- `LAST_YEAR = 2050`
- `FIRST_YEAR = get_first_year()` from `Config_country_codes.yaml`
- Replaces country code `JAM` with `BRB`
- PWR code parsing assumes `PWR` + 3-char fuel + 3-char country + optional 2-char region
- NGS unification maps `CCG` and `OCG` toward `NGS` when enabled
- Optional region consolidation is controlled by `Config_region_consolidation.yaml`
- Optional `FORCE_EMPTY_MAX_CAP_INV_PWR` comes from `force_empty_max_capacity_investment_pwr`

Pseudocode:

```text
read all OG CSV files
normalize temporal profiles:
    SpecifiedDemandProfile by REGION/FUEL/YEAR
    YearSplit by YEAR
    DaySplit by YEAR/DAYTYPE/REGION where present
replace JAM with BRB
filter years to FIRST_YEAR..LAST_YEAR
load Tech_Country_Matrix
filter technologies by country matrix when enabled
unify CCG/OCG to NGS when enabled
if region consolidation enabled: aggregate selected countries/regions
clean/merge PWR technologies according to pwr_cleanup_mode
for each A1_Outputs_* scenario folder:
    update demand workbook
    update parametrization workbook
    update emissions workbook
    update AR base-year workbook
    update AR projections workbook
    update storage extra workbook
    update YAML structures
write normalized CSVs back to OG_csvs_inputs
```

### 2.4 Stage A2: `A2_AddTx.py`

CLI:

| Argument | Default | Effect |
|---|---|---|
| `--yaml` | `Config_country_codes.yaml` | Country and generated technology config. |
| `--base` | scenario `A-O_AR_Model_Base_Year.xlsx` | Base-year AR workbook. |
| `--proj` | scenario `A-O_AR_Projections.xlsx` | AR projection workbook. |
| `--param` | scenario `A-O_Parametrization.xlsx` | Parametrization workbook. |
| `--demand` | scenario `A-O_Demand.xlsx` | Demand workbook. |

Reads:

- `Config_country_codes.yaml`
- all `A1_Outputs/A1_Outputs_<scenario>/` folders
- `A-O_AR_Model_Base_Year.xlsx`
- `A-O_AR_Projections.xlsx`
- `A-O_Parametrization.xlsx`
- `A-O_Demand.xlsx`

Writes:

- Modified scenario workbooks in place
- `_post_a2_snapshot_<scenario>` directories. Current code removes an existing snapshot directory before copying.

Key transformations:

```text
for each A1_Outputs_* scenario:
    parse country/region list from Config_country_codes.yaml
    add RNWTRN/RNWRPO/RNWNLI/PWRTRN/TRNRPO/TRNNLI rows
    if enable_dsptrn:
        rewrite TRN interconnection fuels to dispatch-ready/imported ELC codes
        add DSPTRN technologies with two modes
    add demand-side transmission rows to parametrization
    add/update demand workbook rows
    copy scenario folder to _post_a2_snapshot_<scenario>
```

Hardcoded filters:

- `TRN_INTERCONNECTION = ^TRN[A-Z]{5}[A-Z]{5}$`
- `PWRTRN`, `TRNRPO`, `TRNNLI` are not treated as interconnectors by that regex.
- Demand technology parameter list is `CapitalCost`, `FixedCost`, `ResidualCapacity`, `TotalAnnualMinCapacityInvestment`, `TotalAnnualMaxCapacity`.

### 2.5 Stage A3: `A3_process.py`

CLI:

| Argument | Effect |
|---|---|
| `--scenario` | Scenario key to materialize and process. |
| `--soasia` | Override SOASIA template path. |
| `--rules-script` | Override the scenario's rules script list. |
| `--inherit-from` | Override restriction inheritance source scenario. |
| `--input-dir` | Override input workbook directory. |
| `--output-dir` | Override final workbook output directory. |
| `--keep-workdir` | Keep `_run_YYYYMMDD_HHMMSS` work directory. |

Observed important behavior:

- Scenario input is restored from `_post_a2_snapshot_BAU` for each scenario, then scenario differences are applied through SOASIA template restrictions and rule scripts.
- The work directory is `t1_confection/A3_process/_run_<timestamp>`.
- Final workbooks are copied back to the scenario output directory.

Execution order:

```text
0. materialize SOASIA v18 scenario template through _scenarios.py
   set OSTRAM_TEMPLATE_PATH for downstream stage 1
1. copy input workbooks into stage1
2. run fix_rnwbio_restore.py against RNWBIO reference workbook
3. run 1_merge_timeslices_into_WV.py
4. run 2_extract_ao_extensions.py
5. copy OSTRAM_AO_Extensions_FILLED.xlsx to OSTRAM_AO_Extensions.xlsx
6. run 3_update_ao_from_extensions.py
7. run 4_apply_manual_fixes.py
8. run 5_propagate_timeslice_fabric.py
9. copy stage1 outputs to stage1b
10. run A0_insert_reserve_margin.py
11. run add_max_capacity_investment_rule_OLD_8ee8056.py
12. run add_max_capacity_investment_rule_NEW_2be1616.py
13. run fix_elc_pmode_revert.py
14. run B1b_Pre_solver_validation.py --auto-fix-all
15. run patch_ao_c2a.py
16. run fix_pwrpet_clear.py
17. run fix_trn_residuals.py --mode min --cutoff-year 2023
18. run clear_stale_unbinding_caps.py
19. run cap_trn_to_residual.py
20. consolidate final workbooks into stage5
21. inherit scenario Restrictions from SOASIA template via _scenarios.py
22. run scenario rule scripts from A3_process/rules_scripts
23. run 6_sync_og_to_ts20.py
24. persist CHANGES.json into SOASIA v18 Restrictions
25. copy final workbook set to A1_Outputs/A1_Outputs_<scenario>
```

Rule scripts use per-scenario YAML from `A3_process/rules_scripts/configs/<scenario>/`.

### 2.6 Stage B1: `B1_Run_Compiler.py` and `B1_Compiler.py`

`B1_Run_Compiler.py` CLI:

| Argument | Effect |
|---|---|
| `--scenarios` | Optional comma-separated scenario subset. |

`B1_Run_Compiler.py` behavior:

```text
discover A1_Outputs/A1_Outputs_* scenario directories
skip suffixes containing backup, snapshot, pre_experiment, or date-like suffixes
backup Config_MOMF_T1_A.yaml to .bak
for each scenario:
    set xtra_scen.Main_Scenario in Config_MOMF_T1_A.yaml
    run B1_Compiler.py
restore Config_MOMF_T1_A.yaml from .bak
```

`B1_Compiler.py` reads:

- `Config_MOMF_T1_A.yaml`
- scenario `A-O_AR_Model_Base_Year.xlsx`
- scenario `A-O_AR_Projections.xlsx`
- scenario `A-O_Demand.xlsx`
- scenario `A-O_Parametrization.xlsx`
- `A2_Extra_Inputs/A-Xtra_Emissions.xlsx`
- `A2_Extra_Inputs/A-Xtra_Storage.xlsx`
- optional `A2_Extra_Inputs/A-Xtra_Projections.xlsx`
- `OG_csvs_inputs/EMISSION.csv` when `Use_OG_module` is true

`B1_Compiler.py` writes:

- `A2_Output_Params/<scenario>/*.csv`
- `A2_Structure_Lists.xlsx`
- `A-O_Demand_COMPLETED.xlsx`
- `A-O_Parametrization_COMPLETED.xlsx`
- `A-O_Parametrization_Natural_COMPLETED.xlsx`
- `A-O_AR_Projections_COMPLETED.xlsx`

Core compiler logic:

```text
load YAML structure and scenario settings
load workbook sheets
derive sets: REGION, YEAR, TECHNOLOGY, FUEL, EMISSION, MODE_OF_OPERATION, TIMESLICE, STORAGE
compile InputActivityRatio and OutputActivityRatio from AR base-year/projection workbooks
compile demand parameters from demand workbook
compile technology parameters from parametrization workbook
compile emissions from A-Xtra_Emissions
compile storage from A-Xtra_Storage
compile Conversionls/ld/lh from YAML for 20-timeslice mode
compile ReserveMargin from optional System Parameters sheet
apply projection modes:
    Flat
    Yearly percent change
    User defined
    Interpolate to stated end value from projection parameter
    Zero
write one CSV per set/parameter to A2_Output_Params/<scenario>
write completed user-facing workbooks
```

### 2.7 Stage B2: `B2_Executing_OG_Model.py`

CLI:

| Argument | Effect |
|---|---|
| `--scenarios` | Optional comma-separated scenario subset. |

Reads:

- `Config_MOMF_T1_AB.yaml`
- `Config_MOMF_T1_A.yaml`
- `Miscellaneous/templates/*.csv`
- `Miscellaneous/conversion_format.yaml`
- `Miscellaneous/osemosys_fast_preprocessed.txt`
- `A2_Output_Params/<scenario>/*.csv`
- solver datafile patch scripts

Writes:

- `A2_Outputs_Params_otoole/<scenario>/*.csv`
- `Executables/<scenario>_0/<scenario>_0.txt`
- `Executables/<scenario>_0/Pre_processed_<scenario>_0*.txt`
- `Executables/<scenario>_0/<scenario>_0.lp`
- `Executables/<scenario>_0/<scenario>_0.sol`
- `Executables/<scenario>_0/Outputs/*.csv`
- `Executables/<scenario>_0/<scenario>_0_Input.csv`
- `Executables/<scenario>_0/<scenario>_0_output.csv`
- root-level `OSTRAM_Inputs.csv`, `OSTRAM_Outputs.csv`, `OSTRAM_Combined_Inputs_Outputs.csv` and dated copies when scenario concatenation is enabled

Execution order:

```text
load B2 YAML
discover/resolve scenarios
for each scenario:
    copy/fill template CSVs using A2_Output_Params/<scenario>
    normalize CSV column names for otoole
    sort CSVs
    run otoole convert csv datafile
    run Miscellaneous/preprocess_data.py
    run inject_DaysInDayType.py
    optionally run patch_storage_delay.py
    optionally run strip_storage.py
    optionally run open_pwrbck_caps.py
    optionally run reserve margin repair scripts
    concatenate input CSVs for traceability
    optionally create LP matrix with glpsol
    run selected solver
    run otoole results <solver> csv
    concatenate result CSVs with concatenate_ostram.py
after scenarios:
    concatenate all scenario input/output CSVs
    optionally annualize capital
```

CPLEX command assembled by code:

```text
cplex -c
  "set logfile <output_file>.cplex.log"
  "read <output_file>.lp"
  "set threads <cplex_threads>"
  "set randomseed <cplex_random_seed>"
  "set parallel 1"
  "optimize"
  "feasopt all"
  "write <output_file>.feasopt.sol"
  "write <output_file>.sol"
```

GLPK is still used to create the LP matrix when `create_matrix: true`, even when the selected solver is CPLEX.

## 3. Excel Workbook Data Flow Map

### 3.1 Main Workbook Sheet Inventory

`A-O_Parametrization.xlsx` sheets observed:

| Sheet | Type | Main parameters or role | Read/write scripts |
|---|---|---|---|
| `Fixed Horizon Parameters` | parameter data | `CapacityToActivityUnit`, `OperationalLife` | A1 writes, A3 modifies, B1 reads |
| `Primary Techs` | parameter data | `CapitalCost`, `FixedCost`, `ResidualCapacity`, `TotalAnnualMaxCapacity`, `TotalAnnualMaxCapacityInvestment`, `TotalAnnualMinCapacityInvestment`, `AvailabilityFactor`, activity bounds, reserve margin tags | A1 writes, A3 modifies, B1 reads |
| `Secondary Techs` | parameter data | Same technology parameters as Primary for secondary technologies | A1/A2/A3/D2 write, D1 reads, B1 reads |
| `Capacities` | time-slice data | `CapacityFactor` | A1/A3 write, D2 may inspect, B1 reads |
| `Yearsplit` | time-slice data | `YearSplit` | A1/A3 write, B1 reads |
| `DaySplit` | time-slice data | `DaySplit` | A1/A3 write, B1 reads |
| `VariableCost` | mode parameter data | `VariableCost` | A1/A3 write, B1 reads |
| `Other_Techs` | generic parameter data | Miscellaneous technology parameters outside primary/secondary/demand sheets | A1/A3 write, B1 reads |
| `Demand Techs` | parameter data | Demand transmission technology costs/capacities | A2 writes, B1 reads |
| `Vehicle Techs` | transport data | Transport technology parameters when enabled | B1 reads if `Use_Transport` |
| `Vehicle Groups` | transport data | Transport grouping | B1 reads if `Use_Transport` |
| `Transport Fuel Distribution` | transport data | Transport fuel distribution | B1 reads if `Use_Transport` |
| `System Parameters` | system data | `ReserveMargin` | A3 writes, B1 reads when present |

Observed headers for the active BAU `A-O_Parametrization.xlsx`:

| Sheet | Key headers |
|---|---|
| `Fixed Horizon Parameters` | `Tech.Type`, `Tech.ID`, `Tech`, `Tech.Name`, `Parameter.ID`, `Parameter`, `Unit`, `Value` |
| `Primary Techs` | `Tech.ID`, `Tech`, `Tech.Name`, `Parameter.ID`, `Parameter`, `Unit`, `Projection.Parameter`, `Projection.Mode`, yearly columns `2023`..`2049` |
| `Secondary Techs` | Same pattern as `Primary Techs` |
| `Capacities` | `Tech.ID`, `Tech`, `Tech.Name`, `Timeslices`, `Parameter.ID`, `Parameter`, `Unit`, yearly columns |
| `Yearsplit` | `Timeslices`, `Parameter.ID`, `Parameter`, `Unit`, yearly columns |
| `DaySplit` | `DayType`, `DailyTimeBracket`, `Parameter.ID`, `Parameter`, `Unit`, yearly columns |
| `VariableCost` | `Tech.ID`, `Tech`, `Tech.Name`, `Mode.Operation`, `Parameter.ID`, `Parameter`, `Unit`, `Projection.Parameter`, `Projection.Mode`, yearly columns |
| `Other_Techs` | `Application`, `Tech`, `Tech.Name`, `Fuel`, `Parameter`, `Unit`, `Projection.Mode`, yearly columns |
| `Demand Techs` | `Tech.ID`, `Tech`, `Tech.Name`, `Parameter.ID`, `Parameter`, `Unit`, `Projection.Parameter`, `Projection.Mode`, yearly columns |
| `System Parameters` | `Parameter`, `Unit`, yearly columns |

Observed note: in `A1_Outputs/A1_Outputs_BAU/A-O_Parametrization_COMPLETED.xlsx`, the sampled `Yearsplit` and `DaySplit` sheets contained `CapacityFactor`-style headers/rows. This is an inventory observation from workbook sampling, not a claim about intended design.

### 3.2 Demand Workbook

`A-O_Demand.xlsx` sheets:

| Sheet | Maps to | Read/write scripts |
|---|---|---|
| `Demand_Projection` | `SpecifiedAnnualDemand` | A1/A2/D2 write; B1 reads |
| `Profiles` | `SpecifiedDemandProfile` | A1 writes; B1 reads |

Completed workbook `A-O_Demand_COMPLETED.xlsx` has a consolidated `A-O_Demand` sheet generated by B1.

### 3.3 Activity Ratio Workbooks

`A-O_AR_Model_Base_Year.xlsx` and `A-O_AR_Projections.xlsx` sheets:

| Sheet | Maps to | Read/write scripts |
|---|---|---|
| `Primary` | `InputActivityRatio`, `OutputActivityRatio` | A1 writes; B1 reads |
| `Secondary` | `InputActivityRatio`, `OutputActivityRatio` | A1/A2/D2 write; D1 reads interconnectors; B1 reads |
| `Demand Techs` | `InputActivityRatio`, `OutputActivityRatio` | A2 writes; B1 reads |
| `Distribution Transport` | transport AR | B1 reads if transport enabled |
| `Transport` | transport AR | B1 reads if transport enabled |
| `Transport Groups` | transport AR | B1 reads if transport enabled |

### 3.4 Extra Input Workbooks

| Workbook | Sheets | Maps to | Scripts |
|---|---|---|---|
| `A-Xtra_Emissions.xlsx` | `GHGs`, `Externalities` | `EmissionActivityRatio`, `EmissionsPenalty`; `EMISSION` set may be read from OG CSV when `Use_OG_module` | A1 writes, B1 reads |
| `A-Xtra_Storage.xlsx` | `Fixed Horizon Parameters`, `CapitalCostStorage`, `TechnologyStorage` | `StorageLevelStart`, `OperationalLifeStorage`, `CapitalCostStorage`, `ResidualStorageCapacity`, `TechnologyToStorage`, `TechnologyFromStorage` | A1 writes, B1 reads |
| `A-Xtra_Projections.xlsx` | `Projections`, `Projection_Mode` | Projection controls for extra parameters | B1 reads where applicable |
| `A-Xtra_Battery_Replacement.xlsx` | `Replacements` | Battery replacement support data | Auxiliary/input workbook; not part of core B1/B2 path observed |

### 3.5 A3 Template Workbooks

| Workbook | Sheets | Role |
|---|---|---|
| `SOASIA_OSeMOSYS_Template_v18.xlsx` | `README`, `Control`, `Restrictions`, `Fixed_Horizon_Parameters`, `Primary_Techs`, `Secondary_Techs`, `Capacities_CF`, `VariableCost`, `Demand_Projection`, `Demand_Profiles`, `Demand_Techs`, `Emissions`, `Yearsplit_Template`, `DaySplit`, `Interconnectors`, `Interconnector_Params`, `Existing_Generation`, `Planned_Generation`, `Technology_Costs`, `RE_Targets_Policies` | Scenario restriction and template source for A3 |
| `OSTRAM_Timeslice_Outputs.xlsx` | `YearSplit`, country demand/CF sheets for BGD, BTN, INDEA, INDNE, INDNO, INDSO, INDWE, LKA, MDV, NPL, plus `Summary`, `Config` | 20-timeslice demand and capacity factor fabric |
| `OSTRAM_AO_Extensions_FILLED.xlsx` | `1_Extensions_To_Add`, `2_Parameter_Rows_To_Replicate`, `3_Signal_Disagreements` | A3 extension instructions copied to `OSTRAM_AO_Extensions.xlsx` |

## 4. YAML Config Anatomy

### 4.1 `environment.yaml`

| Key | Value/control | Consumed by |
|---|---|---|
| `name` | `OSTRAM-env` | `run.py` |
| `channels` | conda channels | conda |
| `dependencies` | Python and package requirements | conda/pip |

Dependencies include Python 3.10, Git, pandas, numpy, openpyxl, pyyaml, xlsxwriter, glpk, coincbc, and pip packages `dvc` and `otoole>=1.1.1`.

### 4.2 `dvc.yaml`

| Key | Control | Consumed by |
|---|---|---|
| `stages.preprocess.cmd` | B1 command | DVC |
| `stages.preprocess.deps` | B1 dependencies | DVC |
| `stages.preprocess.outs` | B1 tracked outputs | DVC |
| `stages.executing.cmd` | B2 command | DVC |
| `stages.executing.deps` | B2 dependencies | DVC |
| `stages.executing.outs` | B2 tracked outputs | DVC |

### 4.3 `Config_country_codes.yaml`

| Key/block | Control | Consumed by |
|---|---|---|
| `country_data` | Country names and OSTRAM naming metadata for BGD, BTN, IND, NPL, LKA, MDV | config loader, A0, A1, A2, D scripts |
| `special_regions` / `INT` | Special region metadata | config loader |
| `first_year` | Model first year, currently 2023 | A1 |
| `add_missing_countries_from_ostram` | Whether to add countries from OSTRAM data not already in config | A1/A0 helpers |
| `pwr_cleanup_mode` | PWR cleanup behavior; current value is `merge` | A1 |
| `force_empty_max_capacity_investment_pwr` | Forces selected PWR max-investment projection modes to `EMPTY` | A1 |
| `ostram_tech_mapping` | OSTRAM source technology labels to model fuel/tech codes | A0/A1/D scripts |
| `shares_tech_mapping` | SHARES workbook technology names to codes | D2 |
| `renewable_fuels` | Renewable fuel codes such as BIO/HYD/CSP/GEO/SPV/WAS/WON/WOF | D2 and helpers |
| `countries` | Country-region keys: BGD, BTN, INDEA, INDNE, INDNO, INDSO, INDWE, NPL, LKA, MDV | A2 |
| `template_generation` | Template country generation settings, including MDV from LKA | country template helpers |
| `implausible_combinations` | Country/technology combinations to avoid | matrix generation/filtering |
| `RNWTRN`, `RNWRPO`, `RNWNLI`, `PWRTRN`, `TRNRPO`, `TRNNLI`, `DSPTRN` | Generated technology parameter blocks: `CapacityToActivityUnit`, `OperationalLife`, `CapitalCost`, `FixedCost`, `ResidualCapacity`, `TotalAnnualMaxCapacityInvestment` | A2 |
| `enable_dsptrn` | Enables dispatch transmission technology injection | A2, D2 |

### 4.4 `Config_region_consolidation.yaml`

| Key | Control | Consumed by |
|---|---|---|
| `enabled` | Master switch, currently false | A1 |
| `countries` | Mapping from source regions to consolidated region | A1 |
| `aggregation_rules.average_parameters` | Parameters averaged during consolidation | A1 |
| `aggregation_rules.sum_parameters` | Parameters summed during consolidation | A1 |
| `aggregation_rules.disabled_parameters` | Parameters not consolidated | A1 |

### 4.5 `Config_tech_equivalences.yaml`

This YAML is a reference/migration mapping for older LATAM/RELAC technology codes into the current model. It contains:

| Key/block | Control |
|---|---|
| `model_versions` | Old and new model naming context |
| `countries` | Country code mappings |
| `suffix_rules` | Country/region suffix rules |
| `gas_unification` | Gas technology mappings such as CCG/OCG to NGS |
| `direct_mappings` | Direct old-code to new-code relationships |
| `removed_demand_techs` | Demand technologies removed from current structure |
| `unchanged_techs` | Technologies kept as-is |
| `tech_availability` | Country/technology availability |
| `aggregation_rules` | Aggregation behavior |
| `direct_mapping_rules` | Additional mapping rules |

No direct primary-pipeline consumer was found in the main A1/A2/A3/B1/B2 code path.

### 4.6 `Config_MOMF_T1_A.yaml`

This is the B1 compiler and structure configuration. A copy also exists in `A3_process/Config_MOMF_T1_A.yaml`.

| Key/block | Control | Consumed/modified by |
|---|---|---|
| `A1_outputs` | A1 output directory | B1 |
| `A2_extra_inputs` | Extra input directory | B1 |
| `A2_output` | B1 parameter CSV output directory | B1 |
| `A2_output_main_scen` | Main scenario output naming | B1 |
| workbook and sheet-name keys | File/sheet names for demand, AR, parametrization, emissions, storage | B1 |
| `base_year`, `initial_year`, `final_year` | Horizon controls; currently 2023 to 2050 | A1 updates, B1 reads |
| `Use_Transport` | Transport-module switch; currently false | B1 |
| `Use_OG_module` | OG module switch; currently true | B1 |
| `xtra_scen.Main_Scenario` | Scenario compiled by B1 | B1_Run writes, B1 reads, B2 reads |
| `xtra_scen.Other_Scenarios` | Other scenario list | B1 |
| `xtra_scen.Region` | Region set | B1 |
| `xtra_scen.Mode_of_Operation` | Mode set, currently 1 and 2 | B1 |
| `xtra_scen.Season` | Season set, currently 1..4 | B1 |
| `xtra_scen.DayType` | DayType set, currently `[1]` | B1 |
| `xtra_scen.DailyTimeBracket` | Daily time bracket set, currently 1..5 | B1 |
| `xtra_scen.Timeslice` | Timeslice mode, current 20-timeslice `Some` mode | B1 |
| `xtra_scen.Storage` | Storage set, LDS/SDS by country-region | B1 |
| `Conversionls` | Season-to-timeslice conversion values | B1 |
| `Conversionld` | Daytype-to-timeslice conversion values | B1 |
| `Conversionlh` | day-bracket-to-timeslice conversion values | B1 |
| header list keys | Column names expected in Excel sheets | B1 |
| parameter list keys | OSeMOSYS parameters compiled from workbook sheets | B1 |
| projection mode strings | Accepted projection modes | B1 |
| `columns4` | Four-column output formatting helper | B1 |

### 4.7 `Config_MOMF_T1_AB.yaml`

| Key | Control | Consumed by |
|---|---|---|
| `A2_output` | B1 CSV parameter input directory | B2 |
| `A2_output_otoole` | otoole-ready CSV output directory | B2 |
| `Miscellaneous` | Miscellaneous directory | B2 |
| `templates` | CSV template directory | B2 |
| `executables` | Solver work/output directory | B2 |
| `outputs` | otoole result output folder name | B2 |
| `concatenate_folder` | Folder with concatenation script | B2 |
| `otoole_config` | otoole config path/name | B2 |
| `preprocess_data` | preprocessing script name | B2 |
| `osemosys_model` | GNU MathProg model file | B2 |
| `conv_format` | conversion format YAML | B2 |
| `concat_csvs` | concatenation script | B2 |
| `inputs_file`, `outputs_file` | final concatenated file labels | B2 |
| `prefix_final_files` | root final file prefix, `OSTRAM_` | B2 |
| `preprocess_data_name` | preprocessed datafile prefix | B2 |
| `output_files` | solver output suffix | B2 |
| `base_scenario` | base scenario, `BAU` | B2, D1, D2 |
| `solver` | active solver, currently `cplex` | B2 |
| `iteration_time` | time limit for CBC path | B2 |
| `cbc_random_seed`, `gurobi_random_seed`, `cplex_random_seed` | solver reproducibility controls | B2 |
| `cplex_threads`, `gurobi_threads` | solver thread count | B2 |
| `del_files` | whether to delete intermediate files | B2 |
| `only_main_scenario` | scenario selection control | B2 |
| `parallel`, `max_x_per_iter` | parallel execution controls | B2 |
| `A2_otoole_outputs` | generate otoole input CSVs | B2 |
| `write_txt_model` | run otoole csv->datafile | B2 |
| `create_matrix` | create LP matrix through GLPK | B2 |
| `execute_model` | run solver | B2 |
| `reuse_existing_sol` | skip solver if `.sol` exists | B2 |
| `concat_otoole_csv`, `concat_scenarios_csv` | result concatenation controls | B2 |
| `annualize_capital` | run annualization helper | B2 |
| `storage_delay_*` | storage delay patch controls; currently inactive | B2, `patch_storage_delay.py` |
| `strip_storage_*` | strip storage controls; currently active, mode `all` | B2, `strip_storage.py` |
| `open_pwrbck_*` | PWRBCK cap-opening controls; currently active with value 9999 | B2, `open_pwrbck_caps.py` |
| `reserve_margin_repair_*` | old reserve margin repair controls; currently inactive | B2 |
| `reserve_margin_repair_careful_xlsx_*` | workbook-driven reserve margin repair; currently active | B2, `patch_reserve_margin_repair_careful_xlsx.py` |

### 4.8 `Miscellaneous/conversion_format.yaml`

This is the otoole schema. For each set, parameter, and result it defines index columns, dtype, and often default values. B2 passes it to both `otoole convert` and `otoole results`.

Active parameter names in the schema are the same parameter set listed in section 1.5. Result variables include the standard OSeMOSYS output variables converted by otoole, such as `ProductionByTechnology`, `ProductionByTechnologyAnnual`, `TotalCapacityAnnual`, `RateOfActivity`, capacity investment, emissions, cost, storage, and demand/result variables.

### 4.9 A3 Rule YAMLs

| YAML | Keys | Consumed by | Effect |
|---|---|---|---|
| `retirement_schedule.yaml` and scenario variants | `base_year`, optional `fuel_slice`, `age_based.<fuel>.lifetime_years`, `age_based.<fuel>.retirement_profile`, `scheduled`, `exempt` | `set_retirement_schedule.py` | Writes `ResidualCapacity` retirement paths. Current B/C schedules include COA, NGS, PET, NUC and exemptions such as PWRHYD, PWRSPV, PWRWON, PWRBCK. |
| `lid_rule.yaml` and variants | `rule_mode`, `percentage_default`, `percentage_by_year`, `security_factor`, `exempt_prefixes`, `relaxation_schedule` | `add_max_cap_investment_lid_rule.py` | Writes `TotalAnnualMaxCapacityInvestment` limits. B scenario uses proportional mode with security factor 1.1 and relaxation factors through 2050. C scenario uses uniform percentage ranges. |
| `relax_interconnectors.yaml` | `mode`, `headroom_factor`, `overrides` | `relax_interconnectors.py` | Relaxes interconnector investment caps. B uses factor 3.0; C uses factor 2.0. |
| `bau_calibration.yaml` | `warn_on_untie`, `floors`, `ceilings`, each with `cr`, `tech`, `param`, `schedule` | `set_min_capacity_floors.py` | Writes capacity/activity/investment floors and ceilings for calibration. |
| `set_vre_targets.yaml` | `bau_results_path`, `constraint_type`, `max_floor_share`, `targets` | `set_vre_targets.py` | Writes VRE target constraints. Current file contains a hardcoded `bau_results_path` under `C:/Users/ClimateLeadGroup/Desktop/CLG_repositories/OSTRAM/...`, not this workspace. |

## 5. Python Script Transformation Inventory

| Script | Inputs -> logic -> outputs | Key constants, filters, or flags |
|---|---|---|
| `run.py` | Environment/DVC/repo state -> ensure conda/DVC, optional A1/A2, per-scenario A3, B1, B2 -> full pipeline outputs | CLI `--skip-*`, `--scenarios`; skips A1/A2 when `_post_a2_snapshot_*` exists |
| `concatenate_files/concatenate_ostram.py` | Folder of otoole result CSVs -> add scenario/file metadata and concatenate -> scenario output CSV | Used by B2; expects result CSV naming from otoole |
| `A0_generate_tech_country_matrix.py` | OG CSVs and config mappings -> create tech-country availability matrix -> `Tech_Country_Matrix.xlsx` | Produces matrix/reference sheets used by A1 |
| `A1_Pre_processing_OG_csvs.py` | OG CSVs, templates, config YAMLs, tech matrix -> normalize/filter/merge/update workbooks -> A1 scenario workbooks and normalized OG CSVs | `LAST_YEAR=2050`; `JAM` to `BRB`; CCG/OCG to NGS; optional region consolidation |
| `A2_AddTx.py` | Scenario workbooks plus country config -> inject transmission/demand/dispatch tech structures -> updated workbooks and snapshots | `TRN[A-Z]{5}[A-Z]{5}` interconnector regex; `enable_dsptrn` |
| `A3_process.py` | Post-A2 BAU snapshot, SOASIA template, A3 rules -> staged scenario workbook patches -> final scenario workbooks | Workdir `_run_YYYYMMDD_HHMMSS`; `--keep-workdir` |
| `_build_v18_from_v17.py` | v17 SOASIA workbook -> build v18 workbook structure -> `SOASIA_OSeMOSYS_Template_v18.xlsx` | A3 template migration utility |
| `_scenarios.py` | SOASIA Control/Restrictions plus CLI -> list/materialize/apply scenario restrictions -> scenario-specific template/restriction effects | Used by `run.py` and `A3_process.py` |
| `_test_scenarios_lite.py` | Scenario config/template fixtures -> lightweight scenario checks -> console/test output | Test helper |
| `_xlsx_validation_core.py` | Workbooks -> structural validation helpers -> validation messages/fixes | Used by validation/pre-solver code |
| `1_merge_timeslices_into_WV.py` | A-O workbooks plus `OSTRAM_Timeslice_Outputs.xlsx` -> merge 20-timeslice fabric -> stage workbook outputs | Uses `OSTRAM_TEMPLATE_PATH` from A3 |
| `2_extract_ao_extensions.py` | A-O workbooks/template -> extract extension rows/disagreements -> `OSTRAM_AO_Extensions.xlsx` style output | Uses extension instruction sheets |
| `3_update_ao_from_extensions.py` | A-O workbooks plus filled extensions -> add/replicate/update parameter rows -> stage workbooks | Handles row replication and extension application |
| `4_apply_manual_fixes.py` | Stage A-O workbooks -> hardcoded repair rules -> stage workbooks | Manual fix script; keep in A3 sequence |
| `5_propagate_timeslice_fabric.py` | Updated A-O workbooks -> propagate timeslice values consistently -> stage workbooks | Ensures workbook timeslice fabric consistency |
| `6_sync_og_to_ts20.py` | Final A3 workbooks and YAML/CSV structures -> synchronize OG inputs and config to 20 timeslices -> final A3 outputs | Final A3-to-B1 bridge |
| `A0_insert_reserve_margin.py` | Parametrization workbook -> insert/update `System Parameters` ReserveMargin -> workbook | Adds system parameter expected by B1 |
| `add_max_capacity_investment_rule_OLD_8ee8056.py` | Parametrization workbook -> legacy max investment cap repairs -> workbook | Kept in A3 stage before new rule |
| `add_max_capacity_investment_rule_NEW_2be1616.py` | Parametrization workbook -> newer max investment cap repairs -> workbook | Kept in A3 stage before pre-solver validation |
| `B1b_Pre_solver_validation.py` | A3 workbook set -> validate and optional autofix -> workbook/log output | A3 calls `--auto-fix-all` |
| `cap_trn_to_residual.py` | Parametrization workbook -> copy/cap TRN capacity values into residual-capacity style constraints -> workbook | Applied after stale cap clearing |
| `clear_stale_unbinding_caps.py` | Parametrization workbook -> clear stale unconstraining cap values -> workbook | A3 stage 3 |
| `fix_elc_pmode_revert.py` | Parametrization workbook -> repair ELC projection modes -> workbook | A3 stage 1b |
| `fix_pwrpet_clear.py` | Parametrization workbook -> clear PWRPET-specific issue after C2A patch -> workbook | A3 stage 2.5 |
| `fix_rnwbio_restore.py` | Workbooks plus RNWBIO reference -> restore RNWBIO rows/values -> workbook | Runs before time-slice merge |
| `fix_trn_residuals.py` | Parametrization workbook -> repair TRN residual capacities -> workbook | CLI `--mode min --cutoff-year 2023` in A3 |
| `patch_ao_c2a.py` | Parametrization workbook -> patch `CapacityToActivityUnit` in A-O workbook -> workbook | A3 stage 2 |
| `rules_scripts/add_max_cap_investment_lid_rule.py` | Workbook plus `lid_rule.yaml` -> set `TotalAnnualMaxCapacityInvestment` LID rows -> workbook and `CHANGES.json` | CLI `--input-dir`, `--sheets`, `--skip-backup`, `--self-test`, `--restore`, `--restore-from`, `--yaml`; skips TRN interconnector patterns and exempt prefixes |
| `rules_scripts/add_max_cap_investment_lid_rule_BACKUP_20260525.py` | Backup copy of LID rule script -> same purpose as above | Backup file currently present |
| `rules_scripts/relax_interconnectors.py` | Workbook plus `relax_interconnectors.yaml` -> relax interconnector investment caps -> workbook and changes | CLI includes self-test/restore options; uses `TECH_TYPES.csv` where available |
| `rules_scripts/set_min_capacity_floors.py` | Workbook plus `bau_calibration.yaml` -> apply floors/ceilings -> workbook and changes | `PARAM_MAP` maps YAML names to workbook parameters |
| `rules_scripts/set_retirement_schedule.py` | Workbook plus `retirement_schedule.yaml` -> write `ResidualCapacity` retirement paths -> workbook and changes | Age-based and scheduled retirements; exemption list |
| `rules_scripts/set_vre_targets.py` | Workbook plus `set_vre_targets.yaml` and BAU results -> set activity/capacity VRE targets -> workbook and changes | Uses BAU `ProductionByTechnology.csv`; has hardcoded path in YAML |
| `B1_Compiler.py` | Scenario Excel workbooks and A config -> compile OSeMOSYS set/parameter CSVs -> `A2_Output_Params/<scenario>` and completed workbooks | Projection-mode engine; `Use_Transport`; `Use_OG_module` |
| `B1_Run_Compiler.py` | Scenario folders -> temporarily set YAML main scenario and run B1 -> B1 outputs for each scenario | Skips backup/snapshot/date-like dirs |
| `B2_Executing_OG_Model.py` | B1 CSVs, templates, conversion schema, model file -> otoole datafile, patches, LP, solver, results -> scenario/root outputs | Solver switch: glpk/cbc/cplex/gurobi; CPLEX command includes `feasopt all` |
| `D1_generate_editor_template.py` | Scenario `Secondary Techs` sheets and base AR workbook -> generate controlled editor workbook -> `Secondary_Techs_Editor.xlsx` | Reads `base_scenario`; creates dropdown/validation sheets and interconnection controls |
| `D2_update_secondary_techs.py` | `Secondary_Techs_Editor.xlsx`, OSTRAM/SHARES/demand/trade workbooks -> apply edits to scenario workbooks -> updated A-O workbooks and log | Can update residuals, demand, activity bounds, interconnections; creates timestamped backups/log |
| `inject_DaysInDayType.py` | Preprocessed datafile -> inject/repair `DaysInDayType` block -> patched datafile | B2 always runs this after preprocessing |
| `Miscellaneous/preprocess_data.py` | otoole datafile -> clean/preprocess MathProg datafile -> preprocessed datafile | Called by B2 |
| `open_pwrbck_caps.py` | Datafile -> set PWRBCK cap patterns to high value -> patched datafile | B2 active when `open_pwrbck_active: true` |
| `patch_reserve_margin_repair_careful.py` | Datafile -> reserve-margin repair based on text/datafile rules -> patched datafile | Old path inactive by current B2 config |
| `patch_reserve_margin_repair_careful_xlsx.py` | Datafile plus fallback workbook -> reserve-margin repairs -> patched datafile | Current B2 config active |
| `patch_storage_delay.py` | Datafile -> delay storage build/operation for first N years -> patched datafile | Current config inactive |
| `strip_storage.py` | Datafile -> remove/disable storage structures -> patched datafile | Current config active, mode `all` |
| `templates/MDV/merge_into_inputs.py` | MDV template files -> merge into inputs -> updated input/template files | Country template utility |
| `test_strip_storage.py` | Strip-storage fixtures -> test strip behavior -> test output | Local test script |
| `Z_AUX_capital_annualization_script.py` | Final concatenated results -> annualized capital-cost additions -> updated final CSV | Optional B2 post-process |
| `Z_AUX_config_loader.py` | `Config_country_codes.yaml` -> mapping/helper functions -> imported config dictionaries | Shared by A/D/Z scripts |
| `Z_AUX_D1b_set_trn_limits_from_flows.py` | Transmission flow data/results -> derive TRN limits -> workbook/output edits | Auxiliary calibration script |
| `Z_AUX_fix_excel_profiles.py` | Excel profile workbooks -> repair profile formatting/values -> workbook | Auxiliary repair |
| `Z_AUX_generate_interactive_dashboards_aggregated.py` | Result CSVs -> interactive aggregated dashboards -> HTML/figures | Visualization output |
| `Z_AUX_generate_RES_diagram.py` | Model structure/results -> RES diagram -> figure/output | Visualization output |
| `Z_AUX_generate_transmission_maps.py` | Capacity/distance/result data -> transmission map outputs -> figures/maps | Visualization output |
| `Z_AUX_interconnections_dashboard.py` | Interconnection result data -> dashboard -> HTML/figures | Visualization output |
| `Z_AUX_sort_csv.py` | CSV folder/file -> sorted CSV outputs -> sorted files | Utility used by model prep |
| `Z_AUX_united_regions.py` | Region/country data -> region unification outputs -> data/workbook updates | Auxiliary region handling |
| `Z_generate_country_template.py` | Config and base country data -> country template workbook/files -> template outputs | Supports adding new country structures |
| `Z_validate_country_data.py` | Country template/input files -> validation report -> console/report output | Template/input validator |

## 6. OSeMOSYS Parameter Coverage

### 6.1 Parameters Compiled Directly By B1 From Excel/YAML

The B1 compiler actively derives these from scenario workbooks or YAML:

```text
AvailabilityFactor
CapacityFactor
CapacityToActivityUnit
CapitalCost
CapitalCostStorage
Conversionld
Conversionlh
Conversionls
DaySplit
EmissionActivityRatio
EmissionsPenalty
FixedCost
InputActivityRatio
OperationalLife
OperationalLifeStorage
OutputActivityRatio
ReserveMargin
ReserveMarginTagFuel
ReserveMarginTagTechnology
ResidualCapacity
ResidualStorageCapacity
SpecifiedAnnualDemand
SpecifiedDemandProfile
StorageLevelStart
TechnologyFromStorage
TechnologyToStorage
TotalAnnualMaxCapacity
TotalAnnualMaxCapacityInvestment
TotalAnnualMinCapacityInvestment
TotalTechnologyAnnualActivityLowerLimit
TotalTechnologyAnnualActivityUpperLimit
VariableCost
YearSplit
```

B1 also writes model sets such as `YEAR`, `REGION`, `TECHNOLOGY`, `FUEL`, `EMISSION`, `MODE_OF_OPERATION`, `TIMESLICE`, `SEASON`, `DAYTYPE`, `DAILYTIMEBRACKET`, and `STORAGE`.

### 6.2 Parameters Defaulted By Templates/Otoole Schema

These appear in `A2_Outputs_Params_otoole/<scenario>` through `Miscellaneous/templates` and `conversion_format.yaml` even when not materially populated by B1:

```text
AccumulatedAnnualDemand
AnnualEmissionLimit
AnnualExogenousEmission
CapacityOfOneTechnologyUnit
DaysInDayType
DepreciationMethod
DiscountRate
DiscountRateStorage
MinStorageCharge
ModelPeriodEmissionLimit
ModelPeriodExogenousEmission
REMinProductionTarget
RETagFuel
RETagTechnology
StorageMaxChargeRate
StorageMaxDischargeRate
TotalAnnualMinCapacity
TotalTechnologyModelPeriodActivityLowerLimit
TotalTechnologyModelPeriodActivityUpperLimit
TradeRoute
```

Selected defaults in `conversion_format.yaml` include:

| Parameter | Default observed |
|---|---:|
| `AnnualEmissionLimit` | `-1` |
| `AnnualExogenousEmission` | `0` |
| `AvailabilityFactor` | `1` |
| `CapacityFactor` | `1` |
| `CapacityOfOneTechnologyUnit` | `0` |
| `CapacityToActivityUnit` | `1` |
| `CapitalCost` | `0.001` |
| `CapitalCostStorage` | `0.001` |
| `DaysInDayType` | schema default exists, then patched by `inject_DaysInDayType.py` |
| `DaySplit` | `0.0068` |
| `DepreciationMethod` | `2` |
| `DiscountRate` | `0.1` |
| `EmissionActivityRatio` | `0` |
| `EmissionsPenalty` | `0` |
| `FixedCost` | `0` |
| `InputActivityRatio` | `0` |
| `ModelPeriodEmissionLimit` | `-1` |
| `ModelPeriodExogenousEmission` | `0` |
| `OutputActivityRatio` | `0` |
| `ReserveMargin` | `1` |
| `TotalAnnualMaxCapacity` | `-1` |
| `TotalAnnualMaxCapacityInvestment` | `-1` |
| `TotalTechnologyAnnualActivityUpperLimit` | `-1` |
| `VariableCost` | `0.001` |
| `YearSplit` | `0` |

### 6.3 Parameters Absent From Custom Excel Compilation

The following are not compiled from a user-facing Excel sheet by B1 in the current path; they are supplied as defaults/templates or left structurally present:

```text
AccumulatedAnnualDemand
AnnualEmissionLimit
AnnualExogenousEmission
CapacityOfOneTechnologyUnit
DaysInDayType
DepreciationMethod
DiscountRate
DiscountRateStorage
MinStorageCharge
ModelPeriodEmissionLimit
ModelPeriodExogenousEmission
REMinProductionTarget
RETagFuel
RETagTechnology
StorageMaxChargeRate
StorageMaxDischargeRate
TotalAnnualMinCapacity
TotalTechnologyModelPeriodActivityLowerLimit
TotalTechnologyModelPeriodActivityUpperLimit
TradeRoute
```

### 6.4 Results Coverage

B2 converts solver results using `otoole results` and `conversion_format.yaml`. The concatenation step preserves one row per output variable CSV with scenario/file metadata. Downstream scripts explicitly reference outputs such as `ProductionByTechnology.csv`; the full otoole result set is available in each `Executables/<scenario>_0/Outputs/` directory before optional cleanup.

## 7. Known Naming And Operational Conventions

### 7.1 Scenario Naming

- Scenario workbook folders: `A1_Outputs/A1_Outputs_<scenario>`.
- B1 parameter folders: `A2_Output_Params/<scenario>`.
- otoole-ready folders: `A2_Outputs_Params_otoole/<scenario>`.
- solver folders: `Executables/<scenario>_0`.
- final B2 per-scenario names use `<scenario>_0`.
- current scenario names: `BAU`, `A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE`.

### 7.2 Technology And Fuel Code Patterns

Observed code conventions:

| Pattern | Meaning |
|---|---|
| `PWR<fuel><country><region>` | Power generation technology, for example `PWRHYDINDEA`. |
| `TRN<origin_country><origin_region><dest_country><dest_region>` | Interconnector, 13 characters, for example `TRNBGDXXINDEA`. |
| `MIN...` | Mining/extraction style technology. |
| `RNW...` | Renewable or renewable-delivery helper technology family in A2. |
| `SDS...`, `LDS...` | Short/long duration storage technology or storage set labels. |
| `ELC<country><region><suffix>` | Electricity fuel code; suffixes such as `01`, `02`, `03`, `04` distinguish bus/import/dispatch uses. |
| `PWRBCK` | Backstop technology family opened by B2 patch when configured. |

Country-region keys:

```text
BGD, BTN, INDEA, INDNE, INDNO, INDSO, INDWE, NPL, LKA, MDV
```

Single-region countries use region `XX`; India is split into `EA`, `NE`, `NO`, `SO`, and `WE`.

Fuel/technology families observed in config include:

```text
BIO, HYD, CSP, GEO, SPV, WAS, WON, WOF, COA, NGS, PET, OIL, URN/NUC
```

### 7.3 Timestamp And Backup Conventions

- A3 work directories: `_run_YYYYMMDD_HHMMSS`.
- Rule backups and logs are timestamped or `*.bak`.
- D2 logs: `secondary_techs_update_log_YYYYMMDD_HHMMSS.txt`.
- A2 snapshots: `_post_a2_snapshot_<scenario>`.
- Root final output dated copies are generated by B2 for combined input/output CSVs.

### 7.4 Non-Destructive Patching Pattern

Most patch scripts avoid editing the original solver datafile in place. B2 passes the output of each patch as the input to the next patch and keeps suffixes such as:

```text
StorageDelayN5
NoStorage
OpenBCK
RMCarefulXLSX
```

A3 rule scripts also write `CHANGES.json` and support backup/restore flows. Several rule scripts expose:

```text
--skip-backup
--self-test
--restore
--restore-from
--yaml
```

### 7.5 Dry-Run, Self-Test, And Restore Controls

- `run.py`: skip flags for A3/B1/B2 and DVC pull.
- `A3_process.py`: `--keep-workdir`.
- Rule scripts: `--self-test`, `--restore`, `--restore-from`, `--skip-backup`.
- `B2_Executing_OG_Model.py`: `reuse_existing_sol`, `execute_model`, `create_matrix`, `write_txt_model` in YAML.
- `strip_storage.py`: covered by `test_strip_storage.py`.

## 8. External Dependencies And Local Tooling

### 8.1 Python And Conda

`environment.yaml` expects:

```text
Python 3.10
Git >= 2.40
pandas >= 2.1
numpy >= 1.26
openpyxl >= 3.1
pyyaml >= 6.0
xlsxwriter >= 3.2.4
glpk
coincbc
pip: dvc, otoole >= 1.1.1
```

Observed local shell details during inventory:

- `py` is available and reports Python 3.12.0.
- `python` resolves to the Windows Store stub in this shell.
- `conda` was not on the current PATH during inventory.
- `openpyxl 3.1.5` and `yaml` were importable through `py`.

### 8.2 Solvers And Model Tools

B2 supports:

| Solver path | How code invokes it |
|---|---|
| GLPK | `glpsol -m <model> -d <datafile> --wlp <lp> --check`; also direct GLPK solve path |
| CBC | `cbc <lp> randomSeed <seed> randomCbcSeed <seed> -seconds <iteration_time> solve -solu <sol>` |
| CPLEX | `cplex -c ... read <lp> ... optimize ... feasopt all ... write <sol>` |
| Gurobi | `gurobi_cl Threads=<threads> Seed=<seed> ResultFile=<sol> <lp>` |

Observed local commands on PATH:

```text
cplex.exe -> C:\Program Files\IBM\ILOG\CPLEX_Studio2212\cplex\bin\x64_win64\cplex.exe
glpsol.exe -> C:\glpk-4.65\w64\glpsol.exe
```

`cbc` and `otoole` were not found on the current shell PATH during this inventory; they may be available inside the intended conda environment.

## 9. Training-Course Component Map

A course can be organized around these concrete components:

1. Repository orientation: file classes, scenario folders, and generated outputs.
2. A1 raw OSeMOSYS CSV normalization and workbook generation.
3. A2 transmission/demand technology injection.
4. A3 SOASIA template materialization, timeslice merge, workbook patching, and scenario rule YAMLs.
5. B1 Excel-to-OSeMOSYS CSV compilation.
6. B2 otoole conversion, datafile preprocessing, patch chain, LP generation, CPLEX solve, result conversion.
7. Editor workflow with D1/D2 for controlled secondary technology edits.
8. Configuration anatomy: country codes, compiler config, solver config, rule YAMLs, otoole schema.
9. Parameter coverage and default behavior.
10. Debugging and reproducibility: snapshots, backups, self-tests, solver artifacts, and generated logs.

## 10. Notable Observations And Risks

- The current pipeline is snapshot-aware: because `_post_a2_snapshot_BAU` exists, `run.py` skips A1/A2 unless snapshots are removed or the code is changed.
- `A3_process.py` restores scenario inputs from the BAU post-A2 snapshot, then applies scenario-specific restrictions/rules.
- `set_vre_targets.yaml` contains a hardcoded BAU result path outside this workspace.
- `Config_MOMF_T1_AB.yaml` currently selects `solver: cplex`, `create_matrix: true`, `execute_model: true`, `strip_storage_active: true`, `open_pwrbck_active: true`, and `reserve_margin_repair_careful_xlsx_active: true`.
- CPLEX local path suggests IBM CPLEX Studio 22.1.2 is installed.
- The active shell did not expose conda, CBC, or otoole, although the intended environment file specifies them.
- Several generated/output directories are tracked or present locally; running the full pipeline can modify many workbooks, CSVs, solver files, and final root CSVs.
