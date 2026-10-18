# Glossary

Each term below has one meaning in this documentation. The table at the end
maps the terms to the names used in the code.

| Term | Meaning |
|---|---|
| A-O workbooks | The four Excel workbooks that carry one scenario's model inputs from A1 to B1: `A-O_AR_Model_Base_Year.xlsx`, `A-O_AR_Projections.xlsx`, `A-O_Demand.xlsx` and `A-O_Parametrization.xlsx`. |
| A-result seed | The folder of a solved `A_Calibrated_BAU` result. `C_Target_VRE` reads it to size its generation floors. It is passed with `--a-result-seed`. |
| Adequacy top-up | The last step of the build-cap rule. When the build caps of a country-region in one year add up to less than a set share of its demand-scaled fleet, every cap is multiplied by one factor so that the sum reaches that share. |
| Backstop | A very expensive generation technology, `PWRBCK` followed by the country-region, that the solver uses only when the other limits leave demand unmet. |
| Base-year pin | The A3 rule that writes reviewed values for 2023 to 2026 from `config/scenarios/rules/pwr_min_2023_2026_pin.csv`. It runs after the shared rules of `A_Calibrated_BAU`, `B_Optimised_VRE` and `C_Target_VRE`, the roots listed in `PWR_MIN_PIN_ROOT_SCENARIOS`, and again after the scenario edits of a derived scenario. Support `BAU` does not run it. |
| Build cap | The most new capacity of one technology that the model may build in one country-region in one year. Model parameter `TotalAnnualMaxCapacityInvestment`. |
| Cap multiplier | A factor on the build cap that rises over time. |
| Country-region | One of the ten nodes of the South Asia model: BGDXX, BTNXX, INDEA, INDNE, INDNO, INDSO, INDWE, LKAXX, MDVXX and NPLXX. File and column names abbreviate it as `cr`. |
| Decision scenario | A scenario whose results the study compares. The decision scenarios are listed under `decision_scenarios` in `config/setups/south_asia.yaml`. |
| Demand-scaled fleet | The residual capacity of a country-region multiplied by the growth of its demand since 2024. The build-cap rule computes it. |
| Derived scenario | A scenario built from a root's finished A3 output by the shared VRE ceilings, its scenario edits, the base-year pin and, for the two direction variants, an overlay. Its parent is always a root. |
| Edit log | The JSON file written beside a derived scenario's A-O workbooks. It lists every cell changed, every row created and every edit skipped. Its name is `apply_patches_CHANGES_<stamp>.json`. |
| Family | The three-letter technology family code in characters 4 to 6 of a technology name, for example COA in `PWRCOAINDNO`. |
| Family limit | The largest cap multiplier that a technology family may take. A family limit of 1.0 or less makes a locked family. |
| Gas and oil allowance | The model-file adjustment that replaces a build cap of 0 or 9999 for gas and oil plants with 0.5 GW per year in each country-region. |
| Locked family | A family with a family limit of 1.0 or less. It keeps its proportional build cap and receives no technology top-up. |
| Margin factor | The factor 1.1 in the build-cap formula. It leaves 10 per cent to spare above each technology's share of the required growth. It is a model margin and has no link to energy security. |
| Model-file adjustments | The B2 changes to the solver datafile and the model file after the otoole conversion: days in day type, the storage delay, backstop opening and the gas and oil allowance. |
| Overlay | A YAML file under `config/scenarios/<scenario>/` that sets each listed interconnector to forward, reverse or bidirectional flow. It is applied after the scenario edits of `B_Opt_DirBidir` and `B_Opt_DirContractual`. Also called the direction overlay. |
| Parent | The root scenario from which a derived scenario is built. |
| Setup | A named set of inputs, configuration and scenario registry. `south_asia` is the South Asia model and `unescap_training` is the two-region training model. `config/setups/registry.json` maps each name to its setup file. |
| Reserve margin | The requirement that the capacity tagged as firm exceeds demand by a set fraction. OSeMOSYS parameter `ReserveMargin`. |
| Residual capacity | The capacity that exists in the base year and survives in each later year after retirements. Model parameter `ResidualCapacity`. |
| Root scenario | A scenario built by A1, A2 and the A3 rules named in the Control sheet of the scenario workbook. The roots are `BAU`, `A_Calibrated_BAU`, `B_Optimised_VRE` and `C_Target_VRE`. |
| Scenario edits | The ordered changes in a derived scenario's file `config/scenarios/<name>/scenario.yaml` that build the scenario from its parent's A-O workbooks. With `reuse_edits_from`, the edits of the named scenario come first. |
| Scenario registry | The `scenario_registry` section of `config/setups/south_asia.yaml`: the roots and the decision scenarios. Each derived scenario's parent and overlay are in its own file, `config/scenarios/<name>/scenario.yaml`. |
| Shared VRE ceilings | Twenty fixed values of `TotalAnnualMaxCapacity` for solar and onshore wind technologies, applied in every year to every derived scenario. Section `vre_ceilings` of `config/setups/south_asia.yaml`. |
| Solver datafile | The final text file that `glpsol` reads: `Executables/<S>_0/Pre_processed_<S>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt` under `workspace/execution/`. |
| Source of record | The tracked file whose values the pipeline reads for an input. A generated copy in the workspace is never the source of record. |
| Stage | One of the five pipeline steps. A1 prepares the base inputs, A2 adds transmission, A3 builds scenarios, B1 compiles model inputs and B2 adjusts the model files, solves and collects results. |
| Storage delay | The model-file adjustment that sets both capacity caps of the storage-linked technologies to 0 in the first five model years and to -1 afterwards. |
| Support scenario | `BAU`, a root whose generated restrictions `A_Calibrated_BAU` inherits. Its results are not part of the comparison. |
| Technology taxonomy | `config/scenarios/technology_types.csv`, which assigns every technology to a category such as GENERATION or INTERCONNECTORS. |
| Technology top-up | The step of the build-cap rule that raises any cap of an unlocked family that is below a set value to that value. |
| Workspace | The folder that holds generated files: `workspace/` under the clone, or the folder passed with `--workspace`. Git ignores it. |

## Docs terms and code names

| Docs term | Code name | Where |
|---|---|---|
| build cap | `build_cap` | `ostram/pipeline/scenarios/rules/build_cap_rule.py` |
| cap multiplier | `cap_multiplier_schedule` | `config/scenarios/<root>/build_cap_rule.yaml` |
| family limit | `family_limits` | `config/scenarios/<root>/build_cap_rule.yaml` |
| technology top-up | `technology_top_up_gw` | `config/scenarios/<root>/build_cap_rule.yaml` |
| adequacy top-up | `adequacy_share_schedule` | `config/scenarios/<root>/build_cap_rule.yaml` |
| demand-scaled fleet | `scaled_pool` | `build_cap_rule.py` |
| margin factor | `margin_factor` | `config/scenarios/<root>/build_cap_rule.yaml` |
| security factor | `margin_factor` | The earlier name of the margin factor. |
| scenario edits | `edits`, `scenario_edits.py` | `config/scenarios/<scenario>/scenario.yaml`, `ostram/pipeline/scenarios/` |
| model-file adjustments | `model_file_adjustments/` | `ostram/pipeline/execution/` |
| edit log | `apply_patches_CHANGES_<stamp>.json` | `workspace/preparation/A1_Outputs/A1_Outputs_<scenario>/` |
| shared VRE ceilings | `vre_ceilings` | `config/setups/south_asia.yaml` |
| overlay | `direction_overlay`, `set_interconnector_direction.yaml` | `config/scenarios/<scenario>/` |
| parent | `base_scenario` | `config/scenarios/<scenario>/scenario.yaml` |
| gas and oil allowance | `reserve_margin_xlsx_*` | `config/execution/Config_MOMF_T1_AB.yaml` |
| storage delay | `storage_delay_*` | `config/execution/Config_MOMF_T1_AB.yaml` |
| backstop | `PWRBCK` | technology names |
| country-region | `cr` | file and column names |
| base-year pin | `apply_base_year_pin.py`, `pwr_min_2023_2026_pin.csv` | `ostram/pipeline/scenarios/rules/`, `config/scenarios/rules/` |
| A-result seed | `--a-result-seed`, `OSTRAM_A_CALIBRATED_BAU_RESULT` | the `run` command, the environment |
