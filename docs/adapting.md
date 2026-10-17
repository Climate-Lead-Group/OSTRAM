# Adapting OSTRAM to another system

OSTRAM is a framework built on OSeMOSYS. The South Asia model is its
reference application, and some files and code lines carry South Asia
values. This page lists what a modeller replaces to apply OSTRAM to another
power system, and the South Asia specifics that sit in code. Terms are in
the [glossary](glossary.md).

## Inputs and configuration to replace

| What | Where | Notes |
|---|---|---|
| The base model data | `inputs/osemosys_global/*.csv` | The 64 OSeMOSYS Global set and parameter files for the new countries. |
| Countries and regions | `config/preparation/Config_country_codes.yaml` | `country_data`, `countries`, the technology mappings, the implausible combinations and the transmission defaults. See [country management](country-management.md) and the [configuration](configuration.md) page. |
| The technology-country matrix | `inputs/preparation/secondary_technologies/Tech_Country_Matrix.xlsx` | Which technology exists in which country. |
| The scenario workbook | `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx` | The Control sheet, the parameter sheets, `Interconnector_Params` and the extension decisions. |
| The timeslice workbook | `inputs/scenarios/OSTRAM_Timeslice_Inputs.xlsx` | Year splits, demand profiles and capacity factors for the chosen timeslices. The generator is in `timeslice_module/`. |
| The technology taxonomy | `config/scenarios/technology_types.csv` | Every technology with its category. |
| The rule settings | `config/scenarios/<root>/*.yaml` | Retirement lifetimes, build caps, interconnector relaxation, storage floors and targets. See the [rules](rules.md) page. |
| The internal transmission residuals | `config/scenarios/rules/internal_tx_residuals.csv` | One row per country-region. |
| The base-year pin table | `config/scenarios/rules/pwr_min_2023_2026_pin.csv` | Reviewed values for 2023 to 2026. The pin code expects the South Asia table, as described below. |
| The gas and oil allowance | `inputs/execution/firm_capacity_fallbacks_by_cr.xlsx` | One row per country-region. |
| The setup | `config/setups/south_asia.yaml` | The sources of record, the scenario registry and the shared VRE ceilings. The scenario edits are in `config/scenarios/<name>/scenario.yaml`. A second setup can be registered in `config/setups/registry.json`, as the training setup `examples/unescap_training/setup.yaml` is. |
| The model file | `model/osemosys_fast_preprocessed.txt` | The demand accounting sets and the historical trade constraints name South Asia fuels and corridors, as described below. |

## South Asia specifics in code

These lines are described as they are. Changing them is a code change.

- `ostram/pipeline/scenarios/rules/apply_base_year_pin.py`, line 75, allows the
  countries BGD, BTN, IND, LKA and NPL.
- `ostram/pipeline/scenarios/transformations/update_ao_from_extensions.py`,
  lines 103 to 116, map the region codes BGDXX, BTNXX, INDEA, INDNE, INDNO,
  INDSO, INDWE, LKAXX, MDVXX, MMRXX and NPLXX to their names, and define
  those codes as the valid regions.
- `ostram/pipeline/preparation/secondary_techs/editor_template.py`, lines 128
  to 152, parse the country and region from technology names with the
  examples `PWRHYDINDEA`, `TRNBGDXXINDEA`, `INDEA` and `BGD`, and lines 195
  to 215 document the interconnector discovery with the example
  `TRNBGDXXINDEA` and its fuels `ELCBGDXX02` and `ELCINDEA01`.
- `model/osemosys_fast_preprocessed.txt`, lines 402 to 404, list the nine
  demand fuels of the upstream demand accounting, from `ELCBGDXX03` to
  `ELCINDNE03`, with the Maldives left out. Lines 593 onwards hold the 50
  historical trade constraints on named South Asia corridors for 2023 to
  2025, which B2 activates for the scenarios of the roots listed in
  `historical_trade_roots`.
