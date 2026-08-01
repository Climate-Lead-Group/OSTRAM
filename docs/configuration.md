# Configuration Reference

OSTRAM keeps maintained YAML under `config/preparation/`, `config/scenarios/`,
`config/compilation/`, and `config/execution/`. Scenario and timeslice workbook
authorities live under `inputs/scenarios/`. This page documents their options.

:::{warning}
All data entered in the configuration files (technologies, years, countries, codes) **must match values that exist in the model**. Using technology codes, country codes, year ranges, or any other identifiers that are not present in the model data can cause the pipeline to fail during execution.
:::

## Config_country_codes.yaml

**Location:** `config/preparation/Config_country_codes.yaml`

The single source of truth for all country, region, and technology definitions. Used by most scripts in the project.

### `country_data`

Master registry of countries. Each entry is keyed by a 3-letter ISO code and contains:

```yaml
country_data:
  BGD:
    english_name: "Bangladesh"
    ostram_name: "Bangladesh"
  BTN:
    english_name: "Bhutan"
    ostram_name: "Bhutan"
  IND:
    english_name: "India"
    ostram_name: "India"
  NPL:
    english_name: "Nepal"
    ostram_name: "Nepal"
  LKA:
    english_name: "Sri Lanka"
    ostram_name: "Sri Lanka"
  MDV:
    english_name: "Maldives"
    ostram_name: "Maldives"
```

- `english_name`: Display name used in reports and documentation.
- `ostram_name`: Name used for matching against OSTRAM source data Excel files.

Note that `IND` is registered once as a country, but the model runs it as 5 separate regions (see `countries` below) -- `country_data` is the ISO/naming registry, not the list of model-active region keys.

### `special_entries`

Non-country codes used in the model:

```yaml
special_entries:
  INT: "International Markets"
```

### `countries`

Ordered list of active country/region codes for the current model run:

```yaml
countries:
  - BGD
  - BTN
  - INDEA    # India East
  - INDNE    # India North-East
  - INDNO    # India North
  - INDSO    # India South
  - INDWE    # India West
  - NPL
  - LKA
  - MDV
```

Countries with sub-regions (like India split into 5 regions) use extended codes (e.g., `INDEA` for India-East). `Z_AUX_config_loader.get_multi_region_map()` derives `{iso3: [region_codes]}` (e.g. `{'IND': ['EA','NE','NO','SO','WE']}`) from this list for scripts that need to group regions back under their parent country.

### `first_year`

Reference/start year of the model time horizon:

```yaml
first_year: 2023
```

### `pwr_cleanup_mode`

Controls how duplicate PWR (power) technology entries are handled during preprocessing:

```yaml
pwr_cleanup_mode: "merge"
```

| Value | Behavior |
|-------|----------|
| `"drop"` | Drop PWR00 when PWR01 exists, rename PWR01 to PWR |
| `"merge"` | Sum PWR00 values into PWR01, drop PWR00, rename PWR01 to PWR |
| `false` | Skip PWR cleanup entirely |

### `force_empty_max_capacity_investment_pwr`

Forces `Projection.Mode` to `"EMPTY"` for `TotalAnnualMaxCapacityInvestment` on all PWR technologies, overriding the auto-detected mode:

```yaml
force_empty_max_capacity_investment_pwr: true
```

### `add_missing_countries_from_ostram`

Whether the preprocessing step should fill missing country data from OSTRAM source files:

```yaml
add_missing_countries_from_ostram: false
```

### `ostram_tech_mapping`

Maps technology names from the OSTRAM source Excel files to 3-character model codes:

```yaml
ostram_tech_mapping:
  Nuclear: URN
  "Natural gas": NGS
  "Mineral coal": COA
  Hydro: HYD
  Geothermal: GEO
  Wind: WON
  Solar: SPV
```

:::{note}
BIO (Biomass) is a special case: it is the sum of Biogas + Solid biomass + Liquid biofuels from the source data.
:::

### `code_to_energy`

Maps technology codes to human-readable descriptions (includes both source/fuel codes and structural prefixes):

```yaml
code_to_energy:
  MIN: "Mining tradable commodity"
  RNW: "Mining non-tradable (renewable) commodity"
  BIO: "Biomass"
  GAS: "Natural Gas"
  COA: "Coal"
  GEO: "Geothermal"
  HYD: "Hydroelectric"
  OIL: "Oil"
  OTH: "Other"
  PET: "Petroleum"
  SPV: "Solar Photovoltaic"
  URN: "Nuclear"
  WAV: "Wave"
  WAS: "Waste"
  WOF: "Offshore Wind"
  WON: "Onshore Wind"
  CCG: "Combined Cycle Natural Gas"
  COG: "Cogeneration"
  CSP: "Concentrated Solar Power"
  NGS: "Natural Gas"
  OCG: "Open Cycle Natural Gas"
  TRN: "Transmission technology"
  LDS: "Long duration storage"
  SDS: "Short duration storage"
  PWR: "Power generator"
  ELC: "Electricity"
  BCK: "Backstop"
  CCS: "Carbon Capture Storage with Coal"
```

### `renewable_fuels`

List of technology codes classified as renewable energy:

```yaml
renewable_fuels:
  - BIO
  - HYD
  - CSP
  - GEO
  - SPV
  - WAS
  - WON
  - WOF
```

### `shares_tech_mapping`

Maps technology names from `Shares_Power_Generation_Technologies.xlsx` to model codes:

```yaml
shares_tech_mapping:
  Biomass: BIO
  Bunker: OIL
  Coal: COA
  Diesel: PET
  Wind: WON
  "Fuel oil": OIL
  "Natural gas": NGS
  Geothermal: GEO
  Hydroelectric: HYD
  Nuclear: URN
  "Solar (DG)": SPV
  "Solar (utility scale)": SPV
```

### `implausible_combinations`

Technology-country pairs where a technology is physically infeasible (a combination is considered implausible when `ResidualCapacity`, `CapitalCost`, `TotalAnnualMaxCapacity`, and `TotalAnnualMaxCapacityInvestment` are all zero for every year in the source data). These are marked as NO (red) in the Tech-Country Matrix:

```yaml
implausible_combinations:
  BIO: [BTN, MDV]                        # No biomass resources or grid-connected capacity
  CCS: [BTN, NPL, LKA, MDV]               # No coal infrastructure to retrofit
  COA: [BTN, NPL, MDV]                    # No coal plants, resources, or plans
  COG: [BTN, NPL, MDV]                    # No industrial base for cogeneration
  CSP: [BGD, BTN, NPL, LKA, MDV]           # Unsuitable climate/terrain or insufficient land
  GAS: [BGD, BTN, IND, NPL, LKA, MDV]      # No PWRGAS tech in model; modeled via CCG/OCG -> NGS
  GEO: [BGD, BTN, IND, NPL, LKA, MDV]      # No geothermal resources in any modeled country
  HYD: [MDV]                              # Flat coral atolls, no rivers or elevation
  NGS: [BGD, BTN, NPL, MDV]                # No CCG/OCG/NGS power data
  OIL: [BTN, NPL, MDV]                    # No oil-fired grid power plants
  URN: [BTN, NPL, LKA, MDV]                # No nuclear program or plans
  WAS: [BTN, MDV]                          # No waste-to-energy plants or plans
  WAV: [BGD, BTN, IND, NPL, LKA, MDV]       # No capacity/plans; BTN and NPL are landlocked
  WOF: [BTN, NPL, MDV]                     # BTN/NPL landlocked; MDV scale economically unviable
  WON: [MDV]                              # No country lacks onshore wind data except MDV
```

### `template_generation`

Configuration for the country template generator (`Z_generate_country_template.py`). This section is a **list**, so multiple countries can be generated in a single run. Each entry defines one country to create:

```yaml
template_generation:
  - new_country: MDV
    reference_country: LKA
    region: XX
    centerpoint_lat: 1.924992
    centerpoint_lon: 73.399658
    interconnections:
      - LKA
```

| Key | Required | Description |
|-----|----------|-------------|
| `new_country` | Yes | 3-letter ISO code for the country to create |
| `reference_country` | Yes | Existing country to clone data from |
| `region` | No | Region suffix (default: `XX`) |
| `centerpoint_lat` | No | Latitude for the country's geographic centerpoint |
| `centerpoint_lon` | No | Longitude for the country's geographic centerpoint |
| `interconnections` | No | List of neighbor country codes for TRN links. **`[]` (empty list)** = generate the country with zero interconnections. **Omitting the key entirely** = legacy mode, which blindly copies the reference country's TRN rows verbatim without remapping (only correct if the new country's topology is identical to the reference's) |

Multiple countries -- or multiple regions of the same new country -- can be defined as separate list entries. The `countries:` comment block in the live YAML documents two illustrative patterns beyond the single-country case above:

```yaml
template_generation:
  # Multiple independent countries in one run:
  - new_country: MMR
    reference_country: BGD
    region: XX
    centerpoint_lat: 16.871311
    centerpoint_lon: 96.199379
    interconnections:
      - BGD
      - IND
  - new_country: PAK
    reference_country: IND
    region: XX
    centerpoint_lat: 24.87
    centerpoint_lon: 66.99
    interconnections:
      - INDNO
      - INDWE

  # A country with multiple regions (like India): one entry per region,
  # each with its own centerpoint. Use 5-letter codes to interconnect
  # to a specific region of another multi-region country.
  - new_country: CHN
    reference_country: IND
    region: SO                # China South -> CHNSO
    centerpoint_lat: 23.145
    centerpoint_lon: 113.325
    interconnections:
      - MMRXX
      - CHNNO                 # connected to China North
  - new_country: CHN
    reference_country: IND
    region: NO                # China North -> CHNNO
    centerpoint_lat: 39.929
    centerpoint_lon: 116.388
    interconnections:
      - CHNSO                 # connected to China South
      - MNGXX
```

The only entry actually active in the current YAML is the single MDV-from-LKA example above; MMR/PAK/CHN are documentation-only illustrations.

The package implementation consumes these entries through the preparation
workflow. Run it with `python -m ostram run`; source files under `ostram/` are
not public script entrypoints.

| CLI Flag | Description |
|----------|-------------|
| `--new`, `-n` | New country code (3 letters) |
| `--ref`, `-r` | Reference country code |
| `--region` | Region code (2 letters) |
| `-i`, `--interconnections` | Neighbor country codes (space-separated) |
| `--lat` | Centerpoint latitude |
| `--lon` | Centerpoint longitude |
| `-o`, `--output` | Output directory (default: `templates/<new_code>`) |

### Transmission Technology Parameters

Seven sections define default parameters for transmission and dispatch technologies:

| Section | Description |
|---------|-------------|
| `RNWTRN` | Renewable transmission (existing) |
| `RNWRPO` | Renewable transmission (repowered) |
| `RNWNLI` | Renewable transmission (new lines) |
| `PWRTRN` | Non-renewable transmission (existing) |
| `TRNRPO` | Non-renewable transmission (repowered) |
| `TRNNLI` | Non-renewable transmission (new lines) |
| `DSPTRN` | Dispatch (interconnection routing) |

Each section contains:

```yaml
RNWTRN:
  CapacityToActivityUnit: 31.536
  OperationalLife: 20
  CapitalCost: 100
  FixedCost: 4
  ResidualCapacity: 5
  TotalAnnualMaxCapacityInvestment: 5
```

DSPTRN is a virtual dispatch node with zero costs and high capacity, used to route electricity to/from cross-border interconnections (`ELC*02` → `ELC*03` on export, `ELC*04` → `ELC*03` on import -- see {doc}`pipeline` Stage A2 for the full fuel-tier routing):

```yaml
DSPTRN:
  CapacityToActivityUnit: 31.536
  OperationalLife: 20
  CapitalCost: 0
  FixedCost: 0
  ResidualCapacity: 9999
  TotalAnnualMaxCapacityInvestment: 9999
```

### `enable_dsptrn`

Master switch for whether A2 injects the `DSPTRN` dispatch technology and rewrites TRN fuel codes to the 4-tier `ELC*00`-`ELC*04` scheme:

```yaml
enable_dsptrn: true
```

---

## Config_MOMF_T1_A.yaml

**Location:** `config/compilation/Config_MOMF_T1_A.yaml`

The primary compiler configuration. Defines the data model for the Excel-to-OSeMOSYS compilation step.

### Key Settings

| Key | Value | Description |
|-----|-------|-------------|
| `base_year` | `"2023"` | Base year of the energy model |
| `initial_year` | `"2023"` | First year of the time horizon |
| `final_year` | `"2050"` | Last year of the time horizon |
| `Use_Transport` | `false` | Enable/disable the transport sub-module |
| `Use_OG_module` | `true` | Enable/disable the OSeMOSYS-Global module pathway |

### Temporal Structure (`xtra_scen`)

```yaml
xtra_scen:
  Main_Scenario: BAU
  Other_Scenarios: []
  Region: GLOBAL
  Mode_of_Operation: [1, 2]
  Season: ['1', '2', '3', '4']
  DayType: ['1']
  DailyTimeBracket: ['1', '2', '3', '4', '5']
  Timeslice: Some
  Timeslices: [S1D1, S1D2, S1D3, S1D4, S1D5, S2D1, S2D2, S2D3, S2D4, S2D5,
               S3D1, S3D2, S3D3, S3D4, S3D5, S4D1, S4D2, S4D3, S4D4, S4D5]
  Storage: [LDSBGDXX01, SDSBGDXX01, ...]
```

The model uses **20 timeslices** (4 seasons x 5 daily time brackets), a single region (`GLOBAL`), and 2 modes of operation. This was expanded from an earlier 12-timeslice (4x3) structure -- see {doc}`data-reference` for the full `Conversionls`/`Conversionld`/`Conversionlh` mapping.

### Directory and File Paths

The configuration retains logical input/output names used by the compiler.
`ostram.paths` maps maintained inputs to the project bundle and mutable outputs
to the selected compilation workspace:

- `A1_inputs` / `A1_outputs`: Stage A1 directories
- `A2_extra_inputs` / `A2_output`: Stage A2 directories
- `Print_*`: Output Excel file name templates (e.g., `Print_Paramet: "/A-O_Parametrization.xlsx"`)

### OSeMOSYS Parameters

The file lists all OSeMOSYS parameters organized by technology category:

- `tech_param_list_primary`: Parameters for primary supply technologies
- `tech_param_list_secondary`: Parameters for secondary (power) technologies
- `tech_param_list_demands`: Parameters for demand technologies
- `tech_param_list_disttrn` / `_trn` / `_trngroups`: Transport parameters

---

## Config_MOMF_T1_AB.yaml

**Location:** `config/execution/Config_MOMF_T1_AB.yaml`

The execution/runtime configuration for the model solver.

### Solver Configuration

```yaml
solver: 'cplex'
cplex_threads: 4
cplex_random_seed: 12345
cbc_random_seed: 12345
iteration_time: 20000
gurobi_threads: 3
gurobi_seed: 12345
```

| Key | Description |
|-----|-------------|
| `solver` | Active solver: `glpk`, `cbc`, `cplex`, or `gurobi` |
| `cplex_threads` | Number of threads for CPLEX |
| `cplex_random_seed` | Random seed for CPLEX reproducibility |
| `cbc_random_seed` | Random seed for CBC |
| `iteration_time` | Time limit for CBC in seconds |
| `gurobi_threads` | Number of threads for Gurobi |
| `gurobi_seed` | Random seed for Gurobi |

### Pipeline Control Flags

Current values (these change often between experiments -- always check the live file):

```yaml
del_files: False
only_main_scenario: False
parallel: False
max_x_per_iter: 4
A2_otoole_outputs: True
write_txt_model: True
create_matrix: True
execute_model: True
reuse_existing_sol: False
concat_otoole_csv: True
concat_scenarios_csv: True
annualize_capital: False
```

| Flag | Description |
|------|-------------|
| `del_files` | Delete intermediate files after execution |
| `only_main_scenario` | Run only the main scenario (skip others) |
| `parallel` | Run scenarios in parallel |
| `max_x_per_iter` | Maximum scenarios per parallel batch |
| `A2_otoole_outputs` | Write otoole-format output CSVs |
| `write_txt_model` | Generate the `.txt` model file for the solver |
| `create_matrix` | Create the optimization matrix (always via `glpsol`, even when the active solver isn't GLPK) |
| `execute_model` | Run the solver |
| `reuse_existing_sol` | Skip solving if a `.sol` already exists at the expected path (regenerates outputs from a previous solve); falls back to a normal solve if missing |
| `concat_otoole_csv` | Concatenate otoole CSVs across scenarios |
| `concat_scenarios_csv` | Concatenate scenario result CSVs |
| `annualize_capital` | Run capital cost annualization post-processing (`Z_AUX_capital_annualization_script.py`) |

### Other Settings

| Key | Value | Description |
|-----|-------|-------------|
| `base_scenario` | `"BAU"` | Name of the base/reference scenario |
| `prefix_final_files` | `"OSTRAM_"` | Prefix for final output file names (overridden when `storage_delay_active: true` -- see below) |
| `osemosys_model` | `"osemosys_fast_preprocessed.txt"` | OSeMOSYS model file (GMPL) |

### Datafile Patch Chain

B2 applies a fixed chain of patches to the preprocessed OSeMOSYS datafile before solving, each gated by its own flag. See {doc}`pipeline` (Stage B2) for the exact application order and how they interact.

**Storage delay** -- blocks storage builds/operation for the first N years, then releases them. Mutually exclusive with storage stripping (storage delay silently wins if both are `true`):

```yaml
storage_delay_active: True
storage_delay_first_n_years: 5
storage_delay_storage_prefixes: [SDS, LDS]
storage_delay_allowed_value: "-1"          # unconstrained PWR cap in open years
storage_delay_suffix: "StorageDelayN5"
storage_delay_model_output: "osemosys_fast_preprocessed_storage_delay.txt"
storage_delay_prefix_final_files: "OSTRAM_StorageDelay_"
storage_delay_root_datafile: "OSTRAM_data_storage_delay.txt"
```

**Storage stripping** -- a diagnostic that removes storage entirely (`mode: "all"`), a single technology (`mode: "tech"`), or a whole storage class (`mode: "class"`):

```yaml
strip_storage_active: True
strip_storage_mode: "all"
strip_storage_targets: []
strip_storage_suffix: "NoStorage"
```

**PWRBCK cap opening** -- opens `TotalAnnualMaxCapacity`/`TotalAnnualMaxCapacityInvestment` for backstop technologies, so an infeasible run can be diagnosed without the backstop artificially binding:

```yaml
open_pwrbck_active: True
open_pwrbck_value: 9999
open_pwrbck_pattern: "PWRBCK"
open_pwrbck_suffix: "OpenBCK"
```

**Reserve-margin repair** -- two alternative implementations; only one should be active at a time. The older, blunt version is kept for compatibility but disabled by default in favor of the workbook-driven one:

```yaml
reserve_margin_repair_active: False        # older, blunt text/datafile-rule version

reserve_margin_xlsx_active: True           # current: workbook-driven, via patch_reserve_margin_repair_careful_xlsx.py
reserve_margin_xlsx_suffix: "RMCarefulXLSX"
reserve_margin_xlsx_workbook: "firm_capacity_fallbacks_by_cr.xlsx"
reserve_margin_xlsx_sheet: "fallbacks"
reserve_margin_xlsx_backstop_credit: 1.0   # 1.0 = PWRBCK gets full reserve-capacity credit
reserve_margin_xlsx_ccs_credit: 0.9        # 0.9 = 90% of PWRCCS capacity counts for reserve
reserve_margin_xlsx_target_prefixes: [PWRPET, PWROIL, PWRNGS]
reserve_margin_xlsx_sentinel_values: [0, 9999]
```

---

## Config_region_consolidation.yaml

**Location:** `config/preparation/Config_region_consolidation.yaml`

Controls optional consolidation of sub-regional data into unified country-level data. This is relevant when a country is modeled with multiple sub-regions (e.g., India with 5 regions in the current model).

### Enable/Disable

```yaml
enabled: false
```

Currently `false` -- the model keeps India's 5 regions (`INDEA`, `INDNE`, `INDNO`, `INDSO`, `INDWE`) separate rather than consolidating them into one. The `countries:` block is currently empty (commented-out examples only); this mechanism is available but unused in the active configuration.

### Country Definitions

```yaml
countries:
  # MEX:
  #   regions: ["NO", "CE", "SU"]
  #   unified_region: "XX"
```

Each entry specifies:
- `regions`: List of sub-region codes to merge.
- `unified_region`: Target code for the merged region.

To consolidate India's 5 regions into one, for example, you would add:

```yaml
enabled: true
countries:
  IND:
    regions: ["EA", "NE", "NO", "SO", "WE"]
    unified_region: "XX"
```

### Aggregation Rules

Defines how parameters are combined when merging sub-regions:

**Averaged parameters** (`aggregation_rules.avg`): `AvailabilityFactor`, `CapacityFactor`, `CapacityToActivityUnit`, `CapitalCost`, `CapitalCostStorage`, `DiscountRateStorage`, `EmissionActivityRatio`, `FixedCost`, `InputActivityRatio`, `MinStorageCharge`, `OperationalLife`, `OperationalLifeStorage`, `OutputActivityRatio`, `ReserveMarginTagFuel`, `ReserveMarginTagTechnology`, `SpecifiedDemandProfile`, `StorageMaxChargeRate`, `StorageMaxDischargeRate`, `TechnologyFromStorage`, `TechnologyToStorage`, `VariableCost`.

**Summed parameters** (`aggregation_rules.sum`): `ResidualCapacity`, `ResidualStorageCapacity`, `StorageLevelStart`, `SpecifiedAnnualDemand`, `TotalAnnualMaxCapacity`, `TotalAnnualMaxCapacityInvestment`, `TotalAnnualMinCapacityInvestment`, `TotalTechnologyAnnualActivityLowerLimit`, `TotalTechnologyAnnualActivityUpperLimit`.

**Disabled parameters** (`aggregation_rules.disabled`): `CapacityOfOneTechnologyUnit`, `RETagTechnology`, `TotalAnnualMinCapacity`, `TotalTechnologyModelPeriodActivityLowerLimit`, `TotalTechnologyModelPeriodActivityUpperLimit`.

---

## A3 Scenario Rule YAMLs

**Location:** `config/scenarios/<Scenario>/`

Each active scenario (`A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE`) has its own subfolder with the YAML files its rule-script chain consumes. See {doc}`pipeline` (Stage A3) for which rule scripts each scenario runs and a plain-English summary of what each scenario represents. The main YAML anatomy:

| YAML | Consumed by | Key fields |
|------|--------------|------------|
| `retirement_schedule.yaml` | `set_retirement_schedule.py` | `base_year`, `age_based.<fuel>.lifetime_years`, `age_based.<fuel>.retirement_profile`, `scheduled`, `exempt` |
| `lid_rule.yaml` | `add_max_cap_investment_lid_rule.py` | `rule_mode` (`uniform`/`proportional`), `percentage_default`, `percentage_by_year`, `security_factor`, `exempt_prefixes`, `relaxation_schedule` |
| `relax_interconnectors.yaml` | `relax_interconnectors.py` | `mode` (`multiplicative`/`additive`/`unconstrained`), `headroom_factor`, `overrides` (absolute per-corridor schedules) |
| `bau_calibration.yaml` (or `storage_floors.yaml`) | `set_min_capacity_floors.py` | `floors`, `ceilings`, each with `cr` (country-region), `tech`, `param`, `schedule` |
| `set_vre_targets.yaml` | `set_vre_targets.py` | `bau_results_path` (relative path to a prior solved scenario's output, typically `A_Calibrated_BAU`), `constraint_type`, `max_floor_share`, `targets`, `pin_generation_to_target`, `cap_envelope` |

Each `<Scenario>/` folder may also contain a `deprecate/` subfolder and `*_v2`-suffixed duplicates -- these are inert backups, not read by the pipeline (only the exact filename each rule script expects, e.g. `lid_rule.yaml`, is loaded).

:::{warning}
Rule scripts resolve `bau_results_path` **relative to their own staged working directory** inside the A3 run, not relative to the repo root. Don't "fix" it to an absolute path from a different machine -- verify it still resolves correctly if you move or rename `Executables/` output folders.
:::

---

## Excel-Based Configuration

### Tech_Country_Matrix.xlsx

Generated by `A0_generate_tech_country_matrix.py`. Contains 5 sheets:

| Sheet | Purpose |
|-------|---------|
| **Matrix** | YES/NO grid for each technology-country combination |
| **NGS_Unification** | ON/OFF toggle per country for merging CCG+OCG into NGS |
| **Aggregation_Rules** | Rules for averaging, summing, or disabling parameters |
| **Tech_Reference** | Technology code to description mapping |
| **Country_Reference** | Country code to name mapping |

### Secondary_Techs_Editor.xlsx

Generated by `D1_generate_editor_template.py`. See {doc}`secondary-techs-editor` for full documentation.
