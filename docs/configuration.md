# Configuration Reference

OSTRAM uses four YAML configuration files and two Excel-based configuration files. This page documents every option.

:::{warning}
All data entered in the configuration files (technologies, years, countries, codes) **must match values that exist in the model**. Using technology codes, country codes, year ranges, or any other identifiers that are not present in the model data can cause the pipeline to fail during execution.
:::

## Config_country_codes.yaml

**Location:** `t1_confection/Config_country_codes.yaml`

The single source of truth for all country, region, and technology definitions. Used by most scripts in the project.

### `country_data`

Master registry of countries. Each entry is keyed by a 3-letter ISO code and contains:

```yaml
country_data:
  BGD:
    english_name: "Bangladesh"
    olade_name: "Bangladesh"
  IND:
    english_name: "India"
    olade_name: "India"
```

- `english_name`: Display name used in reports and documentation.
- `olade_name`: Name used for matching against OLADE data files (Spanish/local spelling).

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
  - INDEA
  - INDNE
  - INDNO
  - INDSO
  - INDWE
  - NPL
  - LKA
  - MDV
```

Countries with sub-regions (like India split into 5 regions) use extended codes (e.g., `INDEA` for India-East).

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

### `add_missing_countries_from_olade`

Whether the preprocessing step should fill missing country data from OLADE files:

```yaml
add_missing_countries_from_olade: false
```

### `olade_tech_mapping`

Maps OLADE technology names (in Spanish) to 3-character model codes:

```yaml
olade_tech_mapping:
  "Nuclear": URN
  "Carbon mineral": COA
  "Eolica": WON
  "Solar": SPV
  "Hidroelectrica": HYD
  # ... etc.
```

:::{note}
BIO (Biomass) is a special case: it is the sum of Biogas + Solid biomass + Liquid biofuels from the OLADE data.
:::

### `code_to_energy`

Maps 3-character technology codes to human-readable descriptions:

```yaml
code_to_energy:
  BCK: "Backstop"
  BIO: "Biomass"
  SPV: "Solar Photovoltaic"
  WON: "Onshore Wind"
  # ... (24 entries total)
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

Maps technology names from the Shares Excel file to model codes:

```yaml
shares_tech_mapping:
  "Biomasa": BIO
  "Diesel": PET
  "Hidroelectrica": HYD
  # ... (11 entries)
```

### `implausible_combinations`

Technology-country pairs where a technology is physically infeasible. These are marked as NO (red) in the Tech-Country Matrix:

```yaml
implausible_combinations:
  CSP:
    - BGD
    - BTN
    - NPL
    - LKA
    - MDV
  WAV:
    - BGD
    - BTN
    - IND
    - NPL
    - LKA
    - MDV
```

### `template_generation`

Configuration for the country template generator (`Z_generate_country_template.py`):

```yaml
template_generation:
  new_country: MDV
  reference_country: LKA
  region: XX
  interconnections: []
```

| Key | Description |
|-----|-------------|
| `new_country` | 3-letter code for the country to create |
| `reference_country` | Existing country to clone data from |
| `region` | Region suffix (default: `XX`) |
| `interconnections` | List of neighbor country codes for TRN links. Empty = no interconnections |

### Transmission Technology Parameters

Six sections define default parameters for transmission technologies:

| Section | Description |
|---------|-------------|
| `RNWTRN` | Renewable transmission (existing) |
| `RNWRPO` | Renewable transmission (repowered) |
| `RNWNLI` | Renewable transmission (new lines) |
| `PWRTRN` | Non-renewable transmission (existing) |
| `TRNRPO` | Non-renewable transmission (repowered) |
| `TRNNLI` | Non-renewable transmission (new lines) |

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

---

## Config_MOMF_T1_A.yaml

**Location:** `t1_confection/Config_MOMF_T1_A.yaml`

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
  DailyTimeBracket: ['1', '2', '3']
  Timeslice: Some
  Timeslices: [S1D1, S1D2, S1D3, S2D1, S2D2, S2D3, S3D1, S3D2, S3D3, S4D1, S4D2, S4D3]
  Storage: [LDSBGDXX, SDSBGDXX, ...]
```

The model uses 12 timeslices (4 seasons x 3 daily brackets), a single region (`GLOBAL`), and 2 modes of operation.

### Directory and File Paths

The configuration defines all input/output paths and Excel file names used by the compiler. These are relative to `t1_confection/`:

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

**Location:** `t1_confection/Config_MOMF_T1_AB.yaml`

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

```yaml
del_files: True
only_main_scenario: False
parallel: True
max_x_per_iter: 4
A2_otoole_outputs: True
write_txt_model: True
create_matrix: True
execute_model: True
concat_otoole_csv: True
concat_scenarios_csv: True
annualize_capital: True
```

| Flag | Description |
|------|-------------|
| `del_files` | Delete intermediate files after execution |
| `only_main_scenario` | Run only the main scenario (skip others) |
| `parallel` | Run scenarios in parallel |
| `max_x_per_iter` | Maximum scenarios per parallel batch |
| `A2_otoole_outputs` | Write otoole-format output CSVs |
| `write_txt_model` | Generate the `.txt` model file for the solver |
| `create_matrix` | Create the optimization matrix |
| `execute_model` | Run the solver |
| `concat_otoole_csv` | Concatenate otoole CSVs across scenarios |
| `concat_scenarios_csv` | Concatenate scenario result CSVs |
| `annualize_capital` | Run capital cost annualization post-processing |

### Other Settings

| Key | Value | Description |
|-----|-------|-------------|
| `base_scenario` | `"BAU"` | Name of the base/reference scenario |
| `prefix_final_files` | `"OSTRAM_"` | Prefix for final output file names |
| `osemosys_model` | `"osemosys_fast_preprocessed.txt"` | OSeMOSYS model file (GMPL) |

---

## Config_region_consolidation.yaml

**Location:** `t1_confection/Config_region_consolidation.yaml`

Controls optional consolidation of sub-regional data into unified country-level data. This is relevant when a country is modeled with multiple sub-regions (e.g., India with 5 regions).

### Enable/Disable

```yaml
enabled: false
```

### Country Definitions

```yaml
countries:
  BRA:
    regions: ["CN", "NW", "NE", "CW", "SO", "SE", "WE"]
    unified_region: "XX"
```

Each entry specifies:
- `regions`: List of sub-region codes to merge.
- `unified_region`: Target code for the merged region.

### Aggregation Rules

Defines how parameters are combined when merging sub-regions:

**Averaged parameters** (`aggregation_rules.avg`):

Parameters where values are averaged across regions: `AvailabilityFactor`, `CapacityFactor`, `CapacityToActivityUnit`, `CapitalCost`, `CapitalCostStorage`, `FixedCost`, `InputActivityRatio`, `OutputActivityRatio`, `VariableCost`, and others.

**Summed parameters** (`aggregation_rules.sum`):

Parameters where values are summed: `ResidualCapacity`, `ResidualStorageCapacity`, `SpecifiedAnnualDemand`, `TotalAnnualMaxCapacity`, `TotalAnnualMaxCapacityInvestment`, and others.

**Disabled parameters** (`aggregation_rules.disabled`):

Parameters skipped during consolidation: `CapacityOfOneTechnologyUnit`, `RETagTechnology`, `TotalAnnualMinCapacity`, and model-period activity limits.

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
