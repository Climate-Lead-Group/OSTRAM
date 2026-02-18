# OSTRAM - OSeMOSYS Transmission Model

Energy system optimization modeling based on OSeMOSYS.

## Description

This project implements an automated pipeline for running energy optimization models using OSeMOSYS. The system supports multiple solvers (GLPK, CBC, CPLEX, Gurobi) and is designed to ensure full reproducibility of results.

## Key Features

- **Automated Pipeline**: Complete workflow management with DVC
- **Multiple Solvers**: Support for GLPK, CBC, CPLEX, and Gurobi
- **Guaranteed Reproducibility**: Configurable seeds for deterministic results
- **Performance Monitoring**: Built-in timer for tracking execution times
- **Automatic Environment Management**: Automatic creation and update of the Conda environment

## System Requirements

- Windows 10 or higher
- Git for Windows
- Miniconda or Anaconda
- At least one solver: GLPK, CBC, CPLEX, or Gurobi

## Quick Start

```bash
# Clone the repository
git clone https://github.com/Climate-Lead-Group/OSTRAM.git
cd OSTRAM

# Run the model (from Anaconda Prompt)
python run.py
```

The `run.py` script automatically handles:
- Conda environment creation
- Dependency installation
- Full pipeline execution
- Output file generation

## Documentation

For detailed installation and configuration instructions, see the full guide:
- **Installation and Execution Guide**: `OSTRAM_Guia_instalacion_ejecucion.md`

## Output File Structure

Results are generated in `t1_confection/` with the following files:
- `OSTRAM_Inputs.csv` / `OSTRAM_Inputs_YYYY-MM-DD.csv`
- `OSTRAM_Outputs.csv` / `OSTRAM_Outputs_YYYY-MM-DD.csv`
- `OSTRAM_Combined_Inputs_Outputs.csv` / `OSTRAM_Combined_Inputs_Outputs_YYYY-MM-DD.csv`

Date-stamped files maintain a complete execution history.

## Configuration

The main configuration file is `t1_confection/MOMF_T1_AB.yaml`, where you can adjust:
- Solver to use (`solver: 'cplex'`)
- Number of threads for commercial solvers
- Seeds for reproducibility
- Capital annualization (`annualize_capital`)

## Technology-Country Matrix

The system includes a configurable matrix that allows you to specify which technology-country combinations should be processed, as well as unify CCG and OCG technologies into NGS.

### Usage

1. **Generate the matrix**:
   ```bash
   python t1_confection/A0_generate_tech_country_matrix.py
   ```
   This creates the file `Tech_Country_Matrix.xlsx` with the following sheets:
   - **Matrix**: YES/NO matrix for each technology-country combination
   - **NGS_Unification**: Configuration for unifying CCG + OCG → NGS
   - **Aggregation_Rules**: Aggregation rules (avg/sum/disabled)
   - **Tech_Reference**: Technology descriptions
   - **Country_Reference**: Country descriptions

2. **Configure the matrix**:
   - In the **Matrix** sheet: Change YES/NO to enable/disable combinations
   - In the **NGS_Unification** sheet: Change to YES/NO to enable CCG+OCG→NGS unification

3. **Run preprocessing**:
   ```bash
   python t1_confection/A1_Pre_processing_OG_csvs.py
   ```
   The script will automatically apply:
   - Technology-country matrix filtering
   - NGS unification (if enabled)
   - Region consolidation
   - PWR technology cleanup

### Technologies in the Matrix

| Code | Description |
|------|-------------|
| BCK | Backstop |
| BIO | Biomass |
| CCS | Carbon Capture Storage with Coal |
| COA | Coal |
| COG | Cogeneration |
| CSP | Concentrated Solar Power |
| GAS | Natural Gas |
| GEO | Geothermal |
| HYD | Hydroelectric |
| LDS | Long duration storage |
| NGS | Natural Gas (CCG + OCG unified) |
| OIL | Oil |
| OTH | Other |
| PET | Petroleum |
| SDS | Short duration storage |
| SPV | Solar Photovoltaic |
| URN | Nuclear |
| WAS | Waste |
| WAV | Wave |
| WOF | Offshore Wind |
| WON | Onshore Wind |

**Note:** Structural prefixes (ELC, MIN, PWR, RNW, TRN) are not included in the matrix because they are combined with the codes above to form full technology names (e.g., PWRBIOARGXX, MINCOAARGXX).

## Secondary Technologies Editor

The project includes a system to facilitate editing secondary technologies (Secondary Techs) in parameterization files, with support for automatic OLADE data integration.

### Editor Usage

1. **Generate the editor template**:
   ```bash
   python t1_confection/D1_generate_editor_template.py
   ```
   This creates the file `Secondary_Techs_Editor.xlsx` with two sheets:
   - **Instructions**: For manual editing with dropdown lists
   - **OLADE_Config**: Configuration for automatic OLADE data integration

2. **Manual Editing** (Sheet "Instructions"):
   - Select: Scenario (BAU, NDC, NDC+ELC, NDC_NoRPO, or ALL)
   - Select: Country, Technology (Tech.Name), and Parameter
   - Enter values for the desired years (2021-2050)
   - The "Tech" column is automatically populated via VLOOKUP

3. **OLADE Integration** (Sheet "OLADE_Config"):

   Allows automatic population of parameters using OLADE data.

   | Parameter | Description |
   |-----------|-------------|
   | `ResidualCapacitiesFromOLADE` | YES/NO - Enable OLADE integration for installed capacity (ResidualCapacity) |
   | `PetroleumSplitMode` | OIL_only or Split_PET_OIL - Petroleum split mode |
   | `DemandFromOLADE` | YES/NO - Enable OLADE integration for electricity demand |
   | `ActivityLowerLimitFromOLADE` | YES/NO - Enable OLADE integration for TotalTechnologyAnnualActivityLowerLimit |
   | `ActivityUpperLimitFromOLADE` | YES/NO - Enable OLADE integration for TotalTechnologyAnnualActivityUpperLimit |

   **PetroleumSplitMode**:
   - `OIL_only`: Assigns all petroleum capacity to OIL (Fuel oil)
   - `Split_PET_OIL`: Splits between PET (Diesel) and OIL (Fuel oil + Bunker) using proportions from `Shares_PET_OIL_Split.xlsx`

   **DemandFromOLADE**:
   - When enabled, updates electricity demand in `A-O_Demand.xlsx` using OLADE generation data
   - Configure growth rates per country in the `Demand_Growth` sheet
   - Formula: `Demand(year) = Demand(2023) × (1 + rate × (year - 2023))`

   **ActivityLowerLimit and ActivityUpperLimit**:
   - When enabled, automatically populate activity limits in `A-O_Parametrization.xlsx`
   - Uses OLADE electricity generation data combined with technology shares from `Shares_Power_Generation_Technologies.xlsx`
   - Configure optional renewability targets in the `Renewability_Targets` sheet
   - Configure custom technology weights in the `Technology_Weights` sheet
   - Formula: `ActivityLimit(tech,year) = Total_Generation(PJ) × (1 + rate × (year - 2023)) × Share(tech,year)`
   - Includes automatic validation against available capacities
   - See the `Documentation` sheet in the editor for full calculation and validation details

4. **Additional Editor Sheets**:

   The `Secondary_Techs_Editor.xlsx` file also includes:
   - **Renewability_Targets**: Defines renewable % targets per year for each country/scenario (used by Activity Limits)
   - **Technology_Weights**: Allows defining custom distribution of renewable and non-renewable technologies
   - **Scenarios_Demand_Growth**: Configures scenario- and country-specific demand growth rates
   - **Documentation**: Full technical documentation on Activity Limits calculation and validation

5. **Apply changes**:
   ```bash
   python t1_confection/D2_update_secondary_techs.py
   ```

### System Features

- **Dropdown lists**: Facilitate selection of scenarios, countries, technologies, and parameters
- **Tech.Name → Tech mapping**: Automatic conversion from descriptive names to technical codes
- **OLADE Capacity Integration**: Automatic population of ResidualCapacity from installed capacity data
- **OLADE Demand Integration**: Automatic population of electricity demand from generation data
- **OLADE Activity Limits Integration**: Automatic population of TotalTechnologyAnnualActivityLowerLimit and UpperLimit
- **Unit conversion**: MW → GW (capacity), GWh → PJ (demand and activity)
- **Flat values (capacity)**: The same capacity value is used for all years
- **Linear growth (demand and activity)**: Configurable growth rate per country
- **Activity Limits validation**: Automatically verifies that limits do not exceed available capacity
- **Renewability targets**: Interpolation system to reach renewable % goals
- **Automatic backups**: One backup per scenario before applying changes
- **Projection.Mode**: Automatically updated to "User defined" when values are modified
- **Detailed logs**: Full logging with country identification for each operation

### Related Files

| File | Description |
|------|-------------|
| `A0_generate_tech_country_matrix.py` | Generates the technology-country matrix |
| `D1_generate_editor_template.py` | Generates the Excel template |
| `D2_update_secondary_techs.py` | Applies changes to scenarios |
| `Tech_Country_Matrix.xlsx` | Technology-country matrix (generated) |
| `Secondary_Techs_Editor.xlsx` | Editor template (generated) |
| `OLADE - Capacidad instalada por fuente - Anual.xlsx` | OLADE source data (installed capacity) |
| `OLADE - Generación eléctrica por fuente - Anual.xlsx` | OLADE source data (electricity generation) |
| `Shares_PET_OIL_Split.xlsx` | Petroleum split proportions (Diesel, Fuel oil, Bunker) per scenario |
| `Shares_Power_Generation_Technologies.xlsx` | Power generation technology proportions per country/scenario/year |

### OLADE → Model Country Mapping

Some country codes differ between OLADE and the model:

| Country | OLADE | Model |
|---------|-------|-------|
| Barbados | BAR | JAM |
| Chile | CHI | CHL |
| Costa Rica | CRC | CRI |

## Country Management Tools

### Country Data Validator

Verifies that a country has all required data in the OSeMOSYS input CSV files.

```bash
python t1_confection/Z_validate_country_data.py                  # Validate all countries
python t1_confection/Z_validate_country_data.py --country ARG    # Validate a specific country
python t1_confection/Z_validate_country_data.py --country NCC --report  # Generate detailed report
```

**Validations performed:**
- Presence in sets (TECHNOLOGY, FUEL, EMISSION, STORAGE)
- Minimum number of technologies per prefix (PWR, MIN, RNW)
- Data in all required parameters (costs, capacity, factors, ratios, etc.)
- Expected fuel patterns per country

### New Country Template Generator

Creates a set of CSV files with the minimum structure needed to add a new country, using an existing country as a reference.

```bash
python t1_confection/Z_generate_country_template.py                              # Read config from YAML
python t1_confection/Z_generate_country_template.py --new NCC --ref ARG -i BOL PRY  # CLI override
```

**Configuration** (`template_generation` section in `Config_country_codes.yaml`):

| Parameter | Description |
|-----------|-------------|
| `new_country` | 3-letter code for the new country |
| `reference_country` | Existing country to clone data from |
| `region` | Region code (default: XX) |
| `interconnections` | List of neighbors for interconnections (empty = no interconnections) |

**Features:**
- Generates CSVs in `templates/{CODE}/` without modifying original files
- Dynamic interconnection handling: supports more, fewer, equal, or zero interconnections relative to the reference country
- Correct transformation of fuel and mode-of-operation codes for TRN technologies
- Includes a `merge_into_inputs.py` script in the generated folder for easy integration

### Related Files

| File | Description |
|------|-------------|
| `Z_validate_country_data.py` | Validates country data in OG_csvs_inputs |
| `Z_generate_country_template.py` | Generates CSV template for adding a country |
| `Config_country_codes.yaml` | Centralized configuration (includes `template_generation` section) |

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

Copyright 2025 Climate Lead Group

This project is developed by Climate Lead Group for energy system analysis.
