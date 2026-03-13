# Pipeline Workflow

OSTRAM processes energy system data through a multi-stage pipeline. This page documents each stage in detail, including its inputs, outputs, and configuration.

## Pipeline Overview

```
┌──────────────────────────────────────────────────────────────────────┐
│                        DATA PREPARATION                             │
│                                                                      │
│  A0: Generate Tech-Country Matrix                                    │
│   ↓                                                                  │
│  A1: Preprocess Raw CSVs → Excel Model Files                        │
│   ↓                                                                  │
│  A2: Add Transmission Technologies                                   │
│   ↓                                                                  │
│  (Optional) A3: Migrate Old Inputs                                   │
│   ↓                                                                  │
│  D1: Generate Secondary Techs Editor Template                        │
│   ↓                                                                  │
│  Manual Editing of Secondary_Techs_Editor.xlsx                       │
│   ↓                                                                  │
│  D2: Apply Edits to Model Files (Parametrization, Demand,            │
│      Interconnections)                                               │
├──────────────────────────────────────────────────────────────────────┤
│                    SCENARIO CREATION & EDITING                       │
│                                                                      │
│  To create a new scenario:                                           │
│    1. Duplicate the BAU folder (A1_Outputs_BAU → A1_Outputs_NEW)     │
│    2. Use D2 to parametrize A-O_Parametrization.xlsx and toggle      │
│       interconnections ON/OFF                                        │
│    3. Edit A-O_Demand.xlsx directly for demand changes               │
│                                                                      │
│  To edit parameters in an existing scenario:                         │
│    • Edit A1_Outputs files directly, OR                              │
│    • Use Secondary_Techs_Editor.xlsx + D2 for supported parameters   │
├──────────────────────────────────────────────────────────────────────┤
│                        MODEL EXECUTION                               │
│                                                                      │
│  B1: Compile Excel → OSeMOSYS CSVs                                  │
│   ↓                                                                  │
│  B2: Execute Solver → Results                                        │
└──────────────────────────────────────────────────────────────────────┘
```

---

## Stage A0: Generate Technology-Country Matrix

**Script:** `t1_confection/A0_generate_tech_country_matrix.py`

Generates `Tech_Country_Matrix.xlsx`, which controls which technology-country combinations are included in the model.

### Usage

From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python t1_confection/A0_generate_tech_country_matrix.py
```

### What It Does

1. Reads the country list from `Config_country_codes.yaml`.
2. Creates a matrix of 21 technology codes against all countries.
3. Marks implausible combinations (from `implausible_combinations` in the YAML) as **NO** with red highlighting.
4. All other combinations default to **YES**.
5. Writes the matrix to `Tech_Country_Matrix.xlsx` with 5 sheets.

### Technology Codes

| Code | Description |
|------|-------------|
| BCK | Backstop |
| BIO | Biomass |
| CCS | Carbon Capture & Storage (Coal) |
| COA | Coal |
| COG | Cogeneration |
| CSP | Concentrated Solar Power |
| GAS | Natural Gas |
| GEO | Geothermal |
| HYD | Hydroelectric |
| LDS | Long Duration Storage |
| NGS | Natural Gas (CCG + OCG unified) |
| OIL | Oil |
| OTH | Other |
| PET | Petroleum |
| SDS | Short Duration Storage |
| SPV | Solar Photovoltaic |
| URN | Nuclear |
| WAS | Waste |
| WAV | Wave |
| WOF | Offshore Wind |
| WON | Onshore Wind |

:::{note}
Structural prefixes (`ELC`, `MIN`, `PWR`, `RNW`, `TRN`) are **not** included in the matrix. They combine with the codes above to form full technology names (e.g., `PWRBIOBGDXX`, `MINCOABGDXX`).
:::

### After Generation

Edit `Tech_Country_Matrix.xlsx` to customize:
- In the **Matrix** sheet: change YES/NO for any technology-country pair.
- In the **NGS_Unification** sheet: toggle YES/NO to enable CCG+OCG merging into NGS.

---

## Stage A1: Preprocess Raw CSVs

**Script:** `t1_confection/A1_Pre_processing_OG_csvs.py`

The largest processing step. Reads raw OSeMOSYS CSV files and produces structured Excel model files for each scenario.

### Usage

From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python t1_confection/A1_Pre_processing_OG_csvs.py
```

### Input Files

- `OG_csvs_inputs/*.csv` -- All standard OSeMOSYS parameter and set CSV files.
- `Tech_Country_Matrix.xlsx` -- Technology filtering configuration.
- `Config_country_codes.yaml` -- Country definitions and settings.
- `Config_region_consolidation.yaml` -- Region consolidation rules.

### Output Files (per scenario)

Written to `A1_Outputs/A1_Outputs_{scenario}/`:

| File | Content |
|------|---------|
| `A-O_Parametrization.xlsx` | All technology parameters (costs, capacities, limits, etc.) |
| `A-O_Demand.xlsx` | Demand data, profiles, and projections |
| `A-O_AR_Model_Base_Year.xlsx` | Base year activity ratios (InputActivityRatio, OutputActivityRatio) |
| `A-O_AR_Projections.xlsx` | Projection activity ratios |

### Processing Steps

The script performs these operations in order:

1. **Read all CSV files** into memory as DataFrames.
2. **Replace country codes** according to the configuration.
3. **Filter by first year** -- removes data before `first_year`.
4. **Normalize temporal profiles** -- ensures SpecifiedDemandProfile sums to 1.0 per fuel/tech/year.
5. **Consolidate regions** (if enabled) -- merges sub-regional data using avg/sum rules.
6. **Remove internal interconnections** after consolidation.
7. **Clean PWR technologies** -- handles PWR00/PWR01 duplicates based on `pwr_cleanup_mode`.
8. **Apply Tech-Country Matrix filtering** -- removes technology-country pairs marked NO.
9. **Unify NGS technologies** -- merges CCG+OCG into NGS where enabled.
10. **Write Excel output files** with formatted sheets and human-readable names.
11. **Update demand profiles and projections**.
12. **Update parametrization capacities and temporal splits**.

---

## Stage A2: Add Transmission Technologies

**Script:** `t1_confection/A2_AddTx.py`

Adds transmission (TRN) and dispatch (DSPTRN) technology entries to the Excel model files. Creates 6 transmission types plus 1 dispatch type per country for renewable and non-renewable power routing, and updates interconnection fuel codes.

### Usage

From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python t1_confection/A2_AddTx.py
```

### Transmission Technology Types

| Code | Description |
|------|-------------|
| `RNWTRN` | Renewable transmission (existing) |
| `RNWRPO` | Renewable transmission (repowered) |
| `RNWNLI` | Renewable transmission (new lines) |
| `PWRTRN` | Non-renewable transmission (existing) |
| `TRNRPO` | Non-renewable transmission (repowered) |
| `TRNNLI` | Non-renewable transmission (new lines) |
| `DSPTRN` | Dispatch (interconnection routing, 2 modes) |

### Fuel Routing

The script assigns fuel codes for the energy flow:

- `ELC*00` -- Renewable electricity
- `ELC*01` -- Non-renewable electricity
- `ELC*02` -- Transmission output / Demand
- `ELC*03` -- Dispatch-ready for interconnection

### What It Does

1. Reads the country list from `Config_country_codes.yaml`.
2. For each scenario's Excel files:
   - Classifies power plant output as renewable (`ELC*00`) or non-renewable (`ELC*01`) in **Secondary** sheets.
   - Updates TRN interconnection fuel codes from `ELC*02`/`ELC*01` to `ELC*03` in **Secondary** sheets.
   - Adds transmission technology entries (RNWTRN, PWRTRN, etc.) to **Demand Techs** sheets.
   - Adds DSPTRN dispatch technology (Mode 1: `ELC*02` → `ELC*03`, Mode 2: `ELC*03` → `ELC*01`) to **Demand Techs** sheets.
   - Adds parameter entries to `A-O_Parametrization.xlsx` (sheets: **Fixed Horizon Parameters**, **Demand Techs**).

### Command-Line Options

| Flag | Default | Description |
|------|---------|-------------|
| `--yaml` | Auto-detected | Path to YAML configuration |
| `--base` | `A-O_AR_Model_Base_Year.xlsx` | Base year filename |
| `--proj` | `A-O_AR_Projections.xlsx` | Projections filename |
| `--param` | `A-O_Parametrization.xlsx` | Parametrization filename |

---

## Stage A3: Migrate Old Inputs (Optional)

**Script:** `t1_confection/A3_migrate_old_inputs_CLG.py`

Migrates data from an older input format (`Old_Inputs/` directory) into the current model structure. This is only needed when transitioning from a legacy data format.

### Usage

From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python t1_confection/A3_migrate_old_inputs_CLG.py
python t1_confection/A3_migrate_old_inputs_CLG.py --dry-run  # Preview without writing
```

### Required Setup

The `Old_Inputs/` folder **must be created manually** inside `t1_confection/`. It is not generated by any script. The folder must contain the following files, which must have **the same format** as the current model files:

```
Old_Inputs/
├── A2_Extra_Inputs/
│   └── A-Xtra_Storage.xlsx
└── A1_Outputs/
    └── A1_Outputs_{scenario}/
        ├── A-O_Parametrization.xlsx
        ├── A-O_Demand.xlsx
        ├── A-O_AR_Projections.xlsx
        └── A-O_AR_Model_Base_Year.xlsx
```

If the `Old_Inputs/` folder does not exist, the script will exit with an error.

### What It Does

- Applies technology name transformations (CCG+OCG to NGS, suffix removal).
- Imports and normalizes profiles from old files.
- Reads equivalence rules from `Config_tech_equivalences.yaml`.

---

## Stage D1: Generate Secondary Techs Editor Template (Optional)

**Script:** `t1_confection/D1_generate_editor_template.py`

Creates the `Secondary_Techs_Editor.xlsx` workbook, which provides a user-friendly interface for editing secondary technology parameters across all scenarios. This step is only needed when you want to regenerate the editor template (e.g., after adding a new scenario or country).

### Usage

From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python t1_confection/D1_generate_editor_template.py
```

### What It Does

1. Reads all `A-O_Parametrization.xlsx` files across scenario folders.
2. Collects TRN interconnection technologies from the base year file.
3. Generates `Secondary_Techs_Editor.xlsx` with dropdown lists, auto-fill formulas, and configuration sheets.

### Output Sheets

| Sheet | Purpose |
|-------|---------|
| **Instructions** | User guide |
| **OSTRAM_Config** | Toggle switches for OSTRAM data integration |
| **Editor** | Main editing area (Scenario, Country, Technology, Parameter, Year columns) |
| **Demand_Growth** | Demand growth rates per country |
| **Scenarios_Demand_Growth** | Scenario-specific demand growth overrides |
| **Renewability_Targets** | Renewable percentage targets per year/country |
| **Technology_Weights** | Custom distribution weights for technologies |
| **Interconnections** | Transmission interconnection ON/OFF controls (if TRN technologies exist) |

---

## Stage D2: Apply Edits to Model Files (Optional)

**Script:** `t1_confection/D2_update_secondary_techs.py`

Reads the filled `Secondary_Techs_Editor.xlsx` and applies all changes to the model files. This step is only needed when you want to apply changes to the model -- for example, when parametrizing a scenario, toggling interconnections, or updating demand growth rates.

### Usage

From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python t1_confection/D2_update_secondary_techs.py
```

### What It Does

1. Reads manual edit instructions from the **Editor** sheet.
2. Reads OSTRAM configuration toggles from the **OSTRAM_Config** sheet.
3. For each scenario:
   - Creates a **backup** of the files before modifying them.
   - Applies manual edits to `A-O_Parametrization.xlsx` (Secondary Techs sheet).
   - If OSTRAM integration is enabled: populates ResidualCapacity, demand, activity limits, and petroleum split.
   - Applies interconnection ON/OFF controls from the **Interconnections** sheet to base year and projection activity ratios.
   - Updates `A-O_Demand.xlsx` (if demand integration is enabled).
4. Sets **Projection.Mode** to "User defined" for modified parameters.
5. Generates a detailed log file (`secondary_techs_update_log_*.txt`).

### What Can Be Edited with Secondary_Techs_Editor + D2

| What | Where in Editor |
|------|-----------------|
| Technology cost and capacity parameters (CapitalCost, FixedCost, ResidualCapacity, etc.) | **Editor** sheet |
| Interconnection ON/OFF per direction | **Interconnections** sheet |
| Demand growth rates | **Demand_Growth** / **Scenarios_Demand_Growth** sheets |
| Renewable energy targets | **Renewability_Targets** sheet |
| Technology activity distribution weights | **Technology_Weights** sheet |
| OSTRAM automatic data integration toggles | **OSTRAM_Config** sheet |

### What Requires Direct Editing of A1_Outputs Files

| What | File to Edit Directly |
|------|-----------------------|
| Demand values and profiles | `A-O_Demand.xlsx` |
| Activity ratios (InputActivityRatio, OutputActivityRatio) | `A-O_AR_Model_Base_Year.xlsx` / `A-O_AR_Projections.xlsx` |
| Parameters not in Secondary Techs sheet | `A-O_Parametrization.xlsx` (other sheets) |

### Creating a New Scenario

To create a new scenario (e.g., NDC based on BAU):

1. **Duplicate** the base scenario folder: copy `A1_Outputs/A1_Outputs_BAU/` to `A1_Outputs/A1_Outputs_NDC/`.
2. **Regenerate the editor template** by running D1 (so the new scenario appears in dropdowns).
3. **Parametrize** the new scenario using `Secondary_Techs_Editor.xlsx`:
   - In the **Editor** sheet, select the new scenario and modify technology parameters.
   - In the **Interconnections** sheet, toggle interconnections ON/OFF for the new scenario.
4. **Edit demand** directly in `A1_Outputs/A1_Outputs_NDC/A-O_Demand.xlsx` for scenario-specific demand changes.
5. **Run D2** to apply all changes.

### Editing Parameters in an Existing Scenario

There are two approaches:

1. **Via Secondary_Techs_Editor + D2**: For parameters available in the editor (Secondary Techs parameters, interconnections, demand growth, renewability targets). Run D1 first if the template is outdated, fill in the editor, then run D2.
2. **Direct editing**: Go directly to the Excel files in `A1_Outputs/A1_Outputs_{scenario}/` and modify values manually. This is needed for demand profiles, activity ratios, and other parameters not covered by the editor.

---

## Stage B1: Compile to OSeMOSYS Format

**Script:** `t1_confection/B1_Compiler.py` (invoked via `B1_Run_Compiler.py`)

Reads the Excel model files and compiles them into OSeMOSYS-format CSV parameter files.

### Usage

This stage runs automatically via the DVC pipeline. From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python -u t1_confection/B1_Run_Compiler.py
```

### Input Files

- `A1_Outputs/A1_Outputs_{scenario}/A-O_*.xlsx` -- All Excel model files.
- `A2_Extra_Inputs/A-Xtra_*.xlsx` -- Extra inputs (storage, emissions, projections).
- `Config_MOMF_T1_A.yaml` -- Compiler configuration.

### Output Files

- `A2_Output_Params/{scenario}/*.csv` -- One CSV per OSeMOSYS parameter.
- `A2_Structure_Lists.xlsx` -- Generated structure/set listings.

### Compilation Logic

The compiler handles:

- **Projection modes**: Flat, yearly percentage change, user-defined, interpolation to stated value, zero.
- **Activity ratios**: InputActivityRatio and OutputActivityRatio from base year and projection sheets.
- **Parametrization**: All technology parameters (costs, capacities, operational life, etc.).
- **Transport** (when enabled): Fleet calculations, distance handling, occupancy rates.
- **Capacity limits**: Hard limits, lower limits, continuous residual capacity.

---

## Stage B2: Execute the Model

**Script:** `t1_confection/B2_Executing_OG_Model.py`

Runs the OSeMOSYS optimization model using the configured solver.

### Usage

This stage runs automatically via the DVC pipeline. From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python -u t1_confection/B2_Executing_OG_Model.py
```

### Execution Steps

1. **CSV to datafile conversion** via otoole.
2. **Preprocessing** -- runs the OSeMOSYS preprocessor.
3. **Solver execution** -- runs the selected solver (GLPK/CBC/CPLEX/Gurobi).
4. **Result extraction** -- converts solver output back to CSV.
5. **Post-processing** -- capital annualization, scenario concatenation.

### Solver Configuration

Configure in `Config_MOMF_T1_AB.yaml`:

```yaml
solver: 'cplex'        # glpk | cbc | cplex | gurobi
cplex_threads: 4
cplex_random_seed: 12345
```

### Parallel Execution

When `parallel: True` and `only_main_scenario: False`, multiple scenarios are solved simultaneously:

```yaml
parallel: True
max_x_per_iter: 4  # Max scenarios per batch
```

### Output Files

| Directory/File | Content |
|----------------|---------|
| `A2_Outputs_Params_otoole/{scenario}/` | otoole-format CSVs (one per parameter) |
| `Executables/` | Compiled solver data files |
| `OSTRAM_Inputs.csv` | Combined inputs (all scenarios) |
| `OSTRAM_Outputs.csv` | Combined outputs (all scenarios) |
| `OSTRAM_Combined_Inputs_Outputs.csv` | Merged inputs and outputs |

### Reproducibility

Deterministic results are ensured through:

- `PYTHONHASHSEED=0` (set by `run.py`).
- Configurable random seeds per solver.
- Sorted CSV files for consistent ordering.
