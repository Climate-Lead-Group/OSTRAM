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
│  (Optional) D1 → Manual Editing → D2: Secondary Techs Editing       │
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

Adds transmission (TRN) technology entries to the Excel model files. Creates 6 transmission types per country for renewable and non-renewable power routing.

### Usage

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

### Fuel Routing

The script assigns fuel codes for the energy flow:

- `ELC*00` -- Renewable electricity
- `ELC*01` -- Non-renewable electricity
- `ELC*02` -- Transmission output

### What It Does

1. Reads the country list from `Config_country_codes.yaml`.
2. For each scenario's Excel files:
   - Adds TRN entries to the **base year** activity ratios.
   - Adds TRN entries to the **projections** activity ratios.
   - Adds TRN entries to the **parametrization** (costs, capacities, lifetimes).

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

```bash
python t1_confection/A3_migrate_old_inputs_CLG.py
python t1_confection/A3_migrate_old_inputs_CLG.py --dry-run  # Preview without writing
```

### What It Does

- Applies technology name transformations (CCG+OCG to NGS, suffix removal).
- Imports and normalizes profiles from old files.
- Reads equivalence rules from `Config_tech_equivalences.yaml`.

---

## Stage B1: Compile to OSeMOSYS Format

**Script:** `t1_confection/B1_Compiler.py` (invoked via `B1_Run_Compiler.py`)

Reads the Excel model files and compiles them into OSeMOSYS-format CSV parameter files.

### Usage

This stage runs automatically via the DVC pipeline:

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

This stage runs automatically via the DVC pipeline:

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
