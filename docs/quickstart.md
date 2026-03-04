# Quick Start

This page walks you through running OSTRAM for the first time.

## 1. Run the Full Pipeline

From an **Anaconda Prompt** (or any terminal with conda available):

```bash
cd OSTRAM
python run.py
```

The `run.py` launcher automatically:

1. Creates the Conda environment (`OG-MOMF-env`) if it does not exist.
2. Installs any missing dependencies.
3. Injects today's date into the DVC pipeline for date-stamped output files.
4. Runs `dvc repro` to execute all pipeline stages.
5. Restores the original `dvc.yaml` after completion.

### Command-Line Options

| Flag | Default | Description |
|------|---------|-------------|
| `--env-name` | `osemosys_env` | Conda environment name |
| `--env-file` | `environment.yml` | Path to the Conda environment file |
| `--dvc-file` | `dvc.yaml` | Path to the DVC pipeline file |
| `--date` | Today (`YYYY-MM-DD`) | Date stamp for output files |

Example with a custom date:

```bash
python run.py --date 2026-01-15
```

## 2. Pipeline Stages

The DVC pipeline has two stages that run in sequence:

### Stage 1: Preprocess (Compile)

```
python -u t1_confection/B1_Run_Compiler.py
```

Reads the Excel model files in `A1_Outputs/` and compiles them into OSeMOSYS-compatible CSV parameter files in `A2_Output_Params/`.

### Stage 2: Execute

```
python -u t1_confection/B2_Executing_OG_Model.py
```

Converts CSVs to the solver's data format, runs the optimization, and produces combined result files.

## 3. Output Files

After a successful run, results are generated in `t1_confection/`:

| File | Description |
|------|-------------|
| `OSTRAM_Inputs.csv` | Compiled model inputs (all scenarios) |
| `OSTRAM_Outputs.csv` | Optimization results (all scenarios) |
| `OSTRAM_Combined_Inputs_Outputs.csv` | Merged inputs and outputs |
| `OSTRAM_Inputs_YYYY-MM-DD.csv` | Date-stamped copy of inputs |
| `OSTRAM_Outputs_YYYY-MM-DD.csv` | Date-stamped copy of outputs |
| `OSTRAM_Combined_Inputs_Outputs_YYYY-MM-DD.csv` | Date-stamped combined file |

Date-stamped files preserve a complete execution history so you can compare runs over time.

## 4. Directory Structure Overview

```
OSTRAM/
├── run.py                          # Main launcher
├── dvc.yaml                        # DVC pipeline definition
├── environment.yaml                # Conda environment spec
├── concatenate_files/              # Post-processing scripts
│   └── concatenate_ostram.py
└── t1_confection/                  # Core model directory
    ├── Config_MOMF_T1_A.yaml       # Compiler configuration
    ├── Config_MOMF_T1_AB.yaml      # Execution configuration
    ├── Config_country_codes.yaml   # Country & technology definitions
    ├── Config_region_consolidation.yaml
    ├── OG_csvs_inputs/             # Raw OSeMOSYS CSV inputs
    ├── A1_Outputs/                 # Excel model files (per scenario)
    │   └── A1_Outputs_BAU/
    ├── A2_Extra_Inputs/            # Extra inputs (storage, emissions, etc.)
    ├── A2_Output_Params/           # Compiled parameter CSVs
    ├── A2_Outputs_Params_otoole/   # otoole-format CSVs for solver
    ├── Executables/                # Solver data files
    ├── Miscellaneous/              # Templates and auxiliary files
    ├── Tech_Country_Matrix.xlsx    # Technology-country config
    ├── Secondary_Techs_Editor.xlsx # Parameter editor
    ├── A0_generate_tech_country_matrix.py
    ├── A1_Pre_processing_OG_csvs.py
    ├── A2_AddTx.py
    ├── B1_Compiler.py
    ├── B1_Run_Compiler.py
    ├── B2_Executing_OG_Model.py
    ├── D1_generate_editor_template.py
    ├── D2_update_secondary_techs.py
    └── Z_*.py                      # Auxiliary tools
```

## 5. Typical Workflow

A typical modeling workflow follows these steps:

1. **Configure countries and technologies** in `Config_country_codes.yaml`.
2. **Generate the Tech-Country Matrix** (`A0`).
3. **Preprocess raw CSVs into Excel model files** (`A1`).
4. **Add transmission technologies** (`A2`).
5. **Optionally edit secondary technologies** (`D1` + manual editing + `D2`).
6. **Run the full pipeline** with `python run.py`.
7. **Analyze results** using the output CSVs or the interactive dashboards.

See {doc}`pipeline` for a detailed walkthrough of each stage.
