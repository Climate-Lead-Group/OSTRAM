# Quick Start

This page walks you through running OSTRAM for the first time.

## 1. Run the Full Pipeline

From an **Anaconda Prompt** (or any terminal with conda available):

```bash
cd OSTRAM
python run.py
```

The `run.py` launcher automatically:

1. Creates the Conda environment (`OSTRAM-env`) if it does not exist, and installs any missing dependencies into it.
2. Initializes the DVC repository if it does not exist yet (`dvc init`), and runs `dvc pull` if a DVC remote is configured.
3. Runs **A1 + A2** as a combo, but only if no `_post_a2_snapshot_*` folder exists yet in `t1_confection/A1_Outputs/` (A2 creates the snapshot at the end of its run). If a snapshot already exists, A1/A2 are skipped and A3 restores from it instead -- delete the snapshot folder if you need to force a full A1/A2 re-run from raw CSVs.
4. Discovers the active scenarios (via `t1_confection/A3_process/_scenarios.py list-active`, in dependency order) and runs **A3** once per scenario.
5. Runs **B1** (`B1_Run_Compiler.py`) and then **B2** (`B2_Executing_OG_Model.py`), passing the same scenario filter to both.

:::{note}
Unlike older versions of this pipeline, `run.py` does **not** call `dvc repro`. It invokes each stage script directly as a subprocess. `dvc.yaml` and `dvc pull` are only used for pulling versioned data, not for orchestrating execution.
:::

### Command-Line Options

| Flag | Default | Description |
|------|---------|--------------|
| `--env-name` | Read from `environment.yaml`, else `OSTRAM-env` | Conda environment name |
| `--env-file` | `environment.yaml` | Path to the Conda environment file |
| `--dvc-file` | `dvc.yaml` | Path to the DVC pipeline file (used only for the optional `dvc pull` check) |
| `--skip-pull` | off | Skip `dvc pull` even if a remote is configured |
| `--skip-a3` | off | Skip the A3 scenario-processing stage |
| `--skip-b1` | off | Skip the B1 compiler stage |
| `--skip-b2` | off | Skip the B2 execution stage |
| `--scenarios` | all active scenarios | Comma-separated scenario subset, e.g. `--scenarios "B_Optimised_VRE,C_Target_VRE"`. Propagated to A3, B1, and B2. |

Example running only two scenarios, skipping the solver stage:

```bash
python run.py --scenarios "A_Calibrated_BAU,B_Optimised_VRE" --skip-b2
```

## 2. Pipeline Stages Invoked by `run.py`

### A1 + A2 (conditional, raw-data preprocessing)

```
python -u t1_confection/A1_Pre_processing_OG_csvs.py
python -u t1_confection/A2_AddTx.py
```

Reads the raw OSeMOSYS CSVs in `OG_csvs_inputs/` and produces the per-scenario Excel model files in `A1_Outputs/`, then adds transmission/dispatch technologies. Runs only when no post-A2 snapshot exists yet.

### A3 (once per active scenario)

```
python -u t1_confection/A3_process.py --scenario <name>
```

Restores the BAU post-A2 snapshot, applies the SOASIA template + scenario-specific rule scripts, and writes the finished workbook set back into `A1_Outputs/A1_Outputs_<name>/`. See {doc}`pipeline` for the full stage breakdown.

### B1 (compile)

```
python -u t1_confection/B1_Run_Compiler.py [--scenarios "A,B,C"]
```

Compiles the Excel model files into OSeMOSYS-format CSV parameter files in `A2_Output_Params/`.

### B2 (execute)

```
python -u t1_confection/B2_Executing_OG_Model.py [--scenarios "A,B,C"]
```

Converts CSVs to the solver's data format, applies the configured patch chain (storage delay, storage stripping, backstop cap opening, reserve-margin repair -- see {doc}`pipeline`), runs the optimization, and produces combined result files.

## 3. Output Files

After a successful run, results are generated in `t1_confection/`. The exact filename prefix depends on `storage_delay_active` in `Config_MOMF_T1_AB.yaml`:

| `storage_delay_active` | Prefix used | Example combined file |
|---|---|---|
| `False` | `prefix_final_files` (`OSTRAM_`) | `OSTRAM_Combined_Inputs_Outputs.csv` |
| `True` | `storage_delay_prefix_final_files` (`OSTRAM_StorageDelay_`) | `OSTRAM_StorageDelay_Combined_Inputs_Outputs.csv` |

A run only ever writes **one** of these sets (the prefix is swapped wholesale, not both at once), so check the current value of `storage_delay_active` before looking for output files.

| File | Description |
|------|--------------|
| `{prefix}Inputs.csv` | Compiled model inputs (all scenarios) |
| `{prefix}Outputs.csv` | Optimization results (all scenarios) |
| `{prefix}Combined_Inputs_Outputs.csv` | Merged inputs and outputs |
| `{prefix}Inputs_YYYY-MM-DD.csv` | Date-stamped copy of inputs |
| `{prefix}Outputs_YYYY-MM-DD.csv` | Date-stamped copy of outputs |
| `{prefix}Combined_Inputs_Outputs_YYYY-MM-DD.csv` | Date-stamped combined file |

Date-stamped files preserve a complete execution history so you can compare runs over time. The date stamp is generated directly inside `B2_Executing_OG_Model.py` (today's date at the moment B2 runs) -- it is unrelated to `dvc.yaml`, which still declares literal `..._DATEPLACEHOLDER.csv` output names from an older mechanism that is no longer wired up to anything.

The compiled solver datafile is exported at the repo root as `OSTRAM_data.txt` (or `OSTRAM_data_storage_delay.txt` when storage delay is active).

## 4. Directory Structure Overview

```
OSTRAM/
├── run.py                          # Main launcher (A1/A2 → A3 → B1 → B2)
├── dvc.yaml                        # DVC data-versioning (not used to orchestrate execution)
├── environment.yaml                # Conda environment spec
├── TECHNICAL_INVENTORY.md          # Generated technical inventory of the whole repo
├── concatenate_files/
│   └── concatenate_ostram.py       # Result concatenation, invoked by B2
└── t1_confection/                  # Core model directory
    ├── Config_MOMF_T1_A.yaml       # Compiler configuration (years, timeslices, sheet/param lists)
    ├── Config_MOMF_T1_AB.yaml      # Execution configuration (solver, patch chain toggles)
    ├── Config_country_codes.yaml   # Country, technology & transmission-tech definitions
    ├── Config_region_consolidation.yaml
    ├── Config_tech_equivalences.yaml   # Reference mapping from the old LATAM/RELAC model (not read by the active pipeline)
    ├── OG_csvs_inputs/             # Raw OSeMOSYS CSV inputs
    ├── A1_Outputs/                 # Excel model files (per scenario)
    │   ├── _post_a2_snapshot_BAU/  # Snapshot A3 restores from for every scenario
    │   ├── A1_Outputs_BAU/
    │   ├── A1_Outputs_A_Calibrated_BAU/
    │   ├── A1_Outputs_B_Optimised_VRE/
    │   └── A1_Outputs_C_Target_VRE/
    ├── A3_process/                 # A3 scenario-processing engine (see Stage A3 below)
    │   ├── SOASIA_OSeMOSYS_Template_v18.xlsx   # Control/Restrictions + scenario template
    │   └── rules_scripts/
    │       └── configs/            # Per-scenario rule YAMLs (A_Calibrated_BAU/, B_Optimised_VRE/, C_Target_VRE/)
    ├── A2_Extra_Inputs/             # Extra inputs (storage, emissions, projections, battery replacement)
    ├── A2_Output_Params/            # B1-compiled parameter CSVs (per scenario)
    ├── A2_Outputs_Params_otoole/    # otoole-format CSVs for the solver
    ├── Executables/                 # Solver data files, LP/.sol, per-scenario result CSVs
    ├── Figures/                     # HTML outputs from the Z_AUX_* visualization scripts
    ├── Miscellaneous/               # Templates, GMPL model file, otoole schema, preprocessing script
    ├── Tech_Country_Matrix.xlsx     # Technology-country config
    ├── Secondary_Techs_Editor.xlsx  # Manual parameter editor (D1/D2)
    ├── firm_capacity_fallbacks_by_cr.xlsx   # Reserve-margin repair fallback data (B2)
    ├── Shares_PET_OIL_Split.xlsx / Shares_Power_Generation_Technologies.xlsx   # D2 source-data inputs
    ├── A0_generate_tech_country_matrix.py
    ├── A1_Pre_processing_OG_csvs.py
    ├── A2_AddTx.py
    ├── A3_process.py
    ├── B1_Compiler.py / B1_Run_Compiler.py
    ├── B2_Executing_OG_Model.py
    ├── D1_generate_editor_template.py
    ├── D2_update_secondary_techs.py
    ├── patch_storage_delay.py / strip_storage.py / open_pwrbck_caps.py / patch_reserve_margin_repair_careful_xlsx.py / inject_DaysInDayType.py   # B2 patch chain
    └── Z_*.py                       # Auxiliary tools (see {doc}`auxiliary-tools`)
```

## 5. Typical Workflow

A typical modeling workflow follows these steps. All terminal commands must be run from an **Anaconda Prompt** (or any terminal with conda available) with the `OSTRAM-env` environment activated.

1. **Configure countries and technologies** in `Config_country_codes.yaml`.
2. **Generate the Tech-Country Matrix** (`A0`).
3. **Preprocess raw CSVs into Excel model files** (`A1`) and **add transmission technologies** (`A2`) -- or just run `python run.py`, which does this automatically the first time.
4. **Define scenarios** in the SOASIA v18 template's `Control` sheet (scenario name, `rules_script` chain, optional `inherit_restrictions_from`), and add the corresponding YAML files under `t1_confection/A3_process/rules_scripts/configs/<Scenario>/`. This is the primary mechanism for creating a new scenario -- see {doc}`pipeline` (Stage A3) for the full YAML anatomy (retirement schedules, investment lids, VRE targets, interconnector relaxation, capacity floors).
5. **Run A3** for each scenario (`python run.py --skip-b1 --skip-b2`, or `A3_process.py --scenario <name>` directly) to materialize the scenario workbooks.
6. **Optional manual touch-ups** with the Secondary Techs Editor (`D1` + manual editing + `D2`) for one-off parameter overrides, interconnection ON/OFF toggles, or OSTRAM-source data integration not covered by a rule script. This step is independent of A3 and can be applied to any scenario's workbook set afterward.
7. **Run the full pipeline** with `python run.py` (or `--skip-a3` if you only need to recompile/re-execute an already-processed scenario set).
8. **Analyze results** using the output CSVs, or the interactive HTML dashboards produced by the `Z_AUX_generate_*`/`Z_AUX_*_dashboard.py` scripts (see {doc}`auxiliary-tools`).

See {doc}`pipeline` for a detailed walkthrough of each stage.
