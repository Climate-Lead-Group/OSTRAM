# Pipeline Workflow

OSTRAM processes energy system data through a multi-stage pipeline. This page documents each stage in detail, including its inputs, outputs, and configuration.

## Pipeline Overview

```
┌──────────────────────────────────────────────────────────────────────┐
│                        DATA PREPARATION                             │
│                                                                      │
│  A0: Generate Tech-Country Matrix                                    │
│   ↓                                                                  │
│  A1: Preprocess Raw CSVs → Excel Model Files (creates BAU folder)    │
│   ↓                                                                  │
│  A2: Add Transmission/Dispatch Technologies (snapshots BAU)          │
├──────────────────────────────────────────────────────────────────────┤
│                    SCENARIO GENERATION (A3)                          │
│                                                                      │
│  For each active scenario declared in the SOASIA v18 Control sheet:  │
│    1. Restore the post-A2 BAU snapshot                               │
│    2. Merge 20-timeslice fabric, apply template extensions/fixes     │
│    3. Run automatic pre-solver validation (auto-fix)                 │
│    4. Apply scenario-specific rule scripts (retirement schedules,    │
│       investment lids, VRE targets, interconnector relaxation, ...)  │
│    5. Deliver the finished workbook set to A1_Outputs_<scenario>/    │
├──────────────────────────────────────────────────────────────────────┤
│         OPTIONAL MANUAL TOUCH-UP (Secondary Techs Editor)             │
│                                                                      │
│  D1: Generate Secondary_Techs_Editor.xlsx from current A1_Outputs     │
│   ↓                                                                  │
│  Manual editing (parameter overrides, interconnection ON/OFF,        │
│      demand growth, OSTRAM source-data integration)                  │
│   ↓                                                                  │
│  D2: Apply edits back to the scenario's model files                  │
├──────────────────────────────────────────────────────────────────────┤
│                        MODEL EXECUTION                               │
│                                                                      │
│  B1: Compile Excel → OSeMOSYS CSVs (per scenario)                    │
│   ↓                                                                  │
│  B2: Patch datafile → Execute Solver → Results                       │
└──────────────────────────────────────────────────────────────────────┘
```

`run.py` runs A0 is **not** included in the automated chain -- it is a one-off setup step you run manually whenever the country/technology configuration changes. A1+A2 run automatically only the first time (or after deleting the post-A2 snapshot). A3, B1, and B2 always run automatically. D1/D2 are always manual and sit outside `run.py` entirely.

---

## Stage A0: Generate Technology-Country Matrix

**Script:** `t1_confection/A0_generate_tech_country_matrix.py`

Generates `Tech_Country_Matrix.xlsx`, which controls which technology-country combinations are included in the model.

### Usage

From an **Anaconda Prompt** (with the `OSTRAM-env` environment activated):

```bash
python t1_confection/A0_generate_tech_country_matrix.py
```

### What It Does

1. Reads the country list from `Config_country_codes.yaml`.
2. Creates a matrix of technology codes against all countries.
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
| GAS | Natural Gas (legacy code, superseded by NGS -- see below) |
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

The largest processing step. Reads raw OSeMOSYS CSV files and produces structured Excel model files for the `BAU` scenario.

### Usage

```bash
python t1_confection/A1_Pre_processing_OG_csvs.py
```

### Input Files

- `OG_csvs_inputs/*.csv` -- All standard OSeMOSYS parameter and set CSV files.
- `Miscellaneous/A-O_*.xlsx`, `A-Xtra_Emissions.xlsx`, `A-Xtra_Storage.xlsx` -- workbook templates.
- `Tech_Country_Matrix.xlsx` -- Technology filtering configuration.
- `Config_country_codes.yaml` -- Country definitions and settings.
- `Config_region_consolidation.yaml` -- Region consolidation rules.

### Output Files

Written to `A1_Outputs/A1_Outputs_BAU/`:

| File | Content |
|------|---------|
| `A-O_Parametrization.xlsx` | All technology parameters (costs, capacities, limits, etc.) |
| `A-O_Demand.xlsx` | Demand data, profiles, and projections |
| `A-O_AR_Model_Base_Year.xlsx` | Base year activity ratios (InputActivityRatio, OutputActivityRatio) |
| `A-O_AR_Projections.xlsx` | Projection activity ratios |

A1 also writes normalized CSVs back to `OG_csvs_inputs/` and updates parts of `Config_MOMF_T1_A.yaml` (years, temporal structure).

### Processing Steps

1. **Read all CSV files** into memory as DataFrames.
2. **Normalize temporal profiles** -- `SpecifiedDemandProfile`, `YearSplit`, `DaySplit` are re-normalized so they sum correctly per fuel/tech/year.
3. **Replace country code** `JAM` with `BRB` (a hardcoded legacy-data fix).
4. **Filter by first year** -- removes data before `first_year` (2023) through the last year (2050).
5. **Apply Tech-Country Matrix filtering** -- removes technology-country pairs marked NO.
6. **Unify NGS technologies** -- merges CCG+OCG into NGS where enabled.
7. **Consolidate regions** (if enabled in `Config_region_consolidation.yaml`) -- merges sub-regional data using avg/sum rules, then removes internal interconnections that become self-loops.
8. **Clean PWR technologies** -- handles PWR00/PWR01 duplicates based on `pwr_cleanup_mode`.
9. **Write Excel output files** for the `BAU` scenario, with formatted sheets and human-readable names.

:::{note}
A1 only produces the `BAU` scenario folder. All other scenarios (`A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE`, ...) are generated by **Stage A3** from the post-A2 BAU snapshot -- they are not separate A1 runs.
:::

---

## Stage A2: Add Transmission Technologies

**Script:** `t1_confection/A2_AddTx.py`

Adds transmission (TRN) and dispatch (DSPTRN) technology entries to the `BAU` Excel model files. Creates 6 transmission types plus 1 dispatch type per country-region pair, and rewrites the electricity fuel tiers to route power through cross-border interconnections. On success, it also creates the `_post_a2_snapshot_BAU/` folder that Stage A3 restores from for every scenario.

### Usage

```bash
python t1_confection/A2_AddTx.py
```

### Transmission Technology Types

| Code | Description |
|------|-------------|
| `RNWTRN` | Renewable transmission (existing) |
| `RNWRPO` | Renewable transmission (repowered) |
| `RNWNLI` | Renewable transmission (new lines, unplanned candidate) |
| `PWRTRN` | Non-renewable transmission (existing) |
| `TRNRPO` | Non-renewable transmission (repowered) |
| `TRNNLI` | Non-renewable transmission (new lines, unplanned candidate) |
| `DSPTRN` | Dispatch (interconnection routing, 2 modes) |

### Fuel Routing (Four Electricity Tiers)

The script assigns fuel codes for the energy flow:

- `ELC*00` -- Renewable power plant output.
- `ELC*01` -- Non-renewable power plant output.
- `ELC*02` -- Transmission line output (available for domestic demand and as DSPTRN Mode 1 input).
- `ELC*03` -- Dispatch-ready for interconnection (DSPTRN output on both modes; TRN interconnector reads this as input on the exporting side).
- `ELC*04` -- Imported electricity (TRN interconnector writes this on the importing side; DSPTRN Mode 2 reads it back to make it dispatch-ready again).

### What It Does

1. Reads the country/region list from `Config_country_codes.yaml`.
2. Classifies power plant output as renewable (`ELC*00`) or non-renewable (`ELC*01`) in the **Secondary** sheets.
3. If `enable_dsptrn: true` (current default): rewrites TRN interconnector fuel codes so their input reads `ELC*03` and their output writes `ELC*04`.
4. Adds transmission technology entries (RNWTRN, PWRTRN, etc.) to the **Demand Techs** sheets, converting `ELC*00`/`ELC*01` into `ELC*02`.
5. Adds the `DSPTRN` dispatch technology (Mode 1: `ELC*02` → `ELC*03`; Mode 2: `ELC*04` → `ELC*03`) to the **Demand Techs** sheets.
6. Adds parameter entries to `A-O_Parametrization.xlsx` (sheets: **Fixed Horizon Parameters**, **Demand Techs**), using the per-technology defaults from `Config_country_codes.yaml` (`RNWTRN`, `PWRTRN`, `DSPTRN`, etc.).
7. Copies the finished `A1_Outputs_BAU/` folder to `_post_a2_snapshot_BAU/`, replacing any previous snapshot.

### Command-Line Options

| Flag | Description |
|------|-------------|
| `--yaml` | Path to `Config_country_codes.yaml` |
| `--base` | Base year AR workbook filename |
| `--proj` | Projections AR workbook filename |
| `--param` | Parametrization workbook filename |
| `--demand` | Demand workbook filename |

---

## Stage A3: Scenario Generation

**Script:** `t1_confection/A3_process.py`, orchestrating scripts under `t1_confection/A3_process/`

This is the primary scenario-creation mechanism. It always restores its input from the post-A2 `BAU` snapshot (never the current scenario folder), then layers automated fixes, Stage-5 scenario rules, late WS3/WS4 transmission transformations, and--for the three canonical roots only--the static non-Maldives 2023--2026 PWR/MIN pin before delivering the finished workbook set to `A1_Outputs/A1_Outputs_<scenario>/`. `run.py` calls it once per active scenario, in dependency order.

### Usage

```bash
python t1_confection/A3_process.py --scenario BAU
python t1_confection/A3_process.py --scenario B_Optimised_VRE
```

### Command-Line Options

| Flag | Default | Description |
|------|---------|-------------|
| `--scenario` | `BAU` | Scenario name. Must exist in the SOASIA v18 `Control` sheet. |
| `--soasia` | `A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx` | Override the template path (useful for testing new scenarios without touching the canonical file). |
| `--rules-script` | From `Control` sheet | Override the scenario's rule-script chain. An empty string skips Stage 5 only; it does not disable the root-gated late-WS4 pin. |
| `--inherit-from` | From `Control` sheet | Override `inherit_restrictions_from` (comma-separated scenario names). |
| `--input-dir` | `A1_Outputs/A1_Outputs_<scenario>` | Override the input workbook directory. |
| `--output-dir` | Same as `--input-dir` | Override the delivery directory. |
| `--keep-workdir` | off | Preserve the `_run_<timestamp>/` working directory for debugging instead of deleting it. |

### Scenario Definition (SOASIA v18 `Control` Sheet)

Scenarios are declared as rows in the `Control` sheet of `SOASIA_OSeMOSYS_Template_v18.xlsx`, with columns:

| Column | Meaning |
|--------|---------|
| `scenario` | Scenario name (matches `A1_Outputs_<scenario>`) |
| `active` | Whether `run.py`/`_scenarios.py list-active` include it |
| `rules_script` | Comma/newline-separated list of `.py` filenames under `rules_scripts/`, run in order in Stage 5 |
| `inherit_restrictions_from` | Comma-separated scenario names whose persisted `Restrictions` rows this scenario reads (creates a dependency edge) |
| `notes` | Free-text |

To add a new scenario: add a `Control` row, create scenario-tagged override rows in the template's parametric sheets if needed, and create `t1_confection/A3_process/rules_scripts/configs/<Scenario>/` with the YAML files the chosen rule scripts expect.

### Execution Order

For a single `--scenario` invocation, `A3_process.py` runs:

```text
0. Restore input from _post_a2_snapshot_BAU (always, regardless of scenario).
1. Build a temporary workdir: t1_confection/A3_process/_run_<timestamp>/.
2. Stage 0: materialize the SOASIA v18 scenario template (_scenarios.py).
3. Stage 0.5: fix_rnwbio_restore.py -- restore RNWBIO reference rows.
4. Stage 1: 1_merge_timeslices_into_WV.py, 2_extract_ao_extensions.py,
   3_update_ao_from_extensions.py, 4_apply_manual_fixes.py,
   5_propagate_timeslice_fabric.py.
5. Stage 1b: A0_insert_reserve_margin.py, add_max_capacity_investment_rule_OLD_8ee8056.py,
   add_max_capacity_investment_rule_NEW_2be1616.py, fix_elc_pmode_revert.py,
   B1b_Pre_solver_validation.py --auto-fix-all (auto-fixes infeasible
   Min/Max capacity investment and activity-limit combinations).
6. Stage 2/2.5: patch_ao_c2a.py, fix_pwrpet_clear.py.
7. Stage 3: fix_trn_residuals.py --mode min --cutoff-year 2023,
   clear_stale_unbinding_caps.py, cap_trn_to_residual.py.
8. Stage 4: consolidate the 4 final workbooks.
9. Stage 4.5: apply inherited Restrictions from inherit_restrictions_from (skipped if empty).
10. Stage 5: run the scenario's rules_script chain, each against the consolidated workbooks.
11. Late WS3/WS4: apply interconnector costs, internal-transmission calibration,
    and internal-transmission losses.
12. Late WS4: for the exact A_Calibrated_BAU, B_Optimised_VRE, and C_Target_VRE
    roots only, validate and apply pwr_min_2023_2026_pin.csv to its explicit
    2023--2026 PWR/MIN workbook cells. BAU is untouched; descendants inherit the
    corrected root before their sensitivity patches.
13. Stage 6: 6_sync_og_to_ts20.py -- sync OG_csvs_inputs/Config_MOMF_T1_A.yaml to the
    20-timeslice fabric; persist each rule script's CHANGES.json into the SOASIA
    v18 Restrictions sheet.
14. Deliver the final workbook set to A1_Outputs/A1_Outputs_<scenario>/; remove the
    workdir unless --keep-workdir.
```

### Scenario Rule Scripts

The scenario-configured scripts live under `t1_confection/A3_process/rules_scripts/`, each reading a scenario-specific YAML from `rules_scripts/configs/<Scenario>/`. Those scripts normally write a timestamped backup and a `*_CHANGES.json` change log before/after modifying the workbook, and accept `--input-dir`, `--sheets`, `--skip-backup`, `--yaml`; most (all except `add_max_cap_investment_lid_rule.py`) also accept `--self-test`, `--restore`, `--restore-from`.

`apply_base_year_pin.py` is a dedicated late-WS4 transformer, not a
scenario-configured Stage 5 rule. It validates the embedded production-allowlist
digest, accepts the exact root scenario explicitly, reads no solver output, and
emits no generic `*_CHANGES.json`, so Stage 6 cannot persist it as a competing
Restrictions authority.

| Script | What it does |
|--------|---------------|
| `set_retirement_schedule.py` | Draws down `ResidualCapacity` for thermal technologies via age-based (linear/logistic) and/or explicit scheduled retirement trajectories. |
| `add_max_cap_investment_lid_rule.py` | Computes a per-year cap ("lid") on `TotalAnnualMaxCapacityInvestment` for generation technologies, anchored to demand growth with a security-factor floor above minimum investment. |
| `relax_interconnectors.py` | Opens `TotalAnnualMaxCapacityInvestment` (and, where overridden, `TotalAnnualMaxCapacity`) for TRN interconnectors beyond their residual capacity, via a uniform headroom factor or explicit per-corridor overrides. |
| `set_min_capacity_floors.py` | Applies exogenous min/max capacity or activity floors/ceilings from a curated YAML (national plans, CCDRs/IRPs), matched by technology pattern and country-region. |
| `set_vre_targets.py` | Sets renewable generation/capacity floors per country-region and technology as a percentage of a prior scenario's solved `ProductionByTechnology.csv` (typically `A_Calibrated_BAU`), with optional capacity-envelope bounding. |
| `apply_base_year_pin.py` | Applies the audited static 2023--2026 PWR/MIN allowlist to exact canonical roots after WS3; it is solver-independent and leaves excluded cells unchanged. |

### Current Scenarios

| Scenario | Rule-script chain | Summary |
|----------|--------------------|---------|
| `BAU` | (none -- produced by A1/A2/A3 with no rule scripts) | Business-as-usual baseline. |
| `A_Calibrated_BAU` | `set_retirement_schedule`, `add_max_cap_investment_lid_rule`, `set_min_capacity_floors`, `relax_interconnectors` | Calibration/validation scenario: reproduces BAU proportions with a flat (non-relaxing) investment lid and no interconnector headroom. Its solved output feeds `set_vre_targets.py` for `C_Target_VRE`. |
| `B_Optimised_VRE` | `set_retirement_schedule`, `add_max_cap_investment_lid_rule`, `relax_interconnectors`, `add_storage_min_investment` | Least-cost VRE-optimized scenario: investment lids relax sharply over time (up to 25x by 2050) for non-locked technologies, interconnector caps get 1.5x headroom plus evidence-based per-corridor overrides, and storage gets small seed investment floors. Solved independently from the BAU snapshot (no inheritance). |
| `C_Target_VRE` | `set_retirement_schedule`, `add_max_cap_investment_lid_rule`, `set_vre_targets`, `relax_interconnectors`, `add_storage_min_investment` | NDC-target-driven VRE scenario: layers explicit country/region renewable-generation floors (from NDC pledges, pinned near target) on top of the same relaxed-lid/loose-interconnector machinery as `B_Optimised_VRE`. |

The static pin is independent of each scenario's Stage-5 rule chain. It dispatches
only for `A_Calibrated_BAU`, `B_Optimised_VRE`, and `C_Target_VRE`; `BAU` and
derived scenarios are never direct pin targets.

:::{warning}
`set_vre_targets.py` (used by `C_Target_VRE`) requires a solved `A_Calibrated_BAU` run to already exist under `Executables/` -- run/solve `A_Calibrated_BAU` through B1/B2 before generating `C_Target_VRE`.
:::

---

## Optional Manual Layer: Secondary Techs Editor (D1/D2)

D1/D2 are a separate, always-manual mechanism for touching up whatever workbook set currently exists in `A1_Outputs_<scenario>/` -- whether produced by plain A1/A2 or by the full A3 rule-script engine. They are not part of `run.py` and A3 never invokes them; use them for one-off parameter overrides, interconnection ON/OFF toggling, or auto-populating parameters from OSTRAM source data that isn't covered by a rule script. See {doc}`secondary-techs-editor` for full details.

```bash
python t1_confection/D1_generate_editor_template.py   # generates Secondary_Techs_Editor.xlsx
# ... fill in the Editor / Interconnections / Demand_Growth / etc. sheets ...
python t1_confection/D2_update_secondary_techs.py      # applies the edits back
```

Neither script takes command-line arguments; both operate on whatever scenario folders currently exist under `A1_Outputs/`.

---

## Stage B1: Compile to OSeMOSYS Format

**Script:** `t1_confection/B1_Compiler.py` (invoked via `B1_Run_Compiler.py`)

Reads the Excel model files and compiles them into OSeMOSYS-format CSV parameter files.

### Usage

```bash
python -u t1_confection/B1_Run_Compiler.py
python -u t1_confection/B1_Run_Compiler.py --scenarios "BAU,B_Optimised_VRE"
```

`B1_Run_Compiler.py` discovers scenario folders by their `A1_Outputs_*` suffix (skipping any containing `backup`, `snapshot`, `pre_experiment`, or an 8-digit datestamp). With `--scenarios` omitted, it compiles **every** discovered scenario; an unknown name in `--scenarios` aborts with an error. For each scenario it temporarily rewrites `xtra_scen.Main_Scenario` in `Config_MOMF_T1_A.yaml`, runs `B1_Compiler.py` (which itself takes no arguments), and restores the YAML from a `.bak` backup afterward regardless of success.

Discovery is alphabetical and filtering preserves that order; the order in a comma list
does not control execution. The runner logs an individual compiler failure and continues,
so verify every requested scenario artifact instead of trusting only the overall process
exit code.

### Input Files

- `A1_Outputs/A1_Outputs_{scenario}/A-O_*.xlsx` -- All Excel model files.
- `A2_Extra_Inputs/A-Xtra_*.xlsx` -- Extra inputs (storage, emissions, projections, battery replacement).
- `Config_MOMF_T1_A.yaml` -- Compiler configuration.
- `OG_csvs_inputs/EMISSION.csv` -- when `Use_OG_module: true`.

### Output Files

- `A2_Output_Params/{scenario}/*.csv` -- One CSV per OSeMOSYS parameter/set.
- `A2_Structure_Lists.xlsx` -- Generated structure/set listings.
- `A-O_Demand_COMPLETED.xlsx`, `A-O_Parametrization_COMPLETED.xlsx`, `A-O_Parametrization_Natural_COMPLETED.xlsx`, `A-O_AR_Projections_COMPLETED.xlsx` -- user-facing completed workbooks per scenario.

### Compilation Logic

- **Projection modes**: `Flat`, `Yearly percent change`, `User defined`, `Interpolate to stated end value from projection parameter`, `Zero` -- plus demand-specific modes (`GDP`, `GDP joint <tech>`, `Flat after final year`, `Interpolate to final value`, `Percent growth of incomplete years`).
- **Activity ratios**: `InputActivityRatio`/`OutputActivityRatio` from the base-year and projection AR workbooks.
- **Parametrization**: All technology parameters (costs, capacities, operational life, etc.).
- **Storage**: `StorageLevelStart`, `OperationalLifeStorage`, `CapitalCostStorage`, `ResidualStorageCapacity`, `TechnologyToStorage`/`FromStorage` from `A-Xtra_Storage.xlsx`.
- **System parameters**: an optional `System Parameters` sheet in `A-O_Parametrization.xlsx` supplies `ReserveMargin`.
- **Transport**: only compiled `if Use_Transport` (currently `false` -- dead code path in the active configuration).
- **Timeslice conversions**: `Conversionls`/`Conversionld`/`Conversionlh` are derived from the `xtra_scen` block in `Config_MOMF_T1_A.yaml` -- see {doc}`data-reference` for the current 20-timeslice structure.

---

## Stage B2: Execute the Model

**Script:** `t1_confection/B2_Executing_OG_Model.py`

Runs the OSeMOSYS optimization model using the configured solver, after applying a chain of datafile patches.

### Usage

```bash
python -u t1_confection/B2_Executing_OG_Model.py
python -u t1_confection/B2_Executing_OG_Model.py --scenarios "BAU,B_Optimised_VRE"
```

With `--scenarios` omitted, B2 runs every subfolder of `A2_Output_Params/` (except `Default`), or only the YAML `Main_Scenario` if `only_main_scenario: true`.

### Patch Chain

Applied in this fixed order to the preprocessed datafile before solving, each stage gated by its own YAML flag (current values from `Config_MOMF_T1_AB.yaml` shown):

| Step | Script | Flag | Current value |
|------|--------|------|----------------|
| 1 | otoole CSV → datafile conversion | -- | always |
| 2 | `Miscellaneous/preprocess_data.py` | -- | always |
| 3 | `inject_DaysInDayType.py` | -- | always (no gate) |
| 4 | `patch_storage_delay.py` | `storage_delay_active` | **True** -- delays storage builds for the first `storage_delay_first_n_years` (5) years |
| 5 | `strip_storage.py` | `strip_storage_active` | **False** (and forced off whenever storage delay is active) |
| 6 | `open_pwrbck_caps.py` | `open_pwrbck_active` | **True** -- opens PWRBCK capacity caps to `open_pwrbck_value` (9999) |
| 7 | `patch_reserve_margin_repair_careful.py` (older, blunt) | `reserve_margin_repair_active` | False |
| 8 | `patch_reserve_margin_repair_careful_xlsx.py` (current) | `reserve_margin_xlsx_active` | **True** -- uses `firm_capacity_fallbacks_by_cr.xlsx` |

:::{important}
`storage_delay_active` and `strip_storage_active` are mutually exclusive: if both are `True` in the YAML, B2 forces `strip_storage_active` to `False` at startup. Storage delay always wins. Check the live YAML rather than assuming a flag combination -- these values change frequently between experiments.
:::

Each active patch appends a suffix to the preprocessed filename (e.g. `Pre_processed_BAU_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt`), so the datafile passed to the solver reflects exactly which patches ran.

### Output File Prefix

When `storage_delay_active: True` (the current default), B2 overrides `prefix_final_files` for the whole run from `OSTRAM_` to `storage_delay_prefix_final_files` (`OSTRAM_StorageDelay_`), and exports the root datafile as `storage_delay_root_datafile` (`OSTRAM_data_storage_delay.txt`) instead of `OSTRAM_data.txt`. A single run only ever produces one of these two output sets -- check which files are actually current (by modification date) rather than assuming the plain `OSTRAM_*` names are up to date.

### Solver Configuration

Configure in `Config_MOMF_T1_AB.yaml`:

```yaml
solver: 'cplex'        # glpk | cbc | cplex | gurobi
cplex_threads: 4
cplex_random_seed: 12345
```

For any non-GLPK solver, the LP matrix is still built through `glpsol --wlp ... --check` first. The current CPLEX path runs `optimize` and writes the resulting `.sol`:

```
cplex -c "set logfile ..." "read ....lp" "set threads N" "set randomseed S" "set parallel 1" "optimize" "write ....sol"
```

FeasOpt is deliberately off. B2 deletes stale `.feasopt.sol` files before a CPLEX run;
solver feasibility/status must be read from the CPLEX log and the expected `.sol`, not
inferred from the presence of an old FeasOpt artifact.

### Parallel Execution

```yaml
parallel: False        # currently disabled
max_x_per_iter: 4       # max scenarios per batch, when enabled
```

### Output Files

| Directory/File | Content |
|----------------|---------|
| `A2_Outputs_Params_otoole/{scenario}/` | otoole-format CSVs (one per parameter) |
| `Executables/{scenario}_0/` | Compiled datafiles, `.lp`, `.sol`, per-scenario result CSVs |
| `{prefix}Inputs.csv` / `{prefix}Outputs.csv` / `{prefix}Combined_Inputs_Outputs.csv` | Combined result files, with `{prefix}` depending on `storage_delay_active` (see Quick Start) |

### Reproducibility

The pipeline reduces avoidable nondeterminism through:

- `PYTHONHASHSEED=0` (set by `run.py`).
- Configurable random seeds per solver.
- Sorted CSV files for consistent ordering.
- `reuse_existing_sol: True` lets you regenerate outputs from a previous `.sol` file without re-solving (falls back to a normal solve if the `.sol` is missing).

These controls do not prove numerical equivalence across solver/toolchain versions. Use
the regression evidence policy in {doc}`regression`; full behavioral equivalence requires
a coherent solver-backed baseline.
