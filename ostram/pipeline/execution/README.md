# ostram/pipeline/execution — stage B2: solver files, solve, results and cross-scenario tables

B2 turns each compiled scenario (CSV parameters under `<workspace>/compilation/A2_Output_Params/`)
into an OSeMOSYS datafile, applies the datafile patch chain in `patches/`, builds the LP with
`glpsol`, runs the configured solver, converts results with otoole, reconciles costs and writes
the cross-scenario `Inputs` / `Outputs` / `Combined_Inputs_Outputs` CSVs. Configuration comes from
`config/execution/Config_MOMF_T1_AB.yaml` (`docs/configuration.md`); the stage contract is in
`docs/pipeline.md`, "Run the model and collect results (B2)". Everything is written under
`<workspace>/execution/`.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `runner.py` | B2 entry point. Holds the concrete operations (otoole conversion, preprocessing and patch subprocesses, combined input CSV, root datafile export, cleanup) and injects them into `orchestrator.orchestrate_b2` | `Config_MOMF_T1_AB.yaml`, `Config_MOMF_T1_A.yaml`, `<A2_output_otoole>/<scenario>/*.csv`, the otoole config, `model/osemosys_fast_preprocessed.txt` (trade marker), profile policies | `<Executables>/<scenario>_0/<scenario>_0.txt`, `Pre_processed_<scenario>_0[_chain].txt` siblings, `<scenario>_0_Input.csv`, the root datafile `OSTRAM_data.txt` / `OSTRAM_data_storage_delay.txt` at the project root | `python -m ostram run [--compile-only]` (stage B2); directly `python -B -m ostram.pipeline.execution.runner --scenarios A,B [--compile-only]`, the command `ostram/pipeline/orchestration.py` launches |
| `orchestrator.py` | Dependency-injected B2 orchestration: run plan from the YAMLs, per-scenario patch chain order, solver commands (glpk, cbc, cplex, gurobi), otoole results conversion, final post-processing, compile-only overrides | Execution and compilation YAML (via the injected loader), `<A2_output>` scenario folders, `.sol` files for CBC validation | `.lp`, `.sol`, `.cplex.log`, `.glp` under `<Executables>/<scenario>_0/`; `Outputs/` result CSVs via otoole; dated copies of the combined CSV | Imported only (from `runner.py`) |
| `preprocess.py` | Appends the `MODExTECHNOLOGYperFUEL*`, `MODEperTECHNOLOGY`, storage and emission helper sets to a datafile so matrix generation is faster | The otoole datafile (argv 1) | The preprocessed datafile (argv 2) | `python -m ostram.pipeline.execution.preprocess <in.txt> <out.txt>`; called by `runner.run_preprocessing_script` |
| `concatenate.py` | Combines one scenario's otoole result CSVs into a single output table and, when given the compiled datafile and the exporter's `VariableCost.csv`, reconciles the frames in `RECONCILED_COST_FRAMES` | `<scenario>_0/Outputs/*.csv`, optionally the final patched datafile, `RateOfActivity.csv`, `AnnualVariableOperatingCost.csv`, `VariableCost.csv` | `<output_file>.csv` next to the datafile (`Pre_processed_<scenario>_0[_chain]_output.csv`) and `<output_file>_cost_reconciliation.json` | `python -B -m ostram.pipeline.execution.concatenate <Outputs dir> <output_file> [--datafile F --parameter-csv F]`; called by `orchestrator.run_scenario_output_stage` when `concat_otoole_csv` is true |
| `scenario_concatenation.py` | Streams every `<Executables>/<scenario>_<future>/` input and output CSV into the cross-scenario `Inputs`, `Outputs` and `Combined_Inputs_Outputs` files, one scenario at a time, to bound memory | `<scenario>_0_Input.csv` and the active-chain output CSV of every scenario folder | `<prefix_final_files>Inputs.csv`, `...Outputs.csv`, `...Combined_Inputs_Outputs.csv` plus dated copies, under `<workspace>/execution/` | Imported only (from `runner.py`, run in the final post-processing stage when `concat_scenarios_csv` is true) |
| `annualization.py` | Converts lump-sum `CapitalInvestment` / `CapitalInvestmentStorage` rows of the combined CSV into capital-recovery-factor payment streams appended as `CapitalInvestmentAnnualized` / `CapitalInvestmentStorageAnnualized` | The combined CSV (chunked) and each scenario's original `<scenario>_<future>.txt` for `DiscountRate`, `DiscountRateStorage` and default `OperationalLife` | The combined CSV, rewritten in place through a `.tmp` file | `python -m ostram.pipeline.execution.annualization [--input F] [--executables DIR] [--discount-rate R] [--asset-lifetime N] [--quiet]`; in the pipeline only when `annualize_capital` is true |
| `gross_import_cap.py` | Renders the GMPL block for an aggregate annual received-import cap (`set D4ImportCountry`, `D4ImportDemand`, `D4ImportFlow`, `param D4ImportFraction`) from the scenario's `gross_import_cap` policy | `config/profiles/full.yaml` (`scenario_inputs.<scenario>.gross_import_cap`), `SpecifiedAnnualDemand.csv`, `OutputActivityRatio.csv` | Nothing; returns text that `runner.run_otoole_conversion` inserts before `end;` of `<scenario>_0.txt` | Imported only (from `runner.py`) |

## How to run

From the repo root in `OSTRAM-env`:

```bash
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU   # datafile + patch chain, no matrix, no solver
python -m ostram run --skip-pull --scenarios A_Calibrated_BAU --verbose        # full B2 with the configured solver (approval required)
```

`--compile-only` sets `execute_model`, `create_matrix`, `reuse_existing_sol`, `concat_otoole_csv`,
`concat_scenarios_csv`, `annualize_capital` and `del_files` to `False` and returns before the
execution stage (`orchestrator.apply_compile_only_overrides`).

## Order / dependencies

Per scenario, when `write_txt_model` is true (`orchestrator.run_compiled_input_stage`):

1. `runner.run_otoole_conversion` writes `<scenario>_0.txt` (and inserts the `gross_import_cap` block).
2. `preprocess.py` writes `Pre_processed_<scenario>_0.txt`.
3. `patches/days_in_day_type.py` rewrites that file in place (always, no flag).
4. `patches/storage_delay.py` if `storage_delay_active` (suffix `StorageDelayN5`).
5. `patches/strip_storage.py` if `strip_storage_active` (suffix `NoStorage`).
6. `patches/open_pwrbck_caps.py` if `open_pwrbck_active` (suffix `OpenBCK`).
7. `patches/reserve_margin_repair.py` if `reserve_margin_repair_active` (suffix `RMRepair`).
8. `patches/reserve_margin_repair_xlsx.py` if `reserve_margin_xlsx_active` (suffix `RMCarefulXLSX`).
9. `runner.generate_combined_input_file` writes `<scenario>_0_Input.csv`.

Each patcher reads the previous active patcher's output, so the solved datafile is named
`Pre_processed_<scenario>_0_<active suffixes>.txt` (with the committed YAML:
`..._StorageDelayN5_OpenBCK_RMCarefulXLSX.txt`). Then, per scenario: `glpsol` matrix build (non-glpk
solvers), solver, otoole results conversion into `Outputs/`, `concatenate.py`. After the last
scenario: `scenario_concatenation.py`, optional `annualization.py`, root datafile export, cleanup
when `del_files` is true.

## Gotchas

- `storage_delay_active` forces `strip_storage_active` to `False` with only a printed notice, and
  swaps `osemosys_model` for the patched model written by `storage_delay.py`
  (`orchestrator.apply_configuration_overrides`).
- `reuse_existing_sol: True` with an existing `.sol` skips the solve; check the mtime of
  `<output_file>.cplex.log` to know whether a solve happened (`orchestrator.execute_scenario`).
- The suffix chain order is duplicated in `runner.py`, `orchestrator.resolve_scenario_execution_paths`
  and `scenario_concatenation.active_output_csv_candidates`; they must stay in sync.
- `runner.run_reserve_margin_repair_patcher` passes `--open-capacity-value`,
  `--open-capacity-prefixes` and `--skip-capacity-opening`, which `patches/reserve_margin_repair.py`
  does not define; with `reserve_margin_repair_active: True` argparse would reject the command. The
  committed YAML keeps that flag `False` and uses the xlsx repair instead. Not fixed here.
- `concatenate.py` requires the accounting region to be exactly `GLOBAL` and hard-codes the 2023
  base year with mid-year discounting; reconciliation runs only when the profile policy
  `reconcile_variable_cost_exports` is true. Absent cost frames are skipped with a `RuntimeWarning`.
- `annualization.py` refuses to run twice on the same file, has no numeric defaults for life or
  rate, and any exception in it is caught and logged so the run continues without annualization.
- `scenario_concatenation.py` derives `Scenario` and `Future` from the folder name with
  `rpartition("_")`; a folder without `_` gets an empty future. It rereads every scenario folder
  present in `Executables/`, so campaign time grows with the folders on disk (`docs/pipeline.md`).
- CBC solutions are validated by the first `.sol` line (`optimal - objective value`) and flagged
  bound violations are tolerated only between `-1e-6` and `0`.
- `gross_import_cap.py` reads `config/profiles/full.yaml` directly, not the active profile, and
  requires exactly one `end;` in the otoole datafile.
- Several module docstrings still use pre-refactor script names (`inject_DaysInDayType.py`,
  `patch_storage_delay.py`, `patch_reserve_margin_repair*.py`, `main_executer`).
