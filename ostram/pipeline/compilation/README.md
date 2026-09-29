# ostram/pipeline/compilation — stage B1: materialized workbooks to OSeMOSYS parameter CSVs

B1 reads each materialized scenario's `A-O_*` workbooks under
`<workspace>/preparation/A1_Outputs/A1_Outputs_<scenario>/`, copies the maintained compiler
configuration into `<workspace>/compilation/`, sets the selected scenario in that runtime copy and
writes one CSV per OSeMOSYS parameter and set under `<workspace>/compilation/A2_Output_Params/`.
The maintained `config/compilation/Config_MOMF_T1_A.yaml` is never mutated (`docs/pipeline.md`,
"Compile model inputs (B1)"). The pure transform steps the compiler calls live in `transforms/`.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `runner.py` | Public B1 CLI: parses `--scenarios`, builds the default `B1Paths` and calls `orchestrator.orchestrate` with the real discoverer, updater and compiler runner | Nothing directly | Nothing directly | `python -m ostram compile-inputs [--scenarios a,b]`; also launched by `python -m ostram run` as `python -B -m ostram.pipeline.compilation.runner --scenarios ...` |
| `orchestrator.py` | Discovers `A1_Outputs_*` scenario folders, creates the runtime config copy when absent and rewrites its `A1_outputs` / `A2_*` keys, then for each scenario sets `xtra_scen.Main_Scenario` and runs `compiler.py` as a subprocess with cwd `<workspace>/compilation`, restoring the config from a `.bak` afterwards | `config/compilation/Config_MOMF_T1_A.yaml`, `<workspace>/compilation/Config_MOMF_T1_A.yaml`, `<workspace>/preparation/A1_Outputs/A1_Outputs_*` | The runtime config copy and its `.bak` | Imported only (`runner.py`) |
| `compiler.py` | The compiler itself: opens the scenario workbooks (base year, projections, demand, parametrization, emissions, storage), validates timeslices, builds the parameter tables and structure lists, allocates country fuel costs and delivers the CSVs. All work happens at import time | `Config_MOMF_T1_A.yaml` in the cwd, `<A1_outputs>/A1_Outputs_<scenario>/A-O_*.xlsx`, `<A2_extra_inputs>/A-Xtra_{Emissions,Storage}.xlsx`, `og_csvs_inputs/EMISSION.csv` (else `inputs/osemosys_global/EMISSION.csv`), `OSTRAM_PROFILE_POLICIES` (`validation_no_mutation`) | Completed workbooks (`*_Completed.xlsx`), `A2_Structure_Lists.xlsx`, `<A2_output_main_scen>/<scenario>/<Parameter>.csv` and set CSVs, `_fuel_cost_allocation_<scenario>.csv`, copies for `Other_Scenarios` | `python -B -m ostram.pipeline.compilation.compiler` with cwd `<workspace>/compilation` (launched by `orchestrator.py`; no flags) |

## How to run

From the repo root in `OSTRAM-env`, after A3 has materialized the scenarios:

```bash
python -m ostram compile-inputs --scenarios A_Calibrated_BAU
python -m ostram compile-inputs --scenarios "A_Calibrated_BAU,B_Optimised_VRE"
```

`dvc.yaml` stage `preprocess` runs `python -m ostram compile-inputs` for every discovered scenario.

## Order / dependencies

`runner.py` → `orchestrator.orchestrate` → per scenario: update `Main_Scenario` → subprocess
`compiler.py` → (inside the compiler, in order) `transforms.effects.read_config`,
`transforms.planning.build_transform_plan`, optional report-only pre-solver validation, workbook
opens, `transforms.validation` checks (demand, capacity, yearsplit, daysplit),
`transforms.tables` builders, `transforms.effects` writers, `transforms.delivery` CSV delivery.
B2 consumes the CSVs under `A2_Output_Params/`.

## Gotchas

- `compiler.py` has an authorship-only header instead of a module docstring and executes at import;
  never import it from other code. Its cwd must contain `Config_MOMF_T1_A.yaml`.
- The orchestrator skips scenario folders whose name contains `backup`, `snapshot` or
  `pre_experiment` or ends in `_<8 digits>`; a `--scenarios` filter naming an unknown scenario exits
  with code 1, and no `A1_Outputs_*` folder at all exits 0 with a warning.
- The runtime config copy holds the last scenario's `Main_Scenario` until the `.bak` restore in
  `finally`; a failed backup never triggers a restore (deliberate).
- Legacy GDP-based demand projection and retired transport workbooks fail closed
  (`docs/pipeline.md`); `Use_Transport` paths are still present in the compiler but off in the
  maintained config.
- Under profile policy `validation_no_mutation`, the compiler runs `pre_solver_validation` in
  report-only mode and exits early if the final workbook is invalid.
- `compiler.py` line 1113 compares a list with `1` (`DailyTimeBracket == 1`), which is never true;
  `transforms/validation.py` states the one-bracket branch is deliberately not repaired.
- Log text still names the compiler `B1_Compiler.py`; the orchestrator accepts only that name or
  `compiler.py` when building the command.
