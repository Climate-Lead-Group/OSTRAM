# Core workflow characterization

This note records the pre-refactor, no-solver contract for OSTRAM's core workflow.
The tests in `tests/regression/test_core_workflow_characterization.py` enforce the
contract with AST inspection, temporary fixtures, and mocked process boundaries.
They do not run a model stage, DVC, a batch file, or a solver. This is
characterization evidence, not a claim of behavioral or numerical equivalence.

## Entrypoint classification

The primary core entrypoints are protected operational code:

- `run.py`
- `t1_confection/A0_generate_tech_country_matrix.py`
- `t1_confection/A1_Pre_processing_OG_csvs.py`
- `t1_confection/A2_AddTx.py`
- `t1_confection/A3_process.py`
- `t1_confection/B1_Run_Compiler.py`
- `t1_confection/B1_Compiler.py`
- `t1_confection/B2_Executing_OG_Model.py`

The following are optional model-writing entrypoints. They are core/protected even
though `run.py` does not call them:

- `t1_confection/D1_generate_editor_template.py`
- `t1_confection/D2_update_secondary_techs.py`
- `t1_confection/Z_AUX_D1b_set_trn_limits_from_flows.py`

The following are canonical analysis utilities, not core workflow stages. Each retains
a compatibility wrapper at its former `t1_confection/` path:

- `tools/analysis/check_combined.py` and `t1_confection/check_combined.py`
- `tools/analysis/ostram_scenario_analysis.py` and
  `t1_confection/ostram_scenario_analysis.py`
- `tools/analysis/ostram_trn_plotter.py` and `t1_confection/ostram_trn_plotter.py`
- `tools/analysis/slice_by_country.py` and `t1_confection/slice_by_country.py`
- `tools/analysis/analyse_sensitivity.py` and
  `t1_confection/analyse_sensitivity.py`
- `tools/analysis/concat_all_scenarios.py` and
  `t1_confection/concat_all_scenarios_2.py`
- `tools/analysis/reproduce_A1_A6.py` and `t1_confection/reproduce_A1_A6.py`
- `t1_confection/sensitivity_expansion/analyse_ws4_vs_phaseB.py` (analysis-only,
  but retained in place because the protected-tree gate covers this path)
- `tools/analysis/visualization/Z_AUX_generate_interactive_dashboards_aggregated.py`
  and `t1_confection/Z_AUX_generate_interactive_dashboards_aggregated.py`
- `tools/analysis/visualization/Z_AUX_generate_RES_diagram.py` and
  `t1_confection/Z_AUX_generate_RES_diagram.py`
- `tools/analysis/visualization/Z_AUX_generate_transmission_maps.py` and
  `t1_confection/Z_AUX_generate_transmission_maps.py`
- `tools/analysis/visualization/Z_AUX_interconnections_dashboard.py` and
  `t1_confection/Z_AUX_interconnections_dashboard.py`

Validation code, fail-closed stubs, and material under `docs/archive/` are neither
core stages nor analysis utilities. A helper with a historical-looking name remains
core when A3 or B2 still calls it.

## Frozen call paths

`run.py` owns environment/DVC setup and the outer stage sequence. When no
`_post_a2_snapshot_*` directory exists, its path is
`A1_Pre_processing_OG_csvs.py` -> `A2_AddTx.py` -> per-scenario `A3_process.py` ->
`B1_Run_Compiler.py` -> `B2_Executing_OG_Model.py`. With a snapshot, A1 and A2 are
skipped. Active A3 scenarios come from `_scenarios.py list-active` in topological
order. An explicit scenario filter preserves that active order for A3, while the
original comma-separated filter order is forwarded to B1 and B2. A failed delegated
command propagates out of the launcher.

`B1_Run_Compiler.py` follows `list_scenario_suffixes` -> `update_main_scenario` -> `run_compiler`
for every selected scenario. Discovery sorts `A1_Outputs_*` directories and excludes
backup, snapshot, pre-experiment, and dated names. Filtering preserves discovery
order. The runner backs up `Config_MOMF_T1_A.yaml` before iteration and restores it in
a `finally` block. A config-update failure skips that scenario; a nonzero compiler
return is reported and iteration continues. `run_compiler` launches
`B1_Compiler.py` with the current interpreter and the `t1_confection/` working
directory.

`A3_process.py` restores every selected input directory from
`_post_a2_snapshot_BAU` before building its work directory. Its frozen stage path is
`stage_0_5_rnwbio` -> `stage_1_scripts_1_to_5` -> `stage_1b` ->
`stage_2_and_2_5` -> `stage_3_fix_2` -> `stage_4_consolidate` ->
`stage_4_5_apply_inherited_restrictions` -> `stage_5_rules_scripts` ->
`stage_ws3_interconnector_costs` -> `stage_ws3_internal_transmission` ->
`stage_ws3_internal_tx_losses` -> `stage_6_sync_og_to_ts20` ->
`stage_6_persist_restrictions` -> `deliver_outputs`. A scenario-specific rules YAML
takes precedence over the default rules YAML. Helper failure is fail-fast through
`run_subproc`; work-directory and environment cleanup currently occur only after the
ordered path succeeds.

`B2_Executing_OG_Model.py` discovers sorted entries under the configured A2 output,
removes `Default`, optionally reduces the list to `Main_Scenario`, and then applies a
CLI filter without reordering the discovered list. Its compiled-input path is
`run_otoole_conversion` -> `run_preprocessing_script` -> `run_days_in_day_type_patcher` ->
`run_storage_delay_patcher` -> `run_strip_storage_patcher` ->
`run_open_pwrbck_patcher` -> `run_reserve_margin_repair_patcher` ->
`run_reserve_margin_xlsx_patcher`. `main_executer` is the only solve boundary for
both linear calls and `multiprocessing.Process`; it constructs GLPK, CBC, CPLEX, or
Gurobi commands and launches them only under the existing execution/configuration
flags. Solver execution remains outside this test phase.

## `run.py` launcher contract

Importing `run.py` defines constants and functions but does not parse arguments,
inspect the repository, start a process, or run a model stage. The existing
`argparse` definition remains inline at the beginning of `main()` and reads the
process argument list when `main()` is called.

The accepted command line is:

| Option | Parsed default | Current interpretation |
|---|---|---|
| `--env-name VALUE` | `None` | Any string is accepted. A false/omitted value falls back to the `name:` in `--env-file`, then `OSTRAM-env`. |
| `--env-file VALUE` | `environment.yaml` | Passed as a current-working-directory-relative string to environment-name lookup and environment creation. |
| `--dvc-file VALUE` | `dvc.yaml` | Resolved against the current working directory and printed. It does not select a DVC command or stage. |
| `--skip-pull` | `False` | Bypasses both the DVC remote check and `dvc pull`; environment and DVC setup still run. |
| `--skip-a3` | `False` | Skips active-scenario enumeration, validation, and all A3 calls. It does not skip snapshot-gated A1/A2. |
| `--skip-b1` | `False` | Skips the B1 runner call only. |
| `--skip-b2` | `False` | Skips the B2 runner call only. |
| `--scenarios VALUE` | `None` | Accepts one unrestricted string, later split on commas, trimmed, and stripped of empty items. |

There are no option choices or positional arguments. Standard `argparse` behavior
applies: help exits zero, while an unknown option or missing option value exits two
before launcher setup. Those `SystemExit` results are not handled by the script's
`Exception` handlers.

Launcher setup always precedes stage selection. It prints the selected environment
and resolved DVC path, checks `conda --version`, creates the Conda environment when
absent, checks/installs dependencies, initializes DVC when absent, and optionally
pulls from a configured remote. These are unsafe external mutation boundaries and
are mocked in the characterization suite. `run.py` does not call `dvc repro`.

After setup, the outer order is fixed:

1. test for any `_post_a2_snapshot_*` directory;
2. if none exists, call A1 and then A2, even when every `--skip-a3`/`--skip-b1`/
   `--skip-b2` flag is present;
3. unless skipped, enumerate active scenarios and call A3 once per selected scenario;
4. unless skipped, call B1 once; and
5. unless skipped, call B2 once.

Scenario propagation deliberately remains asymmetric. With no `--scenarios`, A3
uses `_scenarios.py list-active` in topological order, while B1 and B2 receive no
filter and perform their own independent discovery. With an explicit filter and A3
enabled, unknown names are rejected only after setup and any required A1/A2 calls;
A3 then follows active topological order and removes duplicate occurrences. B1 and
B2 instead receive the trimmed original comma order, including duplicates. With
`--skip-a3`, active-name validation is bypassed and the trimmed filter is forwarded
unchanged. An explicitly supplied string containing only commas/whitespace selects
no A3 scenarios but is falsey after parsing, so B1 and B2 receive no `--scenarios`
option. These behaviors are characterized, not corrected.

All paths in `run.py` are relative constants resolved against the caller's current
working directory, not against the location of `run.py`. Child stages inherit that
working directory because no `cwd` is supplied. `run_pipeline_script` also formats
its display path with `script_path.relative_to(Path.cwd())`; a script outside the
current directory raises `ValueError` before command execution.

The model-stage command boundaries are string commands passed with `shell=True`:

- A1, A2, B1, and B2 use
  `conda run -n <env> python -u "<absolute-script>"`;
- B1 and B2 append `--scenarios "<trimmed-original-order>"` only when the parsed
  list is nonempty;
- each A3 call uses
  `conda run -n <env> python -u "<absolute-A3-script>" --scenario "<name>"`; and
- active-scenario enumeration uses
  `conda run -n <env> python -u "<absolute-_scenarios.py>" list-active`.

The environment name is interpolated without additional quoting. The shared `run`
boundary copies the current environment, overwrites only `PYTHONHASHSEED=0`, and uses
`subprocess.check_call`; direct Conda availability/discovery and scenario-enumeration
checks do not use that environment override. The launcher does not change or restore
the working directory.

The A1/A2/A3 and B1 calls can overwrite model inputs or compiled parameter artifacts.
The B2 call is the boundary that can generate the final solver-consumed `.txt` input
and reach GLPK, CBC, CPLEX, or Gurobi through `main_executer`. No characterization
test crosses any of these process boundaries.

`main()` returns `None` after a successful sequential run. A delegated
`CalledProcessError` propagates immediately and prevents later stages; the script
entrypoint reports it and exits with the child's return code. Other exceptions stop
the sequence and become exit code one. Missing stage files fail before command
execution. Environment-discovery helpers retain their narrower existing fallback
behavior (for example, failed Conda environment listing is treated as not found, and
a failed DVC remote-list command is treated as no remote).

## Discovery and safety assertions

The characterization suite freezes the exact 20 preserved definitions, the 16
static cleanup-acceptance scenarios, and the 15 decision-relevant compiled-input
scenarios. Plain `BAU` remains in the first two scopes only; the four superseded
scenarios remain preservation-only.

Temporary fixtures cover B1 directory discovery, B1 config restoration, A3 rules
YAML precedence, and the shared country-config loader's script-relative path, sorted
country accessor, preserved configured order, and cache. Mocked launcher tests prove
stage/scenario propagation without starting a child process. AST checks cover the
guarded B2 and A3 orchestration paths.

For this phase, `tests/regression/` is the maintained no-solver-safe path. AST
inspection rejects process-launch APIs there except the regression harness's single
`subprocess.run` call in `_git`, whose command is a non-shell `git -C ...` metadata
read. Archived batch originals are never executed, and their retained stubs must
contain only notices followed by `exit /b 2`.
