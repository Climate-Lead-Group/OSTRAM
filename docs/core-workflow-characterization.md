# Core workflow characterization

This note records the pre-refactor behavioral contract and current isolated
boundaries for OSTRAM's core workflow. The tests in
`tests/regression/test_core_workflow_characterization.py` enforce the contract with
AST inspection, temporary fixtures, and mocked process boundaries.
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

The repository-local canonical command package is also protected operational code:

- `ostram/__init__.py`
- `ostram/__main__.py`

From the repository root, `python -m ostram` provides three exact compatibility
dispatches:

| Canonical route | Historical route | Shared callable boundary |
|---|---|---|
| `python -m ostram run ...` | `python run.py ...` | `run.main()` plus the direct launcher guard's exit translation |
| `python -m ostram transform ...` | `python t1_confection/A3_process.py ...` | `A3_process.main()` |
| `python -m ostram compile-inputs ...` | `python t1_confection/B1_Run_Compiler.py ...` | `B1_Run_Compiler.main()` |

The dispatcher examines only the first token. It passes every remaining token once,
in the same order and without an extra `--`, through a temporary historical
`sys.argv`, then restores the original object even after `SystemExit`, an ordinary
exception, or `KeyboardInterrupt`. It does not change the caller working directory
or environment, redirect streams, add a process boundary, or import an unselected
route. The selected historical `main()` retains all of its existing effects,
including any subprocesses it normally starts. Importing either canonical module is
side-effect free.

Canonical `run` retains the historical direct guard: a delegated
`CalledProcessError` is reported on stderr and returns the child code, another
ordinary exception is reported on stderr and returns one, and `SystemExit` or
`KeyboardInterrupt` propagates. A3 propagates its `main()` result through the process
exit; B1 retains its natural successful exit zero. Downstream parser help and errors
therefore remain byte-for-byte those of the selected historical entrypoint.

There is no canonical command for the import-executing, per-scenario
`t1_confection/B1_Compiler.py`; `compile-inputs` correctly targets the public B1
runner. `prepare-model` and `solve` are deliberately deferred because
`t1_confection/B2_Executing_OG_Model.py` exposes only one configuration-driven public
workflow spanning compiled-input generation, optional matrix/solver execution,
cleanup, and postprocessing. Naming one portion would invent a public boundary.

`t1_confection/B1_Run_Compiler.py` remains the public B1 CLI, while the protected
core implementation now lives in `t1_confection/b1_runner.py`. The wrapper retains
the former callable helper surface and delegates explicitly to that module.

`t1_confection/A3_process.py` remains the public A3 CLI and retains its callable
stage/helper surface. The protected planning and effect sequence now lives in
`t1_confection/a3_orchestrator.py`; the entrypoint binds its explicit effect
seams, including the authorized root-gated `stage_ws4_pwr_min_pin`. The static
transformer and its frozen allowlist remain under `A3_process/rules_scripts/`.

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

`B1_Run_Compiler.py` delegates to the isolated B1 orchestration boundaries in
`b1_runner.py`: argument parsing, `B1Paths` construction, scenario discovery and
pure resolution, configuration preservation, compiler-command construction, an
injectable command executor, and top-level orchestration. The per-scenario effect
path remains `list_scenario_suffixes` once, then `update_main_scenario` ->
`run_compiler` for each selected scenario. Discovery sorts `A1_Outputs_*`
directories and excludes backup, snapshot, pre-experiment, and dated names. Filtering
preserves discovery order and collapses requested duplicates. The runner backs up
`Config_MOMF_T1_A.yaml` before iteration and restores it in a `finally` block. A
config-update failure skips that scenario; a nonzero compiler return is reported and
iteration continues. `run_compiler` launches
`B1_Compiler.py` with the current interpreter and the `t1_confection/` working
directory.

`A3_process.py` delegates read-only planning and the effect sequence to
`a3_orchestrator.py`, while retaining all existing stage helpers. Every selected
input directory is restored from `_post_a2_snapshot_BAU` before the work directory
is built. The frozen stage path is
`stage_0_5_rnwbio` -> `stage_1_scripts_1_to_5` -> `stage_1b` ->
`stage_2_and_2_5` -> `stage_3_fix_2` -> `stage_4_consolidate` ->
`stage_4_5_apply_inherited_restrictions` -> `stage_5_rules_scripts` ->
`stage_ws3_interconnector_costs` -> `stage_ws3_internal_transmission` ->
`stage_ws3_internal_tx_losses` -> `stage_ws4_pwr_min_pin` ->
`stage_6_sync_og_to_ts20` ->
`stage_6_persist_restrictions` -> `deliver_outputs`. A scenario-specific rules YAML
takes precedence over the default rules YAML. Helper failure is fail-fast through
`run_subproc`; work-directory and environment cleanup currently occur only after the
ordered path succeeds.

`stage_ws4_pwr_min_pin` is an exact-root-only seam for `A_Calibrated_BAU`,
`B_Optimised_VRE`, and `C_Target_VRE`. It validates and applies a static audited
2023--2026 PWR/MIN allowlist after the general WS3 transformations and before
Stage 6. It neither depends on solver results nor contributes a generic
`*_CHANGES.json`; descendants inherit the corrected root through the existing
sensitivity-expansion path.

`B2_Executing_OG_Model.py` retains the public entrypoint and computational helpers;
`t1_confection/b2_orchestrator.py` now owns explicit argument/scenario resolution,
compiled-input generation, matrix preparation, solver invocation, per-scenario
output handling, and final postprocessing boundaries. Both modules are import-safe.
Solver execution remains outside this test phase.

## A3 orchestration contract

The accepted direct A3 command keeps the seven predecessor options:
`--scenario`, `--soasia`, `--rules-script`, `--inherit-from`, `--input-dir`,
`--output-dir`, and `--keep-workdir`. Their defaults and standard `argparse` exit
behavior are unchanged. The default scenario remains `BAU`; an explicitly relative
SOASIA path remains caller-working-directory-relative, while relative input/output
overrides remain anchored to `t1_confection/`. Importing either the entrypoint or
the isolated orchestrator performs no planning, filesystem mutation, workbook
operation, or process launch.

`A3Paths` makes the entrypoint roots explicit. `resolve_plan` produces an immutable
`A3Plan` containing the scenario, ordered rule chain, ordered inherited restrictions,
SOASIA path, input/output destinations, fixed BAU snapshot, workdir base, and cleanup
choice. The existing `_resolve_scenario_config` remains on the public helper surface
and is injected into planning, so legacy mode, Control-sheet validation, CLI
override precedence, duplicate rule/inheritance values, and error text remain
unchanged. Direct A3 still processes one named Control scenario; multi-scenario
active discovery, topological order, duplicate collapse, and filtering remain owned
by `_scenarios.py` and `run.py`.

The canonical Control sheet has exactly four active definitions in this order:
`BAU`, `A_Calibrated_BAU`, `B_Optimised_VRE`, and `C_Target_VRE`. The only active
dependency edge is `BAU` -> `A_Calibrated_BAU`. Filtering follows the already
topologically ordered active list and does not add prerequisites. The other 16
protected scenario definitions are downstream/post-A3 configurations, not inactive
Control rows, and are not promoted into the A3 loop.

`A3Dependencies` is the explicit effect boundary. It receives the materializer,
workdir builder, 14 ordered stage helpers, delivery helper, snapshot copy/remove
operations, environment mapping, clocks, banner, and output emitter.
`execute_plan` first checks the BAU snapshot, then preserves the predecessor order
while inserting the authorized static pin after internal-transmission losses and
before either Stage-6 operation:
plan banner -> input deletion/restore -> workdir construction -> optional scenario
materialization and `OSTRAM_TEMPLATE_PATH` assignment -> four input copies -> the
frozen stage chain -> four-file delivery -> optional workdir deletion -> environment
removal -> completion banner. It never invokes B1, B2, the compiler, a matrix tool,
or a solver itself; computational A3 helpers retain their existing subprocess
boundary in `A3_process.py`.

Failure handling is deliberately unchanged. Missing prerequisites and failed helper
processes remain fail-fast. An unexpected exception, `SystemExit`, or
`KeyboardInterrupt` after work begins propagates immediately; workdir cleanup and
`OSTRAM_TEMPLATE_PATH` removal remain success-path-only. This known partial-state
hazard is characterized rather than repaired in this structural refactor.

## B2 orchestration contract

The accepted B2 command line has one optional value, `--scenarios`, with parsed
default `None`. Standard `argparse` help exits zero; an unknown option or missing
value exits two before configuration or pipeline work. Discovery sorts every entry
under the configured A2 output, removes only the exact `Default` entry, optionally
replaces discovery with `Main_Scenario`, and then filters without reordering.
Unknown requested names retain duplicates in the error and exit one. Valid requested
duplicates collapse because selection follows the discovered list.

The fixed top-level order is configuration and scenario planning, compiled-input
generation, optional execution, optional cleanup, first timing, optional
cross-scenario concatenation and annualization/dated-copy handling, then final
timing. For each scenario, the compiled-input path is
`process_scenario_folder` when enabled. After that optional step, the successful
conversion and patch path remains `run_otoole_conversion` -> `run_preprocessing_script` -> `run_days_in_day_type_patcher` -> `run_storage_delay_patcher` ->
`run_strip_storage_patcher` -> `run_open_pwrbck_patcher` ->
`run_reserve_margin_repair_patcher` ->
`run_reserve_margin_xlsx_patcher`, followed by the combined final text generation.
A failed conversion skips the remaining work for that scenario and continues with
the next. Root data-file export remains conditional on text generation and selection
of the configured main scenario.

Exactly two `main_executer` routes remain: a `multiprocessing.Process` target when
`parallel` is true and a direct linear call otherwise. The unchanged outer guard for
both is exactly `execute_model or create_matrix`; neither flag defaults to true.
Consequently, a configuration with both false never enters `main_executer`, never
constructs a solver adapter, and never reaches a matrix or solver process boundary.
Parallel child exit codes remain unchecked; direct exceptions still stop later
cleanup and postprocessing.

Within a represented per-scenario execution, non-GLPK matrix preparation remains
conditional on `create_matrix` and absence of a reusable solution. GLPK, CBC, CPLEX,
and Gurobi command preparation and invocation are isolated behind `SolverAdapter`
and `invoke_solver_command`; this boundary is reached only when `execute_model` is
true and a solution is not reused. When matrix creation and solving are both active,
solver-specific stale-file removal and environment checking still occur before the
matrix command, while the matrix subprocess still runs before the solver subprocess.
Commands retain inherited working directory/environment plus `shell=True` and
`check=True`. Per-scenario otoole conversion remains conditional on `execute_model`;
per-scenario concatenation remains conditional on `concat_otoole_csv` but is
dominated by entry into one of the two executor routes. Cross-scenario concatenation
is independently controlled by `concat_scenarios_csv`.

The source retains its existing Unicode status output. Captured Windows validation
must set `PYTHONIOENCODING=utf-8` before the first B2 invocation; the refactor does
not reconfigure console encoding or change user-visible status text.

## B1 runner contract

The accepted B1 command line has one optional value, `--scenarios`. Its parsed
default is `None`; standard `argparse` help exits zero, while an unknown option or a
missing value exits two. A missing config or compiler exits one. If discovery finds
no scenarios, B1 warns and exits zero before validating an explicit filter.

With no filter, every eligible discovered scenario runs in sorted discovery order.
A truthy filter is comma-split, trimmed, and stripped of empty values. Unknown names
are reported in requested order, including duplicates, and abort before backup.
Valid names are selected in discovery order, so `C,A,C` runs `A` then `C` once each.
A comma/whitespace-only truthy filter selects no scenarios but still enters the
backup/restore scope and finishes successfully; an empty string is treated as no
filter.

The compiler plan contains exactly the current `sys.executable` and the absolute
`B1_Compiler.py` path. Execution uses a token list, sets `cwd` to the B1 script
directory, and omits `env`, `shell`, `check`, timeout, and output-capture arguments.
The compiler therefore inherits the complete B1 environment and console streams.
There is no B2 or solver call in this path.

The configuration scope deliberately preserves all predecessor outcomes. Backup is
`shutil.copy2(config, config.yaml.bak)` and occurs before the scope's `try/finally`;
a backup error propagates without a restore attempt. Success, compiler launch
exceptions, and `KeyboardInterrupt` all attempt byte-exact restoration with
`shutil.move`. An ordinary restore error is warned and swallowed, the backup is
reported when present, and B1 still prints `All done` and exits zero even though the
live config may remain modified when the body otherwise finishes. A body exception
still propagates after the failed restore, without `All done`. This hazard is
characterized, not corrected by the isolation refactor. The last-resort
no-YAML-library regex behavior is likewise preserved despite its known
malformed-replacement edge case.

A compiler nonzero exit is reported but does not stop later scenarios or change the
final successful B1 exit. An ordinary config-update exception skips only that
scenario. A compiler-launch or other unexpected exception stops iteration,
restores, and propagates without printing `All done`.

Importing the public wrapper and helper is side-effect free. `B1_Compiler.py` remains
a top-level executable and is never imported by the runner. After B1 has printed
restoration and `All done`, no B1 code remains except return from `main()`; each
compiler child has already been synchronously awaited. The previously observed
external-wrapper timeout after those messages is therefore not treated as a B1 wait
or as success without the separate process/config/artifact checks.

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

Temporary fixtures cover B1 CLI/filter/order behavior, command construction,
compiler-error continuation, every safely testable restoration route, A3 rules YAML
precedence, and the shared country-config loader's script-relative path, sorted
country accessor, preserved configured order, and cache. Mocked launcher tests prove
stage/scenario propagation without starting a child process. They also cover exact
root-only pin dispatch, rejection of `BAU` and every registered descendant,
static-asset/CLI binding, fail-closed behavior, and the
losses -> pin -> Stage 6 -> delivery order. AST checks cover the isolated B1
boundaries and guarded B2 and A3 orchestration paths.

For this phase, `tests/regression/` is the maintained no-solver-safe path. AST
inspection rejects process-launch APIs there except the regression harness's single
`subprocess.run` call in `_git`, whose command is a non-shell `git -C ...` metadata
read. Archived batch originals are never executed, and their retained stubs must
contain only notices followed by `exit /b 2`.
