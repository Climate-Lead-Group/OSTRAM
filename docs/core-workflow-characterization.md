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

The following are analysis utilities, not core workflow stages. The first four
canonical paths retain compatibility wrappers at their old `t1_confection/` paths:

- `tools/analysis/check_combined.py` and `t1_confection/check_combined.py`
- `tools/analysis/ostram_scenario_analysis.py` and
  `t1_confection/ostram_scenario_analysis.py`
- `tools/analysis/ostram_trn_plotter.py` and `t1_confection/ostram_trn_plotter.py`
- `tools/analysis/slice_by_country.py` and `t1_confection/slice_by_country.py`
- `t1_confection/analyse_sensitivity.py`
- `t1_confection/concat_all_scenarios_2.py`
- `t1_confection/reproduce_A1_A6.py`
- `t1_confection/sensitivity_expansion/analyse_ws4_vs_phaseB.py`
- `t1_confection/Z_AUX_generate_interactive_dashboards_aggregated.py`
- `t1_confection/Z_AUX_generate_RES_diagram.py`
- `t1_confection/Z_AUX_generate_transmission_maps.py`
- `t1_confection/Z_AUX_interconnections_dashboard.py`

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
