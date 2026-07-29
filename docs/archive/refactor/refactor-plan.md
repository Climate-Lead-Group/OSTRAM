# Next-phase script and workflow refactor plan

> **Historical plan:** Execution status is superseded by the accepted final
> 15-scenario baseline. The tables remain design rationale, not current
> authorization. See the maintained [regression policy](../../regression.md).
>
All branches and PRs proposed here must follow the
[`cleanup and refactor governance rules`](../../refactor-rules.md).

## Status and evidence boundary

This is a planning document. It does not authorize moving a script or changing
model behaviour. The inventory is based on the tracked tree after the repository
structure cleanup and on static call-site, import, configuration, and documentation
searches. A script's primary category below is a maintenance classification; it is
not a claim that the script has only one use.

The current offline evidence has three distinct scopes:

- 20 preserved scenario directories and A1/config files;
- 16 scenarios covered by the cleanup-acceptance static baseline; and
- 15 decision-relevant scenarios with byte-exact and normalized-exact final
  compiled-input equivalence: `A_Calibrated_BAU`, `A_Calibrated_BAU_Clipped`,
  `B_Optimised_VRE`, `B_Opt_Clipped`, `B_Opt_DirBidir`,
  `B_Opt_DirContractual`, `B_Opt_IndiaCosts`, `B_Opt_IndiaCostsFuel`,
  `B_Opt_SolarCapex130`, `B_Opt_SolarCapexHi`, `B_Opt_SolarCapexSpike`,
  `B_Opt_TradeCap15`, `B_Opt_TxCap150`, `C_Target_VRE`, and
  `C_Target_VRE_Clipped`.

Plain `BAU` remains retained support evidence, but it is not part of the
decision-relevant 15-scenario comparison. None of these offline results is a
solver-backed equivalence claim. Core workflow changes therefore require a future,
separately authorized solver-backed branch before they can be accepted. See
[`regression.md`](../../regression.md) for the current scenario and evidence policy.

## Classification rules

### Core pipeline

Treat these files as operational and protected from opportunistic relocation:

- the root entrypoint `run.py`;
- `t1_confection/A0_generate_tech_country_matrix.py`,
  `A1_Pre_processing_OG_csvs.py`, `A2_AddTx.py`, `A3_process.py`,
  `B1_Run_Compiler.py`, `B1_Compiler.py`, and `B2_Executing_OG_Model.py`;
- the executable stages, actively invoked helpers, and rules under
  `t1_confection/A3_process/` (the test/validation helpers are classified below);
- `t1_confection/D1_generate_editor_template.py`,
  `D2_update_secondary_techs.py`, and `Z_AUX_D1b_set_trn_limits_from_flows.py`,
  which are optional model-writing workflow tools;
- the B2 patch helpers under `t1_confection/`, including storage, reserve-margin,
  capacity, `DaysInDayType`, and preprocessing scripts;
- `t1_confection/Z_AUX_config_loader.py`, which is imported by multiple workflow
  stages, and `Z_AUX_capital_annualization_script.py`, which is imported by B2 and
  named in `dvc.yaml`;
- `concatenate_files/concatenate_ostram.py`, which supports B2 result assembly and
  is referenced by configuration, DVC, and documentation; and
- `t1_confection/sensitivity_expansion/apply_patches.py` and
  `gen_sensitivity_patches.py`, which can write scenario workbooks or protected
  sensitivity definitions.

Names alone are not evidence of obsolescence. In particular, A3 still copies or
invokes helpers whose filenames include `OLD`, `NEW`, `fix`, or `patch`; they stay
in this category until call-site characterization proves otherwise.

### Analysis utilities

These read model inputs or existing outputs and produce analysis, figures, slices,
or reports. Canonical moved implementations are under `tools/analysis/`, with
compatibility wrappers at the former `t1_confection/` paths. The WS-4 report remains
at `t1_confection/sensitivity_expansion/analyse_ws4_vs_phaseB.py` because the
protected-tree gate covers that path even though its behavior is analysis-only.

### Validation and regression

These provide checks or evidence and must remain separate from production
orchestration:

- `tests/regression/` and its discovery, gate, protected-tree, and strict-compare
  commands;
- `tests/validation/test_strip_storage_cli.py`;
- `tests/validation/test_scenarios_lite.py` and the retained production helper
  `t1_confection/A3_process/_xlsx_validation_core.py`;
- `t1_confection/Z_validate_country_data.py`;
- `t1_confection/sensitivity_expansion/desk_check.py` and
  `validate_sensitivity_configs.py`;
- the read-only WS-3/WS-4 audit and consistency scripts under
  `ws3_transmission_audit/`; and
- `cleanroom_tests/cleanroom_check.py`, which is retained as historical cleanup
  evidence rather than treated as the maintained regression entrypoint.

### Generated or ignored artifacts

These are outputs, not source scripts. They must not be swept into a source move or
used as primary provenance merely because they are near a script:

- combined input/output CSVs, solver LP/SOL/log/output files, storage-delay files,
  plots, slices, temporary A3 run directories, workbook backups, and patch-change
  manifests covered by `.gitignore`; and
- intentionally tracked workbooks, compiled CSVs, validation reports, sensitivity
  tables, and change manifests that serve as derived evidence.

Tracked derived evidence must remain byte-stable unless a dedicated evidence update
explicitly explains and verifies the change. For VRE ceilings, the operational JSON,
the documentary CSV mirror, and the accompanying provenance note have the distinct
roles documented in `OSTRAM_METHODOLOGY.md`; generated tables do not replace them.

### Historical or archive

This category includes the fail-closed runner stubs at their former
`t1_confection/` paths and their originals under `docs/archive/legacy-runners/`,
historical WS-3/WS-4 audit material, superseded methodology and cleanup reports, and
the original cleanroom consolidation harness.
Historical material may still be valuable evidence. Moving it requires link and
manifest checks even when it is no longer executed.

### Obsolete or unsafe

These are not safe general-purpose entrypoints in their present form:

- `t1_confection/concat_all_scenarios.py`: its successor documents that the older
  merge strategy duplicated rows;
- root `run_*.bat` files: they contain machine-specific interpreter paths and invoke
  B2/solver execution directly;
- `t1_confection/Z_AUX_united_regions.py`: it is hard-coded to a Brazil workflow and
  has no active inbound code reference;
- `t1_confection/Z_AUX_fix_excel_profiles.py`: it writes workbooks and carries stale,
  scenario-specific assumptions; and
- `ws3_transmission_audit/set_final_v18_interconnector_values.py`:
  it contains an absolute local path and writes a template in place.

“Unsafe” means retain for provenance or replace with a fail-closed stub; it does not
mean delete. No file in this category should be executed as part of cleanup checks.

## Standard no-solver gate

Every future move below must run the applicable move-specific checks plus this
common offline gate:

1. `python -m unittest discover -s tests/regression -p "test_*.py" -v`;
2. the preservation and cleanup-acceptance discovery commands documented in
   `docs/regression.md`;
3. the committed-evidence gate;
4. the protected-tree verification;
5. strict baseline self-comparison;
6. `git diff --check`, tracked-file hygiene, and relative documentation-link checks;
7. Python syntax/AST checks for moved code and import-boundary checks that do not
   execute model stages; and
8. a clean worktree after the commit.

Tests that need temporary files must use isolated fixtures, never tracked workbooks,
templates, scenario directories, configs, or evidence outputs.

## Analysis utilities round 2 status

The `refactor/analysis-utilities-round2` phase implements seven analysis-only moves
from the inventory below:

- `tools/analysis/concat_all_scenarios.py`;
- `tools/analysis/analyse_sensitivity.py`;
- `tools/analysis/reproduce_A1_A6.py`; and
- the four scripts under `tools/analysis/visualization/`.

The old user-facing paths remain forwarding wrappers for a deprecation cycle. Inputs
and generated-output locations remain anchored to `t1_confection/` where they were
script-relative before the move; the aggregated dashboard retains its intentional
current-working-directory behavior. No validation, maintenance, archive, or core move
is included in this phase. The WS-4 analysis utility is deliberately deferred: moving
it would change the protected-tree hash and therefore cannot be accepted as Tier 1.

## Utility, validation, and archive move inventory

The analysis-only rows other than the protected WS-4 utility are implemented by the
round-2 phase. The WS-4, validation, maintenance, and archive rows remain proposals
for later branches. “Wrapper” means that the old user-facing path remains as a small
forwarding entrypoint for at least one deprecation cycle.

| Old path | Canonical or proposed path | Reason | Affected imports, config, or docs | Risk | Required no-solver checks beyond the standard gate | Solver-backed verification |
|---|---|---|---|---|---|---|
| `t1_confection/concat_all_scenarios.py` | `docs/archive/legacy-tools/concat_all_scenarios_merge.py` | Preserve the row-duplicating predecessor without presenting it as maintained. | References in `concat_all_scenarios_2.py` and historical sensitivity documentation. No active inbound import was found. | Low–medium | Assert no production import; archive-link check; fixture demonstrating why the old merge is not canonical. | No. |
| `t1_confection/concat_all_scenarios_2.py` | `tools/analysis/concat_all_scenarios.py`, plus old-path wrapper | Give the maintained output concatenator a descriptive canonical home. | `analyse_sensitivity.py`, sensitivity methodology/README material, and user invocations. | Medium | Fixture covering scenario discovery, column order, row counts, and byte/normalized output equality through both entrypoints. | No; it consumes existing outputs. |
| `t1_confection/analyse_sensitivity.py` | `tools/analysis/analyse_sensitivity.py`, plus old-path wrapper | Separate report generation from workflow orchestration. | Sensitivity methodology, historical README, concatenator invocation, baseline/ceiling path resolution. | Medium | Run against immutable fixture CSVs; compare report text and comparison-table values; test invocation from repository and non-repository working directories. | No; it consumes existing outputs. |
| `t1_confection/sensitivity_expansion/analyse_ws4_vs_phaseB.py` | `tools/analysis/analyse_ws4_vs_phaseB.py`, plus old-path wrapper | Consolidate comparison-only analysis. | `RUN_ORDER.md`, `OSTRAM_METHODOLOGY.md`, and historical Phase-B paths. | Medium | Fixture metric equality, missing-input failure test, and working-directory independence. | No; it consumes existing outputs. |
| `t1_confection/reproduce_A1_A6.py` | `tools/analysis/reproduce_A1_A6.py`, plus old-path wrapper | It reads compiled/output data and produces plots/audit CSVs; it is not a compile stage. | User invocations and any figure/report documentation discovered immediately before the move. | Low–medium | Fixture assertions for selected rows, audit CSV values, plot filenames, and no writes outside a temporary output directory. | No. |
| `t1_confection/Z_AUX_generate_interactive_dashboards_aggregated.py`, `Z_AUX_generate_RES_diagram.py`, `Z_AUX_generate_transmission_maps.py`, `Z_AUX_interconnections_dashboard.py` | `tools/analysis/visualization/`, with compatibility wrappers | Group standalone visualization and resource-analysis entrypoints. | `docs/auxiliary-tools.md`, optional plotting dependencies, user commands, and hard-coded relative paths. | Medium | CLI/help smoke tests, path-resolution tests, immutable small fixtures, output-name/value assertions, and wrapper parity. | No. |
| `t1_confection/test_strip_storage.py` | `tests/validation/test_strip_storage_cli.py` | Put a subprocess validation harness with tests rather than production scripts. | Default path to `strip_storage.py` and any developer test commands. | Medium | Temporary text fixtures; identical exit status and transformed content; explicit assertion that tracked files are untouched. | No. |
| `t1_confection/Z_validate_country_data.py` and `Z_generate_country_template.py` | `tools/maintenance/country/`, with compatibility wrappers | Separate country-maintenance commands from model execution. | Country-management docs, README, generated-template comments, `Config_country_codes.yaml`, and template path assumptions. | Medium | Validator fixture; generator confined to a temporary tree; stable generated text; wrapper parity; no tracked config/template writes. | No. |
| `t1_confection/Z_AUX_sort_csv.py` | `tools/maintenance/sort_csv.py`, plus old-path wrapper | Make an in-place maintenance operation explicit and testable. | User invocations; no active inbound import was found. | Medium | Temporary CSV fixture covering headers, ordering, encodings, and failure behaviour; prove no default write to tracked data. | No. |
| `t1_confection/Z_AUX_united_regions.py` | `docs/archive/legacy-tools/Z_AUX_united_regions.py` | Preserve the hard-coded Brazil helper as historical code. | `docs/auxiliary-tools.md` and any newly discovered historical references. | Low | Repeat inbound-reference search; syntax check only; archive-link check. | No. |
| `t1_confection/Z_AUX_fix_excel_profiles.py` | `docs/archive/legacy-tools/Z_AUX_fix_excel_profiles.py`, with a fail-closed old-path notice if compatibility requires it | Prevent accidental mutation while preserving provenance. | `docs/auxiliary-tools.md`, workbook/path literals, and user invocations. | Medium | Static check that the old path cannot write; syntax check archived copy; never open or rewrite tracked workbooks during validation. | No for archival; any revived transformation would require separate model verification. |
| root `run_baselines.bat`, `run_directions.bat`, `run_sensitivities.bat` | originals under `docs/archive/legacy-runners/root/`, with fail-closed root stubs | Remove machine-specific, solver-invoking commands from the apparent supported surface without losing history. | `RUN_ORDER.md`, regression utility-layout assertions, and existing nested runner-stub/archive conventions. | Medium | Static content assertions only; verify stubs fail before invoking Python; do not execute the originals or any batch file. | No for archival; solver execution remains a separately authorized workflow. |
| `cleanroom_tests/cleanroom_check.py` and `hero_refs.yaml` | `docs/archive/cleanroom/harness/` | Record the one-time consolidation harness as historical evidence instead of a maintained gate. | Cleanroom reports, regression fixtures/manifests, branch-name text, and protected-tree scope. | Medium–high | First prove no maintained gate imports it; preserve hashes and links; update protection policy only in an explicitly reviewed evidence commit. | No, but do not move while the protected-tree gate treats the current path as fixed. |
| `ws3_transmission_audit/*.py` | `docs/archive/ws3-ws4/scripts/` after separating read-only audits from writers | Keep dated audit evidence together and quarantine the in-place writer. | WS-3 reports, `apply_patches.py` provenance comments, configuration comments, absolute paths, and report output paths. | Medium–high | Classify every script as read-only or writing; hash archived files; link/provenance checks; never execute the writer; use fixtures for any retained audit test. | No for archival; yes if any script is revived to alter model inputs. |

Before each move, repeat the inbound-reference search. The table records the
references visible at this snapshot, not permission to ignore a reference added by a
later commit.

## Proposed core implementation extractions

Core entrypoint paths must remain in place. A future refactor may move implementation
behind them only after characterization tests exist and a solver baseline can be
reproduced in an authorized environment.

The A3 orchestration phase now implements one such boundary:
`t1_confection/A3_process.py` retains the direct CLI and callable transformation
helpers, while `t1_confection/a3_orchestrator.py` owns explicit paths, a read-only
run plan, injected effects, and the preserved stage sequence. It does not move or
rewrite computational transformations. Acceptance still requires the Tier 3
compiled-input and solver evidence described below.

| Existing entrypoint | Proposed implementation path | Reason and preserved boundary | Affected imports, config, or docs | Risk | Required no-solver checks beyond the standard gate | Solver-backed verification |
|---|---|---|---|---|---|---|
| `run.py` | `t1_confection/workflow/launcher.py`; retain root wrapper | Isolate path resolution, dependency checks, command construction, and stage selection without changing the public command. | All run documentation, DVC/Conda detection, A1–B2 command lines, environment handling, exit propagation. | High | Mocked command-construction and dispatch tests for every flag/solver; current exit/failure characterization; invocation from another working directory; no subprocess model stage. | Yes, for each supported stage sequence selected for release. |
| `t1_confection/B1_Run_Compiler.py` | `t1_confection/workflow/b1_runner.py`; retain old wrapper | Isolate scenario discovery, temporary config editing, compiler dispatch, restoration, and failure aggregation. | `run.py`, `Config_MOMF_T1_A.yaml`, `B1_Compiler.py`, scenario ordering, logs, and current continue-on-error semantics. | High | Fixture-only discovery/filter/order tests; config restoration after success and exception; characterize current exit status before deciding any change. | Yes; exact 15-scenario compiled inputs first, then solver-backed comparison. |
| `t1_confection/B2_Executing_OG_Model.py` | modules under `t1_confection/workflow/b2/`; retain old wrapper | Separate configuration, scenario discovery, patch ordering, solver dispatch, result concatenation, and annualization from the entrypoint. | `run.py`, both configs, all B2 patch helpers, `concatenate_ostram.py`, solver adapters, output layout, and annualization. | Critical | Mock every external command; assert scenario and patch order, cwd restoration, arguments, failure propagation, and output-path construction; no solver invocation. | Yes; full solver-backed equivalence for the decision-relevant 15, with plain BAU diagnostic only. |
| `t1_confection/A3_process.py` | `t1_confection/a3_orchestrator.py`; retain public entrypoint and helper surface | Isolate snapshot/restore, scenario planning, rule dispatch, validation, and persistence while preserving file order and workbook semantics. | `run.py`, `_scenarios.py`, every active A3 helper/rule, restrictions, workbook names, and scenario directory layout. | Critical | Mocked predecessor/candidate trace; temporary filesystem fixtures; CLI, path, environment, stage-order, failure, and interruption tests; protected-tree check; exact compiled-input comparison without solving. | Yes before accepting changed operational output. |
| `t1_confection/A1_Pre_processing_OG_csvs.py` | incrementally extracted pure transforms under `t1_confection/workflow/transforms/a1/`; retain entrypoint | Reduce a monolith without changing CSV reads, defaults, ordering, or workbook schema. | A0/A2 expectations, config loader, input CSVs, output workbooks, DVC stages, and downstream compiler assumptions. | Critical | One pure transform at a time; immutable fixtures; exact table/schema/order/dtype comparisons; protected-tree verification. | Yes after exact 15-scenario compiled-input equivalence. |
| `t1_confection/B1_Compiler.py` | incrementally extracted pure transforms under `t1_confection/workflow/transforms/b1/`; retain entrypoint | Make compilation rules testable while preserving scenario-specific overrides and emitted text. | B1 runner, both configs, A1/A2/A3 outputs, OSeMOSYS text schema/order, and scenario folders. | Critical | Golden fixtures per extracted transform; exact and normalized-exact compiler output; all 15 decision-relevant scenarios before integration. | Yes; compiled equivalence is necessary but not sufficient for this core change. |

Do not begin with shared “cleanup” helpers across A1, A3, B1, and B2. First preserve
each entrypoint and characterize its current path, ordering, mutation, and failure
semantics. A convenience abstraction is not evidence of equivalent behaviour.

## Recommended PR sequence

1. `test/core-workflow-characterization`: tests and documentation only. Capture
   command construction, discovery/order, config restoration, mutation boundaries,
   and current failure semantics without running model stages.
2. `cleanup/archive-unsafe-entrypoints`: archive one obsolete/unsafe family at a
   time, retain fail-closed stubs where users may call old paths, and run the offline
   gate after each commit. Split root runners, legacy utilities, cleanroom evidence,
   and WS-3 audit scripts if their protection constraints differ.
3. `refactor/analysis-utilities-round2`: move output-only analysis utilities in small
   commits with wrappers and fixture-based output equality. No core imports or model
   inputs may change.
4. `validation/solver-baseline-15`: in a separately authorized environment, record
   toolchain versions and solver-backed reference results for the 15
   decision-relevant scenarios. Retain plain BAU as a diagnostic, non-decision run.
5. `refactor/run-orchestration-seams`: extract the root launcher implementation while
   preserving its command and entrypoint contract; compare against the established
   solver baseline before merge.
6. `refactor/b1-runner-isolation`: isolate discovery/config/restore behaviour, prove
   exact compiled-input equivalence, then run solver-backed comparison.
7. `refactor/b2-orchestration`: split B2 in dependency-order, one seam per PR, with
   full solver-backed verification.
8. `refactor/a3-orchestration`: the structural isolation is implemented on its
   dedicated branch; retain its compiled-input evidence and solver-backed comparison
   as acceptance requirements.
9. `refactor/a1-b1-transforms`: extract one pure transformation at a time only after
   the orchestration boundaries and solver baseline are stable.

## Stop conditions

Stop a proposed move and open a narrower investigation if it changes a protected
hash, alters a tracked workbook or generated-evidence file, requires editing rules or
configs, changes scenario discovery/order, breaks old-path invocation, makes a
previously fail-closed entrypoint executable, or cannot reproduce its fixture output.
For core work, also stop if the authorized solver/toolchain baseline is unavailable;
offline compiled-input equality must not be reported as solver-backed equivalence.
