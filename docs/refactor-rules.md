# Cleanup and refactor governance rules

These rules apply to every OSTRAM cleanup or refactor branch and pull request (PR).
They are the acceptance policy for the proposals in
[`refactor-plan.md`](refactor-plan.md); the evidence commands and current baseline are
documented in [`regression.md`](regression.md). A plan, passing offline check, or
historical result does not override these rules.

In this document, **must** is a requirement. **Protected** means that a file or tree
must never change incidentally. A branch may touch a protected code boundary only
when that boundary is its declared purpose and the required evidence tier is
completed. Model inputs, scenario artifacts, generated outputs, evidence files, and
templates must not be rewritten to make a cleanup or refactor pass.

## Branch and PR discipline

- Give each branch one stated purpose and one reviewable risk boundary. Split work
  when it can be accepted independently.
- Do not mix documentation changes with core behavioral changes unless the PR
  explains why the documentation is inseparable from that behavior change. Purely
  explanatory follow-up documentation belongs in a separate branch.
- Do not perform broad formatting, renaming, dependency, typing, import, or dead-code
  cleanup while refactoring behavior. Limit edits to the declared seam and its tests.
- Classify the changed files and select the highest applicable test tier before
  editing. A lower-risk file does not lower the tier triggered by another file in the
  same PR.
- Review `git status` before work, after every check that could write, and before each
  commit. Preserve pre-existing user changes and stop on unexpected drift.
- Do not commit while an applicable check is failing. A Tier 3 candidate that must be
  bound to a commit for solver validation may be committed only after Tier 2 passes;
  it must remain explicitly **solver-pending** until the required solver evidence
  passes.

Current accepted solver baseline: correction `d295dcc`, protected manifest
SHA-256
`778b4706522bc2b29911e74d5b31d24355c84cbe4c0c7d11d1c9680b2ddc9916`.
A new candidate does not inherit solver equivalence merely by ancestry.
Repository housekeeping follows the guarded protocol in
[`regression.md`](regression.md#maintained-no-solver-byte-identity-contract),
uses raw byte identity, and does not run a solver.

## Scenario policy

Discovery and preservation always cover all 20 definitions in
`tests/regression/scenarios.yaml`. The inventory must continue to match both the A1
snapshot directories and the A3 rule/config directories. Missing, renamed,
duplicated, or additional scenario definitions are failures, not inferred coverage.

The final compiled-input equivalence scope contains these 15 decision-relevant
scenarios:

- `A_Calibrated_BAU` (the decision-relevant baseline);
- `A_Calibrated_BAU_Clipped`;
- `B_Optimised_VRE`;
- `B_Opt_Clipped`;
- `B_Opt_DirBidir`;
- `B_Opt_DirContractual`;
- `B_Opt_IndiaCosts`;
- `B_Opt_IndiaCostsFuel`;
- `B_Opt_SolarCapex130`;
- `B_Opt_SolarCapexHi`;
- `B_Opt_SolarCapexSpike`;
- `B_Opt_TradeCap15`;
- `B_Opt_TxCap150`;
- `C_Target_VRE`; and
- `C_Target_VRE_Clipped`.

Plain `BAU` is retained, discovered, and protected. It remains part of the
16-scenario static cleanup-acceptance scope, but it is a legacy support scenario and
is not decision-relevant final compiled-input or solver-acceptance evidence.

The four superseded scenarios remain protected and visible in preservation reports:
`B_Opt_LinkFreeze`, `B_Opt_SolarHi10`, `B_Opt_TradeCap30`, and
`B_Opt_TradeCap50`. Their exclusion from cleanup acceptance does not authorize their
deletion, renaming, regeneration, or silent promotion into an acceptance scope.

## Protected paths and artifacts

At minimum, treat the following as protected. The maintained protected-tree hash in
`tests/regression/ostram_regression.py` is an executable minimum, not permission to
modify an unlisted workbook, template, model input, scenario artifact, or generated
output.

- **A1 snapshots:** `t1_confection/A1_Outputs/`.
- **A2 and otoole inputs:** `t1_confection/A2_Output_Params/` and
  `t1_confection/A2_Outputs_Params_otoole/`.
- **Original model CSV inputs:** `t1_confection/OG_csvs_inputs/`, together with
  `t1_confection/A2_Extra_Inputs/` and other tracked model-input trees.
- **A3 rules, configs, and templates:**
  `t1_confection/A3_process/rules_scripts/`, the operational Python files directly
  under `t1_confection/A3_process/`, both `Config_MOMF_T1_*.yaml` files, and A3
  workbooks/data templates including `SOASIA_OSeMOSYS_Template_v17.xlsx`,
  `SOASIA_OSeMOSYS_Template_v18.xlsx`, `Interconnectors.xlsx`, and `TECH_TYPES.csv`.
- **Core A0-B2 workflow:** `run.py` and the operational entrypoints and helpers from
  `t1_confection/A0_generate_tech_country_matrix.py` through
  `t1_confection/B2_Executing_OG_Model.py`, including A1, A2, A3, B1, compiler,
  config-loader, patch, solver-dispatch, and result-concatenation helpers.
- **DVC metadata:** `.dvc/`, `dvc.yaml`, `dvc.lock`, and every tracked `*.dvc` file.
- **Workbooks and templates:** tracked `*.xlsx`/`*.xlsm` files,
  `t1_confection/templates/`, `t1_confection/Miscellaneous/`, and template or
  reference trees such as `t1_confection/sensitivity_expansion/reference/`.
- **Compiled and generated artifacts:** solver-consumed
  `t1_confection/osemosys_fast_preprocessed*.txt` files, scenario artifacts, reports,
  output CSVs, logs, plots, workbook backups, and temporary pipeline outputs.

Protected code may be changed only on a narrowly scoped branch whose declared
purpose requires it. Such a change uses at least Tier 2; a core workflow or
solver-boundary change uses Tier 3. Protected data, templates, scenario definitions,
or generated evidence must not be changed as a side effect of cleanup/refactoring.
If an intentional model-data or evidence update is required, stop and move it to a
separate, explicitly authorized PR with its own provenance and acceptance policy.

## Test tiers

Use the highest tier triggered by any changed path or behavior. Passing a tier means
all required commands completed successfully and left the primary working tree free
of unexpected changes.

### Tier 0: documentation only

Run all of the following:

1. regression unit tests;
2. 20-scenario preservation discovery;
3. 16-scenario cleanup-acceptance discovery and the committed-evidence gate;
4. protected-tree hash verification;
5. Markdown link and repository-path checks for the changed documentation;
6. AST parsing of every touched Python file, if a Python file is touched despite the
   documentation-only classification; and
7. `git diff --check` plus a final changed-file and working-tree review.

Any Python change normally means the branch is not documentation-only and must be
reclassified unless it is limited to a non-executing documentation example or test
fixture and the PR explains why.

### Tier 1: archive or non-core utilities

Run all Tier 0 checks. Also run compatibility-wrapper, old-path invocation, import
boundary, fixture-output, or fail-closed checks wherever they apply. Do not regenerate
compiled inputs merely because an archive or output-only utility moved. If a pipeline
path is touched or the utility can affect solver-consumed input, reclassify the PR as
Tier 2 or Tier 3.

### Tier 2: pipeline-adjacent changes

Run all Tier 0 checks and applicable focused unit/characterization tests. In a
disposable checkout, regenerate the canonical solver-consumed `.txt` files for all 15
decision-relevant scenarios and compare them with the accepted baseline using the
documented byte-exact or normalized-exact rules. Do not execute a solver.

For repository housekeeping, the stricter maintained rule applies: all 15 raw
bytes must match the accepted record exactly. Normalization, tolerance, and
waivers are not permitted.

### Tier 3: core workflow or solver-boundary changes

Run all Tier 2 checks. In addition, document and review the affected
`B1`/`B2`/`run.py`/A3 call paths, including arguments, working directories, config
mutation and restoration, scenario propagation and order, failure behavior, output
paths, and the solver boundary. Solver-backed validation is required before the PR
may claim behavioral or numerical equivalence.

Tier 3 classification is not permission to run a solver. Without explicit solver
authorization, complete the safe checks, mark solver validation pending, and make no
behavioral-equivalence claim.

## Canonical compiled-input rule

Codex and human contributors must perform the 15-scenario compiled-input comparison
whenever a change touches any of the following:

- B1 runner or compiler behavior;
- A1-to-B1 transforms;
- config loading or config interpretation;
- scenario discovery, filtering, ordering, or propagation;
- A3 rules, configs, or generation behavior;
- `run.py` orchestration;
- path handling for generated model inputs; or
- any script that can affect the final solver-consumed `.txt` files.

The regeneration must run in a disposable checkout or disposable worktree bound to
the exact candidate commit. It must never run in the primary working tree. Before
running it, audit the proposed command and its transitive call path; stop if it might
invoke B2, CPLEX, GLPK, `run.py`, DVC reproduction, or a batch file.

Compare all 15 candidate file sets with a source-bound accepted baseline. Record the
comparison rule and provenance. A byte-exact pass requires identical bytes; a
normalized-exact pass requires the approved deterministic normalizer and must report
the raw mismatch rather than conceal it. Missing or extra files are mismatches.

Raw regenerated inputs and other generated artifacts stay out of Git. Only compact,
reviewable reports or hashes with the candidate commit, baseline identity, scenario
scope, normalizer/version, commands, and result may be committed.

## Solver rule and claim boundary

- Do not run CPLEX or GLPK unless the user explicitly requests and authorizes the
  solver, scenario scope, and command for that work.
- Do not execute B2, `run.py`, DVC reproduction, or batch files as a shortcut to a
  no-solver check. Audit unfamiliar wrappers before execution.
- No PR may claim numerical or full behavioral equivalence from unit tests, static
  hashes, protected-tree checks, or compiled-input equality alone.
- Compiled-input equality is necessary pre-solver evidence for Tier 2 and Tier 3; it
  is not a solver result. A solver-backed claim must identify the exact source commit,
  toolchain, scenarios, statuses/logs, outputs, tolerances, and comparison result.

## Stop conditions

Stop without committing, regenerating in place, installing dependencies, or
weakening the checks when any of these occurs:

- `git status` shows an unexpected dirty or untracked file;
- the protected-tree hash changes;
- any of the 20 scenarios is missing, duplicated, renamed, or unexpectedly added;
- a canonical generated `.txt` file is missing, extra, byte-different, or not
  accepted by the documented normalizer;
- the candidate, baseline, artifact, config, or command provenance is unclear;
- a required check needs a dependency that is not already available; or
- a proposed command, wrapper, or transitive call path might launch CPLEX, GLPK, B2,
  `run.py`, DVC reproduction, or a batch file.

Report the condition and narrow or move the work to a dedicated branch. Do not update
a baseline, normalize away a difference, delete a scenario, or edit a protected
artifact to turn a failure into a pass.

## PR checklist

Copy this checklist into every cleanup or refactor PR:

```text
- Branch purpose:
- Changed-file category (docs / archive / utility / pipeline-adjacent / core / protected):
- Highest test tier used (0 / 1 / 2 / 3) and why:
- Checks run, with commands and results:
- Scenario coverage (20 preservation / 16 static / 15 compiled, as applicable):
- Protected-tree result:
- Generated artifacts policy (where generated; confirmation raw artifacts are untracked):
- Baseline, candidate commit, and comparison rule, if applicable:
- Solver limitation (not run / explicitly authorized result / pending):
- Behavioral or numerical claims made, if any:
- Stop conditions encountered and resolution:
- Next safe branch or deliberately deferred follow-up:
```
