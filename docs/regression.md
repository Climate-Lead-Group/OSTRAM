# Offline Regression Evidence

Future cleanup and refactor PRs must follow the
[`cleanup and refactor governance rules`](refactor-rules.md).

The solver-free regression harness lives in `tests/regression/`. It never calls OSTRAM
stages, DVC, GLPK, CPLEX, or another optimizer. Its only child process is `git`, used for
read-only metadata and tracked-file discovery.

The pre-refactor entrypoint, discovery, call-order, mutation, and process-boundary
contract is recorded in
[`core-workflow-characterization.md`](core-workflow-characterization.md) and enforced
by the regression unit suite.

Focused A3 predecessor/candidate coverage lives in
`tests/regression/test_a3_orchestration.py`. It uses only disposable fixtures and
in-process doubles. The suite verifies CLI and import behavior, exact active order
and dependency semantics, path/environment/artifact boundaries, the full ordered
stage trace, success and failure/interruption behavior, and that no real A3
transformation, B1, B2, compiler, matrix, or solver process is invoked.

## Current static inventory

- Authoritative A1 snapshot scenarios: 20/20.
- Authoritative rule/config scenario directories: 20/20.
- Tracked A2 and otoole snapshots: 17/20.
- Historical reference output folders: 16/20.
- Safely regenerated candidate artifacts: none in the no-solver audit.

The three scenarios without tracked A2/otoole snapshots are `B_Opt_LinkFreeze`,
`B_Opt_SolarHi10`, and `B_Opt_TradeCap50`. Historical direct outputs are also absent for
`B_Opt_TradeCap30`. These are explicit coverage gaps, not inferred passes.

The four active scenarios in the SOASIA v18 `Control` sheet (`BAU`,
`A_Calibrated_BAU`, `B_Optimised_VRE`, and `C_Target_VRE`) define normal A3
execution, not regression acceptance. The repository uses three separate evidence
scopes:

- **20-scenario preservation:** every A1 and rule/config scenario remains mandatory.
- **16-scenario static cleanup acceptance:** the non-superseded scenarios with complete
  A1, config, A2, and otoole coverage.
- **15-scenario final compiled-input equivalence:** the decision-relevant scenarios have
  byte-exact final solver-consumed `.txt` files. `A_Calibrated_BAU` is the baseline.
  Plain `BAU` remains retained, discovered, and protected, but is excluded because it is
  a legacy non-decision support scenario and its available reference uses the older
  `NoStorage` chain rather than the current `StorageDelayN5` chain.

The 16-scenario static scope is plain `BAU` plus these 15 decision-relevant scenarios:
`A_Calibrated_BAU`, `A_Calibrated_BAU_Clipped`, `B_Optimised_VRE`,
`B_Opt_Clipped`, `B_Opt_DirBidir`, `B_Opt_DirContractual`, `B_Opt_IndiaCosts`,
`B_Opt_IndiaCostsFuel`, `B_Opt_SolarCapex130`, `B_Opt_SolarCapexHi`,
`B_Opt_SolarCapexSpike`, `B_Opt_TradeCap15`, `B_Opt_TxCap150`, `C_Target_VRE`,
and `C_Target_VRE_Clipped`.

The four superseded scenarios remain protected and visible in reports. They are excluded
only from the cleanup acceptance decision, not removed from the inventory.

`tests/regression/scenarios.yaml` is the machine-readable preservation and static-scope
policy. The committed static decision is recorded in
`tests/regression/baselines/5ce4e66480e1-static-nosolver/cleanup_acceptance_16.json`;
the narrower compiled-input result is recorded in
`tests/regression/reports/final_compiled_input_equivalence_15.json`.

## Run the checks

```powershell
$Py = 'python'  # or an existing OSTRAM-env Python executable

& $Py -m unittest discover -s tests\regression -p 'test_*.py' -v
& $Py tests\regression\ostram_regression.py discover --repo . --scope preservation
& $Py tests\regression\ostram_regression.py discover --repo . --scope cleanup-acceptance
& $Py tests\regression\ostram_regression.py gate `
  --scope cleanup-acceptance `
  --evidence tests\regression\baselines\5ce4e66480e1-static-nosolver
& $Py tests\regression\ostram_regression.py verify-protected `
  --repo . `
  --manifest tests\regression\baselines\5ce4e66480e1-static-nosolver\manifest.json
```

## What the evidence proves

1. **Exact pre-solver/static equivalence:** format-aware hashes compare artifacts that
   exist in both the tracked checkout and reference material. Generated backups and
   documented volatile timestamp fields are excluded from exact comparison.
2. **Available historical-output evidence:** hashes record the direct output CSVs that
   exist in the read-only reference checkout, with missing scenarios reported.
3. **Full solver-backed behavioral equivalence:** pending. No offline hash or static check
   establishes complete CPLEX numerical equivalence. The recorded 15/15 byte-exact final
   compiled inputs are strong pre-solver evidence, not literal CPLEX behavioral proof.

See `tests/regression/README.md` for capture and comparison commands and the normalization
policy. Compact final hashes are recorded in
`tests/regression/reports/final_compiled_input_equivalence_15.json`.

## Deferred solver-backed work

- The approved 16-scenario offline gate permits only isolated, non-core Stage 3 utility
  organization. All 20 scenario definitions remain under the preservation gate.
- The current `run.py --scenarios` propagation behavior is documented but unchanged. A
  future behavior change must be evaluated with a source-bound candidate and the
  solver-backed acceptance baseline.
- Full CPLEX behavioral equivalence, DVC pipeline reconstruction, and any core
  A0-B2/A3/configuration changes remain outside the no-solver evidence level.
- A future solver-backed validation branch must bind its candidate to an exact source
  commit and toolchain, solve the 15 decision-relevant scenarios, retain status/log
  evidence, and compare agreed numerical outputs under documented tolerances. Byte-exact
  compiled inputs remain a prerequisite, not a substitute for that solve evidence.
- Plain `BAU` may be retained as a support diagnostic, but it is not a decision-scope
  solver acceptance scenario unless the policy is explicitly changed. The four
  superseded scenarios likewise remain preservation-only unless a later scope decision
  promotes them.
- Changes to scenario propagation, execution ordering, failure semantics, solver
  invocation, or core A0-B2/A3 transformations must be evaluated on future
  solver-backed refactor branches. No such equivalence is claimed by the current
  cleanup branch.
