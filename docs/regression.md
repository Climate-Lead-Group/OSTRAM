# Offline Regression Evidence

Future cleanup and refactor PRs must follow the
[`cleanup and refactor governance rules`](refactor-rules.md).

The solver-free regression harness lives in `tests/regression/`. It never calls OSTRAM
stages, DVC, GLPK, CPLEX, or another optimizer. Child processes are limited to `git`
for read-only metadata/tracked-file discovery and the current Python interpreter for
exact `--help` or unknown-command CLI smoke checks. The CLI tests also install
fail-closed in-process sentinels at each first workflow-effect boundary.

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

`tests/regression/scenarios.yaml` is the machine-readable preservation and
static-scope policy. The committed static decision is recorded in
`tests/regression/baselines/5ce4e66480e1-static-nosolver/cleanup_acceptance_16.json`.
The current accepted compiled-input and source-bound solver identity is
[`accepted_compiled_solver_baseline_15.json`](../tests/regression/reports/accepted_compiled_solver_baseline_15.json).
The earlier `41a54e5` compiled-input check remains byte-preserved as
[`pre_correction_41a54e5_compiled_input_equivalence_15.json`](../tests/regression/reports/pre_correction_41a54e5_compiled_input_equivalence_15.json);
it is historical evidence, not the current baseline.

The accepted source baseline is solver-backed: its protected manifest has
SHA-256
`778b4706522bc2b29911e74d5b31d24355c84cbe4c0c7d11d1c9680b2ddc9916`
and records 15/15 scenarios as primal feasible with accepted optimal CPLEX
termination status. This offline harness does not rerun that solve and cannot
extend its solver claim to a future candidate. Housekeeping candidates must
instead reproduce the 15 raw compiled-input bytes exactly.

## Maintained no-solver byte-identity contract

This is the acceptance path for repository housekeeping. It must not be
replaced with `run.py`, a convenience `--skip-solver` invocation, normalized
comparison, or a newly solved result.

### Fixed identities and environment

- Reference commit:
  `8dd3361a1fc7f2c9ea4df51d5b2d0e50e0ce8554`, tree
  `65cbef7e977084b0a45f7bd8fca958d69ca916ce`.
- Accepted correction:
  `d295dcccca6c62e88f74484d7b8201e950881c3f`; it is an ancestor
  and the reference merge commit's second parent, with the same tree.
- Use new detached disposable worktrees for the reference and candidate
  commits. Never reuse a primary, benchmark, evidence, correction, audit, or
  earlier validation worktree.
- Use one pre-existing offline environment for both sides. Record Python and
  package versions, including the exact installed `ruamel.yaml`; do not install
  or update anything. Set `PYTHONHASHSEED=0`, `PYTHONUTF8=1`, and
  `PYTHONIOENCODING=utf-8`.
- Require a clean initial status and byte-back up both
  `t1_confection/Config_MOMF_T1_A.yaml` and
  `t1_confection/Config_MOMF_T1_AB.yaml`. Restore any temporary configuration
  bytes in `finally` and verify the restored hashes.

### Scenario and seed identities

Use this full literal canonical list for B1 and guarded B2:

```text
A_Calibrated_BAU,A_Calibrated_BAU_Clipped,B_Optimised_VRE,B_Opt_Clipped,B_Opt_DirBidir,B_Opt_DirContractual,B_Opt_IndiaCosts,B_Opt_IndiaCostsFuel,B_Opt_SolarCapex130,B_Opt_SolarCapexHi,B_Opt_SolarCapexSpike,B_Opt_TradeCap15,B_Opt_TxCap150,C_Target_VRE,C_Target_VRE_Clipped
```

The A-dependent C transformation requires the immutable accepted A combined
output:

- `Pre_processed_A_Calibrated_BAU_0_StorageDelayN5_OpenBCK_RMCarefulXLSX_output.csv`;
- 44,743,620 bytes;
- SHA-256
  `762a7b926f91710846dc37e474747f5d670aed3d8746d7b74117ee978e645f5a`;
- protected source below
  `G/s15-e7f99ab54fda-20260728T052124Z-9ccc09/source/t1_confection/Executables/A_Calibrated_BAU_0/`.

Resolve that full source path, require exactly one matching filename, verify its
size and hash, and copy it into each disposable worktree. Never edit or remove
the protected source.

The accepted 15-scenario outputs also depend on two frozen external,
scenario-specific post-A3 validation workbooks. They are validation inputs,
not replacements for the repository's production source-of-truth files, and
must not be added to Git:

- `A_Calibrated_BAU`:
  `C:\Users\luisfernando\Desktop\OSeMOSYS\G\g15-e7f99ab54fda-20260727T133603Z-3e7150\source\t1_confection\A1_Outputs\A1_Outputs_A_Calibrated_BAU\A-O_Parametrization.xlsx`;
  450,758 bytes; SHA-256
  `44b147d83bd13b287faa5ec722bd059d6ddd74b240a0f501827cf347925a54c9`.
- `B_Optimised_VRE`:
  `C:\Users\luisfernando\Desktop\OSeMOSYS\G\g15-e7f99ab54fda-20260727T133603Z-3e7150\source\t1_confection\A1_Outputs\A1_Outputs_B_Optimised_VRE\A-O_Parametrization.xlsx`;
  450,427 bytes; SHA-256
  `300cca3541a9caddbe2700092e3325342f5fbd69aa0f0fbaacb2861f36d7ff1e`.

Stage these two workbooks only inside each disposable validation worktree,
immediately before B1, and verify both identities after copying. Leave every
other A1 file and the other thirteen scenario directories unchanged. Do not
rerun A3 for A or B; A3 remains required only for `C_Target_VRE`, using the
immutable accepted A-result seed above.

A clean `8dd3361a1fc7f2c9ea4df51d5b2d0e50e0ce8554` checkout alone contains
older tracked A/B snapshots. Without the two external validation workbooks it
compiles the known alternate A hash
`63c5f0745269f695648d0ee4e1ae0bd583f2fba77c8921d697d2c21e5982267f`
and B hash
`5e1a313b87725c422360b55c2eb8eae09240c20328f6f75f71f88f3c3bb7af72`.
This validation-input correction changes no accepted output hash, scenario,
modelling decision, or production model input.

### Barriers and generation

Before generation, install and monitor process barriers that fail closed on:

- `glpsol`, `cplex`, `cbc`, `gurobi`, or any solver adapter;
- matrix/LP/MPS creation, including `--wlp`, `.lp`, and `.mps`;
- `otoole results`, result conversion, post-solve concatenation, or a result
  route;
- DVC, Conda/pip installation, batch files, and unrestricted `run.py`.

Log every permitted child command. Stop before generation if the barriers
cannot be installed and observed.

For each detached worktree:

1. Verify and copy the accepted A seed, then run A3 only for C:

   ```powershell
   & $Py -u t1_confection\A3_process.py --scenario C_Target_VRE
   ```

2. Immediately before B1, copy the two hash-bound external post-A3
   `A-O_Parametrization.xlsx` validation workbooks above into their
   corresponding A/B scenario directories. Verify the staged sizes and
   SHA-256 values. Do not run A3 for A or B.
3. Run B1 for the exact literal 15:

   ```powershell
   & $Py -u -m ostram compile-inputs --scenarios "A_Calibrated_BAU,A_Calibrated_BAU_Clipped,B_Optimised_VRE,B_Opt_Clipped,B_Opt_DirBidir,B_Opt_DirContractual,B_Opt_IndiaCosts,B_Opt_IndiaCostsFuel,B_Opt_SolarCapex130,B_Opt_SolarCapexHi,B_Opt_SolarCapexSpike,B_Opt_TradeCap15,B_Opt_TxCap150,C_Target_VRE,C_Target_VRE_Clipped"
   ```

4. Before B2, require every exact final target listed by the accepted record to
   be absent. Use an empty isolated output root where supported; otherwise
   remove only those exact 15 target `.txt` files inside the disposable
   worktree. Record each as absent. Never remove from a primary, benchmark, or
   protected evidence location.
5. Parse `Config_MOMF_T1_AB.yaml` with a YAML-aware driver and set real YAML
   booleans: `execute_model: false`, `create_matrix: false`,
   `concat_otoole_csv: false`, `concat_scenarios_csv: false`, and
   `parallel: false`; retain `A2_otoole_outputs: true` and
   `write_txt_model: true`. Redirect storage-delay/runtime outputs to a
   disposable validation folder. Parse the file again and independently assert
   those values and types.
6. Run guarded compile-only B2 with the same full literal list:

   ```powershell
   & $Py -u t1_confection\B2_Executing_OG_Model.py --scenarios "A_Calibrated_BAU,A_Calibrated_BAU_Clipped,B_Optimised_VRE,B_Opt_Clipped,B_Opt_DirBidir,B_Opt_DirContractual,B_Opt_IndiaCosts,B_Opt_IndiaCostsFuel,B_Opt_SolarCapex130,B_Opt_SolarCapexHi,B_Opt_SolarCapexSpike,B_Opt_TradeCap15,B_Opt_TxCap150,C_Target_VRE,C_Target_VRE_Clipped"
   ```

7. Restore the two original tracked workbook files and all temporary config
   bytes in `finally`. Require every final target to
   exist and record that it was newly created by this run.

### Exact comparison and stop conditions

Require reference, candidate, and the accepted record to have the same
canonical scenario order and the exact 15 relative paths, filenames, sizes,
and SHA-256 values. After hashes match, compare raw bytes directly between
reference and candidate. No normalization, tolerance, or waiver is permitted.

Stop immediately on any missing, extra, inherited, reordered, resized, or
rehashed target; solver/matrix/result request or process; `.lp`/`.mps` or
solver artifact; configuration-restoration failure; protected-source drift;
test failure; or dirty detached reference. Do not commit generated
A1/A2/otoole/`Executables` outputs.

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
3. **Source-bound solver baseline:** the current accepted source baseline is
   solver-backed and recorded by manifest SHA-256 `778b...9916`. This offline
   harness does not rerun it and cannot extend that claim to a future candidate.
   A housekeeping candidate must match all 15 raw compiled-input bytes; no
   solver rerun belongs in housekeeping.

See `tests/regression/README.md` for capture and comparison commands and the
static-evidence normalization policy. The accepted raw compiled-input
identities are recorded in
`tests/regression/reports/accepted_compiled_solver_baseline_15.json`.

## Future solver-backed changes

- The approved 16-scenario offline gate permits only isolated, non-core Stage 3 utility
  organization. All 20 scenario definitions remain under the preservation gate.
- The current `run.py --scenarios` propagation behavior is documented but unchanged. A
  future behavior change must be evaluated with a source-bound candidate and the
  solver-backed acceptance baseline.
- A future candidate does not inherit solver equivalence from the accepted
  baseline by ancestry. Full CPLEX behavioral equivalence for that candidate,
  DVC pipeline reconstruction, and any core A0-B2/A3/configuration changes
  remain outside the no-solver evidence level.
- A future solver-backed validation branch must bind its candidate to an exact source
  commit and toolchain, solve the 15 decision-relevant scenarios, retain status/log
  evidence, and compare agreed numerical outputs under documented tolerances. Byte-exact
  compiled inputs remain a prerequisite, not a substitute for that solve evidence.
- Plain `BAU` may be retained as a support diagnostic, but it is not a decision-scope
  solver acceptance scenario unless the policy is explicitly changed. The four
  superseded scenarios likewise remain preservation-only unless a later scope decision
  promotes them.
- Changes to scenario propagation, execution ordering, failure semantics,
  solver invocation, or core A0-B2/A3 transformations must be evaluated on
  future solver-backed refactor branches. The current housekeeping branch
  claims only raw compiled-input byte identity against the accepted record.
