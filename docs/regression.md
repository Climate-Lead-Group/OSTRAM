# Regression and reproduction acceptance

Repository housekeeping uses a solver-free, byte-exact compiled-input gate.
It does not rerun the solver or infer numerical equivalence from normalized
files.

## Accepted decision scenarios

`config/profiles/full.yaml::scenario_registry` (validated by
`ostram/pipeline/scenarios/registry.py`) declares four roots (`BAU`,
`A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE`) and these 17 decision
scenarios, in canonical order:

1. `A_Calibrated_BAU`
2. `A_Calibrated_BAU_Clipped`
3. `B_Optimised_VRE`
4. `B_Opt_Clipped`
5. `B_Opt_DirBidir`
6. `B_Opt_DirContractual`
7. `B_Opt_IndiaCosts`
8. `B_Opt_IndiaCostsFuel`
9. `B_Opt_SolarCapex130`
10. `B_Opt_SolarCapexHi`
11. `B_Opt_SolarCapexSpike`
12. `B_Opt_TradeCap15`
13. `B_Opt_TxCap150`
14. `C_Target_VRE`
15. `C_Target_VRE_Clipped`
16. `A_Calibrated_BAU_Clipped_TxFreeze2026`
17. `B_Opt_Clipped_GrossImportCap15`

`BAU` is a separate support scenario and receives its own smoke check.
`C_Target_VRE` depends on a completed `A_Calibrated_BAU` result
([pipeline](pipeline.md#accepted-17-scenario-portfolio)).

Compiled-input acceptance is bound by a governed comparator manifest: a CSV with
columns `Scenario`, `AuthorityClass`, `SHA256`, `ByteSize`, `LineCount` and
`Provenance` that records, for every root and derived scenario, the final
preprocessed datafile
`Pre_processed_<scenario>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt`. The
manifest and the accepted A-result seed used for the C dependency are kept with
the release evidence outside Git. They are read-only validation inputs, not
production dependencies.

## Maintained checks

Validate the registry and the ignored-workspace contract without a solver:

```powershell
python -B tests\regression\accepted_baseline.py
```

Validate freshly generated compile-only outputs against the governed manifest:

```powershell
python -B tests\regression\accepted_baseline.py `
  --governed-manifest <governed comparator manifest CSV> `
  --outputs-root <disposable regeneration root>
```

Run the unit suites (recorded counts and known failures are in `AGENTS.md`):

```powershell
$env:PYTHONPATH = "tests"; python -m pytest tests -q
python -B -m unittest discover -s tests -p "test_*.py"
```

Focused A3 orchestration coverage uses disposable fixtures and in-process
doubles. It checks CLI behavior, scenario/dependency resolution, ordered stage
dispatch, path and environment boundaries, and failure propagation without
running A3 transformations, B1, B2, a matrix writer, or a solver.

## No-solver byte-identity contract

A cleanup acceptance run must:

1. start from a disposable clean checkout;
2. materialize the exact accepted scenario set from maintained inputs;
3. use the external A seed only for the declared C dependency;
4. compile without a solver or matrix/result route;
5. require all 17 expected target files to be newly generated;
6. compare exact relative paths, filenames, sizes, SHA-256 values, and line
   counts against the governed manifest;
7. require the three decision roots to remain byte-identical to their frozen
   comparators;
8. run the separate BAU smoke check;
9. verify authoritative workbooks are byte-unchanged; and
10. finish with only expected ignored runtime products.

Any missing, extra, inherited, reordered, resized, or rehashed target is a
failure. Normalization, tolerance, waivers, and solver reruns are not permitted
for this housekeeping gate.

## Compare a reproduced portfolio

Keep comparison evidence outside the checkout and production workspace. Use
the explicit selected portfolio and solver-chain records to pair all 17 cases.
Compare compiled data and LP hashes exactly, and compare effective parameter
values, defaults and sets independently. Require CPLEX optimal status 1,
primal/dual feasibility, and maximum primal/dual infeasibility at most `1e-6`.
For objective and reconciled cost, the reproduction uses
`abs(fresh - accepted) <= 1e-6 + 1e-10 * abs(accepted)` MUSD, a small allowance
for floating-point summation roundoff.

Compare every raw exported CSV. Byte-identical files establish exact physical
output equality; any different table remains unresolved until its differences
are examined. Equal objectives alone do not establish an acceptable alternative
optimum. A complete solution-XML comparison may normalize line endings and
exclude only the `problemName` path attribute. Keep per-case source paths,
hashes, solver certificates, objective/cost deltas and reconciliation JSON with
the comparison record. Wait for the runner's final cross-scenario
post-processing and successful exit before declaring the workflow complete.
