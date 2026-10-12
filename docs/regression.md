# Regression and cleanup acceptance

Repository housekeeping uses a solver-free, byte-exact compiled-input gate.
It does not rerun CPLEX or infer numerical equivalence from normalized files.

## Accepted decision scenarios

The accepted compiled-input manifest, `STAGE_2_GOVERNED_COMPARATOR_MANIFEST.csv`,
is kept outside the repository with the Run 3 evidence. It binds each
current root-plus-declared-rule output by SHA-256, byte size, and line
count. It is the only compiled-input acceptance source of record used by the
comparator utility `tests/regression/accepted_baseline.py`. Earlier
source-bound capture reports are historical external evidence.

The 17 decision scenarios, in the order of `EXPECTED_DECISION_SCENARIOS` in
`tests/regression/accepted_baseline.py`, are:

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

The A-result seed for the C dependency is the solved `A_Calibrated_BAU`
result folder passed with `--a-result-seed`. The comparator and seed
locations are recorded with the external evidence. They are read-only
validation inputs, not production dependencies.

## Maintained checks

Validate the repository-side scenario and generated/ignored contract without
invoking a solver, from the repository root with the environment's Python:

```powershell
python -B tests\regression\accepted_baseline.py
```

Stages 8 and 14 must validate freshly generated outputs against the accepted
external manifest:

```powershell
$Evidence = '<authenticated Run 3 evidence directory>'
python -B tests\regression\accepted_baseline.py `
  --governed-manifest "$Evidence\STAGE_2_GOVERNED_COMPARATOR_MANIFEST.csv" `
  --outputs-root '<disposable regeneration root>'
```

Run the current unit suite:

```powershell
python -B -m unittest discover -s tests -p 'test_*.py' -v
```

Focused A3 orchestration coverage uses disposable fixtures and in-process
doubles. It checks CLI behavior, scenario/dependency resolution, ordered stage
dispatch, path and environment boundaries, and failure propagation without
running A3 transformations, B1, B2, a matrix writer, or a solver.

## Maintained no-solver byte-identity contract

The final cleanup acceptance run must:

1. start from a disposable clean checkout;
2. materialize the exact accepted scenario set from maintained inputs;
3. use the external A seed only for the declared C dependency;
4. compile without a solver or matrix/result route;
5. require all 17 expected target files to be newly generated;
6. compare exact relative paths, filenames, sizes, SHA-256 values, and line
   counts against the accepted Stage 2 manifest;
7. require the three decision roots to remain byte-identical to their frozen
   Stage 0 comparators;
8. run the separate BAU smoke check;
9. verify authoritative workbooks are byte-unchanged; and
10. finish with only expected ignored runtime products.

Any missing, extra, inherited, reordered, resized, or rehashed target is a
failure. Normalization, tolerance, waivers, and solver reruns are not permitted
for this housekeeping gate.
