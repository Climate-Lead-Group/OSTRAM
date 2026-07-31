# Regression and cleanup acceptance

Repository housekeeping uses a solver-free, byte-exact compiled-input gate.
It does not rerun CPLEX or infer numerical equivalence from normalized files.

## Accepted decision scenarios

The governed Stage 2 comparator manifest is
`STAGE_2_GOVERNED_COMPARATOR_MANIFEST.csv` in the authenticated Run 3 evidence
directory. It binds each current root-plus-declared-rule output by SHA-256,
byte size, and line count. It is the only compiled-input acceptance authority
used by the maintained comparator utility. Earlier source-bound capture reports
are historical external evidence, not tracked runtime or test inputs.

The governed manifest's exact ordered scenarios are:

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

`BAU` is a separate support scenario and receives its own smoke check.

The accepted A-result seed for the C dependency is bound to:

- 44,743,620 bytes
- SHA-256
  `762a7b926f91710846dc37e474747f5d670aed3d8746d7b74117ee978e645f5a`

The complete external comparator and seed locations are authenticated in the
Stage 0 evidence report. They are read-only validation inputs, not production
dependencies.

## Maintained checks

Validate the repository-side scenario and generated/ignored contract without
invoking a solver:

```powershell
$Py = 'C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe'
& $Py -B tests\regression\accepted_baseline.py
```

Stages 8 and 14 must validate freshly generated outputs against the governed
external manifest:

```powershell
$Evidence = '<authenticated Run 3 evidence directory>'
& $Py -B tests\regression\accepted_baseline.py `
  --governed-manifest "$Evidence\STAGE_2_GOVERNED_COMPARATOR_MANIFEST.csv" `
  --outputs-root '<disposable regeneration root>'
```

Run the current unit suite:

```powershell
& $Py -B -m unittest discover -s tests\regression -p 'test_*.py' -v
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
5. require all 15 expected target files to be newly generated;
6. compare exact relative paths, filenames, sizes, SHA-256 values, and line
   counts against the governed Stage 2 manifest;
7. require the three decision roots to remain byte-identical to their frozen
   Stage 0 comparators;
8. run the separate BAU smoke check;
9. verify authoritative workbooks are byte-unchanged; and
10. finish with only expected ignored runtime products.

Any missing, extra, inherited, reordered, resized, or rehashed target is a
failure. Normalization, tolerance, waivers, and solver reruns are not permitted
for this housekeeping gate.
