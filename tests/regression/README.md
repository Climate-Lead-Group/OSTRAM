# Current regression checks

This directory contains solver-free tests for the maintained runtime and the
portable accepted compiled-input contract.

The authoritative scenario inventory is `scenarios.yaml`. The exact accepted
decision-scenario identities are recorded in
`reports/accepted_compiled_solver_baseline_15.json`; plain `BAU` remains a
separate support scenario.

Run with the existing OSTRAM interpreter:

```powershell
$Py = 'C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe'
& $Py tests\regression\accepted_baseline.py
& $Py -m unittest discover -s tests\regression -p 'test_*.py' -v
```

These checks do not run a solver. Final housekeeping acceptance regenerates
the exact 15 compiled inputs in a disposable checkout and compares their
paths, filenames, sizes, SHA-256 values, and raw bytes with the frozen external
comparator documented by Stage 0.
