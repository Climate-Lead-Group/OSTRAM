# Offline Regression Evidence

The solver-free regression harness lives in `tests/regression/`. It never calls OSTRAM
stages, DVC, GLPK, CPLEX, or another optimizer. Its only child process is `git`, used for
read-only metadata and tracked-file discovery.

## Current static inventory

- Authoritative A1 snapshot scenarios: 20/20.
- Authoritative rule/config scenario directories: 20/20.
- Tracked A2 and otoole snapshots: 17/20.
- Historical reference output folders: 16/20.
- Safely regenerated candidate artifacts: none in the no-solver audit.

The three scenarios without tracked A2/otoole snapshots are `B_Opt_LinkFreeze`,
`B_Opt_SolarHi10`, and `B_Opt_TradeCap50`. Historical direct outputs are also absent for
`B_Opt_TradeCap30`. These are explicit coverage gaps, not inferred passes.

## Run the checks

```powershell
$Py = 'python'  # or an existing OSTRAM-env Python executable

& $Py -m unittest discover -s tests\regression -p 'test_*.py' -v
& $Py tests\regression\ostram_regression.py discover --repo . --scope regression
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
   establishes complete CPLEX numerical equivalence.

See `tests/regression/README.md` for capture and comparison commands and the normalization
policy.

## Deferred solver-backed work

- Stage 3 utility reorganization is not attempted in this offline cleanup because the
  all-20 pre-solver gate is incomplete: tracked A2/otoole material covers 17 scenarios
  and no safely regenerated candidate set exists.
- The current `run.py --scenarios` propagation behavior is documented but unchanged. A
  future behavior change must be evaluated with a source-bound all-20 candidate and the
  solver-backed acceptance baseline.
- Full all-20 CPLEX behavioral equivalence, DVC pipeline reconstruction, and any core
  A0-B2/A3/configuration changes remain outside the no-solver evidence level.
