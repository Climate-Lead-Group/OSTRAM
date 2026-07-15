# Offline regression evidence

This directory provides the solver-free acceptance layer for repository cleanup. It
does not run OSTRAM stages, DVC, GLPK, CPLEX, or another optimizer. The harness only
reads files and uses `git` for repository metadata; `capture` writes compact evidence
to the explicitly supplied output directory.

The authoritative inventory is the 20 entries in `scenarios.yaml`. It is intentionally
JSON-compatible YAML so the harness has no third-party runtime dependency. The same is
true of `tolerances.yaml`.

## Commands

Run with an existing Python 3.10+ interpreter:

```powershell
$Py = 'python'  # or the Python executable from an existing OSTRAM-env

& $Py -m unittest discover -s tests\regression -p 'test_*.py' -v
& $Py tests\regression\ostram_regression.py discover --repo . --scope regression

& $Py tests\regression\ostram_regression.py capture `
  --repo . `
  --reference-repo ..\OSTRAM_mainredo `
  --out tests\regression\baselines\5ce4e66480e1-static-nosolver

& $Py tests\regression\ostram_regression.py compare `
  --baseline tests\regression\baselines\5ce4e66480e1-static-nosolver `
  --candidate tests\regression\baselines\5ce4e66480e1-static-nosolver `
  --profile strict
```

`capture` records five compact files: `manifest.json`, `coverage.csv`, `comparisons.csv`,
`metrics.csv`, and `hashes.csv`. Raw model artifacts remain where they already are and
are never copied into Git. `comparisons.csv` distinguishes raw exact matches,
format-normalized exact matches, missing evidence, file-set drift, and normalized drift.

## Evidence levels

1. **Exact pre-solver/static equivalence** uses raw and format-aware normalized hashes
   for tracked A1 snapshots, scenario configs, A2 CSVs, otoole CSVs, and any available
   compiled text. XLSX member hashing ignores ZIP timestamps; CSV normalization rejects
   duplicate keys and invalid values. Exact comparisons exclude generated
   `*_PRE_*`/`*_PREPATCH_*` backups and normalize only the
   `Restrictions.source_run_timestamp` workbook field.
2. **Available historical-output evidence** records normalized hashes for direct
   `Executables/<scenario>_0/Outputs/*.csv` files found in the read-only reference repo.
   Its coverage and provenance limitations are explicit in `coverage.csv`.
3. **Full solver-backed behavioral equivalence** is not established by this harness.
   It remains pending a coherent, source-bound all-20 CPLEX baseline.

Missing artifacts are evidence, not an invitation to fabricate coverage. A candidate
column remains empty until a safely regenerated candidate actually exists.
