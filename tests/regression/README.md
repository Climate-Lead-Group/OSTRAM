# Offline regression evidence

This directory provides the solver-free acceptance layer for repository cleanup. It
does not run OSTRAM stages, DVC, GLPK, CPLEX, or another optimizer. The harness only
reads files and uses `git` for repository metadata; `capture` writes compact evidence
to the explicitly supplied output directory.

The authoritative preservation inventory is the 20 entries in `scenarios.yaml`. A
separate `cleanup_acceptance` field selects the 16 non-superseded scenarios with complete
offline evidence. The four excluded scenarios remain discoverable and protected, and
each has an explicit exclusion reason. The inventory is intentionally JSON-compatible
YAML so the harness has no third-party runtime dependency. The same is true of
`tolerances.yaml`.

## Commands

Run with an existing Python 3.10+ interpreter:

```powershell
$Py = 'python'  # or the Python executable from an existing OSTRAM-env

& $Py -m unittest discover -s tests\regression -p 'test_*.py' -v
& $Py tests\regression\ostram_regression.py discover --repo . --scope preservation
& $Py tests\regression\ostram_regression.py discover --repo . --scope cleanup-acceptance

& $Py tests\regression\ostram_regression.py gate `
  --scope cleanup-acceptance `
  --evidence tests\regression\baselines\5ce4e66480e1-static-nosolver `
  --out tests\regression\baselines\5ce4e66480e1-static-nosolver\cleanup_acceptance_16.json

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
   It remains pending a coherent, source-bound CPLEX baseline for the accepted scope.

Missing artifacts are evidence, not an invitation to fabricate coverage. A candidate
column remains empty until a safely regenerated candidate actually exists.

## Scope policy

- **Preservation:** all 20 A1/config scenario definitions must remain present exactly.
- **Cleanup acceptance:** 16 non-superseded scenarios must have working and reference A1,
  config, A2, and otoole evidence with exact or normalized-exact static comparisons. The
  read-only reference must also contain compiled text and direct output CSVs for each.
- **Excluded but protected:** `B_Opt_LinkFreeze`, `B_Opt_SolarHi10`,
  `B_Opt_TradeCap30`, and `B_Opt_TradeCap50` are not inferred passes and are not deleted.

Passing the cleanup gate authorizes only offline-safe repository organization. It does
not establish solver or numerical equivalence.
