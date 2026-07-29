# Offline regression evidence

This directory provides the solver-free acceptance layer for repository cleanup. It
does not run OSTRAM stages, DVC, GLPK, CPLEX, or another optimizer. The evidence
harness reads files and uses `git` for repository metadata; `capture` writes compact
evidence to the explicitly supplied output directory. The unit suite additionally
uses the current Python interpreter for tightly constrained CLI `--help` and
unknown-command smoke checks guarded by fail-closed workflow-effect tests.

The authoritative preservation inventory is the 20 entries in `scenarios.yaml`. A
separate `cleanup_acceptance` field selects the 16 non-superseded scenarios with complete
static offline evidence. Final solver-consumed compiled-input equivalence has a narrower
15-scenario decision scope: it includes `A_Calibrated_BAU` as the baseline and excludes
plain `BAU`, which is a retained legacy support scenario whose available reference uses
the older `NoStorage` chain. The four superseded scenarios remain discoverable and
protected, and each has an explicit exclusion reason. The inventory is intentionally
JSON-compatible YAML so the harness has no third-party runtime dependency. The same is
true of `tolerances.yaml`.

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
3. **Source-bound solver baseline:** the current accepted 15-scenario source
   baseline is solver-backed and recorded by protected manifest SHA-256
   `778b4706522bc2b29911e74d5b31d24355c84cbe4c0c7d11d1c9680b2ddc9916`.
   This offline harness does not rerun it and cannot extend that solver claim
   to a future candidate.

Missing artifacts are evidence, not an invitation to fabricate coverage. A candidate
column remains empty until a safely regenerated candidate actually exists.

## Scope policy

- **Preservation:** all 20 A1/config scenario definitions must remain present exactly.
- **Static cleanup acceptance:** 16 non-superseded scenarios must have working and
  reference A1, config, A2, and otoole evidence with exact or normalized-exact static
  comparisons.
- **Accepted compiled-input identity:** 15 decision-relevant scenarios must
  have byte-exact final solver-consumed `.txt` files. Plain `BAU` remains under
  preservation and static cleanup acceptance but is excluded from this
  decision scope. The current portable identities are recorded in
  [`reports/accepted_compiled_solver_baseline_15.json`](reports/accepted_compiled_solver_baseline_15.json).
  The byte-preserved `41a54e5` predecessor remains historical evidence in
  [`reports/pre_correction_41a54e5_compiled_input_equivalence_15.json`](reports/pre_correction_41a54e5_compiled_input_equivalence_15.json).
- **Excluded but protected:** `B_Opt_LinkFreeze`, `B_Opt_SolarHi10`,
  `B_Opt_TradeCap30`, and `B_Opt_TradeCap50` are not inferred passes and are not deleted.

Passing a housekeeping gate requires the maintained no-solver protocol in
[`docs/regression.md`](../../docs/regression.md): reproduce all 15 raw compiled
inputs in isolated worktrees and compare exact paths, filenames, sizes, hashes,
and bytes. No normalization or solver rerun is part of that gate. Passing it
authorizes only offline-safe repository organization; it does not establish a
new solver or numerical-equivalence claim.
