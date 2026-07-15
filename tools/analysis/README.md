# Standalone analysis utilities

These scripts inspect existing combined outputs and create summaries, filtered CSVs, or
plots. They do not compile or solve the OSTRAM model and are not imported by the core
A0-B2/A3 pipeline.

| Utility | Purpose |
|---|---|
| `check_combined.py` | Print a compact scenario summary from a combined CSV |
| `ostram_scenario_analysis.py` | Produce scenario comparison plots and summaries |
| `ostram_trn_plotter.py` | Produce transmission-focused plots |
| `slice_by_country.py` | Filter a combined CSV by country, region, or scenario |

The former `t1_confection/<name>.py` paths remain as compatibility wrappers. New
documentation and automation should use `tools/analysis/<name>.py`.

The plotting utilities require the optional analysis dependencies already declared by
their imports, including matplotlib, NumPy, and pandas. Do not install them merely to run
repository-cleanup checks. `check_combined.py` and `slice_by_country.py` use filenames or
settings embedded in the scripts, so inspect those values before manual execution.
