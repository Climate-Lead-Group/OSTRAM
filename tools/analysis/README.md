# Standalone analysis utilities

These scripts inspect existing inputs or solved outputs and create summaries, filtered
CSVs, reports, or plots. They do not compile or solve the OSTRAM model and are not
imported by the core A0-B2/A3 pipeline.

| Utility | Purpose |
|---|---|
| `check_combined.py` | Print a compact scenario summary from a combined CSV |
| `ostram_scenario_analysis.py` | Produce scenario comparison plots and summaries |
| `ostram_trn_plotter.py` | Produce transmission-focused plots |
| `slice_by_country.py` | Filter a combined CSV by country, region, or scenario |
| `concat_all_scenarios.py` | Stack existing per-scenario input and output CSVs without row multiplication |
| `analyse_sensitivity.py` | Build the Phase-B sensitivity comparison CSV and text report |
| `reproduce_A1_A6.py` | Reproduce the A1-A6 figures and their audit CSV |
| `visualization/Z_AUX_generate_interactive_dashboards_aggregated.py` | Build interactive aggregated PWR dashboards |
| `visualization/Z_AUX_generate_RES_diagram.py` | Build the reference-energy-system diagram |
| `visualization/Z_AUX_generate_transmission_maps.py` | Build transmission maps and the dispatch chart |
| `visualization/Z_AUX_interconnections_dashboard.py` | Build the cross-border interconnections dashboard |

The former `t1_confection/` paths remain as compatibility wrappers, including
`t1_confection/concat_all_scenarios_2.py`. New documentation and automation should use
the canonical paths in this directory. The protected
`t1_confection/sensitivity_expansion/analyse_ws4_vs_phaseB.py` analysis utility remains
in place and is deferred from this Tier 1 move.

The plotting utilities require the optional analysis dependencies already declared by
their imports, including matplotlib, NumPy, and pandas. Do not install them merely to run
repository-cleanup checks. Several utilities intentionally continue reading from and
writing to `t1_confection/`; moving their implementations does not relocate model data or
generated reports. `check_combined.py`, `slice_by_country.py`, and the aggregated dashboard
also retain current-working-directory inputs, so inspect their settings before manual
execution.

See [`visualization/README.md`](visualization/README.md) for visualization-specific path
and dependency notes.
