# Analysis visualization utilities

These standalone generators read existing OSTRAM inputs or solved outputs and write HTML
or figure artifacts. They are not imported by the A0-B2/A3 workflow and do not launch a
model stage or solver.

The implementations live here; compatibility wrappers remain at their former
`t1_confection/Z_AUX_*.py` paths. Script-relative inputs and output directories still
resolve to `t1_confection/`, preserving the established locations. The aggregated PWR
dashboard is the exception: it intentionally discovers CSV files in the caller's current
directory and writes its timestamped HTML there.

The scripts require their existing optional analysis dependencies, such as pandas,
Plotly, PyYAML, or openpyxl. Do not install optional dependencies only to perform a
repository cleanup check.
