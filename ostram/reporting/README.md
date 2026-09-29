# ostram/reporting — result reports: training dashboard, summary workbook, interconnector direction

Readers of the combined result CSVs (`*Combined_Inputs_Outputs*.csv`) that produce a
self-contained HTML dashboard for `python -m ostram example report`, an Excel summary workbook,
and the shared interconnector-direction aggregation both use. None of them runs the pipeline or
discovers inputs from the current directory.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `training_dashboard.py` | Aggregates result snapshots (generation, capacity, new capacity, emissions, cost, interconnectors, storage) into the `ostram-profile-report-v1` payload and renders one offline HTML file with inline SVG charts | Snapshot CSVs given as `(label, path)` pairs (columns `Scenario`, `REGION`, `YEAR`, `TECHNOLOGY`, `EMISSION`, `ProductionByTechnologyAnnual`, `TotalCapacityAnnual`, `NewCapacity`, `AnnualEmissions`, `TotalDiscountedCost`, `CapitalInvestment`, ...), profile metadata (`country_regions`, `year_range`, `interconnectors`, `scenarios`, `storage`) | The HTML report file (`generate_report`) | Imported only (`ostram/examples.py` for `example report`; `summary_workbook.py` reuses its aggregation) |
| `summary_workbook.py` | Writes a multi-sheet Excel summary from combined result CSVs (folders, single CSVs or zip archives): generation, capacity, costs including fixed and variable O&M, per-year `TotalDiscountedCost` with a cumulative column, seasonal peaks, corridor flows | `*Combined_Inputs_Outputs.csv` (or every CSV) under the given paths, optional `--manifest` profile YAML (`storage`, `interconnectors`) | `OSTRAM_Summary_<UTC stamp>.xlsx` in the cwd or under `--out` | `python -m ostram.reporting.summary_workbook <folder|zip|csv ...> [--out PATH] [--manifest profile.yaml]` |
| `interconnector_direction.py` | Shared taxonomy of interconnector directions (from the registry manifest entries) and per-year aggregation of capacity and production for corridor series | Manifest entries and a DataFrame passed in; imports `TRN_PREFIX`, `VALID_DIRECTIONS`, `parse_tech_regions` from `pipeline/scenarios/rules/set_interconnector_direction.py` | Nothing | Imported only (`training_dashboard.py`) |

## How to run

From the repo root in `OSTRAM-env`, after a solved run:

```bash
python -m ostram --profile unescap example report unescap --capture baseline   # dashboard via examples.py
python -m ostram.reporting.summary_workbook <workspace>/execution --out reports/  # Excel summary
```

## Gotchas

- Units in `summary_workbook.py`: production in PJ, capacity in GW; TWh = PJ × 0.2777778; one
  GW-year = 31.536 PJ, so timeslice power in GW = PJ / (`YearSplit` × 31.536).
- `summary_workbook.py` refuses to overwrite an existing output file and lists scenarios without
  solver output as `SKIPPED` on its Readme sheet. Corridor flows use `TRN` technology production;
  the `Trade` result cannot express direction in the single-`REGION` combined CSV and is not used.
- `training_dashboard.py` counts a scenario as solved only if one of its output columns is present,
  raises on duplicate snapshot labels, and falls back to a default year range when neither
  metadata nor data provide one.
- `interconnector_direction.py` requires 13-character `TRN` codes and rejects duplicate or invalid
  directions; `interconnector_series` has no caller in `ostram/`.
