# timeslice_module/scripts — timeslice workbook generator, checks and comparisons (training)

Standalone scripts that built the OSTRAM timeslice fabric (4 seasons × 5 dayparts = 20
timeslices) and its 12/16/24-timeslice variants from sub-daily demand and Renewables.ninja data,
plus the checks and comparisons used to select and validate them. They are kept for training and
are not part of `python -m ostram run`; the maintained authority the pipeline reads is
`inputs/scenarios/OSTRAM_Timeslice_Inputs.xlsx`. Module overview and recipe:
`../README.md`; selection methodology: `../docs/ostram_timeslice_handover.md`.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `rebuild_reninja_timeslices_latest.py` | Consolidates the Renewables.ninja hourly solar and wind exports into a long hourly table and a timeslice-aggregated table with the OSeMOSYS mapping (UTC to local offsets; December assigned to S1 of the next model year) | `./ninja_data/solar/*.csv`, `./ninja_data/wind/*.csv` (relative to the cwd; the tracked data is under `../inputs/_Reno_Ninja/ninja_data/`) | `./ninja_data/output_rebuilt/{compiled_reninja_hourly,compiled_reninja_ts,catalog_reninja}.csv` and `timeslice_config.json` | `python rebuild_reninja_timeslices_latest.py` from the folder that contains `ninja_data/`; no flags |
| `build_ostram_timeslices.py` | Generates the timeslice parameters per country and region (demand fractions, capacity factors, year split) and the `OSTRAM_Timeslice_Outputs.xlsx` workbook from the PGCB, PUCSL, CEA, Bhutan, Nepal and Ninja sources | The hard-coded `BASE_DIR` tree (`Bangladesh/`, `Sri Lanka/`, `India/`, `Bhutan/`, `Nepal/`, `_Reno_Ninja/ninja_data/output_rebuilt/`) with compiled fast-path CSVs | `OUTPUT_DIR/OSTRAM_Timeslice_Outputs.xlsx`, `timeslice_config.json`, per-country `*_capacity_factors.csv`, `*_demand_fractions.csv`, `*_metadata.json`, `ostram_cf_selector_template.csv`; may write the compiled BGD CSVs next to the raw files | `python build_ostram_timeslices.py` after editing `BASE_DIR` / `OUTPUT_DIR`; executes at import, no flags |
| `check_daypart_sync.py` | Parses `DAYPART_DEF` out of the generator and the Ninja rebuilder with `ast` and verifies they match and tile 0-24 h | `scripts/build_ostram_timeslices.py`, `scripts/rebuild_reninja_timeslices_latest.py` | Nothing | `python timeslice_module/scripts/check_daypart_sync.py` (exit 0 in sync, 1 otherwise) |
| `check_variant.py` | Variant check suite: sum checks at literal and rounding tolerance, season and daypart structure, per-zone demand and CF shapes, against the adopted 5dp/20ts reference workbook | The variant workbook (positional), `--ref` (default `outputs/OSTRAM_Timeslice_Outputs_REFERENCE_5dp20ts.xlsx`) | Optional `--json` report | `python check_variant.py <variant.xlsx> [--ref REF] [--json OUT]` (exit 0 pass, 1 fail) |
| `compare_workbooks.py` | Cell-by-cell comparison of two workbooks (numeric `rtol` 1e-9, text exact, sheet names, order and used range), not hash-based | Two workbooks (argv 1 and 2) | A JSON report (argv 3) | `python compare_workbooks.py <ref.xlsx> <new.xlsx> <out.json>` |
| `solar_hour_profile.py` | Builds the hour-resolved solar CF truth from the Ninja hourly table and scores every candidate daypart fabric (RMSE and phantom solar in dark hours) | `--hourly` (default `inputs/_Reno_Ninja/ninja_data/output_rebuilt/compiled_reninja_hourly.csv`, not tracked) | Optional `--json` | `python solar_hour_profile.py [--hourly CSV] [--json OUT]` |
| `fabric_menu_summary.py` | Four-fabric summary table (12/16/20/24 ts): dominant solar block and mean CF, row counts and solver-size proxy | `outputs/OSTRAM_Timeslice_Outputs_{3dp12ts,4dp16ts,REFERENCE_5dp20ts,6dp24ts}.xlsx`, optional `outputs/solar_hour_profile.json` | Optional `--json` | `python fabric_menu_summary.py [--json OUT]` |
| `build_companion_html.py` | Patches the hand-built `../Timeslice_Fabric_Companion.html` template with values computed from the four workbooks and the fabric summary, aborting if any anchor is not found exactly once | The template, `../_session_logs/fabric_menu_summary.json` (not tracked), the four workbooks | `--out` HTML (default `../outputs/Timeslice_Fabric_Companion.html`) | `python scripts/build_companion_html.py [--out PATH]` |
| `compare_timeslice_runs.py` | Compares solved runs that differ in their timeslice fabric: intraday dispatch, demand profile and a summary of cost and new capacity by class | Per run: `ProductionByTechnology.csv`, `Demand.csv`, `NewCapacity.csv`, `TotalDiscountedCost.csv`, optional `NetChargeWithinDay.csv`, plus the fabric's `DaySplit.csv` and `YearSplit.csv` | `--out-dir` (default `./ts_compare_out`): `dispatch_intraday.png`, `demand_profile.png`, `summary.csv` | `python compare_timeslice_runs.py --run OUTPUTS_DIR FABRIC_DIR LABEL [--run ...] [--year 2035] [--season 3] [--out-dir DIR]` (needs matplotlib) |
| `sensitivity_timeslice_sweep.py` | Sweeps 2 to 12+ daypart schemes against the sub-daily data per country and season, computing within-block variance, peak preservation, RMSE and solar/wind CF variance | The hard-coded `BASE_DIR` tree (BGD, LKA, IND, NPL, BTN sources and the Ninja hourly table) | `<BASE_DIR>/_sensitivity/`: PNGs, `sensitivity_timeslice_summary.csv`, `sensitivity_ninja_cfs.csv`, `sensitivity_config.json` | `python sensitivity_timeslice_sweep.py` after editing `BASE_DIR`; executes at import; `RUN_*` booleans select countries |
| `rank_timeslice_schemes_v2.py` | Ranks the swept schemes with the v2 three-axis composite (demand 45 %, solar 45 %, wind 10 %) and defensibility penalties | `<BASE_DIR>/_sensitivity/sensitivity_timeslice_summary.csv` and `sensitivity_config.json` | `<BASE_DIR>/_sensitivity/ranking_v2/{ranking_full,ranking_by_budget}.csv`, `ranking_report.txt` (a copy of `ranking_by_budget.csv` is tracked under `../docs/`) | `python rank_timeslice_schemes_v2.py` after editing `BASE_DIR`; defines `main()` with no `__main__` guard |

## How to run

The documented recipe (`../README.md`, "How to generate a new variant"): edit `DAYPART_DEF`
identically in `build_ostram_timeslices.py` and `rebuild_reninja_timeslices_latest.py`, run
`check_daypart_sync.py`, run the Ninja rebuilder, run the generator, then `check_variant.py` on
the new workbook. Use a UTF-8 console (`chcp 65001`, `PYTHONIOENCODING=utf-8`). Requires pandas,
openpyxl, numpy (and matplotlib for the comparison and sweep scripts).

## Order / dependencies

`sensitivity_timeslice_sweep.py` → `rank_timeslice_schemes_v2.py` (scheme selection, historical);
`rebuild_reninja_timeslices_latest.py` → `build_ostram_timeslices.py` → `check_variant.py` /
`compare_workbooks.py`; `solar_hour_profile.py` → `fabric_menu_summary.py` →
`build_companion_html.py`.

## Gotchas

- `build_ostram_timeslices.py`, `sensitivity_timeslice_sweep.py` and `rank_timeslice_schemes_v2.py`
  hard-code another user's absolute Windows paths (`BASE_DIR`, `OUTPUT_DIR`); edit them before
  running. The first two execute at import time ("open in Spyder and press F5").
- If the Ninja rebuild is skipped or `DAYPART_DEF` differs, the generator does not fail: it
  silently remaps the capacity factors (`../README.md`, "Known limitation"). The only confirmation
  is the printed "Ninja CFs loaded: N entries (configs match)". `check_daypart_sync.py` is the gate.
- `rebuild_reninja_timeslices_latest.py` resolves `./ninja_data` against the cwd, so it must run
  from `../inputs/_Reno_Ninja/`, not from `timeslice_module/` as the recipe's `cd` suggests.
- Not tracked: `compiled_reninja_hourly.csv`, the raw PGCB and PUCSL files, `_session_logs/`,
  `outputs/solar_hour_profile.json`. `solar_hour_profile.py`, `sensitivity_timeslice_sweep.py`
  and `build_companion_html.py` need them regenerated first.
- `check_variant.py` assumes four seasons of 90/92/122/61 days, `SxDy` names and hard-coded zones;
  `compare_workbooks.py` reads with `data_only=True`, so formulas without cached values read as
  empty.
- Docs name the rebuilder `build_reninja_timeslices.py` and quote different generation times
  (about 45 s versus about 11 s); the script in the tree is `rebuild_reninja_timeslices_latest.py`.
