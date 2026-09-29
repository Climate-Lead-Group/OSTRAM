# ostram/pipeline/compilation/transforms — import-safe steps of the B1 compiler

Pure and injectable helpers extracted from the legacy B1 compiler so that its configuration,
planning, validation, tabular transformations, effects and output delivery can be tested without
running the compiler. Nothing here does filesystem or pipeline work at import (package
docstring); `../compiler.py` calls these modules in a fixed order. The B1 contract is in
`docs/pipeline.md` ("Compile model inputs (B1)").

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `planning.py` | Builds the compiler's transform plan from the config mapping: years, sets, sorted timeslices, workbook paths per `Print_*` suffix, extra-input and output roots, and the `EMISSION.csv` source | The `params` mapping (from `Config_MOMF_T1_A.yaml`), `resolve_paths()` for the `og_csvs_inputs/EMISSION.csv` fallback | Nothing (sorts the config's `Timeslices` list in place) | Imported only (`compiler.py`) |
| `effects.py` | Injectable effects: read the YAML config, open workbooks, read CSVs, load a pickle, write the completed and structure workbooks with xlsxwriter, and `allocate_country_fuel_costs` (fail-closed shared-pool fuel-cost allocation with an audit CSV) | The config YAML, workbooks, CSVs, the `fuel_cost_allocation` config (`fuel-cost-allocation-v1`) | `*_Completed.xlsx`, the structure workbook, `_fuel_cost_allocation_<scenario>.csv` | Imported only (`compiler.py`) |
| `validation.py` | Legacy timeslice consistency checks for demand, capacity factor, yearsplit and daysplit sheets, with injectable reporting and stopping | Config timeslice list and mode, workbook timeslices | Nothing (prints and `sys.exit` on failure) | Imported only (`compiler.py`) |
| `tables.py` | Pure tabular transformations: `normalize_year_like_columns`, `build_system_parameter_rows`, `build_structure_tables` | DataFrames passed in | Nothing | Imported only (`compiler.py`) |
| `delivery.py` | Cleans the parameter tables and writes `<output_root>/<scenario>/<Parameter>.csv` plus the sorted set CSVs, then the same for each `Other_Scenarios` entry | DataFrames passed in | `<A2_output_main_scen>/<scenario>/*.csv`, copies under `<A2_output>/<other>/` | Imported only (`compiler.py`) |

## How to run

Not run directly. `python -m ostram compile-inputs --scenarios <name>` runs the compiler, which
imports these modules (see `../README.md`). Unit coverage:
`python -m pytest tests/regression/test_a1_b1_transforms.py -q`.

## Order / dependencies

Inside `compiler.py`: `effects.read_config` → `planning.build_transform_plan` → workbook opens
(`effects.open_workbook`) → `validation.validate_demand_timeslices` →
`validation.validate_capacity_timeslices` → `validation.validate_yearsplit_timeslices` →
`validation.validate_daysplit_time_brackets` → `tables.build_system_parameter_rows` →
`effects.write_*` → `tables.build_structure_tables` → `effects.allocate_country_fuel_costs` →
`delivery.clean_parameter_tables` → `delivery.deliver_main_csvs` →
`delivery.deliver_additional_csvs`.

## Gotchas

- Behaviour is deliberately predecessor-identical, quirks included: `planning.py` keeps paths as
  unnormalised strings resolved against the cwd; `tables.py` returns the same object when no header
  qualifies and does not reject duplicate normalised columns; `delivery.clean_parameter_tables`
  returns a mapping rebuilt from the main tables, not the predecessor's accumulated mapping;
  `validation.py` keeps the original message typos and the unrepaired one-bracket branch.
- `effects.write_completed_demand_workbook` and `load_pickle` intentionally have no context manager
  or failure cleanup.
- `allocate_country_fuel_costs` fails closed on unmapped fossil consumers, direct fossil demand,
  non-unit-output `MIN` supply, missing country prices and non-positive prices; it relies on
  positional technology-name slices (`tech[6:9]`, `fuel[:3]`).
- `tables.py` forces object dtype for structure-table columns to avoid pandas 3's nullable string
  dtype.
