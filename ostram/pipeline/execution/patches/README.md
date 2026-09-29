# ostram/pipeline/execution/patches — B2 datafile patch chain

Stand-alone patchers that B2 applies to the preprocessed OSeMOSYS datafile before the solve. Each
one is a module CLI that reads a datafile and writes a sibling with a suffix (only
`days_in_day_type.py` edits in place); the source file is never modified. The wrappers that call
them, the flags that gate them and the order live in `ostram/pipeline/execution/runner.py` and
`orchestrator.py` (see `../README.md`, "Order / dependencies"). Flags and defaults are documented
in `docs/configuration.md`, "Datafile Patch Chain".

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `days_in_day_type.py` | Rewrites the `DaysInDayType` block from `YearSplit` so each season gets its real day count instead of the default 7 | The target datafile: `YearSplit` block, `set DAYTYPE`, timeslice tags `S<season>D<bracket>` | The same file, in place | `python -B -m ostram.pipeline.execution.patches.days_in_day_type <Pre_processed_*.txt>`; always run by B2, no flag |
| `storage_delay.py` | Blocks storage builds for the first N model years (`StorageBuildAllowed` = 0, PWR storage caps 0, blocked-year minima dropped) and adds the `SD1_StorageBuildDelay` constraint to a patched copy of the model | The datafile (sets `REGION`, `YEAR`, `STORAGE`; `TechnologyToStorage`/`FromStorage`; capacity caps and minima) and the model given by `--model-input` | `-o` datafile and `--model-output` model (default name `osemosys_fast_preprocessed_storage_delay.txt`, placed under `<Executables>/<scenario>_0/`) | `python -B -m ... storage_delay <in> -o <out> --model-input M --model-output M2 [--first-n-years 5] [--storage-prefixes ...] [--storages ...] [--allowed-value -1] [--dry-run]`; gated by `storage_delay_active` |
| `strip_storage.py` | Diagnostic: removes selected storage facilities (`--mode all`, `tech` or `class`) and disables their feeding PWR technologies, dropping the minima that would otherwise make the LP infeasible | The datafile: `set STORAGE`, `set YEAR`, storage parameters, `TechnologyToStorage`/`FromStorage`, `TotalAnnualMaxCapacity`, minimum capacity and activity parameters | `-o` datafile (CRLF) | `python -B -m ... strip_storage <in> -o <out> --mode {tech,class,all} [--targets ...] [--no-tech-cap-zero]`; gated by `strip_storage_active` |
| `open_pwrbck_caps.py` | Diagnostic: replaces zero `TotalAnnualMaxCapacity` / `TotalAnnualMaxCapacityInvestment` rows of backstop technologies (`PWRBCK*`) with a high value so infeasibilities can be diagnosed | The datafile blocks named by `--blocks` | `-o` datafile | `python -B -m ... open_pwrbck_caps <in> -o <out> [--pattern PWRBCK] [--value 9999] [--blocks ...]`; gated by `open_pwrbck_active` |
| `reserve_margin_repair.py` | Older reserve-margin repair: sets `ReserveMarginTagTechnology` credits for `PWRBCK` and `PWRCCS` and replaces sentinel firm-capacity caps (`0`, `9999`) of `PWRPET`/`PWROIL`/`PWRNGS` with per-country-region fallbacks from CLI values or a CSV | The datafile (`ReserveMarginTagTechnology`, `ResidualCapacity`, minima, caps, sets `YEAR`/`TECHNOLOGY`), optional `--fallback-csv` | `-o` datafile (CRLF), optional `--warnings-file` | `python -B -m ... reserve_margin_repair <in> -o <out> [--backstop-credit 1.0] [--ccs-credit 0.9] [--skip-*] [--target-prefixes ...] [--sentinel-values ...] [--fallback-csv F] [--stock-fallback CR=V ...] [--flow-fallback CR=V ...]`; gated by `reserve_margin_repair_active` (`False` in the committed YAML) |
| `reserve_margin_repair_xlsx.py` | Current reserve-margin repair: same credits and sentinel replacement, with fallbacks taken from a workbook sheet and reviewed pin rows protected from replacement | The datafile, `--fallback-xlsx` (`inputs/execution/firm_capacity_fallbacks_by_cr.xlsx`, sheet `fallbacks`), the pin rules through `ostram.pipeline.scenarios.rules.apply_base_year_pin` | `-o` datafile (CRLF) and `<out>.warnings.txt` (path compacted when over the Windows path budget) | `python -B -m ... reserve_margin_repair_xlsx <in> -o <out> --fallback-xlsx X [--xlsx-sheet S] [--backstop-credit 1.0] [--ccs-credit 0.9] [--skip-*] [--target-prefixes ...] [--sentinel-values ...] [--warnings-file F]`; gated by `reserve_margin_xlsx_active` |

`...` in the Run column stands for `ostram.pipeline.execution.patches`.

## How to run

The patchers are normally run by B2 (`python -m ostram run ...`), each on the previous active
patcher's output under `<workspace>/execution/Executables/<scenario>_0/`. To run one by hand,
use the module form above from the repo root in `OSTRAM-env`; `reserve_margin_repair_xlsx.py`
uses a relative import and must be run as a module, not as a file.

## Order / dependencies

`days_in_day_type` (in place) → `storage_delay` → `strip_storage` → `open_pwrbck_caps` →
`reserve_margin_repair` → `reserve_margin_repair_xlsx`, each gated by its `*_active` flag and each
adding its suffix to the file name (`runner.py` builds `chain_parts` in this order).

## Gotchas

- `storage_delay` and `strip_storage` are mutually exclusive; delay wins and strip is forced off
  (`orchestrator.apply_configuration_overrides`). The committed YAML has delay on and strip off.
- `storage_delay.py` patches the model too; the solver must use the patched model, which
  `runner.py` sets as the absolute `osemosys_model`. The patched model may not be written inside
  the package or `model/` folders.
- `storage_delay.py` and `strip_storage.py` infer the storage-to-technology mapping as
  `"PWR" + storage_name[:-2]` when the link parameters do not give it.
- `strip_storage.py` and `reserve_margin_repair*.py` assume the single region `GLOBAL`.
- `open_pwrbck_caps.py` only lifts caps; the backstop's high `VariableCost` still keeps it out of
  cost-optimal solutions (module docstring).
- `reserve_margin_repair.py` does not accept the `--open-capacity-*` and `--skip-capacity-opening`
  flags that `runner.run_reserve_margin_repair_patcher` passes; enabling
  `reserve_margin_repair_active` would fail at argument parsing. Not fixed here.
- The xlsx repair's in-code default workbook name (`firm_capacity_fallbacks_by_cr_0p5.xlsx`)
  differs from the tracked file; B2 resolves the tracked
  `inputs/execution/firm_capacity_fallbacks_by_cr.xlsx` before calling it.
- `days_in_day_type.py` hard-codes day type `1` and only warns when more day types exist.
- Module docstrings still carry the pre-refactor script names (`inject_DaysInDayType.py`,
  `patch_storage_delay.py`, `patch_reserve_margin_repair.py`, `patch_reserve_margin_repair_careful.py`).
