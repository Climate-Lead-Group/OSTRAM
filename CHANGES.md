# CHANGES

## 2026-08-13 — Item 2: `ostram/reporting/summary_workbook.py`

New standalone module per `DESIGN_Zixuan_Items_1_2_20260813.md` (Item 2): a
thin xlsx writer over `training_dashboard.aggregate_metrics`, run as
`python -m ostram.reporting.summary_workbook <folder-or-zip-or-csv ...>
[--out PATH] [--manifest profile.yaml]`. Non-destructive: reads inputs,
writes `OSTRAM_Summary_<timestamp>.xlsx`, refuses to overwrite.

### Judgement calls (decision rules, not pause-and-ask)

1. **Category map source.** The design says to copy the substring-pattern map
   from `ostram_report_verify.py`, but that script does not exist in this
   repository (it predates the package refactor). The workbook instead uses
   the dashboard's own `TECH_FAMILIES` prefix taxonomy (already tested by
   `test_training_dashboard.py`), extended with explicit buckets for
   Storage (profile storage prefixes), Interconnector (13-char `TRN*`),
   Backstop (`PWRBCK*`), internal transmission (`PWRTRN*`), and
   `Other (uncategorised)`. The full decision order is printed on the
   Readme sheet, and every uncategorised code observed is listed there —
   nothing is dropped silently.

2. **Trade vs TRN production for corridor flows.** The design prefers the
   `Trade` result when non-empty. Confirmed against
   `ostram/pipeline/execution/concatenate.py`: `Trade` is indexed
   `[REGION, REGION, ...]` but the combined CSV keeps a single REGION
   column, so the counterpart region (corridor direction) cannot survive
   concatenation. The workbook therefore always uses TRN-technology
   production and says so on the Readme sheet.

3. **Emissions sheet dimension.** The design table wants
   scenario × emission × year, but the dashboard's `emissions_series` sums
   across emission codes. Implemented as a small extension
   (`emissions_by_code`) in the dashboard's own dedup pattern; the sheet's
   Total column reproduces the dashboard aggregation.

4. **Capital expenditure by category.** `aggregate_metrics` only carries the
   horizon capex KPI, so the per-year, per-category table is an extension
   (`stacked_by_category` on `CapitalInvestment`) in the same pattern as
   `_stacked_by_family`, but covering ALL technologies (storage, TRN,
   backstop, internal transmission included) so the sheet total reconciles
   with the model's capex. `CapitalInvestmentStorage` is a separate block on
   the same sheet, keyed by STORAGE code, as the design specifies.

5. **YearSplit source.** The combined CSV is inputs+outputs, so `YearSplit`
   is already a column in the same file (confirmed via
   `runner.concatenate_all_scenarios`). Sheet 8 reads it from there instead
   of requiring a separate `YearSplit.csv`, which keeps the single-file
   input contract (snapshots under `reports/snapshots/` work unchanged).
   If YearSplit rows are absent, sheet 8 states the gap instead of failing.

6. **Duplicate scenario ids across input files.** Before/after snapshot
   captures of the same scenario carry the same `Scenario` column value.
   When the same scenario id appears in more than one input file, labels
   are disambiguated as `<file-stem> · <scenario>`; otherwise the plain
   scenario id is used.

7. **Season split.** Timeslices are `S1D1..S4D5` (verified in
   `sync_og_to_ts20.py`), so the season is derived from the `S#` prefix.
   The Readme and sheet 8 note that these are the model's four seasons,
   not a binary rainy/dry split (per the design's Readme-note requirement).

8. **No chunked reader.** The design mentions a chunked reader for the
   bulky timeslice CSV on the full model. The workbook bounds memory with a
   header-probe + `usecols` load (the dashboard's `_load_frame` pattern,
   widened to the workbook's column set); on the two-region training model
   the combined CSV is small. Chunking can be added when the full-model
   route (explicitly deferred by the design) is taken up.

9. **Private-helper imports.** The workbook imports `_dedup` and
   `_scenario_frame` from `training_dashboard` deliberately: the design's
   architecture decision is "workbook numbers equal dashboard numbers by
   construction", which requires the dashboard's own dedup semantics, not a
   re-implementation.

10. **Skipped scenarios.** Combined snapshots accumulate input rows for
    every scenario, so a scenario without solver outputs is skipped
    (mirroring the dashboard's `available` flag) and listed as SKIPPED on
    the Readme sheet.

11. **Cost cross-check.** Sheet 3's per-scenario cumulative total is checked
    against the dashboard's horizon `total_discounted` KPI; a mismatch is
    reported in the Data Gaps block rather than silently accepted.

### Verification (2026-08-13)

- Training model run: `example prepare unescap` +
  `--profile unescap run --scenarios B_Optimised_VRE` completed in 15m36s;
  combined CSV at
  `workspace/profiles/unescap/execution/OSTRAM_StorageDelay_Combined_Inputs_Outputs.csv`
  (17.5 MB). `Trade` is absent from the real output, confirming judgement
  call 2; `YearSplit` is present as a column, confirming call 5.
- Workbook ran on the real execution folder without edits
  (`--manifest examples/unescap/profile.yaml`); 14/14 spot checks passed
  against manual pandas filters of the source CSV, covering all eight data
  sheets (capacity coal/solar/storage, generation solar/coal, per-year and
  cumulative TotalDiscountedCost, coal and all-tech capex, fixed and
  fixed+variable opex, emissions total, corridor GWh, and the sheet-8 peak
  PV value + timeslice, which both the workbook and the manual recompute
  place at S2D2/season S2 in 2050).
- GAP fallback exercised with the real CSV minus its
  `ProductionByTechnology` column: the workbook still writes, sheet 8
  carries a DATA GAPS block naming the missing column, and the annual
  sheets stay fully populated.
- Synthetic smoke test additionally verified: input-only scenarios are
  skipped and listed on the Readme, uncategorised techs are listed, and
  the cumulative cost column accumulates correctly.
- openpyxl `load_workbook` round-trips the output; opening in desktop
  Excel/LibreOffice was not verifiable in this session.
- Test suite (`python -B -m unittest discover -s tests -p "test_*.py"`):
  298 tests, 296 pass, 2 failures. Both failures are pre-existing on this
  branch and unrelated to the workbook (no tracked file was modified;
  both reproduce with the new files absent):
  - `test_b2_orchestration...test_exact_solver_command_for_every_supported_solver`
    (solver='cbc'): the runner now passes `primalTolerance`/`dualTolerance`
    args the characterization does not expect;
  - `test_interconnector_rc_v1...test_protected_cap_and_relax_implementations_are_byte_identical`
    (`relax_interconnectors.py`): the pinned sha256 does not match the
    file on this branch.
- Side effect of the canonical run command, not of the workbook: the B2
  stage's DVC initialisation staged `.dvc/.gitignore` and `.dvc/config`
  into the git index. Left as-is for the maintainers to decide.

### Environment note (fresh clone)

`python -m ostram --profile unescap run` requires the profile workspace to
be prepared first (`python -m ostram example prepare unescap`) and needs
`conda` on PATH (the B2 stage re-invokes tools through the conda
environment).
