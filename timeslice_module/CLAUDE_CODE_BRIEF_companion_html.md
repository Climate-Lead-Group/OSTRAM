# Claude Code session brief — build_companion_html.py: generate the fabric companion HTML from the actual workbooks

## Context

The module at `C:\Users\luisfernando\Desktop\timeslice_module` contains four
computed timeslice workbooks in `outputs/`:

    OSTRAM_Timeslice_Outputs_3dp12ts.xlsx
    OSTRAM_Timeslice_Outputs_4dp16ts.xlsx
    OSTRAM_Timeslice_Outputs_REFERENCE_5dp20ts.xlsx
    OSTRAM_Timeslice_Outputs_6dp24ts.xlsx

Two reference files sit beside this brief:

1. `Timeslice_Fabric_Companion.html` — a hand-built companion page. This is
   the DESIGN TEMPLATE: layout, sections, colours, chart types, and the
   teaching narrative are final. Its zone-level charts currently carry only
   the adopted 20ts data; the fabric grid explorer carries structural data
   only.

2. `Timeslice_Explainer_OSTRAM_SOASIA.html` — the original 20ts explainer.
   Its embedded `const DATA = {...}` JSON defines the DATA CONTRACT: per
   zone, arrays of {ts, df, ys} for demand and {ts, v} per technology for
   capacity factors, plus a config block with seasons and dayparts.

The goal: a script `scripts/build_companion_html.py` that reads the four
workbooks and emits ONE self-contained HTML where every section, including
the zone-level demand and CF explorers, switches per fabric with real
per-zone data extracted from the corresponding workbook.

## Tasks

### Task 1 — Understand the workbook schema

Open the REFERENCE_5dp20ts workbook with openpyxl or pandas. Identify the
sheets and columns that carry, per zone (10 zones: BGD, BTN, INDEA, INDNE,
INDNO, INDSO, INDWE, LKA, MDV, NPL):

- YearSplit per timeslice
- Demand fraction (SpecifiedDemandProfile) per timeslice
- Capacity factors per technology per timeslice (at minimum Solar, Hydro,
  Wind; include others such as Gas/Coal/Oil if present)

Validate against ground truth: the extracted 20ts values MUST match the
DATA object embedded in `Timeslice_Explainer_OSTRAM_SOASIA.html` (e.g. BGD
S1D2 demand fraction 0.107446, YearSplit 0.113014; BGD S1D2 Solar CF
0.3843). If extraction and ground truth disagree beyond rounding, stop and
report the discrepancy; do not proceed with wrong mappings. Document the
schema mapping in a comment block at the top of the script.

### Task 2 — Extract all four fabrics

Apply the same extraction to all four workbooks. Each yields a per-fabric
DATA object in the contract shape (timeslice codes differ per fabric:
S1D1..S4D3 for 12ts, S1D1..S4D4 for 16ts, etc.).

Additional per-fabric metadata comes from
`_session_logs/fabric_menu_summary.json` (phantom solar percentages, peak
solar block CFs, row scaling). Read these values from that file; do NOT
recompute or invent them. If the file is missing or lacks a value, omit
that stat from the HTML and log it.

Daypart definitions per fabric (boundaries and block-type classification
for the grid colouring) are:

    3dp12ts: D1 00-08 night | D2 08-16 solar | D3 16-24 night
    4dp16ts: D1 00-06 night | D2 06-12 solar | D3 12-18 shoulder | D4 18-24 evening
    5dp20ts: D1 00-06 night | D2 06-17 solar | D3 17-20 shoulder | D4 20-22 evening | D5 22-24 night
    6dp24ts: D1 00-05 night | D2 05-08 shoulder | D3 08-17 solar | D4 17-20 shoulder | D5 20-22 evening | D6 22-24 night

### Task 3 — Build the generator script

`scripts/build_companion_html.py`:

- Reads the four workbooks from `outputs/` and the summary JSON from
  `_session_logs/`.
- Emits `outputs/Timeslice_Fabric_Companion.html`, self-contained (all data
  embedded as JSON, Chart.js from CDN as in the template).
- Uses the hand-built companion as the design template. Extend it so the
  fabric tabs also switch the zone-level demand chart, the CF chart, and
  the heatmap to that fabric's data. The summary table and teaching-point
  prose stay as they are in the template.
- Deterministic: same inputs, same output (no timestamps inside the HTML
  body).
- Non-destructive: never modifies the workbooks. If the output HTML exists,
  overwrite it (it is a generated artifact, regeneration is the point).
- Command line: `python scripts/build_companion_html.py` with optional
  `--out` path. No other required arguments; paths default to the module
  layout.

### Task 4 — Verify

- Spot-check at least 10 values across fabrics and zones: open the emitted
  HTML's embedded JSON and compare against the workbook cells directly.
  List the checked values and verdicts in the report.
- The 20ts fabric's zone data in the new HTML must match the original
  explainer's DATA object (same source, so any mismatch means an extraction
  bug).
- Open-in-browser sanity: confirm the HTML parses (no unbalanced braces from
  JSON embedding; escape any </script> sequences in data).

### Task 5 — Close out

Write `_session_logs/COMPANION_BUILD_REPORT.md`: schema mapping found,
values verified, anything omitted and why. Final console output: the
output path, fabric count, zone count per fabric, verification verdict.

## Standing rules

- No invented numbers. Every value in the HTML traces to a workbook cell or
  the summary JSON ("fake charts" rule). Gaps are omitted and logged, never
  filled in.
- The source tree outside the module is not needed and must not be touched.
- British spelling in any prose added; no em-dashes.
