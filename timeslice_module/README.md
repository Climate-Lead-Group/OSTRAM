# OSTRAM Timeslice Module

## What this is

This module contains the OSTRAM timeslice workbook generator and four
pre-computed workbook cases, kept for training on how the fabric (the
number and boundaries of timeslices) affects the modeled workbook. The
adopted case is the 5-daypart, 20-timeslice fabric; the other three are
computed alternatives at coarser or different resolutions.

## Folder layout

- `scripts/` — the generator, the Ninja rebuilder, the sensitivity sweep, the ranking script, and `compare_timeslice_runs.py`, which plots two or more solved runs against each other
- `inputs/` — source data (hourly demand, capacity factors, Ninja raw CSVs)
- `outputs/` — the four pre-computed workbook cases plus per-run evidence
- `docs/` — the selection-methodology handover and the ranking justification (`ranking_by_budget.csv`)

## The four fabric cases

| Workbook file | Fabric | Timeslices | Daypart boundaries | Status |
|---|---|---|---|---|
| `OSTRAM_Timeslice_Outputs_REFERENCE_5dp20ts.xlsx` | 5 dayparts (adopted) | 20 | 00–06 / 06–17 / 17–20 / 20–22 / 22–24 | Canonical |
| `OSTRAM_Timeslice_Outputs_4dp16ts.xlsx` | 4 equal 6-hour blocks | 16 | 00–06 / 06–12 / 12–18 / 18–24 | Computed |
| `OSTRAM_Timeslice_Outputs_3dp12ts.xlsx` | 3 equal 8-hour blocks | 12 | 00–08 / 08–16 / 16–24 | Computed |
| `OSTRAM_Timeslice_Outputs_6dp24ts.xlsx` | 6 dayparts (sweep winner) | 24 | 00–05 / 05–08 / 08–17 / 17–20 / 20–22 / 22–24 | Computed |

The 6dp/24ts fabric won the unconstrained sensitivity sweep (`docs/ostram_timeslice_handover.md`)
but was rejected because solver size scales with timeslice count; 5dp/20ts is the best-scoring
scheme at a five-daypart budget (`docs/ranking_by_budget.csv`).

## The adopted scheme

The model uses 4 seasons × 5 dayparts = 20 timeslices (scheme ID `5dp_D_6_17`).

| Season | Months | Days |
|---|---|---|
| S1 Winter | Dec–Feb | 90 |
| S2 Pre-monsoon | Mar–May | 92 |
| S3 SW Monsoon | Jun–Sep | 122 |
| S4 Post-monsoon | Oct–Nov | 61 |

| Daypart | Hours (local) |
|---|---|
| D1 Night | 00–06 |
| D2 Solar window | 06–17 |
| D3 Evening peak | 17–20 |
| D4 Late evening | 20–22 |
| D5 Late night | 22–24 |

## How to generate a new variant

1. Edit `DAYPART_DEF` in both `scripts/build_ostram_timeslices.py` and
   `scripts/rebuild_reninja_timeslices_latest.py` (must match exactly).
2. Run the Ninja rebuilder. Confirm the "configs match" message appears.
3. Run the generator.
4. The new workbook appears in `outputs/`.
5. Revert `DAYPART_DEF` in both files if the module should stay canonical.

Each variant takes roughly 45 seconds to generate. Build time does not
depend meaningfully on timeslice count.

## Key trade-off

Coarser fabrics (fewer timeslices) are cheaper to solve but misrepresent
when solar energy is available.

| Fabric | Timeslices | Solar block | Mean solar CF | Phantom solar | Build time |
|---|---|---|---|---|---|
| 3dp/12ts | 12 | D2 08–16 | 0.491 | 6.75% | ~41 s |
| 4dp/16ts | 16 | D2 06–12 | 0.402 | 0.15% | ~40 s |
| 5dp/20ts | 20 | D2 06–17 | 0.395 | 0.42% | ~45 s |
| 6dp/24ts | 24 | D3 08–17 | 0.450 | 0.28% | ~46 s |

Phantom solar is the fraction of solar energy a flat-CF block credits to hours
when the sun is down. The 3dp/12ts case credits 16 times more solar energy to
dark hours than the adopted 5dp/20ts fabric, so a solver could dispatch PV
capacity at night. The cost of finer fabrics is paid at solve time, not at
generation time.

## Known limitation

If the Ninja rebuild is skipped, the generator silently uses stale
capacity factors from a previous fabric — always rebuild before generating.
`scripts/check_daypart_sync.py` verifies that the two `DAYPART_DEF` blocks agree.

## Methods paragraph (adopted scheme)

> The OSTRAM model uses four seasons (Winter, Pre-monsoon, SW Monsoon, Post-monsoon) crossed with five dayparts per day, for a total of 20 timeslices per year. The daypart boundaries (00–06, 06–17, 17–20, 20–22, 22–24 local time; scheme `5dp_D_6_17`) were selected from a sensitivity sweep over 23 candidate schemes spanning equal-width, solar-window, morning-ramp, and physics-informed designs, evaluated against regional hourly demand profiles for the seven primary regions (Bangladesh, Sri Lanka, and five Indian sub-regions) and re-aggregated Renewables.Ninja solar and wind capacity-factor series. Schemes were scored on a composite weighted by three independent information axes: demand fit (45%, split across step-reconstruction RMSE, peak-preservation ratio, and worst-region RMSE), solar-hour differentiation (45%, the coefficient of variation of block-mean solar capacity factors), and wind-hour differentiation (10%); the three-axis weighting avoids triple-counting the demand signal carried by the correlated RMSE/peak/worst-region metric cluster. Aggregation used capacity weighting across the seven primary regions, with a small defensibility penalty on blocks shorter than three hours. `5dp_D_6_17` is the highest-scoring scheme at a five-daypart budget (composite 0.80, capacity-weighted RMSE 0.032, solar CV 1.94); it isolates the midday solar window (06–17) from the dark hours and resolves the evening peak (17–20) and post-peak taper (20–22, 22–24) as distinct dispatch regimes. The five-daypart resolution was adopted in place of the unconstrained optimum (`6dp_D_ramp`, 24 timeslices, composite 0.86, RMSE 0.024, solar CV 1.81) to keep the model within the linear solver's variable and constraint budget; `6dp_D_ramp` remains the recommended upgrade should compute capacity increase.

## What this module does not do

It generates the timeslice input workbook only. Wiring a new fabric into the
scenario template, adjusting the model's set definitions to the new timeslice
count, and paying the solver cost of a different resolution are separate steps
on the modeller's side; the pipeline reads `inputs/scenarios/OSTRAM_Timeslice_Inputs.xlsx`.
