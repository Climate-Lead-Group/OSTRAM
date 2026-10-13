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
- `docs/` — handover documentation, the ranking justification, and the adoption rationale README
- `_session_logs/` — build evidence from the automated sessions (not needed for normal use)

## The adopted scheme

The model uses 4 seasons and 5 dayparts, 20 timeslices in all. The scheme
id is `5dp_D_6_17`. The seasons and dayparts are defined in
`scripts/build_ostram_timeslices.py`.

| Season | Months | Days |
|---|---|---|
| S1 Winter | December to February | 90 |
| S2 Pre-monsoon | March to May | 92 |
| S3 SW Monsoon | June to September | 122 |
| S4 Post-monsoon | October to November | 61 |

| Daypart | Hours, local time |
|---|---|
| D1 | 00 to 06 |
| D2 | 06 to 17 |
| D3 | 17 to 20 |
| D4 | 20 to 22 |
| D5 | 22 to 24 |

The sensitivity sweep in `docs/ostram_timeslice_handover.md` ranks
`6dp_D_ramp`, with 24 timeslices, first when the number of dayparts is free.
The model uses the 5-daypart scheme to stay within the solver's variable and
constraint budget. `5dp_D_6_17` is the best scheme at a budget of 5 dayparts
in `docs/ranking_by_budget.csv` and the documented fallback of that analysis.
`6dp_D_ramp` is the upgrade path if solver capacity increases. The model
runs 20 timeslices.

## The four fabric cases

| Workbook file | Fabric | Timeslices | Status |
|---|---|---|---|
| `OSTRAM_Timeslice_Outputs_REFERENCE_5dp20ts.xlsx` | 5 dayparts (adopted) | 20 | Canonical |
| `OSTRAM_Timeslice_Outputs_4dp16ts.xlsx` | 4 equal 6-hour blocks | 16 | Computed |
| `OSTRAM_Timeslice_Outputs_3dp12ts.xlsx` | 3 equal 8-hour blocks | 12 | Computed |
| `OSTRAM_Timeslice_Outputs_6dp24ts.xlsx` | 6 dayparts (sweep winner) | 24 | Computed |

## Methods paragraph

The paragraph below describes the adopted scheme for a report. Its figures
are in `docs/ranking_by_budget.csv` and `docs/ostram_timeslice_handover.md`.

> The OSTRAM model uses four seasons (Winter, Pre-monsoon, SW Monsoon,
> Post-monsoon) crossed with five dayparts per day, for a total of 20
> timeslices per year. The daypart boundaries (00 to 06, 06 to 17, 17 to 20,
> 20 to 22 and 22 to 24 local time; scheme `5dp_D_6_17`) were selected from a
> sensitivity sweep over 23 candidate schemes spanning equal-width,
> solar-window, morning-ramp and physics-informed designs, evaluated against
> regional hourly demand profiles for the seven primary regions (Bangladesh,
> Sri Lanka and five Indian sub-regions) and re-aggregated Renewables.ninja
> solar and wind capacity-factor series. Schemes were scored on a composite
> weighted by three independent information axes: demand fit (45 per cent,
> split across step-reconstruction RMSE, peak-preservation ratio and
> worst-region RMSE), solar-hour differentiation (45 per cent, the coefficient
> of variation of block-mean solar capacity factors) and wind-hour
> differentiation (10 per cent). The three-axis weighting avoids
> triple-counting the demand signal carried by the correlated RMSE, peak and
> worst-region metrics. Aggregation used capacity weighting across the seven
> primary regions, with a small defensibility penalty on blocks shorter than
> three hours. `5dp_D_6_17` is the highest-scoring scheme at a five-daypart
> budget (composite 0.80, capacity-weighted RMSE 0.032, solar CV 1.94). It
> isolates the midday solar window (06 to 17) from the dark hours and resolves
> the evening peak (17 to 20) and the post-peak taper (20 to 22, 22 to 24) as
> distinct dispatch regimes. The five-daypart resolution was adopted in place
> of the unconstrained optimum (`6dp_D_ramp`, 24 timeslices, composite 0.86,
> RMSE 0.024, solar CV 1.81) to keep the model within the linear solver's
> variable and constraint budget; `6dp_D_ramp` remains the recommended
> upgrade should compute capacity increase.

## How to generate a new variant

1. Edit `DAYPART_DEF` in both `scripts/build_ostram_timeslices.py` and
   `scripts/rebuild_reninja_timeslices_latest.py`. The two definitions must
   be identical.
2. Run `python scripts/check_daypart_sync.py`. It exits with code 0 when the
   two definitions agree and are structurally valid, and with code 1
   otherwise.
3. Run the Ninja rebuilder, `python scripts/rebuild_reninja_timeslices_latest.py`.
4. Run the generator, `python scripts/build_ostram_timeslices.py`. It prints
   `configs match` when the rebuilt capacity factors carry the same dayparts
   as its own definition.
5. The new workbook appears in `outputs/`. Rename it with a fabric suffix so
   that it does not overwrite an existing case.
6. Revert `DAYPART_DEF` in both files if the module should stay canonical.

Build time does not depend meaningfully on timeslice count.

## Key trade-off

Coarser fabrics (fewer timeslices) are cheaper to solve but misrepresent
when solar energy is available. The 3dp/12ts case credits more solar energy
to dark hours than the adopted 5dp/20ts fabric.

## Stale capacity factors

If the Ninja rebuild is skipped, the generator does not stop. It remaps the
capacity factors of the previous fabric onto the new dayparts and prints a
remapping note in place of `configs match`. The workbook looks plausible,
and its solar and wind capacity factors are wrong. Run the daypart sync
check and the Ninja rebuilder before every generation.
