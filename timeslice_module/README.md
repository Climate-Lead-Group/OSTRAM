# OSTRAM Timeslice Module

## What this is

This module contains the OSTRAM timeslice workbook generator. The
adopted case is the 5-daypart, 20-timeslice fabric.

## Folder layout

- `scripts/` — the generator, the Ninja rebuilder and the daypart sync check
- `inputs/` — source data (hourly demand, capacity factors, Ninja raw CSVs)
- `outputs/` — `OSTRAM_Timeslice_Outputs_REFERENCE_5dp20ts.xlsx`, the workbook of the adopted fabric

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

The model uses the 5-daypart scheme to stay within the solver's variable and
constraint budget. The model runs 20 timeslices.

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

## Stale capacity factors

If the Ninja rebuild is skipped, the generator does not stop. It remaps the
capacity factors of the previous fabric onto the new dayparts and prints a
remapping note in place of `configs match`. The workbook looks plausible,
and its solar and wind capacity factors are wrong. Run the daypart sync
check and the Ninja rebuilder before every generation.
