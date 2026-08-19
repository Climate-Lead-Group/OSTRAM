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

## The four fabric cases

| Workbook file | Fabric | Timeslices | Status |
|---|---|---|---|
| `OSTRAM_Timeslice_Outputs_REFERENCE_5dp20ts.xlsx` | 5 dayparts (adopted) | 20 | Canonical |
| `OSTRAM_Timeslice_Outputs_4dp16ts.xlsx` | 4 equal 6-hour blocks | 16 | Computed |
| `OSTRAM_Timeslice_Outputs_3dp12ts.xlsx` | 3 equal 8-hour blocks | 12 | Computed |
| `OSTRAM_Timeslice_Outputs_6dp24ts.xlsx` | 6 dayparts (sweep winner) | 24 | Computed |

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
when solar energy is available. The 3dp/12ts case credits 16 times more
solar energy to dark hours than the adopted 5dp/20ts fabric.

## Known limitation

If the Ninja rebuild is skipped, the generator silently uses stale
capacity factors from a previous fabric — always rebuild before generating.
