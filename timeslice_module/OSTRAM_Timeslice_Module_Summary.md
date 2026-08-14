# OSTRAM timeslice module — what was done and how to use it

## Background

Zixuan asked how to extend or change the timeslice resolution in OSTRAM
(e.g. 6-hour or 8-hour daily blocks). The timeslice generation capability
existed but was buried in a crowded, untracked working folder with stale
snapshots, scattered sensitivity outputs, and no clear entry point. This
work extracted it into a clean, portable module with pre-computed workbook
cases for training.

## What was done

The work was completed across four automated sessions, each with full
evidence logs preserved in `_session_logs/`.

**Session 1 — migration and structure.** A migration script
(`ostram_timeslice_migrate.py`) extracted the minimal file set from
`asia_ostram\asia_ostram_data` into a clean `timeslice_module` folder on
the Desktop. Seven files copied and hash-verified; stale FREEZE snapshots
deliberately excluded; ~150 sensitivity diagnostic PNGs left behind. The
source tree was not modified.

**Session 2 — input wiring and parity.** The generator's input paths were
documented. The ~10 genuine input data files (hourly demand, capacity
factors, Renewables Ninja raw CSVs) were copied into `inputs/`. A
discovery: the generator writes into the source tree by default, which is
how the original 16-timeslice workbook was lost. Patched with two path
constants in the module copy. The generator was then run unchanged and
cell-compared against the adopted workbook: 23 sheets, 15,637 cells, zero
differences. This proved the module is portable and the workbook hash
discrepancy was a benign re-save. The 4dp/16ts variant (6-hour blocks) was
produced as a first test of the variant recipe.

**Session 3 — the remaining two variants.** The 3dp/12ts (8-hour blocks)
and 6dp/24ts (the sensitivity sweep winner) were computed, each following
the proven recipe. Key finding: the Ninja capacity-factor rebuild is
mandatory per fabric because the generator silently accepts stale CFs from
a previous fabric without error. A `check_daypart_sync.py` gate script was
created to enforce this. Measured build time: ~45 seconds per fabric
regardless of timeslice count.

**Session 4 — cleanup.** Four copies of a 648 MB intermediate file deleted
(module dropped from 2.6 GB to 47 MB). The two-line output-path patch
applied to the upstream generator in the source tree (the one permitted
source modification, closing the overwrite hazard permanently). All
forensic session files moved to `_session_logs/`. A clean README written.
The OneDrive adoption-rationale document copied into `docs/`.

## Module layout

```
timeslice_module/
    README.md              what this is and how to use it
    scripts/               the generator and supporting tools
    inputs/                source data (hourly demand, capacity factors)
    outputs/               the four pre-computed workbook cases
    docs/                  handover, ranking justification, adoption rationale
    _session_logs/         build evidence (not needed for normal use)
```

## The four fabric cases

All four workbooks sit in `outputs/`. Each is a complete set of
timeslice parameters (YearSplit, demand profiles, capacity factors) ready
to be wired into the SOASIA template.

| Workbook | Fabric | Timeslices | Daypart boundaries | Status |
|---|---|---|---|---|
| `..._3dp12ts.xlsx` | 3 equal 8-hour blocks | 12 | 00–08 / 08–16 / 16–24 | Computed |
| `..._4dp16ts.xlsx` | 4 equal 6-hour blocks | 16 | 00–06 / 06–12 / 12–18 / 18–24 | Computed |
| `..._REFERENCE_5dp20ts.xlsx` | 5 dayparts (adopted) | 20 | 00–06 / 06–17 / 17–20 / 20–22 / 22–24 | Canonical |
| `..._6dp24ts.xlsx` | 6 dayparts (sweep winner) | 24 | 00–05 / 05–08 / 08–17 / 17–20 / 20–22 / 22–24 | Computed |

The 5dp/20ts fabric was adopted for the model. The 6dp/24ts fabric won
the unconstrained sensitivity sweep but was rejected because solver size
scales with timeslice count (justification in `docs/ranking_by_budget.csv`).

## The trade-off in one table

| Fabric | Timeslices | Solar block | Mean solar CF | Phantom solar | Build time |
|---|---|---|---|---|---|
| 3dp/12ts | 12 | D2 08–16 | 0.491 | 6.75% | ~41 s |
| 4dp/16ts | 16 | D2 06–12 | 0.402 | 0.15% | ~40 s |
| 5dp/20ts | 20 | D2 06–17 | 0.395 | 0.42% | ~45 s |
| 6dp/24ts | 24 | D3 08–17 | 0.450 | 0.28% | ~46 s |

**Phantom solar** is the fraction of solar energy a flat-CF block credits
to hours when the sun is down. The 3dp fabric credits 16 times more solar
energy to dark hours than the adopted fabric, meaning a solver could
dispatch PV capacity at 3 a.m. This is the main reason coarse fabrics are
problematic, not because they dim the solar day, but because they invent
solar at night.

Build time does not depend meaningfully on timeslice count. The cost of
finer fabrics is paid at solve time, not generation time.

## How to generate a new variant

**Prerequisites:** Python with pandas, openpyxl, and numpy. The existing
`OSTRAM-env` conda environment has all three. No additional installs
needed.

**Before running, set the console encoding** (Anaconda Prompt):

```
chcp 65001
set PYTHONIOENCODING=utf-8
```

**Recipe** (five steps):

1. Open `scripts/build_ostram_timeslices.py` and
   `scripts/rebuild_reninja_timeslices_latest.py`. Edit `DAYPART_DEF` in
   both files to the desired daypart boundaries. The two definitions must
   match exactly.

2. Run the Ninja rebuilder:
   ```
   cd timeslice_module
   python scripts/rebuild_reninja_timeslices_latest.py
   ```
   Confirm the console prints a message containing "configs match". If it
   does not, the capacity factors will be wrong. Do not proceed without
   this confirmation.

3. Run the generator:
   ```
   python scripts/build_ostram_timeslices.py
   ```

4. The new workbook appears in `outputs/`. Rename it with a fabric suffix
   (e.g. `OSTRAM_Timeslice_Outputs_3dp12ts.xlsx`) to avoid overwriting
   existing cases.

5. Revert `DAYPART_DEF` in both files to the adopted 5dp/20ts values if
   the module should remain in its canonical state.

Each variant takes roughly 45 seconds to generate.

## Known limitation

If the Ninja rebuild (step 2) is skipped, the generator silently uses
stale capacity factors from a previous fabric. There is no error, no
warning. The workbook will look plausible but the solar and wind CFs will
be wrong. Always rebuild before generating. The module includes a gate
script (`scripts/check_daypart_sync.py`) that verifies the two files
agree; run it if in doubt.

## What this does NOT do

This module generates the timeslice input workbook. It does not run the
energy system model. Wiring a new fabric into the SOASIA template,
adjusting the model's set definitions to match the new timeslice count,
and paying the solver cost of a finer or coarser resolution are separate
steps on the modeller's side. The pre-computed workbook cases are ready
to be taken into that process.
