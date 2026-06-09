# OSTRAM Multi-Scenario User Guide

This guide explains how to build, configure and run multiple energy
scenarios in OSTRAM from a single Excel workbook. It targets the
**end user** — modelers who configure runs via Excel and a single
terminal command, without touching Python code.

If you only ever want to run BAU exactly like before, you can keep
doing that — multi-scenario support is additive, BAU still works
identically.

---

## Table of contents

1. [What multi-scenario gives you](#what-multi-scenario-gives-you)
2. [The big picture](#the-big-picture)
3. [The SOASIA Template v18 workbook](#the-soasia-template-v18-workbook)
   - [Sheet `Control`](#sheet-control)
   - [Parametric sheets and the `scenario` column](#parametric-sheets-and-the-scenario-column)
   - [Sheet `Restrictions`](#sheet-restrictions)
4. [Walkthrough: add a new scenario](#walkthrough-add-a-new-scenario)
5. [Running the pipeline](#running-the-pipeline)
6. [Common errors and how to fix them](#common-errors-and-how-to-fix-them)
7. [Reference: identity keys per sheet](#reference-identity-keys-per-sheet)
8. [Reference: rules scripts](#reference-rules-scripts)

---

## What multi-scenario gives you

Before: A3 only processed one scenario (BAU). Any "what if" required
duplicating files and editing Python.

Now: you can define many scenarios — `BAU`, `NDC`, `LMC`, etc. — from
the **`Control`** sheet of `SOASIA_OSeMOSYS_Template_v18.xlsx`. For
each scenario you choose:

- Whether it runs in the next pipeline pass (`active = TRUE`)
- Which rules script writes its restrictions (`rules_script`)
- Whether it inherits restrictions from other scenarios that already
  ran (`inherit_restrictions_from`)

You also tell SOASIA, sheet by sheet, which values to **override**
relative to BAU. Anything you don't override stays equal to BAU, so a
new scenario is cheap to set up.

The pipeline (`python run.py`) then runs A3 once per active scenario
in the correct dependency order, producing `A1_Outputs_<scenario>/`
folders that B1 and B2 already know how to consume.

---

## The big picture

```
                 SOASIA_OSeMOSYS_Template_v18.xlsx
                 +--------------------------------+
                 | [Control]   scenarios + rules  |
                 | [README]                       |
                 | [Restrictions]  persisted out  |
                 | [Primary_Techs]    + overrides |
                 | [Secondary_Techs]  + overrides |
                 | ... (15 parametric sheets)     |
                 +----------------+---------------+
                                  |
                                  v
                    +------------------------------+
                    | run.py                       |
                    | 1. A1 + A2  (BAU only, once) |
                    | 2. For each active scenario: |
                    |     - materialize v18 for it |
                    |     - run A3 stages 0..6     |
                    |       (incl. inheritance and |
                    |        rules_script)         |
                    | 3. B1 + B2 over outputs      |
                    +------------------------------+
                                  |
                                  v
              A1_Outputs/A1_Outputs_BAU/
              A1_Outputs/A1_Outputs_NDC/
              A1_Outputs/A1_Outputs_LMC/
                 ...
```

A2 creates a single snapshot, `_post_a2_snapshot_BAU/`, that ALL
scenarios start from. Divergence between scenarios is expressed only
through:

1. SOASIA `Control` (which rules_script and which inheritances)
2. SOASIA parametric sheets (override rows)
3. SOASIA `Restrictions` (per-scenario restriction values)

This keeps A1/A2 simple and BAU-only.

---

## The SOASIA Template v18 workbook

Path: `t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx`

The first sheets are configuration:

- `Control` — scenario list and per-scenario knobs
- `README` — short in-workbook help
- `Restrictions` — auto-written by the pipeline (you can also edit)

All remaining sheets are data:

- `Yearsplit_Template`, `DaySplit` — timeslice definitions, same for
  every scenario (no `scenario` column)
- 15 **parametric sheets** — each gains a `scenario` column as
  column A

### Sheet `Control`

```
| scenario | active | rules_script                          | inherit_restrictions_from | notes         |
|----------|--------|---------------------------------------|---------------------------|---------------|
| BAU      | TRUE   | add_max_cap_investment_lid_rule.py    |                           | Base scenario |
| NDC      | TRUE   | add_max_cap_investment_lid_rule.py    | BAU                       | example       |
| LMC      | FALSE  | add_max_cap_investment_lid_rule.py    | BAU, NDC                  | example       |
```

Columns:

- **scenario** — any unique identifier without spaces, e.g. `BAU`,
  `NDC`, `LMC`, `NetZero2050`.
- **active** — `TRUE` or `FALSE`. Only `TRUE` scenarios run in the
  next pass. A dropdown enforces the value.
- **rules_script** — file name of a `.py` script under
  `t1_confection/A3_process/rules_scripts/`. The script applies that
  scenario's restrictions (today: a "lid" on
  `TotalAnnualMaxCapacityInvestment`). Leaving the cell blank skips
  that stage entirely.
- **inherit_restrictions_from** — comma-separated list of OTHER
  scenarios whose `Restrictions` rows should be loaded before the
  scenario's own `rules_script` runs. Useful for "start from BAU's
  lid, then tighten." Leave empty if the scenario writes its own
  restrictions from scratch.
- **notes** — free-form text, ignored by the pipeline.

**Inheritance ordering.** When `inherit_restrictions_from = BAU, NDC`
and both bring values for the same cell, the source listed LAST wins
(NDC overrides BAU here). When you list a scenario in
`inherit_restrictions_from`, that scenario must have already produced
rows in `Restrictions` — typically because it ran in a previous
pass. The pipeline also topologically orders active scenarios per pass,
so if NDC inherits from BAU and both are active, BAU runs first.

### Parametric sheets and the `scenario` column

The 15 parametric sheets (`Primary_Techs`, `Secondary_Techs`,
`Capacities_CF`, `VariableCost`, `Demand_Projection`,
`Demand_Profiles`, `Demand_Techs`, `Emissions`, `Interconnectors`,
`Interconnector_Params`, `Fixed_Horizon_Parameters`,
`Existing_Generation`, `Planned_Generation`, `Technology_Costs`,
`RE_Targets_Policies`) each have `scenario` as their **first column**.

The override model: a row tagged `scenario = "BAU"` is the **base**.
A row tagged with another scenario (say `NDC`) is an **override** —
if it shares its identity key with a BAU row, it REPLACES that row
when the NDC scenario runs. If the identity key is new, the row is
ADDED to NDC.

What counts as an "identity key" is sheet-specific — see
[Reference: identity keys per sheet](#reference-identity-keys-per-sheet).

You do **not** copy every BAU row when you create NDC. You only add
the rows you want to change.

The `scenario` column has a dropdown that lists scenario names
sourced from `Control` (positions A2:A101), so you can't typo a
scenario name as long as you've declared it in `Control` first.

### Sheet `Restrictions`

Auto-managed by the pipeline. Each time A3 runs scenario X, the
rules_script's per-cell changes (today: `TotalAnnualMaxCapacityInvestment`
values for generation techs by year) are persisted as rows here:

```
| scenario | source_sheet     | tech         | parameter                            | year | value | rule_applied | source_run_timestamp |
|----------|------------------|--------------|--------------------------------------|------|-------|--------------|-----------------------|
| BAU      | Secondary Techs  | PWRHYDLKAXX  | TotalAnnualMaxCapacityInvestment     | 2024 | 25    | lid_fill     | 2026-05-15T14:30:00   |
| BAU      | Secondary Techs  | PWRHYDLKAXX  | TotalAnnualMaxCapacityInvestment     | 2030 | 35    | lid_fill     | 2026-05-15T14:30:00   |
| BAU      | Secondary Techs  | PWRSPVLKAXX  | TotalAnnualMaxCapacityInvestment     | 2050 | 8557  | lid_fill     | 2026-05-15T14:30:00   |
```

You can **edit** these rows manually — useful when you want to fix
one value without re-running the rules script. Edits to a scenario X
row survive until X is re-run, in which case ALL rows for scenario X
get rewritten (other scenarios' rows are untouched).

When another scenario lists X under `inherit_restrictions_from`, the
pipeline reads these rows and writes the values into that scenario's
`A-O_Parametrization.xlsx` before the scenario's own rules_script
runs on top.

---

## Walkthrough: add a new scenario

Goal: create an "NDC" scenario that starts from BAU, then tightens
the lid using the same rules script.

1. Open `SOASIA_OSeMOSYS_Template_v18.xlsx`.

2. Go to the `Control` sheet, add a new row:

   ```
   | scenario | active | rules_script                          | inherit_restrictions_from | notes                       |
   | NDC      | TRUE   | add_max_cap_investment_lid_rule.py    | BAU                       | Updated NDC commitments     |
   ```

3. (Optional) On any parametric sheet, add override rows for NDC.
   Example — push 2050 demand for Bangladesh up by 1.5x relative to
   BAU. Go to `Demand_Projection`, find the BAU row with
   `Fuel/Tech = ELCBGDXX03`, copy it, and in the new row set
   `scenario = NDC` and change only the 2050 cell. The identity key
   for this sheet is `Fuel/Tech`, so the rest of the year columns
   inherit from BAU automatically (anything you put in the override
   row replaces BAU for that cell).

4. Save and close the file.

5. From the repo root:

   ```
   python run.py
   ```

   You'll see "Active scenarios (topo order): ['BAU', 'NDC']", A3
   running twice, then B1 and B2.

6. After the run, open `Restrictions` in v18 — you'll see new rows
   tagged `scenario = BAU` AND `scenario = NDC`.

---

## Running the pipeline

From the repo root:

```
python run.py
```

This is the only command end users normally need. It:

1. Creates/updates the conda environment from `environment.yaml`.
2. Runs A1 + A2 once for BAU if no `_post_a2_snapshot_BAU/` exists.
3. Reads the `Control` sheet of SOASIA v18 to discover active
   scenarios.
4. Runs A3 once per active scenario, in topological order
   (dependencies first).
5. Runs B1 (compiler) and B2 (solver).

Skip individual stages with `--skip-a3`, `--skip-b1`, `--skip-b2`.
Force a fresh DVC pull is the default if a DVC remote is set.

To run a single scenario without touching `Control`:

```
python t1_confection/A3_process.py --scenario NDC
```

This bypasses run.py's loop and runs A3 for just that scenario, as
long as it's defined in `Control`.

---

## Common errors and how to fix them

**`ERROR: scenario 'X' not in Control sheet of SOASIA_OSeMOSYS_Template_v18.xlsx`**
You passed a scenario name that's not declared in the workbook. Open
v18, add the scenario to `Control`, save.

**`ERROR: SOASIA v18 not found at <path>`**
The workbook is missing. If you have a v17 and just need to migrate,
run `python t1_confection/A3_process/_build_v18_from_v17.py`.

**`ERROR: scenario 'NDC' inherits from unknown scenario 'XYZ'`**
The `inherit_restrictions_from` cell lists a scenario that isn't in
`Control`. Either add it, or remove it from the list.

**`Inheritance cycle detected at scenario 'A'`**
You have A inheriting from B inheriting from A. Break the cycle in
the `inherit_restrictions_from` columns.

**`No Restrictions rows found for scenario(s): ['BAU']`**
You're trying to inherit from a scenario that has never run (so its
restrictions were never persisted). Run that scenario first, or
remove it from `inherit_restrictions_from` until it has run.

**`rules_script 'foo.py' not found at rules_scripts/foo.py`**
You typed a script name that doesn't exist. Either drop the script
into `t1_confection/A3_process/rules_scripts/` or change the cell
to a script that does exist.

**`ERROR: snapshot post-A2 not found: A1_Outputs/_post_a2_snapshot_BAU`**
A1 + A2 have not run yet. Run `python run.py` once without
`--skip-a3` — it will run A1+A2 automatically the first time.

---

## Reference: identity keys per sheet

When a non-BAU row "overrides" a BAU row, the override is matched by
the columns below. Choose the same values in those columns to override;
choose different values to add.

| Sheet                       | Identity key columns                                    |
|-----------------------------|---------------------------------------------------------|
| Fixed_Horizon_Parameters    | Tech, Parameter                                         |
| Primary_Techs               | Tech, Parameter                                         |
| Secondary_Techs             | Tech, Parameter                                         |
| Capacities_CF               | Timeslices, Tech.ID, Parameter                          |
| VariableCost                | Mode.Operation, Tech, Parameter                         |
| Demand_Projection           | Fuel/Tech                                               |
| Demand_Profiles             | Timeslices, Fuel/Tech                                   |
| Demand_Techs                | Tech, Parameter                                         |
| Emissions                   | Tech, Parameter                                         |
| Interconnectors             | NO                                                      |
| Interconnector_Params       | Tech, Parameter                                         |
| Existing_Generation         | Country, Node, Plant_Name, Commissioning_Year, Capacity_MW |
| Planned_Generation          | Country, Node, Project_Name, Expected_COD, Capacity_MW  |
| Technology_Costs            | Technology_Code, Parameter                              |
| RE_Targets_Policies         | Country, Policy_Name                                    |

A few non-obvious notes:

- `Capacities_CF` uses `Tech.ID` (not just `Tech`) because a single
  tech code like `PWRHYDNPLXX` covers 3 distinct units (Reservoir /
  Run-of-River / Pumped Storage) with different CF profiles.
- `Existing_Generation` / `Planned_Generation` keys include
  `Commissioning_Year` / `Expected_COD` and `Capacity_MW` because
  plant lists in v17 contain reused names for different units (e.g.
  "Wind (various)" appears 10 times in Sri Lanka planned).

---

## Reference: rules scripts

Path: `t1_confection/A3_process/rules_scripts/`

Each `.py` file in this folder is a "rules script" — it gets invoked
in stage 5 of A3 against the scenario's working
`A-O_Parametrization.xlsx` and writes restriction values into the
`TotalAnnualMaxCapacityInvestment` cells.

Today, the canonical script is `add_max_cap_investment_lid_rule.py`
(documented in detail in `LID_RULE.md`). To create a variant for a
different scenario:

1. Copy the file in the same folder, rename it (e.g.
   `add_max_cap_investment_lid_rule_v3.py`).
2. Edit the constants near the top (`LID_RULE_MODE`,
   `LID_PERCENTAGE_BY_YEAR`, `LID_SECURITY_FACTOR`, …) to match the
   new scenario's narrative.
3. In SOASIA `Control`, set the scenario's `rules_script` cell to
   the new file name.

Type the filename in the `rules_script` cell. Run-time validation
(`_scenarios.py`) checks the file exists in `rules_scripts/` and
raises a clear error listing available scripts if you mistype.
