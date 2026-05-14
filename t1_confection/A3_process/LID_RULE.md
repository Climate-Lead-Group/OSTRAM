# OSTRAM SOASIA — MaxCapacityInvestment Lid Rule

This document explains what `add_max_cap_investment_lid_rule.py` does, why
it exists, and how to validate that each of its two modes is working
correctly.

---

## Purpose

The OSTRAM SOASIA OSeMOSYS model has many generation technologies whose
`TotalAnnualMaxCapacityInvestment` (henceforth **MaxInv**) was either left
empty or filled with a placeholder `9999` (effectively unbounded). Without
a real upper bound on annual new build:

- The optimizer can over-invest in candidate techs whose annual ramp should
  be physically constrained (e.g. solar in a region that already has a
  small fleet — building 50 GW in one year is unrealistic).
- The model's results become driven by which tech is cheapest in absolute
  cost, not by realistic deployment trajectories.

This script writes a per-tech, per-year **lid** into the MaxInv cells of
allowed generation techs, leaving manually calibrated values untouched and
ensuring `MinCapInvestment < MaxCapInvestment` for LP feasibility.

It runs independently of `add_max_capacity_investment_rule.py`, can run on
a fresh A-O file or on the output of the first-patch script, and is
idempotent (running it twice produces no further changes).

---

## Common quantities (used by both modes)

For every allowed generation tech `t` in country+region `cr` and year `y`:

```
pool(cr, y)         = sum of ResidualCapacity(t, y) for allowed gen techs in cr
mult(cr, y)         = demand(cr, y) / demand(cr, 2024)
scaled_pool(cr, y)  = pool(cr, y) * mult(cr, y)
```

`pool` is the regional installed-capacity baseline. `mult` is how much
demand has grown since 2024. `scaled_pool` is the demand-adjusted pool —
the proxy for "how big should this region's fleet be in year y?"

`country_region(t)` is extracted from the tech code: chars 6..10 of an
11-character `PWR*` code. `PWRHYDLKAXX` → `LKAXX`. Non-PWR techs (e.g.
`TRN*` interconnects) are skipped entirely.

A tech is **allowed** if it has `ResidualCapacity > 0` or
`TotalAnnualMinCapacityInvestment > 0` in any year, AND it appears as
`GENERATION` in `TECH_TYPES.csv`. Storage techs (`STORAGE_SHORT`,
`STORAGE_LONG`) are not allowed and their MaxInv rows are not touched.

---

## Mode 1: `"uniform"` (default)

### Pseudologic

```
base_pct(y)  = LID_PERCENTAGE_BY_YEAR.get(y, LID_PERCENTAGE_DEFAULT)
pct(cr, y)   = base_pct(y) * mult(cr, y)              if ramp on
             = base_pct(y)                            if ramp off
lid(t, y)    = pct(cr, y) * pool(cr, y)
             = base_pct(y) * scaled_pool(cr, y)        (ramp on)
```

**Every allowed tech in the same `cr` gets the same lid value.** Total
fleet ramp in `cr` = `N_techs × lid`.

### The schedule

`LID_PERCENTAGE_BY_YEAR` is a per-year dict of `base_pct(y)`. The default
schedule encodes "tight near-term, loose late-horizon":

| Years         | base_pct | Rationale                                          |
|---------------|----------|----------------------------------------------------|
| 2023–2030     | 0.5%     | Matches current near-term lid; respects national IRPs which have visibility through ~2030. |
| 2031–2040     | 10%      | Planning data thins past 2030; 20× jump frees the optimizer to substitute for storage if needed. |
| 2041–2050     | 50%      | Effectively unbounded; "we don't know what build rates will be in 2045+, let the model decide." |

`mult` then layers on linearly. So a region whose demand triples by 2050
gets a lid that is 3× what the schedule alone implies. Fast-growing
regions get proportionally more headroom.

### Reading the formula in plain English

> "Each year, every allowed gen tech in the region can grow by `pct%` of
> the demand-adjusted regional fleet, where pct comes from a per-decade
> schedule and the demand adjustment scales with each region's own demand
> growth."

### Numerical example (LKAXX, PWRHYDLKAXX, all years)

LKAXX has `pool ≈ 5.08 GW`, `mult(2024)=1.00`, `mult(2050)=3.37`.

| Year | base_pct | mult  | pct      | lid (per tech, MW/yr) |
|------|----------|-------|----------|-----------------------|
| 2024 | 0.5%     | 1.00  | 0.500%   | 25                    |
| 2025 | 0.5%     | 1.05  | 0.525%   | 27                    |
| 2030 | 0.5%     | 1.36  | 0.682%   | 35                    |
| 2031 | **10%**  | 1.42  | 14.2%    | 720                   |
| 2040 | 10%      | 2.14  | 21.4%    | 1119                  |
| 2041 | **50%**  | 2.21  | 110.7%   | 5877                  |
| 2050 | 50%      | 3.37  | 168.5%   | 8557                  |

The 2030→2031 cliff (35 → 720 MW/yr) and the 2040→2041 cliff (1119 → 5877
MW/yr) are intentional — they encode the loss of planning anchor at decade
boundaries. Steps at year boundaries are well-handled by OSeMOSYS as
MaxInv is a per-year inequality, not a coupled equation.

### When to use uniform mode

- "Let the optimizer pick winners" — you don't want to bias which techs
  grow; you only want to cap the speed of any single tech.
- Late-horizon stress tests — uniform mode with 50% in 2050 is effectively
  unbounded, useful for confirming the LP solves cleanly without storage
  carrying load.
- Default for first-pass validation runs.

### How to validate uniform mode

#### Step A — Run the script in uniform mode

In `add_max_cap_investment_lid_rule.py`, confirm:

```python
LID_RULE_MODE = "uniform"
```

Then from the project root:

```
python add_max_cap_investment_lid_rule.py --input-dir A1_Outputs/A1_Outputs_BAU
```

The print summary should display:

```
Rule mode     : uniform
Lid base pct  : 0.500% (anchored at 2024)
Demand ramp   : ON (per country+region)
Year schedule : 2023-2030:0.5%, 2031-2040:10%, 2041-2050:50%
```

If `Rule mode` says `proportional` instead, you forgot to switch back —
edit the file and re-run.

#### Step B — Spot-check the xlsx output

Open `A-O_Parametrization.xlsx`, sheet `Secondary Techs`. Filter to
`PWRHYDLKAXX`, parameter `TotalAnnualMaxCapacityInvestment`. Check three
cells against the table above:

| Year | Expected (GW) | Tolerance       | If off, suspect:                   |
|------|---------------|-----------------|------------------------------------|
| 2024 | ~0.025        | ±20%            | pool calculation                   |
| 2030 | ~0.035        | ±20%            | mult or pool                       |
| 2050 | ~8.56         | ±5% (very tight)| mode wrong, schedule wrong, or `mult^EXP` leak |

If 2050 reads ~0.086 GW, you're still on the old linear rule (something
prevented the schedule from taking effect). If 2050 reads ~0.529 GW,
you're on Rule D with `exp=2.5` (the previous iteration). If 2050 reads
~8.56 GW, uniform mode is working correctly.

Also spot-check that `PWRWONLKAXX` and `PWRSPVLKAXX` 2050 values are
**identical** to `PWRHYDLKAXX` 2050 — uniform mode applies the same lid
to every gen tech in cr.

#### Step C — Run B2.py / OSeMOSYS

With `PWRSDSLKAXX` (Sri Lanka short-duration storage) disabled, run the
LP. Three possible outcomes:

| Outcome                                       | Meaning                          | Next step                               |
|-----------------------------------------------|----------------------------------|------------------------------------------|
| LP solves, build mix looks reasonable         | Uniform mode is your answer      | Commit. Move to writing up.             |
| LP infeasible / still binds                   | Lid not the bottleneck           | Investigate elsewhere (CPLEX IIS)       |
| LP solves but build mix pathological          | Uniform mode too permissive late | Switch to proportional mode (below)     |

A "pathological" build mix would look like: a region's fleet grows from
5 GW to 200 GW in five years, dominated by one tech. Reasonable looks
like: gradual ramp, multiple techs sharing the build, late-horizon mix
diversified.

---

## Mode 2: `"proportional"`

### Pseudologic

```
tech_share(t)     = ResidualCapacity(t, 2024) / pool(cr(t), 2024)
pool_delta(cr, y) = max(0, scaled_pool(cr, y) - scaled_pool(cr, y-1))
lid(t, y)         = LID_SECURITY_FACTOR * tech_share(t) * pool_delta(cr, y)
```

- `tech_share(t)` is each tech's slice of the 2024 fleet. Shares within a
  cr sum to 1.0. Techs with zero residual at 2024 get share=0 and receive
  no proportional allocation (only whatever MinCapInv provides via untie).
- `pool_delta(cr, y)` is how much the demand-scaled pool grew from y-1 to
  y. Guarded with `max(0, ...)` so flat or declining demand years yield
  zero new headroom. For the earliest year in the data (2023), delta = 0
  (no prior year to delta against).
- `LID_SECURITY_FACTOR` is a slack knob, default 1.1. With security=1.0,
  per-tech lids exactly sum to `pool_delta`. Values >1.0 add slack to
  prevent the optimizer from binding at the strict proportional split.

### Reading the formula in plain English

> "Each year, the regional pool needs to grow by `pool_delta` to match
> demand. Distribute that growth among allowed gen techs in proportion to
> their 2024 fleet share, plus a small slack factor for the optimizer."

### When to use proportional mode

- Once you've confirmed feasibility under uniform mode and want to tighten
  back down with a defensible BAU narrative.
- The story is: "fleet evolves proportionally to current composition" —
  appropriate for BAU because it implies path dependency on existing
  infrastructure, not aggressive technology transitions.
- Note: this rule **constrains the renewable transition** by tying each
  tech's build rate to its 2024 weight share. Solar in LKAXX (~19% of
  2024 pool) gets 19% of pool_delta as its lid each year. If you expect
  the optimizer to pick winners that diverge from current weights, this
  mode will fight that.

### Numerical example (LKAXX, all techs, year 2050)

LKAXX 2050: `scaled_pool(2050) = 17.12 GW`, `scaled_pool(2049) ≈ 16.49 GW`,
so `pool_delta(2050) ≈ 0.628 GW`. Security = 1.1. Sum of all per-tech lids
= 1.1 × 0.628 = 0.691 GW.

| Tech         | Share  | Residual_2024 (GW) | lid_2050 (MW/yr) |
|--------------|--------|--------------------|-------------------|
| PWRHYDLKAXX  | 0.361  | 1.835              | 250               |
| PWROILLKAXX  | 0.230  | 1.168              | 159               |
| PWRSPVLKAXX  | 0.187  | 0.947              | 129               |
| PWRCOALKAXX  | 0.160  | 0.810              | 110               |
| PWRWONLKAXX  | 0.052  | 0.264              | 36                |
| PWRBIOLKAXX  | 0.009  | 0.044              | 6                 |
| PWRWASLKAXX  | 0.002  | 0.010              | 1                 |
| PWRNGSLKAXX  | 0.000  | 0.000              | 0                 |
| **Total**    | 1.000  | 5.078              | **691**           |

Compared to uniform mode in 2050 where every tech gets 8557 MW/yr (and
the fleet ramp is 14 × 8557 ≈ 120 GW/yr), proportional mode is roughly
**170× tighter** in total fleet ramp. This is the deliberate trade-off:
proportional is the disciplined narrative, uniform is the headroom valve.

### How to validate proportional mode

#### Step A — Switch modes

In `add_max_cap_investment_lid_rule.py`, change the mode constant:

```python
LID_RULE_MODE = "proportional"
```

Optionally adjust:

```python
LID_SECURITY_FACTOR = 1.1   # 1.0 = strict proportional, 1.5+ = generous
```

#### Step B — Run the script

```
python add_max_cap_investment_lid_rule.py --input-dir A1_Outputs/A1_Outputs_BAU
```

Print summary should display:

```
Rule mode     : proportional  (security factor = 1.1)
Demand ramp   : ON (per country+region)
```

There is no `Year schedule` line in proportional mode — the schedule is
ignored.

#### Step C — Spot-check the xlsx output

The check pattern is different from uniform mode because **per-tech lids
differ within a cr**. Open `A-O_Parametrization.xlsx`, sheet
`Secondary Techs`. Three checks:

1. **`PWRHYDLKAXX` 2050 ≠ `PWRWONLKAXX` 2050.** In uniform mode they
   would be identical. In proportional mode, hyd's lid should be roughly
   7× won's lid (the share ratio: 0.361 / 0.052 ≈ 7).

2. **Sum of LKAXX 2050 lids ≈ `LID_SECURITY_FACTOR × pool_delta(LKAXX, 2050)`.**
   Sum the MaxInv 2050 cells across all 8 LKAXX gen techs. Should be
   ~0.69 GW (with security=1.1) or ~0.63 GW (with security=1.0).

3. **`PWRHYDLKAXX` 2024 = 0.** Reference year has no prior year to delta
   against, so the lid is 0. The untie rule should not have fired here
   because PWRHYDLKAXX has no MinCapInv at 2024. If you see a non-zero
   value, suspect the untie rule or the placeholder-replacement path.

Also check that **shares sum to 1.0 per cr** by reading the JSON change
log (`*_CHANGES.json` in the backup folder's parent). Look for
`"tech_share": {...}` and group by cr — they should sum to 1.0 (or 0.0
for crs whose allowed techs all have zero 2024 residual, an edge case).

#### Step D — Run B2.py / OSeMOSYS

With `PWRSDSLKAXX` disabled, run the LP. Possible outcomes:

| Outcome                              | Meaning                                | Next step                                  |
|--------------------------------------|----------------------------------------|---------------------------------------------|
| LP solves, build mix proportional     | Proportional mode is working          | Compare against uniform mode results       |
| LP infeasible                         | Proportional mode is too tight        | Increase `LID_SECURITY_FACTOR` to 1.25 / 1.5 |
| LP solves but renewable build is shut down | Proportional rigidly mirrors 2024 mix | Expected behavior; this is the BAU narrative |

If the LP infeasibility is in late-horizon years specifically, that's a
sign proportional mode's pool_delta is too small to absorb required
ramps. Increasing security factor is the first response. If even
security=2.0 doesn't unlock it, the bottleneck is elsewhere.

---

## Configuration reference

All config lives at the top of `add_max_cap_investment_lid_rule.py` (no
YAML, no CLI args for these knobs). To change:

1. Open the file in your editor
2. Edit the constant
3. Save
4. Re-run the script

| Constant                   | Default       | Used by         | What it controls                                    |
|----------------------------|---------------|-----------------|-----------------------------------------------------|
| `LID_RULE_MODE`            | `"uniform"`   | both            | Which formula to apply                              |
| `LID_PERCENTAGE_DEFAULT`   | `0.005`       | uniform         | Fallback base_pct for years not in the schedule     |
| `LID_PERCENTAGE_BY_YEAR`   | 0.5/10/50%    | uniform         | Per-year base_pct schedule                          |
| `LID_SECURITY_FACTOR`      | `1.1`         | proportional    | Slack on the proportional split                     |
| `LID_RAMP_FROM_DEMAND`     | `True`        | both            | Whether to apply mult; off = ignore demand growth   |
| `DEMAND_REFERENCE_YEAR`    | `2024`        | both            | Year against which mult is computed                 |
| `RESTRICT_TO_GENERATION`   | `True`        | both            | Whether to require GENERATION category from TECH_TYPES.csv |

---

## V1 untie rule (applies to both modes)

After computing the proposed lid in either mode, the untie rule fires:

```
if MinCapInv(t, y) > 0 and proposed_lid <= MinCapInv(t, y):
    proposed_lid = MinCapInv(t, y) * UNTIE_MULTIPLIER     # default 1.01
```

This guarantees the LP-feasibility invariant `MinCapInv < MaxCapInv`
whenever MinCapInv > 0. The untie reason is logged separately in the
change log so you can distinguish lid-set values from floor-corrected
values.

---

## Test suite

`test_add_max_cap_investment_lid_rule.py` contains 51 tests covering both
modes. Run with:

```
pytest test_add_max_cap_investment_lid_rule.py -v
```

Expected: 51 passed, ~3 minutes runtime.

Notable test classes:

- `TestLidMath` — uniform-mode formula and basic invariants
- `TestPerYearOverride` — schedule entries declare base_pct, mult layers on
- `TestRampDisable` — ramp off → schedule still applies, no per-cr variation
- `TestProportionalMode` — proportional formula, share/delta/security checks

If you change the schedule in `LID_PERCENTAGE_BY_YEAR` or the security
factor, tests should still pass because they re-derive expected values
through the production helpers rather than hard-coding numbers.

---

## Provenance & related files

- **`add_max_cap_investment_lid_rule.py`** — production script
- **`test_add_max_cap_investment_lid_rule.py`** — test suite
- **`TECH_TYPES.csv`** — GENERATION classification, sits next to the script
- **`A-O_Parametrization.xlsx`** — input/output (script edits in place)
- **`A-O_Demand.xlsx`** — read for demand multipliers
- **`measure_lid.py` + `lid_measurement.xlsx`** — analytical workbench used
  to design the rule; not part of the production pipeline

Each run of the production script writes a backup folder (`*_PRE_LID_*`)
and a JSON change log (`*_CHANGES.json`) that records every cell modified
along with the reason (`lid_fill`, `untie_min_inv`, or `preserved_manual`).
