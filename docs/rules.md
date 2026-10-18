# Scenario rules

A3 builds each root scenario by running rule scripts in the order that the
Control sheet of the scenario workbook lists them, followed by three shared
rules and, for three of the roots, the base-year pin. A derived scenario is
then built from its root by the
[scenario edits](scenario-edits.md). Terms are in the [glossary](glossary.md).

## The eleven A3 rules

The rule modules live in `ostram/pipeline/scenarios/rules/`. A root's
settings files live in `config/scenarios/<root>/`.

| Rule module | Settings file | What it does |
|---|---|---|
| `set_retirement_schedule.py` | `retirement_schedule.yaml` | Draws down `ResidualCapacity` of thermal generation over time, by age-based lifetimes per fuel and by explicit schedules. The listed exempt patterns are untouched. |
| `build_cap_rule.py` | `build_cap_rule.yaml` | Writes the build cap of every generation technology that has residual capacity or a minimum investment in at least one year, from residual capacity, demand growth and the five steps below. |
| `set_min_capacity_floors.py` | `bau_calibration.yaml` | Writes minimum capacity or activity floors and maximum ceilings listed in the file. On `main` the A file lists none. The C file lists one ceiling, the annual hydro investment of NPLXX. |
| `set_vre_targets.py` | `set_vre_targets.yaml` | C only. Converts fixed renewable generation shares per country-region into activity floors sized from the solved A result, and writes a bounded capacity envelope for solar and onshore wind in place of their build caps. |
| `relax_interconnectors.py` | `relax_interconnectors.yaml` | Opens the build cap of interconnectors above their residual capacity, by a factor, by a fixed headroom or by per-corridor schedules. It sets `TotalAnnualMaxCapacity` of an overridden corridor with new investment to 9999. |
| `add_storage_min_investment.py` | `storage_floors.yaml` | Writes small minimum investments for the storage families per year. |
| `apply_interconnector_costs.py` | the `Interconnector_Params` sheet of the scenario workbook | Overwrites `CapitalCost`, `FixedCost` and `OperationalLife` of the interconnectors with the sheet's values. |
| `apply_internal_transmission.py` | `Config_country_codes.yaml` and `config/scenarios/rules/internal_tx_residuals.csv` | Writes `ResidualCapacity`, `CapitalCost`, `FixedCost` and `OperationalLife` of the six internal transmission families, 60 technologies in all. |
| `apply_internal_tx_losses.py` | `Config_country_codes.yaml` | Sets the output activity ratio of the internal transmission families to 1 minus the loss setting `internal_transmission.transmission_loss`, which is 0.03 on `main`. |
| `apply_base_year_pin.py` | `config/scenarios/rules/pwr_min_2023_2026_pin.csv` | Writes the reviewed 2023 to 2026 values of the table over the matching workbook rows. |
| `set_interconnector_direction.py` | `config/scenarios/<scenario>/set_interconnector_direction.yaml` | Zeroes one flow mode of each listed interconnector, so that power flows one way. The empty map of `B_Opt_DirBidir` changes nothing. |

## Order of the rules

The Control sheet names the first group for each root. The chain runs in
stage 5 of A3, after the workbook preparation stages and after restriction
inheritance.

| Root | Stage 5 chain, in order |
|---|---|
| `BAU` | build cap |
| `A_Calibrated_BAU` | retirement, build cap, minimum capacity floors, relax interconnectors |
| `B_Optimised_VRE` | retirement, build cap, relax interconnectors, storage minimum investment |
| `C_Target_VRE` | retirement, minimum capacity floors, build cap, VRE targets, relax interconnectors, storage minimum investment |

After the chain, every root runs interconnector costs, internal transmission
and internal transmission losses, in that order. The base-year pin then runs
for the roots listed in `PWR_MIN_PIN_ROOT_SCENARIOS` in
`ostram/pipeline/scenarios/orchestrator.py`: `A_Calibrated_BAU`,
`B_Optimised_VRE` and `C_Target_VRE`. Support `BAU` does not run it. A derived
scenario then receives the shared VRE ceilings, its scenario edits and the
base-year pin again. The two direction variants receive their overlay last.

The order is load-bearing in these places.

- The retirement rule runs before the build-cap rule, because the build-cap
  rule reads the `ResidualCapacity` that retirement writes.
- The build-cap rule skips interconnectors. The relaxation rule sets their
  caps afterwards from their residual capacity, and the direction overlay
  runs after the relaxation rule.
- In C the VRE targets run after the build-cap rule. They overwrite the solar
  and onshore wind build caps that the rule wrote with the capacity envelope,
  for each target marked `cap_envelope`. The reason for that position is not
  recorded.
- The build-cap rule reads `TotalAnnualMinCapacityInvestment` and keeps every
  cap above the minimum investment of the same year. In C the minimum
  capacity floors run before it. The reason for that position is not
  recorded.
- The storage minimum investment rule writes `TotalAnnualMinCapacityInvestment`
  for the storage families, which the build-cap rule exempts. The reason for
  its position is not recorded.
- The internal transmission rule runs after the stage 1 template merge,
  because that merge rewrites `OperationalLife`, and the rule has to win.
- The base-year pin runs last, so that reviewed values are the final values.
  A derived scenario applies it again after its edits, so that the shared
  ceilings and the edits cannot replace a reviewed value.

## 0, -1 and 9999

In the final solver datafile, the model applies the build cap unless its
value is -1. The same holds for `TotalAnnualMaxCapacity`, which caps residual
capacity plus surviving new capacity.

| Value in the final datafile | Build cap | Total capacity cap |
|---|---|---|
| 0 | No new capacity in that year. | Zero total capacity in that year. |
| 9999 | Up to 9999 GW of new capacity. A finite cap. | Up to 9999 GW in total. A finite cap. |
| -1 | The constraint is skipped. | The constraint is skipped. |
| absent | The otoole conversion writes -1. | The otoole conversion writes -1. |

A 0 in a workbook does not always reach the datafile as 0. Each stage
treats it as follows.

| Stage | What happens to a 0 |
|---|---|
| A3 stage 1b | `add_max_capacity_investment_rule` with `--fill-zeros` writes 9999 into the blank and 0 build caps of allowed technologies, those with residual capacity or planned additions, and 0 into the rest. |
| A3 build-cap rule | Treats blank, 0 and 9999 as placeholders and replaces them with the computed cap. Storage is exempt and receives 9999 for both caps. |
| A3 base-year pin | Writes the reviewed values, including zeros. |
| B1 | Exports a row only when its `Projection.Mode` is a text other than `EMPTY`, `Zero`, blank or `According to demand`. A 0 on an exported row stays 0. An omitted row receives -1 from the otoole conversion configuration. |
| B2 storage delay | Sets both caps of the 20 storage-linked technologies to 0 in 2023 to 2027 and to -1 from 2028. |
| B2 backstop opening | Turns 0 into 9999 for both caps of `PWRBCK*`. |
| B2 gas and oil allowance | Replaces a build cap of 0 or 9999 of `PWRPET*`, `PWROIL*` and `PWRNGS*` with 0.5 GW per year, and a total capacity cap of 0 with 10 GW. For a technology with a positive minimum investment in any year, it keeps the build caps and still replaces a total capacity cap of 0. A replaced total capacity cap is raised to the higher of the minimum total capacity and the residual capacity plus the minimum investment of the same year, where that value is higher than 10 GW. It restores the values of the base-year pin. |

Read the final datafile before relying on a value:
`workspace/execution/Executables/<S>_0/Pre_processed_<S>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt`.

## Build caps

The build-cap rule sets each cap in five steps, in a fixed order. The
settings are in `config/scenarios/<root>/build_cap_rule.yaml`.

1. Formula. The required growth of a country-region in a year is the increase
   of its demand-scaled fleet from the year before. It is zero when the fleet
   shrinks. Each technology's cap starts as the required growth multiplied by
   the technology's share of the 2024 fleet and by the margin factor.
2. Cap multiplier. The cap is multiplied by the factor for the year. Years
   between the listed years are interpolated.
3. Family limits. A family takes the cap multiplier only up to its family
   limit. A family limit of 1.0 or less locks the family at its 2024 share.
   Families absent from the list take the full cap multiplier.
4. Technology top-up. Any cap of an unlocked family that is below the top-up
   value is raised to that value.
5. Adequacy top-up. When the caps of a country-region add up to less than the
   adequacy share multiplied by the demand-scaled fleet, every cap is
   multiplied by the same factor so that the sum reaches that amount. The
   adequacy share is interpolated between the listed years.

If a cap would be at or below the minimum investment of the same year, the
rule sets it to 1.01 times that minimum. The storage families are exempt:
both of their caps are set to 9999. Interconnectors are skipped.

| Setting | `A_Calibrated_BAU` | `B_Optimised_VRE` | `C_Target_VRE` |
|---|---|---|---|
| Margin factor | 1.1 | 1.1 | 1.1 |
| Cap multiplier | 1.0 in every year | 1.0 to 2027, 3.0 in 2030, 6.0 in 2035, 10.0 in 2040, 15.0 in 2045, 25.0 in 2050 | 1.0 to 2027, 1.1 in 2030, 1.3 in 2035, 1.5 in 2040, 2.0 in 2045, 3.0 in 2050 |
| Family limits | COA, PET, OIL, BIO and WAS 1.0 | COA, BIO and WAS 1.0; PET and OIL 1.01; NGS 1.2; URN 1.5; HYD and SHP 3.0 | COA, BIO and WAS 1.0; PET and OIL 1.01; NGS 1.2; URN 1.5 |
| Technology top-up | 0.0 GW | 0.1 GW | 0.1 GW |
| Adequacy top-up | none | 0.0 to 2040, 0.08 in 2045, 0.20 in 2050 | 0.0 to 2040, 0.08 in 2045, 0.20 in 2050 |
| Exempt families | PWRSDS, PWRLDS | PWRSDS, PWRLDS | PWRSDS, PWRLDS |

In A the cap multiplier stays at 1.0, so the family limits there only decide
which families are locked, and the technology top-up of 0.0 GW raises
nothing. In B and C, solar and wind have no family limit and take the full
cap multiplier.

### The reason for the top-ups

The rule sets the caps before the model runs, so it can only use the fleet
that is in the input data. The formula ties each cap to the growth of the
demand-scaled fleet, and that fleet counts only plants that already exist.
As those plants retire, the fleet grows more slowly, and late in the horizon
the formula allows very little new building even though demand still rises.
With too little room, the model turns to the backstop. The technology top-up
and the adequacy top-up give back the room lost as plants retire. The
adequacy share is 0.0 until 2040, so the calibrated years keep their caps.
The caps set the most that may be built. The model still chooses what it
builds at least cost.

### Other ways to set build caps

These are alternatives for discussion. None of them is part of OSTRAM.

Caps set by hand. A modeller enters one cap for every technology, country-region
and year from plans and expert judgement. The values are transparent, and
each one must be revised when demand, retirements or committed additions
change.

A growth-rate limit inside the model. A constraint lets each year's new
capacity grow by at most a set fraction of the capacity that the model itself
has built or kept in the year before. The model then sees its own fleet, and
the limit becomes part of the optimisation.

Solving repeatedly while updating the fleet. The model is solved, the caps
are recomputed from the solved fleet, and the model is solved again until
the caps stop changing. Each round is a full solve.

Data-based build rates. Caps come from observed construction rates, for
example the largest annual addition of a technology in a comparable system,
scaled to the size of the fleet. They need a dataset of historical additions
for every technology family.

## Gas and oil allowance

Gas and oil plants that would otherwise be blocked may still add up to
0.5 GW per year in each country-region. This lets each one meet its
reserve margin with real plants instead of the emergency backstop. The value
is a modelling choice. You can change it to explore its effect.

The same value also replaces a build cap of 9999 for these technologies.

The value is in `inputs/execution/firm_capacity_fallbacks_by_cr.xlsx`, sheet
`fallbacks`, column `TotalAnnualMaxCapacityInvestment`, one row per
country-region. Column `TotalAnnualMaxCapacity` of the same sheet holds the
10 GW fallback for a total capacity cap of 0. The settings that switch the
step on and select the technologies and the replaced values are in
`config/execution/Config_MOMF_T1_AB.yaml`: `reserve_margin_xlsx_active`,
`reserve_margin_xlsx_target_prefixes` and `reserve_margin_xlsx_sentinel_values`.

## Storage delay

No new storage can be built before 2028. This reflects the years needed to
plan and build storage; existing storage is kept. The five-year delay is a
modelling choice, the same in every scenario. You can change it to explore
its effect.

The settings are in `config/execution/Config_MOMF_T1_AB.yaml`:
`storage_delay_active` switches the step on, `storage_delay_first_n_years`
sets the number of blocked years from the first model year, and
`storage_delay_storage_prefixes` selects the storage families. The output
file names keep the suffix `StorageDelayN5` whatever number of years is set,
because the suffix is the separate setting `storage_delay_suffix`.
