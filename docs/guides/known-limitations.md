# Known limitations

The South Asia model of OSTRAM compares power generation and cross-border
transmission choices for ten nodes in six countries from 2023 to 2050. Public
data do not give every value the model needs. Where a source stops short, the
model makes a stated assumption. This page lists those places: what the sources
leave open, what the model assumes and why, how far it can move results, and
how to test or improve it. Read results with these items in mind.

## Items that can change results

### India's generation floors in 2023 and 2024

**The data gap.** CEA publishes India's regional demand as the net requirement
at power station bus bars, and its generation as gross generation, which
includes the plants' own use. The two series sit on different boundaries.

**What we assumed and why.** The model takes each series as CEA publishes it.
Demand is net of the plants' own use, and the 2023 and 2024 generation floors
are gross. Each value then stays traceable to its published figure.

**Effect on results.** In 2023 and 2024 the model generates more than India's
demand needs, by about the plants' own use. To compare with statistics, compare
model generation with gross generation.

**How to test or improve.** Put the floors and the demand on one basis: either
the floors as net generation, or the India rows of `Demand_Projection` as gross
requirement. The floors are base-year values; their table is named in
[Where the base-year values come from](../reference/base_year_value_sources.md).

### India's early-year demand projections

**The data gap.** The model needs India's demand by region for every year from
2023. CEA's actual requirements for the early years differ from its projections
region by region, in both directions. The actual figures also leave out
inter-state losses, so they sit on a different boundary.

**What we assumed and why.** India's demand follows the fiscal-year projections
of CEA's 20th Electric Power Survey in every year, including the early years.
The whole series then rests on one source and one boundary.

**Effect on results.** Early-year India demand is a projection. Where a region's
projection is above its actual requirement, the model generates or imports more
there than the region did, and less where it is below. Part of any gap is the
difference of boundary.

**How to test or improve.** Replace the early-year India rows of
`Demand_Projection` with CEA's actual requirements, with inter-state losses
added back so the boundary matches.

### Bangladesh generation in 2024

**The data gap.** BPDB publishes Bangladesh's generation by fiscal year, July
to June. The model's years are calendar years, and its Bangladesh demand is a
calendar-year estimate.

**What we assumed and why.** Each model year takes the fiscal year ending in it.
ASSUMPTION: in 2024, each Bangladesh fuel must generate at least 98 per cent of
BPDB's fiscal-year figure. The 2 per cent allows for rounding and reporting
differences in the published figures
([Where the base-year values come from](../reference/base_year_value_sources.md),
A1 and A3).

**Effect on results.** In 2024 the floors, not demand, set Bangladesh's
generation. Read Bangladesh 2024 generation as set by the floors.

**How to test or improve.** Compare Bangladesh 2024 generation with BPDB's
fiscal-year figures. To test another floor, change the Bangladesh 2024 rows of
the base-year value table named in
[Where the base-year values come from](../reference/base_year_value_sources.md).

### Bhutan hydro availability from 2027

**The data gap.** Bhutan's statistics give what each hydro plant produced. That
output already reflects outages and maintenance. The statistics give no
separate availability.

**What we assumed and why.** From 2027, Bhutan's hydro capacity factors come
from observed output, and an availability of 0.95 applies on top. New Bhutan
hydro capacity carries the same factors. In 2023 to 2026 the base-year capacity
factors divide observed output by 0.95, so the two cancel and existing plants
can produce the published amount
([Where the base-year values come from](../reference/base_year_value_sources.md),
A5).

**Effect on results.** From 2027, the most Bhutan's hydro fleet can produce is
lower than its observed factors allow. This lowers what Bhutan can generate and
export.

**How to test or improve.** Set the `AvailabilityFactor` of `PWRHYDBTNXX` to 1
from 2027, or divide the capacity factors from 2027 by 0.95, as in the base
years.

### Bangladesh hydro output from 2025

**The data gap.** BPDB publishes Bangladesh's hydro output, not the plant's
availability.

**What we assumed and why.** Bangladesh's hydro capacity factors divide
observed output by an availability of 0.9. The model's availability for
Bangladesh hydro is 1, and from 2025 Bangladesh has no generation upper limit
([Where the base-year values come from](../reference/base_year_value_sources.md),
A5 and A16).

**Effect on results.** From 2025, Bangladesh's existing hydro can produce up to
11 per cent more than BPDB reports.

**How to test or improve.** Compare Bangladesh hydro output from 2025 with
BPDB's figures. An `AvailabilityFactor` of 0.9 for `PWRHYDBGDXX` removes the
difference.

### Maldives demand, 2023 to 2025

**The data gap.** Published Maldives figures for 2023 to 2025 are on a utility
basis. They leave out the electricity that resorts generate for themselves.

**What we assumed and why.** ASSUMPTION: Maldives demand, whole-country
consumption including resort self-generation, is 6.770 PJ in 2023, 7.096 PJ in
2024 and 7.444 PJ in 2025. The values are estimated backward from the
roadmap's anchor of 2,400 GWh in 2028 (MCCEE,
[Road Map for the Energy Sector 2024-2033](https://media.environment.gov.mv/20241115-pub-energy-roadmap-maldives-2024-2033.pdf),
Figure 8, p. 41) at the roadmap's growth rates by segment (pp. 24 to 26):
Greater Malé 7.9 per cent, other inhabited islands 9.0 per cent and resorts 2.0
per cent a year. The roadmap gives a whole-country figure and segment growth
rates, so the model works back from them. Utility-basis figures are much lower.

**Effect on results.** Maldives demand in 2023 to 2025 includes resorts, so the
Maldives generates more in the model than utility statistics show.

**How to test or improve.** Change the Maldives row of `Demand_Projection` for
2023 to 2025. A possible future calibration refinement replaces the backward
estimate with observed whole-country consumption, including resorts, when a
statistical series covers it.

### Life of existing solar

**The data gap.** The published year-end tables give the installed solar stock,
not the year each plant entered service. The age of the older stock is open.

**What we assumed and why.** ASSUMPTION: India's solar stock of 2022 and
earlier (22.76 GW North, 25.04 GW South) and Sri Lanka's stock of 2023 and
earlier (0.95 GW) stay in service to 2050, while other solar retires after 25
years. Without commissioning years, a 25-year retirement path for this stock
would rest on a guessed age, so the model keeps it whole
([Where the base-year values come from](../reference/base_year_value_sources.md),
A11).

**Effect on results.** Late in the horizon the model keeps capacity that a
25-year life would retire, so it may build less new solar in the 2040s.

**How to test or improve.** Give that stock a retirement path in
`ResidualCapacity`.

### Reserve credit for storage

**The data gap.** A storage plant's contribution to firm capacity depends on
both its power and the energy it holds. The reserve margin counts capacity
only.

**What we assumed and why.** ASSUMPTION: short-duration storage counts toward
the reserve margin at 0.7 of its power capacity, in every year and node,
whether or not it holds energy. The credit sits below 1 because storage can
discharge only for a limited time.

**Effect on results.** Where storage has very little energy capacity, its power
is the cheapest way to meet the reserve requirement. The model can then hold
storage capacity that barely moves energy.

**How to test or improve.** Read storage capacity together with its energy
throughput. To test, lower the `ReserveMarginTagTechnology` value of the
storage technologies.

### One reserve margin for the whole system

**The data gap.** Each country plans its own reserve. The model's reserve
margin is one requirement for the whole system.

**What we assumed and why.** The reserve margin is written for the model's one
region, per time slice and year. Firm capacity in any country counts toward
demand in every country, as one pool.

**Effect on results.** The model does not check each country's own reserve. A
country can rely on firm capacity elsewhere, even where its links could not
carry that support.

**How to test or improve.** Compare each country's firm capacity with its own
peak demand in the results. The constraint is `RM3_ReserveMargin_Constraint` in
the model file; a check per country would need a reserve margin for each node.

### Total caps on three India-internal corridors

**The data gap.** The model's corridor schedules set how fast each link may
grow from the projects listed with commissioning dates. Three India-internal
corridors have no such schedule: India East to India West, India North-East to
India North and India North to India West.

**What we assumed and why.** The build cap of these three corridors is
`headroom_factor` times their residual capacity
([Interconnector build caps](../rules.md#interconnector-build-caps)). Their
`TotalAnnualMaxCapacity` keeps the value set before the build-cap rule runs.

**Effect on results.** For some of these corridors the total cap, not cost,
limits growth. In B and C cost sets the size only of the 13 listed corridors
with new investment; in A, only of the four listed India-internal corridors.

**How to test or improve.** Compare their capacity in the results with their
`TotalAnnualMaxCapacity`. To let cost set a corridor's size, give it a schedule
in `relax_interconnectors.yaml`.

### India South to West corridor, 2023 to 2026

**The data gap.** CEA states the corridor's capacity on two bases. Table 7.8
of its National Electricity Plan counts HVDC links at their reverse-direction
rating, while its monthly growth statements match Table 7.7 of the plan, which
counts them at their full rating. The plan gives the capacity planned by March
2027, not the year each line enters service.

**What we assumed and why.** The corridor starts at 15.12 GW in 2023, the value
of CEA's Manual on Transmission Planning Criteria 2023. The plan's Table 7.8
prints the same 15,120 MW for South-West at 31 March 2022, rising to 19,320 MW
at 31 March 2024 and 25,120 MW at 31 March 2027 (CEA,
[National Electricity Plan, Volume II, Transmission](http://web.archive.org/web/20241106202726/http://cea.nic.in/wp-content/uploads/notification/2024/10/National_Electricity_Plan_Volume_II_Transmission.pdf),
Table 7.8, printed p. 92, PDF p. 134). ASSUMPTION: the model must build 4.2 GW
in 2024, 4.2 GW in 2025 and 1.6 GW in 2026 on this corridor, the 10 GW the
National Electricity Plan adds to it. The corridor is then at least 15.12,
19.32, 23.52 and 25.12 GW in 2023 to 2026, and at least 25.12 GW from 2026. The
build cap of `A_Calibrated_BAU`, `B_Optimised_VRE` and `C_Target_VRE` starts at
18.12 GW a year, so it leaves room for all three builds.

**Effect on results.** CEA's growth statement gives the West-South total as
18,120 MW to March 2023 and 22,320 MW to March 2024, 2025 and 2026 (CEA,
[Growth of Inter-Regional Transmission Capacity, March 2026](http://web.archive.org/web/20260715234209/https://cea.nic.in/wp-content/uploads/transmission/2026/03/TL_GS_March___2026.pdf),
PDF p. 4). The model is 3.0 GW below it in 2023 and 2024, 1.2 GW above it in
2025 and 2.8 GW above it in 2026. The 3.0 GW in 2023 and 2024 is the gap
between the two rating bases. It matches the difference the plan gives for the
Raigarh to Pugalur HVDC link, 6,000 MW from West to South and 3,000 MW in
reverse (printed p. 92). From 2025 the model builds ahead of the growth
statements and reaches the plan's March 2027 total in 2026.

**How to test or improve.** The minimum builds are `TotalAnnualMinCapacityInvestment`
of `TRNINDSOINDWE`. A possible future calibration refinement aligns the start
value and the build years with CEA's growth statements, on one rating basis.

### Nepal and Bhutan import lines to Bangladesh

**The data gap.** Bangladesh's master plan, IEPMP 2023, names dedicated lines from Nepal and Bhutan only as options to study, with no route, size or year (section 5.5.4 and Figure 5.5-4, p. 95). No signed agreement names either line.

**What we assumed and why.** Neither link has a forced build, because a forced build stands for a committed or officially planned project and neither line is one. In `A_Calibrated_BAU` both links stay at their existing capacity, as A allows only existing capacity plus committed projects: zero for Bhutan to Bangladesh, and for Nepal to Bangladesh 0.04 GW from 2025, the 40 MW sale through the Indian grid that began in November 2024. In B and C both links also stay at their existing capacity until 2030; from 2031 the model may build either link, by at most 1 GW a year, when that lowers cost ([Interconnector build caps](../rules.md#interconnector-build-caps)).

**Effect on results.** Cost decides whether and when B and C build each link; A builds neither line.

**How to test or improve.** When a line becomes committed or planned, add its forced build to the link's `TotalAnnualMinCapacityInvestment` row in `Interconnector_Params`, with its source.

Source: [IEPMP 2023, Power Division copy](https://objectstorage.ap-dcc-gazipur-1.oraclecloud15.com/n/axvjbnqprylg/b/V2Ministry/o/office-powerdivision/2024/12/6aa63b99c7bf4b9083c7dc1b6bb07c28.pdf), section 5.5.4 and Figure 5.5-4 (printed p. 95, PDF p. 148).

### Bangladesh import limits and plant costs

**The data gap.** IEPMP 2023 plans Bangladesh's imports by connection point. It
leaves the projects, their sizes and their years open.

**What we assumed and why.** In `B_Optimised_VRE` and `C_Target_VRE`,
Bangladesh to India East and Bangladesh to India North-East keep their existing
capacity plus the scenario's forced builds until 2030, then may each add at
most 1 GW a year, or a larger forced build of that year, with cost setting
their total size. The two ASSUMPTIONs behind these limits, the lead time to
2030 and the rate of 1 GW a year, are stated with their sources in
[Interconnector build caps](../rules.md#interconnector-build-caps). In
`A_Calibrated_BAU` both links stay at their existing capacity, because
neither has a committed project. No forced build applies to either link in A
and B; in C the planned Katihar-Parbotipur-Bornagar corridor is a forced build
on both links in 2031. In 2023 to 2026 the base-year build lock sets these caps
instead ([Build limits](build-limits.md#the-base-year-build-lock)).

ASSUMPTION: Bangladesh gas plants cost 1,013.1 USD/kW and 17.061 USD/kW-year,
and coal plants 6,336.94 USD/kW and 98.56 USD/kW-year, in constant 2023 USD,
flat from 2023 to 2050. These are the figures of the EIA and Sargent & Lundy
capital cost study for a single-shaft combined-cycle plant, 921 USD/kW and
15.51 USD/kW-year, and an ultra-supercritical coal plant, 4,103 USD/kW and
61.60 USD/kW-year (EIA,
[Capital Cost and Performance Characteristics for Utility-Scale Electric Power Generating Technologies](https://www.eia.gov/analysis/studies/powerplants/capitalcost/pdf/capital_cost_AEO2025.pdf#page=24),
January 2024, Table 1-2, PDF p. 24), times a Bangladesh factor of 1.1 for gas
and about 1.544 (capital) and 1.6 (fixed) for coal. The study prices plants
built in the United States, and the factor carries its figures to Bangladesh.

ASSUMPTION: Bangladesh utility solar costs 840 USD/kW in 2023, falling to
466.5 USD/kW in 2050, in constant 2023 USD. This is 82 USD/kW, about 11 per
cent, above IRENA's 2023 global weighted average of 758 USD/kW (IRENA,
[Renewable Power Generation Costs in 2023](https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2024/Sep/IRENA_Renewable_power_generation_costs_in_2023.pdf)),
to allow for land and grid connection costs in Bangladesh.

**Effect on results.** In B and C the build caps limit how fast Bangladesh's
imports from India can grow; within them, cost sets the size of each link. In A
the two links keep their existing capacity. The plant costs
set how Bangladesh's own gas, coal and solar compare with imports.

**How to test or improve.** Change a link's schedule in the
`relax_interconnectors.yaml` of B or C, its forced builds in
`Interconnector_Params`, or the costs of `PWRNGSBGDXX`, `PWRCOABGDXX` and
`PWRSPVBGDXX` in `Secondary_Techs`. A possible future calibration refinement
replaces the scaled costs with costs of plants built in Bangladesh.

### Dorjilung's export direction in `B_Opt_DirContractual`

**The data gap.** Dorjilung's sponsors, Tata Power and DGPC
([21 November 2025](https://www.tatapower.com/news-and-media/media-releases/tata-power-and-druk-green-power-corporation-limited--dgpc--sign-commercial-agreements-for-the-implementation-of-1125-mw-dorjilung-hydroelectric-power-project-in-bhutan)),
and the World Bank
([5 May 2026](https://www.worldbank.org/en/news/press-release/2026/05/05/the-royal-government-of-bhutan-and-the-world-bank-sign-financing-agreements-for-the-dorjilung-hydroelectric-power-projec))
name India as the project's main export market, with about 80 per cent of its
output going to India. Which way the Bhutan to Bangladesh link would run
under contract is a judgement.

**What we assumed and why.** `B_Opt_DirContractual` is one sensitivity scenario
in which listed links flow in the direction of their contracts. As a judgement
call, it sets the Bhutan to Bangladesh link to flow from Bhutan to Bangladesh
for Dorjilung. `B_Opt_DirBidir`, with both directions open, is the reference
case.

**Effect on results.** In this scenario only, the link carries power only from
Bhutan to Bangladesh, so the model can send Bhutan hydropower to Bangladesh
that the project documents assign to India. Other scenarios are unaffected.

**How to test or improve.** Leave `TRNBTNXXBGDXX` out of the scenario's
`set_interconnector_direction.yaml`, or read `B_Opt_DirBidir`.

### India to Sri Lanka direction in `B_Opt_DirContractual`

**The data gap.** The base case of CEB's
[Long Term Generation Expansion Plan 2025-2044](https://www.ceb.lk/front_img/img_reports/1748839124LTGEP-2025-2044-FINAL_c.pdf)
adds the link as 500 MW in 2039 with Sri Lanka both importing and exporting: in
2039 Sri Lanka imports 2,104 GWh and exports 788 GWh over it (Table 10.6). The
plan sets no single direction.

**What we assumed and why.** As a judgement call in this one sensitivity
scenario, `B_Opt_DirContractual` sets the India to Sri Lanka link to flow only
from India to Sri Lanka. `B_Opt_DirBidir`, with both directions open, is the
reference case.

**Effect on results.** In this scenario only, Sri Lanka cannot export surplus to
India. Other scenarios are unaffected.

**How to test or improve.** To let the link run both ways, leave
`TRNINDSOLKAXX` out of the scenario's `set_interconnector_direction.yaml`, or
read `B_Opt_DirBidir`.

## Standard simplifications

### One internal grid loss for every node

**The data gap.** Countries publish transmission and distribution losses on
their own boundaries. Each node of the model has one internal grid between its
bus bars and its demand.

**What we assumed and why.** ASSUMPTION: every node loses 3 per cent of what
its internal grid carries, in every year, and every bus-bar demand row is
scaled by 0.97. The 3 per cent is a loss between bus bars; distribution losses
sit outside it
([Internal transmission losses and bus-bar demand](../rules.md#internal-transmission-losses-and-bus-bar-demand)).
Bangladesh's transmission losses sit close to it: 3.07 per cent in FY2022-23
and 3.13 per cent in FY2023-24, with distribution losses of 7.65 and 7.25 per
cent (Power Cell, Government of Bangladesh,
[Year wise System Loss Statistics](https://powercell.portal.gov.bd/site/page/d25a46c5-6c0f-47d7-94f2-9ac4351f4404/-),
retrieved 26 October 2026).

**Effect on results.** Real transmission losses differ between countries. A
node with higher losses needs slightly more generation for the same demand
than the model shows, and one with lower losses slightly less.

**How to test or improve.** Change `transmission_loss`. It moves the loss, the
demand scaling and the export ratio together.

### Bangladesh imports in 2023

**The data gap.** Bangladesh publishes imports by fiscal year, and model year
2023 overlaps two fiscal years.

**What we assumed and why.** ASSUMPTION: Bangladesh's 2023 imports lie between
98 and 102 per cent of the published reference. The band allows for rounding
and reporting differences.

**Effect on results.** The model uses the top of the band, so 2023 imports sit
2 per cent above the reference.

**How to test or improve.** Read 2023 imports as an allowed range, not a point.
To test, narrow the band toward the reference.

### Pumped storage in India

**The data gap.** CEA's hydro capacity and generation for India, which the
model uses, include pumped storage.

**What we assumed and why.** The model treats pumped storage as conventional
hydro in India North and India South, so hydro capacity and generation match
CEA's totals.

**Effect on results.** In those regions pumped storage runs on hydro capacity
factors and does not store energy from other plants.

**How to test or improve.** Move pumped storage capacity out of `PWRHYDINDNO`
and `PWRHYDINDSO` into a storage technology.

### India's coal and oil split

**The data gap.** CEA publishes each region's thermal generation as one figure,
not split between coal and oil.

**What we assumed and why.** ASSUMPTION: in 2023 to 2026, each India region's
thermal generation is split between coal and oil by weights on regional
capacity
([Where the base-year values come from](../reference/base_year_value_sources.md),
A8). The weights use coal capacities of North 47.229, South 46.4975, East
38.865, West 84.071 and North-East 0.75 GW, and oil capacities of South 0.43366
and North-East 0.036 GW. The oil capacities equal the model's 2023 oil stock.
The coal capacities differ from the model's 2023 coal stock (North 45.865,
South 44.998, East 36.997, West 85.786 GW) by up to 1.87 GW, in East, which is
5.0 per cent of East's stock.

**Effect on results.** Oil has weight only in South and North-East, so only
those two regions are affected. In South, oil gets 0.357 per cent of 2023
thermal generation instead of 0.369 per cent with the model's stock: about 28
GWh of 237,256 GWh.

**How to test or improve.** Weight the split with the model's own 2023 coal
stock.

### India's base-year solar builds

**The data gap.** MNRE's year-end tables give India's installed solar stock.
CEA's National Electricity Plan 2023 gives the yearly builds it plans. The two
sources split the stock between existing and new capacity differently.

**What we assumed and why.** For 2023 to 2025, the yearly minimum builds follow
CEA's National Electricity Plan 2023, and the installed stock follows
[MNRE's published year-end tables](../reference/base_year_value_sources.md).
The residual capacity of India North and South solar is set so that residual
plus minimum builds equals MNRE's stock in each year. This keeps the plan's
builds unchanged and the total equal to MNRE's stock. Sri Lanka's solar and
hydro residuals are set the same way.

**Effect on results.** Total solar capacity in 2023 to 2025 matches MNRE. The
split between existing and new capacity in those years follows the plan, not
actual additions; read actual additions from MNRE's tables.

**How to test or improve.** Take the 2023 to 2025 builds from actual additions
in MNRE's tables.

### Imports in the reserve margin

**The data gap.** The reserve requirement is plant output. Imports reach a
node through its dispatch technology, which the reserve margin would otherwise
count as supply.

**What we assumed and why.** Imports are left out of the reserve margin. A
country does not count capacity across a border as its own firm supply.

**Effect on results.** Imports do not lower any country's reserve need. Any
other technology that delivers imports onto a tagged fuel would be counted.

**How to test or improve.** The constraint skips every technology whose name
starts with `DSPTRN`. A new import route must use that prefix to stay out of
the reserve margin.

### Transit through a node

**The data gap.** Grid losses depend on the path energy takes inside a country.
Each node of the model has one internal grid, with no separate path for energy
in transit.

**What we assumed and why.** Imports pay the receiving node's 3 per cent loss,
and an export from a bus-bar node draws only 0.97 per unit sent, because a
bus-bar requirement already holds the losses that exports cause. Energy
imported and re-exported in the same time slice therefore ends with no
internal loss.

**Effect on results.** This favours flows through India's regions.

**How to test or improve.** Check transit flows through India's regions in the
results. A loss on transit would need a separate transit path in each node.

### Losses on India's internal links

**The data gap.** CEA publishes each region's requirement at its bus bars. It
does not state whether that requirement includes losses on transfers between
regions.

**What we assumed and why.** ASSUMPTION: a transfer between India's regions
loses only the interconnector's own loss, and each region's demand is CEA's
requirement at its bus bars, in every year.

**Effect on results.** If CEA's requirements already include inter-regional
losses, the model counts that loss twice on these transfers.

**How to test or improve.** Set the `OutputActivityRatio` of India's internal
links to 1, which counts the inter-regional loss only inside the requirements.

### Interconnector life

**The data gap.** Line and converter lives vary by equipment, and the model
needs one life for each link.

**What we assumed and why.** ASSUMPTION: every interconnector lasts 40 years.
One value for every link keeps their costs comparable.

**Effect on results.** No interconnector built from 2023 retires before 2050,
so the life changes only how capital cost is spread and what is credited at
2050.

**How to test or improve.** Change the `OperationalLife` rows of
`Interconnector_Params`.

## Notes on the workbook and files

### Bhutan's energy balance tables

**What you will see.** Two public tables of Bhutan's own statistics do not
match exactly. Generation plus imports less wheeled energy differs from
Bhutan's stated requirement by a few tens of GWh in 2023 and 2024 and by more
in 2025.

**Effect on results.** The model follows the requirement table, so small
differences from the generation table are expected.

**What to do.** Compare Bhutan's generation with the generation table with
that difference in mind.

### Comment heading in two schedule files

**What you will see.** In `config/scenarios/B_Optimised_VRE/relax_interconnectors.yaml`
and `config/scenarios/C_Target_VRE/relax_interconnectors.yaml`, a comment
heading reads "FROZEN CORRIDORS". The group includes Bhutan to Bangladesh and
India South to Sri Lanka, which have new investment. Only the two all-zero
schedules are frozen.

**Effect on results.** None. The model does not read comments.

**What to do.** Read the schedule values, not the heading.

### Comment in the training copy of the C schedule

**What you will see.** Line 2 of
`examples/unescap_training/config/scenarios/C_Target_VRE/relax_interconnectors.yaml`
describes the values as pipeline caps without saying they cap yearly
investment.

**Effect on results.** None. The model does not read comments.

**What to do.** Read each value as a build cap: the most new capacity the link
may add in that year.

### Workbook source cells that cite a home page

**What you will see.** Many source cells of the scenario workbook cite a
website's home page rather than the document that holds the value.

**Effect on results.** None. A reader cannot check those values from the
workbook alone.

**What to do.** Use the reference pages to trace a value. They cite the
documents for the base-year values
([Where the base-year values come from](../reference/base_year_value_sources.md))
and for the two provenance pages
([VRE ceilings](../reference/scenarios/vre_ceiling_provenance.md) and
[interconnector directions](../reference/scenarios/interconnector_direction_references.md)).

### The `Capacities_CF` template sheet

**What you will see.** The `Capacities_CF` sheet holds 12-slice capacity
factors and notes. Stage A3 replaces its values with the 20-slice timeslice
workbook, so its notes describe values that may not reach the model.

**Effect on results.** None. The model reads the timeslice workbook.

**What to do.** Read capacity factors from the compiled `CapacityFactor.csv`.

### Fuel codes of cloned generators

**What you will see.** A cloned generator whose own fuel the Primary sheets do
not produce keeps its template fuel, without a warning.

**Effect on results.** The clone uses its template's fuel, which may not be
its own.

**What to do.** Check the fuels of a new clone in the compiled
`InputActivityRatio.csv`.
