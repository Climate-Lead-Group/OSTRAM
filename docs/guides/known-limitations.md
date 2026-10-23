# Known limitations

Each item says what the limitation is, what it does to results, and what a user
can do. Effects are described in words; read the sizes from your own runs.

## Demand, generation and base years

**India's 2023 and 2024 generation floors are gross generation.** India's
demand rows are the net requirement at power station bus bars, while its
base-year generation floors are gross generation, which includes the plants'
own use. The model therefore generates more than India's demand needs in those
years. To compare with statistics, compare generation with gross generation;
to remove the gap, the floors and the demand would need one boundary.

**India's demand rows are projections, not actuals, for the early years.** The
India rows of `Demand_Projection` follow the 20th Electric Power Survey
fiscal-year projections. Actual requirements of the same years differ region by
region, in both directions, and the actual figures leave out inter-state
losses, so part of any gap is a difference of boundary. Treat early-year India
demand as a projection.

**One uniform internal loss.** Every node loses 3 per cent inside its grid
(`transmission_loss`), and every bus-bar demand row is scaled by the same 0.97.
Real transmission and distribution losses differ between countries. To test
another loss, change `transmission_loss`; it changes the loss, the demand
scaling and the export ratio together.

**Bangladesh 2024 generation follows the fiscal-year floors.** Bangladesh's
generation floors are 98 per cent of BPDB fiscal-year figures, while demand is a
calendar-year estimate. In 2024 the floors, not demand, set generation. Read
Bangladesh 2024 generation as set by the floors.

**Bangladesh 2023 imports sit at the top of their band.** The base-year import
band allows 98 to 102 per cent of the reference, and the model uses the top of
it. Read 2023 imports as an allowed range, not a point.

**Bhutan's published tables do not add up exactly.** Generation plus imports
less wheeled energy differs from Bhutan's stated requirement by a few tens of
GWh in 2023 and 2024 and by more in 2025. The model follows the requirement
table; small mismatches with the generation table are expected.

**Bhutan hydro availability is lowered twice.** Bhutan's hydro capacity factors
come from observed output, and an `AvailabilityFactor` of 0.95 applies on top,
so the most the fleet can produce is lower than the observed factors allow.
New Bhutan hydro capacity carries the same factors.

**Bhutan hydro capacities have no capacity source.** The Bhutan hydro
residual capacities on `Secondary_Techs` are the year-end capacity of the
plants that report production; the cited bulletins give production, not
capacity.

**Pumped storage counts as hydro in India.** India's hydro capacity and
generation, from CEA, include pumped storage, which the model treats as
conventional hydro in `PWRHYDINDNO` and `PWRHYDINDSO`.

**The Maldives' demand is estimated backward.** No public source gives observed
whole-country consumption including resort self-generation for 2023 to 2025;
the model's values are estimated backward from the roadmap's 2028 anchor.
Utility-basis figures are much lower.

**The thermal split in India uses weights.** CEA regional generation is split
into coal and oil with weights on regional capacities. The weights are labelled
ASSUMPTIONs, and the regional capacities have no public locator.

**Some base-year values have no retrievable reference.** The notes of a few
`Secondary_Techs` rows (Bangladesh, Bhutan and Nepal solar, Nepal hydro) say
that their base-year value has no retrievable reference.

## Reserve margin and storage

**Storage earns reserve credit without stored energy.** The reserve margin
counts short-duration storage power at 0.7 of its capacity, whether or not the
storage holds energy. Where storage has very little energy capacity, its power
is the cheapest way to meet the reserve requirement, and the model can hold
storage capacity that barely moves energy. Read storage capacity together with
its energy throughput.

**The reserve margin is one constraint for the whole system.** The constraint
`RM3_ReserveMargin_Constraint` is written for the region `GLOBAL`, per time
slice and year, so firm capacity in any country covers demand in every
country. It does not check each country's own reserve.

**Imports are left out of the reserve margin by name.** The reserve requirement
is plant output, and imports are excluded because the constraint skips every
technology whose name starts with `DSPTRN`. Any other technology that delivers
imports onto a tagged fuel would be counted.

## Interconnector limitations

**Transit through a node has no net internal loss.** Imports pay the receiving
node's 3 per cent loss, and an export from a bus-bar node draws only 0.97 per
unit sent. Energy imported and re-exported in the same time slice therefore
ends with no internal loss, which favours flows through India's regions.

**Three India-internal corridors keep total caps.** India East to India West,
India North-East to India North and India North to India West have no
schedule. Their build cap is `headroom_factor` times residual capacity, and
their `TotalAnnualMaxCapacity` stays as earlier stages wrote it, so for some of
them the total cap, not cost, limits growth. Cost sets the size only of the 13
listed corridors with new investment.

**India link losses are only the links' own.** A transfer between India's
regions pays only the interconnector's loss. Whether CEA's regional bus-bar
requirements already include inter-regional losses is not checked.

**The 40-year interconnector life has no external source.** The
`OperationalLife` rows of `Interconnector_Params` give 40 years for every
interconnector without a source that states an interconnector life. Test
another life by changing those rows.

**Some projects and costs are cited to IEPMP where IEPMP does not list
them.** Several Bangladesh interconnector projects in the `Interconnectors`
sheet, from which the `relax_interconnectors.yaml` schedules are built, and
some Bangladesh plant capital and fixed costs cite the IEPMP 2023, which gives
no such project, capacity, date or cost. A number of other IEPMP citations give
no page or table. Treat those schedules and costs as ASSUMPTIONs.

**The official IEPMP host fails certificate checks.** The IEPMP URL on the
ministry site does not pass a certificate check; the citations also name a
copy hosted by Power Division.

**The India to Sri Lanka design has no source.** The voltage, capacity and
technology stated for the India to Sri Lanka link in the interconnector
direction references have no source that states them.

**Internal text in two sets of source cells.** The rows of
`Interconnector_Params` for `TotalAnnualMinCapacityInvestment` and
`MinimumInvestmentClampBoundary` carry fixed internal wording in their `Source`
cells, because the loader requires that exact text. Their bases are in the
`IEEE Reference` cells for the two Nepal to Bangladesh rows only.

## Text a user may read

**A comment heading in two schedule files.** In the `relax_interconnectors.yaml`
files of `B_Optimised_VRE` and `C_Target_VRE`, a comment heading reads
"FROZEN CORRIDORS", but the group includes Bhutan to Bangladesh and India South
to Sri Lanka, which have new investment. Only the two all-zero schedules are
frozen.

**The training copy of the C schedule.** Line 2 of
`examples/unescap_training/config/scenarios/C_Target_VRE/relax_interconnectors.yaml`
describes the values as pipeline caps without saying they cap yearly
investment.

**The `Capacities_CF` sheet holds a template.** Its 12-slice values are
replaced by the 20-slice timeslice workbook in stage A3; its notes describe
values that may not reach the model. Read capacity factors from the compiled
`CapacityFactor.csv`.

**Clone fuel codes.** A cloned generator whose own fuel the Primary sheets do
not produce keeps its template fuel without a warning. Check the fuels of a new
clone in the compiled `InputActivityRatio.csv`.
