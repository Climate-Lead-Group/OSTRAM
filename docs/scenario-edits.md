# Scenario edits

A derived scenario is built from its parent's A-O workbooks by
`ostram/pipeline/scenarios/scenario_edits.py`, which reads the scenario's
file `config/scenarios/<name>/scenario.yaml`. Terms are in
the [glossary](glossary.md).

## Order of the build

1. The parent's folder `A1_Outputs_<parent>` is copied to `A1_Outputs_<scenario>`
   under `workspace/preparation/A1_Outputs/`. An existing target folder is
   first copied to `A1_Outputs_<scenario>_PREPATCH_<stamp>`, unless backups are
   skipped. The `run` command skips them.
2. The shared VRE ceilings are applied, unless the scenario file sets
   `apply_vre_ceiling_layer: false`.
3. The edits in the file's `edits` list are applied in list order. Each edit
   sees the changes of the earlier ones. When the file names another scenario
   in `reuse_edits_from`, that scenario's edits are applied first.
4. `A-O_Parametrization.xlsx` is saved and the base-year pin is applied again.
5. The edit log is written to the target folder.

For `B_Opt_DirBidir` and `B_Opt_DirContractual` the overlay is applied after
these steps.

## Edit fields

| Field | Meaning |
|---|---|
| `sheet` | The worksheet of `A-O_Parametrization.xlsx`, for example `Secondary Techs` or `Demand Techs`. |
| `param` | The value of the `Parameter` column, for example `CapitalCost`. |
| `tech` | One technology name. |
| `techs` | A list of technology names. |
| `tech_prefix` | Every technology on the sheet whose name starts with the text. |
| `create_if_absent` | With `true`, a technology without a row for the parameter gets a new row, with its identity columns copied from another row of that technology. The default is `false`. |
| `years` | With `multiply`, the years to change. Other years keep their value. |
| `note` | Free text. Nothing reads it. |

An edit names exactly one of `tech`, `techs` and `tech_prefix`, and holds
either `values` or one `op`. The edits recognise year columns from 2020 to
2060. A changed row whose `Projection.Mode` is blank or `EMPTY` is set to
`User defined`, so that B1 exports it.

The `note` of an edit is where its basis goes: an external source, as a URL or
a DOI, or `ASSUMPTION` followed by a short reason. The existing notes are being
sourced separately. Put the note in quotes when it holds a colon followed by a
space, because a bare colon inside an unquoted YAML value stops the file from
loading:

```yaml
  note: "ASSUMPTION: test case, a 20 per cent rise"
```

## Sheets and technology prefixes

The `sheet` field names the worksheet that holds the technology and the
parameter. The delivered scenarios edit `Secondary Techs`, `Demand Techs` and
`VariableCost`. The table shows what the sheets of `A-O_Parametrization.xlsx`
hold after A2. A technology name starts with its prefix; the
[data reference](data-reference.md#technology-prefixes) describes the prefixes.

| Sheet | Prefixes | Parameters |
|---|---|---|
| `Primary Techs` | `MIN`: fuel supply | `CapitalCost`, `FixedCost`, `ResidualCapacity`, `TotalAnnualMaxCapacity`, the investment limits and the annual activity limits |
| `Secondary Techs` | `PWR`: power plants and storage; `TRN`: the interconnector technologies, 13 characters long, such as `TRNBGDXXINDEA`; `ELC`: electricity distribution | The same parameters as `Primary Techs` |
| `Demand Techs` | `RNW`, `TRN` and `PWR`: the per-country grid-tier technologies `RNWNLI`, `RNWRPO`, `RNWTRN`, `TRNNLI`, `TRNRPO` and `PWRTRN`; `DSP`: `DSPTRN` | `CapitalCost`, `FixedCost`, `ResidualCapacity`, `TotalAnnualMaxCapacity`, `TotalAnnualMinCapacityInvestment` |
| `VariableCost` | `MIN`: fuel prices; `PWR`: backstop and storage; `TRN`: interconnectors | `VariableCost` |

An edit changes the first row of its technology and parameter on the sheet. The
sheet `Capacities` holds several rows per technology for `CapacityFactor`, one
for each value of its `Timeslices` column, so it is not a target of the
delivered scenarios.

## Operations

| Operation | Fields | Effect |
|---|---|---|
| `values` | `{year: value}` | Sets the listed year cells. Other years are unchanged. |
| `op: multiply` | `factor`, optional `years` | Multiplies each numeric cell by the factor, rounded to six decimals. Blank and text cells are unchanged. |
| `op: set_flat` | `value` | Sets every year to the value. |
| `op: set_to_residual` | none | Sets every year to the technology's `ResidualCapacity` on the same sheet. |
| `op: clamp_to_residual` | optional `residual_source` | Lowers each numeric cell to the technology's `ResidualCapacity` where the cell is higher. With `residual_source: minimum_investment_boundary`, allowed only for `TotalAnnualMinCapacityInvestment`, the bound comes from the scenario workbook. |
| `op: set_to_residual_factor_floor` | `factor`, `residual_year`, `years`, and either `base_window_floor` or `base_window_result` with `base_window_years`; optional `clamp` | Sets the listed years to the larger of the factor multiplied by `ResidualCapacity` in `residual_year` and the floor, rounded to six decimals. The floor is `base_window_floor`, or the largest `TotalCapacityAnnual` of the technology over `base_window_years` in the solved parent named by `base_window_result`. With `clamp: true` it only lowers cells, and a missing row is skipped: with `base_window_result`, the listed years in order to the headroom the value leaves above the parent's base-year capacity and the earlier listed years; with `base_window_floor`, to the value. Otherwise the target row must exist. |
| `op: from_result` | `result`, `variable`, `years`, optional `result_year` | Sets the listed years to the technology's `variable` (`TotalCapacityAnnual` or `NewCapacity`) in the solved parent named by `result`, in the same year or in `result_year`. |
| `op: import_share_cap` | `techs`, `result`, `share`, `demand_fuel`, `importer`, `years` | Sets each listed technology's value in each listed year to `share` times the parent's demand of `demand_fuel` that year times the technology's share of the parent's imports into `importer` that year. A `demand_fuel` measured at the bus bar (the `Measured_at` column of `Demand_Projection`) holds the published requirement times 1 minus `internal_transmission.transmission_loss`, so its demand is divided by that factor first. A year in which the parent imports nothing gives 0. |
| `op: copy_from_tech` | `source_tech` | Sets every year to the value of the same parameter in the `source_tech` row on the same sheet, as that row stands when the edit runs. |

## Edits that read a solved parent

`op: from_result`, `op: import_share_cap` and `base_window_result` read the
result tables of a solved scenario of the same run, from
`Executables/<parent>_0/Outputs/` under the execution workspace. The scenario
declares that parent under `derived_dependencies` in
`config/setups/south_asia.yaml`, with an environment variable that can point to
another result folder instead. A value missing from a result table is 0, as the
tables leave out zeros. The run stops before A3 when the folder is missing, and
an edit stops when the table it reads is missing or empty; both messages name
the scenario and the parent.

| Scenario | Parent | Environment variable |
|---|---|---|
| `B_Opt_TxCap150` | `B_Optimised_VRE` | `OSTRAM_B_OPTIMISED_VRE_RESULT` |
| `B_Opt_TradeCap15` | `B_Optimised_VRE` | `OSTRAM_B_OPTIMISED_VRE_RESULT` |
| `A_Calibrated_BAU_Clipped_TxFreeze2026` | `A_Calibrated_BAU_Clipped` | `OSTRAM_A_CALIBRATED_BAU_CLIPPED_RESULT` |

So the parent is solved first, as `A_Calibrated_BAU` is for `C_Target_VRE`; the
[portfolio sequence](pipeline.md#24-scenario-portfolio) gives the order. The
factor 1.5 of `B_Opt_TxCap150` and the share 0.15 of `B_Opt_TradeCap15` are the
scenarios' design settings and stay in their files.

## Reusing another scenario's edits

A scenario file can name another derived scenario whose edits it applies before
its own:

```yaml
scenario: B_Opt_IndiaCostsFuel
base_scenario: B_Optimised_VRE
reuse_edits_from: B_Opt_IndiaCosts
edits:
- sheet: VariableCost
  techs: [MINCOABGD, MINCOABTN, MINCOALKA, MINCOAMDV, MINCOANPL]
  param: VariableCost
  op: copy_from_tech
  source_tech: MINCOAIND
```

The named scenario must have its own file under `config/scenarios/`, the same
`base_scenario`, and no `reuse_edits_from` of its own. Otherwise the build
stops before it copies the parent, with a message that says which rule failed.
The named scenario's shared VRE ceilings setting is not used; the reusing
scenario's own setting applies.

## The shared VRE ceilings

Before the edits, for each of the 20 technologies in `vre_ceilings`, the
build sets `TotalAnnualMaxCapacity` to the ceiling in every year, lowers
`TotalAnnualMaxCapacityInvestment` and `TotalAnnualMinCapacityInvestment` to
the ceiling where they are higher, and, where the ceiling lowered the
previous maximum, scales a positive activity lower limit by the same ratio.
The ceiling is written as a flat value. A ceiling above the previous maximum
raises it; the pipeline never takes the lower of the two.

## Gross import cap

A scenario file can hold a `gross_import_cap` block. It is a policy, not an
edit. `ostram/pipeline/execution/gross_import_cap.py` reads it when B2 converts
the compiled files into the solver datafile. The edit stage and the edit log do
not see it. `B_Opt_Clipped_GrossImportCap15` holds one for six countries. A
block for one country reads:

```yaml
gross_import_cap:
  countries: [BGD]
  fraction: 0.15
  first_year: 2027
  last_year: 2050
```

| Field | Meaning |
|---|---|
| `countries` | A list of country codes: `BGD`, `BTN`, `IND`, `LKA`, `MDV` and `NPL`. A repeated code or a code outside the list stops the run. India is one country, with its five regions. |
| `fraction` | The cap as a share of the country's annual demand, from 0 to 1. A value outside that range stops the run. |
| `first_year`, `last_year` | The first and the last year of the cap, both included. `first_year` must not be later than `last_year`. |

In each of these years, the gross electricity that the country receives from
other countries must not exceed `fraction` times its annual demand. The
received electricity is the sum, over the timeslices, of the activity of each
flow multiplied by the year split and by the output activity ratio. The annual
demand is the `SpecifiedAnnualDemand` of the country's demand fuels, the fuels
that start with `ELC`, the country code and end with `03`. For India that is
the sum of its five regions. Demand rows measured at the bus bar hold the
published requirement x (1 - loss), with the loss of
`internal_transmission.transmission_loss`, so for a country whose rows are all
measured at the bus bar `D4ImportFraction` is the fraction divided by
(1 - loss), and the cap is the fraction of the published requirement. For the
Maldives, measured at consumption, it is the fraction. A country whose rows are
measured at both points stops the run.

The flows that count are the modes of the cross-border `TRN` technologies whose
two country codes differ and whose output fuel starts with `ELC` and the code
of the country, with a positive output activity ratio in the compiled
`OutputActivityRatio.csv`. Links between Indian regions are excluded, because
both ends are in India. A flow from India into another country counts for that
country.

The cap is applied at compilation. The final datafile holds it as
`set D4ImportCountry`, `set D4ImportDemand`, `set D4ImportFlow` and
`param D4ImportFraction`, in
`workspace/execution/Executables/<scenario>_0/Pre_processed_<scenario>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt`.
A scenario without the block has an empty `set D4ImportCountry`. The edit log
does not list the cap. The run stops with an error when a scenario holds the
block and the model file lacks the parameter `D4HistoricalTradeActive`.

## The edit log and the backup folder

The edit log is `apply_patches_CHANGES_<stamp>.json` in `A1_Outputs_<scenario>`.
It records the scenario, its parent, the time stamp, the backup folder, every
changed cell with its old and new value, every row created, every
`tech_prefix` member skipped for lack of a row, the number of ceilings
applied, the record of the base-year pin and, for a scenario with
`reuse_edits_from`, the reused scenario and its number of edits. The backup
folder, when one is made, is `A1_Outputs_<scenario>_PREPATCH_<stamp>` beside the target.

The `cells` list also holds the cells of the shared VRE ceilings, so it is
longer than the edits of the file suggest. This command shows the cells of one
parameter:

```powershell
(Get-Content apply_patches_CHANGES_<stamp>.json -Raw | ConvertFrom-Json).cells |
  Where-Object { $_.param -eq 'CapitalCost' } | Format-Table tech, year, old, new
```

The `old` value of a cell is its value when the edit ran. That is the parent's
value, unless the shared VRE ceilings or an earlier edit of the same file
changed the cell first. A comparison with the parent therefore needs no build
of the parent scenario.

## Edits that match nothing

The build stops with an error when an edit changes nothing:

- a `tech` or `techs` edit finds no row for the parameter and
  `create_if_absent` is false;
- a `create_if_absent` edit names a technology that is absent from the
  workbook, or a parameter that is absent from every sheet;
- a `tech_prefix` edit matches no technology;
- a `values` edit, or a `multiply` edit with `years`, names a year that is
  not a year column; the message lists the missing years;
- a shared VRE ceiling finds no `TotalAnnualMaxCapacity` row for its technology;
- a `copy_from_tech` edit finds no row of its `source_tech` for the parameter.

The message names the scenario, the edit and the reason, and suggests the
closest name when one exists:

```text
Scenario B_Opt_MyTest: the edit to TotalAnnualMaxCapacity for "PWRSOLBGXX" changed nothing. No technology has that name. Did you mean PWRSOLBGDXX?
```

The other reasons read `Sheet "<sheet>" has no parameter "<param>".`,
`No technology starts with "<prefix>".`, `None of the years <years> is a
column in the workbook.` and `Sheet "<sheet>" has no <param> row for it.`
A missing `copy_from_tech` source reads `has no source row. Sheet "<sheet>"
has no <param> row for <source_tech>.`
The `run` command prints the message and exits with code 1.

## Rebuilding one derived scenario by hand

From the repository root, with the environment's Python and an existing
`A1_Outputs_<parent>` folder:

```powershell
python -B -m ostram.pipeline.scenarios.scenario_edits --scenario B_Opt_TxCap150
python -m ostram compile-inputs --scenarios B_Opt_TxCap150
python -B -m ostram.pipeline.execution.runner --scenarios B_Opt_TxCap150 --compile-only
```

The first command rebuilds the derived folder and writes the edit log.
`--skip-backup` skips the backup folder, `--restore` restores the latest
backup, and `--source-scenario` must equal the declared parent. The second
command runs B1 and the third runs B2 up to the solver datafile. A direction
variant also needs its overlay, which `python -m ostram run` applies.
