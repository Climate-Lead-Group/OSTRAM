# Scenario edits

A derived scenario is built from its parent's A-O workbooks by
`ostram/pipeline/scenarios/scenario_edits.py`, which reads the scenario's
block under `scenario_inputs` in `config/setups/south_asia.yaml`. Terms are in
the [glossary](glossary.md).

## Order of the build

1. The parent's folder `A1_Outputs_<parent>` is copied to `A1_Outputs_<scenario>`
   under `workspace/preparation/A1_Outputs/`. An existing target folder is
   first copied to `A1_Outputs_<scenario>_PREPATCH_<stamp>`, unless backups are
   skipped. The `run` command skips them.
2. The shared VRE ceilings are applied, unless the block sets
   `apply_vre_ceiling_layer: false`.
3. The edits in the block's `edits` list are applied in list order. Each edit
   sees the changes of the earlier ones.
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

## Operations

| Operation | Fields | Effect |
|---|---|---|
| `values` | `{year: value}` | Sets the listed year cells. Other years are unchanged. |
| `op: multiply` | `factor`, optional `years` | Multiplies each numeric cell by the factor, rounded to six decimals. Blank and text cells are unchanged. |
| `op: set_flat` | `value` | Sets every year to the value. |
| `op: set_to_residual` | none | Sets every year to the technology's `ResidualCapacity` on the same sheet. |
| `op: clamp_to_residual` | optional `residual_source` | Lowers each numeric cell to the technology's `ResidualCapacity` where the cell is higher. With `residual_source: minimum_investment_boundary`, allowed only for `TotalAnnualMinCapacityInvestment`, the bound comes from the scenario workbook. |
| `op: set_to_residual_factor_floor` | `factor`, `residual_year`, `base_window_floor`, `years` | Sets the listed years to the larger of the factor multiplied by `ResidualCapacity` in `residual_year` and the floor, rounded to six decimals. The target row must exist. |

## The shared VRE ceilings

Before the edits, for each of the 20 technologies in `vre_ceilings`, the
build sets `TotalAnnualMaxCapacity` to the ceiling in every year, lowers
`TotalAnnualMaxCapacityInvestment` and `TotalAnnualMinCapacityInvestment` to
the ceiling where they are higher, and, where the ceiling lowered the
previous maximum, scales a positive activity lower limit by the same ratio.
The ceiling is written as a flat value. A ceiling above the previous maximum
raises it; the pipeline never takes the lower of the two.

## The edit log and the backup folder

The edit log is `apply_patches_CHANGES_<stamp>.json` in `A1_Outputs_<scenario>`.
It records the scenario, its parent, the time stamp, the backup folder, every
changed cell with its old and new value, every row created, every
`tech_prefix` member skipped for lack of a row, the number of ceilings
applied and the record of the base-year pin. The backup folder, when one is
made, is `A1_Outputs_<scenario>_PREPATCH_<stamp>` beside the target.

## Edits that match nothing

The build stops with an error when an edit changes nothing:

- a `tech` or `techs` edit finds no row for the parameter and
  `create_if_absent` is false;
- a `create_if_absent` edit names a technology that is absent from the
  workbook, or a parameter that is absent from every sheet;
- a `tech_prefix` edit matches no technology;
- a `values` edit, or a `multiply` edit with `years`, names a year that is
  not a year column; the message lists the missing years;
- a shared VRE ceiling finds no `TotalAnnualMaxCapacity` row for its technology.

The message names the scenario, the edit and the reason, and suggests the
closest name when one exists:

```text
Scenario B_Opt_MyTest: the edit to TotalAnnualMaxCapacity for "PWRSOLBGXX" changed nothing. No technology has that name. Did you mean PWRSOLBGDXX?
```

The other reasons read `Sheet "<sheet>" has no parameter "<param>".`,
`No technology starts with "<prefix>".`, `None of the years <years> is a
column in the workbook.` and `Sheet "<sheet>" has no <param> row for it.`
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
