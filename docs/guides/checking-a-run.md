# Checking a run

A run can finish without an error and still not do what you meant: an edit
that changes the wrong rows, a limit that the model never reaches. These checks confirm a scenario before and after the
solve. Paths are under the workspace folder `workspace/`; `<S>` is the scenario
name.

## Before the solve: compile-only comparison

Build the new scenario together with the scenario it should differ from, for
example its parent or `B_Opt_Clipped`, without solving:

```powershell
python -m ostram run --skip-pull --compile-only --scenarios "B_Opt_Clipped,<S>"
```

The run ends with `OSTRAM compile-only run completed successfully`, and
`$LASTEXITCODE` is 0.

1. **The edit log.** Open
   `preparation/A1_Outputs/A1_Outputs_<S>/apply_patches_CHANGES_<stamp>.json`.
   Every changed cell is listed with its `old` and `new` value; the `old` value
   is the parent's unless the shared VRE ceilings or an earlier edit changed
   it first. The `cells` list also holds the cells of the shared VRE ceilings
   when the scenario applies them. Rows created and `tech_prefix` members
   skipped for lack of a row are listed too. An edit that matches nothing
   stops the build with a message that names the edit.
2. **The compiled files.** Compare
   `execution/A2_Outputs_Params_otoole/<S>/` with the same folder of the
   comparison scenario, file by file, with any file comparison tool. Only the
   parameters and rows you meant to change may differ. A gross import cap
   changes no compiled file.
3. **The solver datafile.** Open
   `execution/Executables/<S>_0/Pre_processed_<S>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt`.
   It holds the values after the B2 steps (base-year build lock, storage delay,
   backstop opening, gas and oil allowance) and the gross import cap. Search
   for the block that starts `param default <value> : <Parameter> :=` and then
   for the technology name, or for `D4Import`.

## After the solve

`python -m ostram run --skip-pull --scenarios <S>` solves the scenario. The
results are in `execution/Executables/<S>_0/Outputs/`, one CSV file per model
variable, with zero values left out. A missing row means zero.

### The solver status

The solver log sits beside the datafile; with CPLEX it is
`Pre_processed_<S>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX_output.cplex.log`.
Require an optimal status. CPLEX may report "optimal with unscaled
infeasibilities". The solution is optimal; after CPLEX undoes its internal
scaling, a few variables sit slightly outside their bounds. Accept the run
when the largest unscaled infeasibility is below 1e-5, and look into it
otherwise. The status is `solutionStatusString` in the header of the solution
file beside the log, `..._output.sol`, and the log prints the largest
unscaled infeasibility near its end, on the lines
`Maximum unscaled bound infeasibility =` and
`Maximum unscaled reduced-cost infeasibility =`, which appear only for a run
with that status. The run log of the whole run is
`<workspace>/logs/<run-id>/run.log`.

### Zero backstop

The backstop `PWRBCK<node>` produces only when the other limits leave demand
unmet. In `ProductionByTechnologyAnnual.csv`, look for rows whose `TECHNOLOGY`
starts with `PWRBCK`. There should be none, or only zero values. A backstop row
names the node and the year where the limits are too tight: a build cap, the
base-year build lock, an import cap or an activity floor.

### The import share

For a country `<C>` with flows listed in `set D4ImportFlow[<C>]` of the
datafile:

```text
received imports(y)       = sum of ProductionByTechnologyAnnual over those flows,
                            rows whose FUEL is the flow's ELC...04 fuel
published requirement(y)  = sum over time slices of Demand for the country's ELC...03 fuels
                            / (1 - 0.03)     bus-bar countries
import share(y)           = received imports(y) / published requirement(y)
```

Under a gross import cap the share is at most `fraction` in each capped year,
equal to it where the cap binds. For the Maldives, leave out the division by
0.97. For a per-corridor cap, compare each link's yearly activity in
`TotalTechnologyAnnualActivity.csv` with its limit in the compiled
`TotalTechnologyAnnualActivityUpperLimit.csv`.

### The build caps reached

To see where a build cap binds, compare `NewCapacity.csv` with the final cap:

```text
NewCapacity(technology, y)  versus  TotalAnnualMaxCapacityInvestment(technology, y) in the datafile
```

A value equal to the cap means the cap binds in that year. For an
interconnector, also compare `TotalCapacityAnnual.csv` with
`TotalAnnualMaxCapacity`. The internal grid has no cap: its
`TotalCapacityAnnual` follows the busiest time slice, which
`RateOfActivity.csv` shows per time slice and mode.

### The internal grid of a node

```text
peak flow(family, y)  = the largest, over time slices, of the sum over modes of RateOfActivity
utilisation           = peak flow / (TotalCapacityAnnual x 31.536)
```

Add the three renewable families together and the three other families
together before dividing. A utilisation of 1 means the grid is full in its
busiest time slice.

## A checklist

The tree shows where this section reads.

```text
OSTRAM/
├── config/
│   ├── scenarios/
│   │   └── <S>/
│   │       └── scenario.yaml                       <- CHECK the scenario file loads
│   └── setups/
│       └── south_asia.yaml
│           └── decision_scenarios:                 <- CHECK the scenario is the last line
└── workspace/
    ├── preparation/
    │   └── A1_Outputs/
    │       └── A1_Outputs_<S>/
    │           └── apply_patches_CHANGES_<stamp>.json
    │                                               <- CHECK only the intended cells
    └── execution/
        ├── A2_Outputs_Params_otoole/
        │   └── <S>/                                <- CHECK only the intended files and rows differ
        └── Executables/
            └── <S>_0/
                ├── Pre_processed_<S>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt
                │                                   <- CHECK the intended values
                └── Outputs/                        <- CHECK no backstop, import share, build caps
```

- The scenario file loads, and the scenario is the last line of
  `decision_scenarios`.
- The edit log lists only the intended cells and no skipped members you did not
  expect.
- Only the intended compiled files and rows differ from the comparison
  scenario.
- The datafile holds the intended values, and for a gross cap a non-empty
  `set D4ImportCountry` with one `D4ImportFraction` line per capped year.
- The solve is optimal, or optimal with unscaled infeasibilities below 1e-5.
- No backstop production in any node or year.
- The import share and the binding build caps are where you expect them.
- The solver-free checks listed under Tests in `CONTRIBUTING.md` pass.
