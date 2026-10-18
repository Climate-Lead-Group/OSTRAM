# Adding a scenario

A derived scenario is one file, `config/scenarios/<name>/scenario.yaml`, and one
line in the scenario list of `config/setups/south_asia.yaml`. It needs no
workbook row, no code change and no test change. The example adds
`B_Opt_SolarCapex120`, which multiplies the solar capital cost by 1.2. Terms are
in the [glossary](glossary.md).

## 1. The scenario file

Start from a scenario that is close to the new one, for example
`B_Opt_SolarCapex130`. From the repository root, with the environment's Python:

```powershell
python -m ostram new-scenario B_Opt_SolarCapex120 --from B_Opt_SolarCapex130
```

```text
Created config/scenarios/B_Opt_SolarCapex120/scenario.yaml (copied from B_Opt_SolarCapex130).

Next:
1. Rewrite the description and edits in that file.
2. Add this line at the end of decision_scenarios in config/setups/south_asia.yaml:
  - B_Opt_SolarCapex120
```

The command copies the folder `config/scenarios/B_Opt_SolarCapex130/` to
`config/scenarios/B_Opt_SolarCapex120/` and sets the `scenario` line of the
copied `scenario.yaml` to the new name. Everything else in the folder is copied
as it is. It stops without copying when the new name is already a folder or a
registered scenario, or when the source is not a derived scenario. It never
edits `config/setups/south_asia.yaml`.

To do the same by hand, copy the folder and change the `scenario` line.

Then edit `scenario.yaml`. The folder name and the `scenario` field must be
equal, and `base_scenario` must be a root.

```yaml
scenario: B_Opt_SolarCapex120
base_scenario: B_Optimised_VRE
description: >-
  The shared VRE ceilings applied to B_Optimised_VRE. The CapitalCost of every
  solar technology is multiplied by 1.2 in every year.
apply_vre_ceiling_layer: true
edits:
- sheet: Secondary Techs
  tech_prefix: PWRSPV
  param: CapitalCost
  op: multiply
  factor: 1.2
```

The [scenario edits](scenario-edits.md) page lists the fields and the
operations. Write the basis of each edit in its `note`, as that page
describes. An edit that matches nothing stops the run with a message that
names the edit.

A scenario that changes interconnector directions also has an overlay file in
its folder, and two more lines in `scenario.yaml`:
`direction_overlay: set_interconnector_direction.yaml` and
`direction_study_start_year: 2027`. `B_Opt_DirContractual` is the folder to
copy. Only direction changes need an overlay.

## 2. The scenario list

Add the name as the last line of `decision_scenarios` in the
`scenario_registry` section of `config/setups/south_asia.yaml`.

```yaml
  decision_scenarios:
  - A_Calibrated_BAU
  ...
  - B_Opt_Clipped_GrossImportCap15
  - B_Opt_SolarCapex120
```

The registry reads the parent and the overlay from the scenario's file. The
order of `decision_scenarios` is the order in which `run` builds the scenarios.

The [scenario catalogue](reference/scenarios/scenario_catalogue.md) documents
the delivered scenarios, and the tests check that they come first in the list.
To add the new scenario to the catalogue as well, add one row for it to the
catalogue's origin table and run, from the repository root:

```powershell
python tools/generate_scenario_catalogue.py
```

## The checks

From the repository root, with the environment's Python:

```powershell
python -m unittest discover -s tests -p "test_*.py"
python -m ostram run --skip-pull --compile-only --env-name <env> --scenarios B_Opt_SolarCapex120
```

The run prepares the four roots when their post-A2 snapshots are missing,
builds the parent root and the new scenario, and stops after the solver
datafile. The edit log is in
`workspace/preparation/A1_Outputs/A1_Outputs_B_Opt_SolarCapex120/` and the
solver datafile is in `workspace/execution/Executables/B_Opt_SolarCapex120_0/`.
A scenario derived from `C_Target_VRE` also needs `--a-result-seed` with a
solved `A_Calibrated_BAU` result.

A new root scenario is a different task. It needs a Control row in the
scenario workbook, its own rule YAMLs under `config/scenarios/<root>/`, root
entries in the registry and, for the base-year pin, a change to the code that
maps scenarios to their calibration root.
