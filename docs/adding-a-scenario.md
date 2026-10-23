# Adding a scenario

To add a scenario, copy a scenario (with `new-scenario` or by hand), write your
change in the new file, and add one line to `decision_scenarios`. The middle
step is the scenario itself.

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

### Which kind of import cap

Two scenarios limit imports, and they work differently. Copy the one that
matches your case:

- `B_Opt_Clipped_GrossImportCap15` caps the gross imports of any set of
  countries as a share of each country's annual demand. It has no edits, only
  a `gross_import_cap` block, which the
  [scenario edits](scenario-edits.md#gross-import-cap) page describes.
- `B_Opt_TradeCap15` sets an upper limit on the yearly flow of each import
  corridor of one country, through an `import_share_cap` edit. Each run
  computes the limits from the solved `B_Optimised_VRE` of the same run, as the
  [scenario edits](scenario-edits.md#edits-that-read-a-solved-parent) page
  describes.

A policy block has no edits, so it has no `note` of an edit. Write its basis in
a `note:` line inside the block, in the same form as the note of an edit: an
external source, or `ASSUMPTION` followed by a short reason, in quotes when it
holds a colon. The cap reader takes only the four keys of the block, so the
note does not change the model.

```yaml
gross_import_cap:
  countries: [BGD]
  fraction: 0.15
  first_year: 2027
  last_year: 2050
  note: "ASSUMPTION: test case, a 15 per cent share"
```

A gross import cap does not appear in the edit log. It appears in the final
datafile, as the sets and the parameter that the
[scenario edits](scenario-edits.md#gross-import-cap) page names.

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
the delivered scenarios, the rows of its origin table, and the tests check that
they come first in the list. A new scenario does not appear in it, and the
generator leaves the page unchanged for it. To add the new scenario to the
catalogue as well, add one row for it to the origin table and run, from the
repository root:

```powershell
python tools/generate_scenario_catalogue.py
```

## The checks

From the repository root, in the activated environment:

```powershell
python -m ostram run --skip-pull --compile-only --scenarios B_Opt_SolarCapex120
```

If you created the environment under another name, add `--env-name <name>`.
The [quickstart](quickstart.md#3-run-a-solver-free-preparation-gate) gives the
run time and says how to read the output. The run prepares the four roots when
their post-A2 snapshots are missing, builds the parent root and the new
scenario, and stops after the solver datafile. A scenario derived from
`C_Target_VRE` also needs `--a-result-seed` with a solved `A_Calibrated_BAU`
result. Then look at these files under `workspace/`:

| What | File |
|---|---|
| The edit log | `preparation/A1_Outputs/A1_Outputs_B_Opt_SolarCapex120/apply_patches_CHANGES_<stamp>.json` |
| A compiled parameter | `execution/A2_Outputs_Params_otoole/B_Opt_SolarCapex120/CapitalCost.csv`, with rows `GLOBAL,<technology>,<year>,<value>` |
| The final datafile | `execution/Executables/B_Opt_SolarCapex120_0/Pre_processed_B_Opt_SolarCapex120_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt` |

Each changed cell in the edit log has an `old` and a `new` value. The `old`
value is the parent's value, unless the shared VRE ceilings or an earlier edit
of the same file changed the cell first. So you can compare the compiled value
with the parent without building a second scenario.
The [scenario edits](scenario-edits.md#the-edit-log-and-the-backup-folder) page
shows how to filter the log. The solver-free checks are listed under
[Tests](https://github.com/Climate-Lead-Group/OSTRAM/blob/main/CONTRIBUTING.md#tests) in `CONTRIBUTING.md`.

The test `tests/regression/test_config_files.py` passes for a new scenario when
its folder `config/scenarios/<name>/` holds the files you mean to add and the
name is in `decision_scenarios`. It fails when the folder is not registered,
and when a file is added anywhere else under `config/`, including the folders
of the delivered scenarios, unless the file is added to the list in the test.

A new root scenario is a different task. It needs a Control row in the
scenario workbook, its own rule YAMLs under `config/scenarios/<root>/`, root
entries in the registry and, for the base-year pin, a change to the code that
maps scenarios to their calibration root.
