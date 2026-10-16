# Adding a scenario

A derived scenario needs one block and two registry entries in
`config/setups/south_asia.yaml`, plus changes to the test contracts that list the
scenarios. It needs no workbook row and no folder, unless it changes
interconnector directions. The example adds `B_Opt_SolarCapex120`, which
multiplies the solar capital cost by 1.2. Terms are in the
[glossary](glossary.md).

## The scenario edits block

Add a block under `scenario_inputs`. The block name and the `scenario` field
must be equal, and `base_scenario` must be a root.

```yaml
  B_Opt_SolarCapex120:
    scenario: B_Opt_SolarCapex120
    base_scenario: B_Optimised_VRE
    description: 'The shared VRE ceilings applied to B_Optimised_VRE, with the solar
      capital cost multiplied by 1.2 in every year.'
    apply_vre_ceiling_layer: true
    edits:
    - sheet: Secondary Techs
      tech_prefix: PWRSPV
      param: CapitalCost
      op: multiply
      factor: 1.2
```

The [scenario edits](scenario-edits.md) page lists the fields and the
operations. An edit that matches nothing stops the run with a message that
names the edit.

## The registry entries

Append the name to `decision_scenarios` and an entry to `derived_scenarios`
in the `scenario_registry` section. The derived list must hold the decision
scenarios in the same order with the roots left out, so append to the end of
both lists.

```yaml
  decision_scenarios:
  - A_Calibrated_BAU
  ...
  - B_Opt_Clipped_GrossImportCap15
  - B_Opt_SolarCapex120
  derived_scenarios:
  - name: A_Calibrated_BAU_Clipped
    base_scenario: A_Calibrated_BAU
  ...
  - name: B_Opt_SolarCapex120
    base_scenario: B_Optimised_VRE
```

A scenario that changes interconnector directions also needs an overlay. Its
entry adds `direction_overlay: <name>/set_interconnector_direction.yaml` and
`direction_study_start_year: 2027`, and the overlay file goes in
`config/scenarios/<name>/`. `B_Opt_DirContractual` is the example to copy.
Only direction changes need an overlay.

## The test contracts

| File | Change |
|---|---|
| `ostram/pipeline/scenarios/registry.py` | Append the name to `DECISION_SCENARIOS`. |
| `tests/regression/accepted_baseline.py` | Append the name to `EXPECTED_DECISION_SCENARIOS`, and raise the two counts of 17 by one: the uniqueness check in `scenario_contract` and the row count of the manifest check. |
| `tests/regression/test_accepted_baseline.py` | Raise the two expected counts of manifest rows by one. |
| `tests/regression/test_scenario_registry.py` | Raise the expected count of decision scenarios by one. |
| `tests/regression/test_b1_runner.py` | Raise the expected count of decision scenarios by one. |
| `tests/regression/test_config_files.py` | Add the overlay file to `EXPECTED_CONFIG_FILES`, when the scenario has one. |

Then regenerate the scenario catalogue and add one row for the scenario to
its origin table:

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
