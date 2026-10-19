# south_asia.yaml section map

`config/setups/south_asia.yaml` is the setup file of the South Asia model.
Its top-level sections, in file order, hold the following. Terms are in the
[glossary](glossary.md).

| Section | What it holds | Read by |
|---|---|---|
| `schema`, `id` | The setup schema `ostram-setup-v1` and the setup id `south_asia`. | The setup loader, for every command. |
| `metadata` | The title and `effective_values`, one documented value of the interconnector `TRNBGDXXINDEA`. | The setup loader and the training reports. |
| `policies` | `support_bau_committed_transmission` sets the transmission ceilings of support `BAU` at the end of its A3 run. `reconcile_variable_cost_exports` adds the omitted default variable costs to the concatenated cost tables in B2. `validation_no_mutation` keeps the A3 stage 1b validation report-only. `historical_trade_roots` names the roots whose scenarios activate the 50 historical trade constraints in B2. `build_cap_replaces_workbook_values` decides what the build-cap rule does with a build-cap cell that already holds a number, when the root lists a cap multiplier schedule: with `true` the computed build cap replaces a number already in the workbook cell; with `false` the number is kept. | A3 and B2. |
| `authorities` | The sources of record: each named input, configuration file or folder and its location, as a `project:` path in the clone or a `package:` path in the installed package. | `ostram.paths`, for every stage. |
| `scenario_registry` | `support_scenarios`, `root_scenarios` with their roles and their result and restrictions dependencies, and `decision_scenarios` in canonical order. A decision scenario that is not a root is a derived scenario: its parent and, for the two direction variants, the overlay file and its study start year are in its own file. | The registry loader, for the `run` command, A3, B1 and B2. |
| `vre_ceilings` | The sheet, the parameter and the 20 values of the shared VRE ceilings. | The scenario edits, for every derived scenario. |

Each derived scenario has its own file, `config/scenarios/<name>/scenario.yaml`.
It holds `scenario`, `base_scenario`, `description`, `apply_vre_ceiling_layer`,
the `edits` list and, optionally, a `gross_import_cap` policy. Any derived
scenario may have one, and `B_Opt_Clipped_GrossImportCap15` is the delivered
case. The scenario edits in A3 read the file, and B2 reads the gross import
policy. The `description` and `note` strings are documentation.
Nothing reads them. The [scenario edits](scenario-edits.md) page describes the
edit fields and operations, and the [scenario catalogue](reference/scenarios/scenario_catalogue.md)
lists every scenario.
