# UNESCAP training operations

The training exercises use only `python -m ostram` commands. This table maps
each legacy script or manual step of the earlier training material to the
current command and to what the command writes under the workspace.

| Legacy script or step | Current command | Writes |
|---|---|---|
| Initial t1 confection and manual folder reset | `example prepare unescap`; guarded replacement with `--reset` | stamped `workspace/profiles/unescap/` bundle |
| Ad-hoc resource/path inspection | `--profile unescap inspect-resources` | JSON containing resolved manifest, prepared authorities, and stage paths |
| `Z_generate_country_template.py` | `--profile unescap country template` | `preparation/country_templates/<ISO3>/` |
| generated `templates/<ISO3>/merge_into_inputs.py` | `--profile unescap country merge <ISO3>` | merged mutable OSeMOSYS CSV authorities plus timestamped backups |
| `Z_validate_country_data.py` | `--profile unescap country validate <ISO3>` | terminal validation summary |
| `populate_v18_new_country.py` | `--profile unescap scenario sync-country --country <ISO3>` | prepared-workspace `OSTRAM_Scenario_Inputs.xlsx` |
| manual scenario-country row synchronization | same `scenario sync-country` route | six mapped scenario sheets synchronized from canonical post-A2 BAU A-O workbook |
| `set_interconnector_direction.py` and per-direction YAML copies | edit the profile scenario YAML, then `--profile unescap run --scenarios B_Optimised_VRE` | materialized B scenario A-O workbook and compiled/solved result |
| `A3_process.py` / scenario-specific direct scripts | `--profile unescap run --scenarios <name> --compile-only` | compiled B1 CSV domain and final `.txt` under the profile workspace |
| direct CBC invocation / legacy B2 solve wrapper | `--profile unescap run --scenarios B_Optimised_VRE --skip-pull` | CBC `.sol`, otoole CSV outputs, and combined result CSV under `execution/` |
| `ostram_training_dashboard.py` | `example report unescap` | `reports/unescap.html` |
| manual copying of combined result files | `example report unescap --capture <label>` | immutable labelled CSV under `reports/snapshots/` |
| `generate_direction_comparison.py` | `example report unescap --compare forward,reverse,bidirectional` | `reports/unescap-interconnector-comparison.html` |
| standalone training-dashboard launch | the same `example report` route | self-contained HTML with embedded `ostram-profile-report-v1` data |

## Exercise commands

The commands in `exercises/training.html`, `exercises/add-country.html`,
`exercises/add-interconnector.html` and `README.md` are checked against the
current parser by `tests/regression/test_release_readiness.py`. The retired
`--rebuild`, the invented `--force`, `country populate-workbook` and
`scenario list` spellings are rejected by that test and are absent from the
live instructions. The workbook synchronization command targets the ignored
prepared copy in the workspace, never the committed seed.

## Retired plotting scripts

The live exercises do not need these legacy plotting families, so they were
not restored: `ostram_trn_plotter.py`, `Z_AUX_interconnections_dashboard.py`,
`Z_AUX_generate_transmission_maps.py`, `Z_AUX_generate_RES_diagram.py`,
`Z_AUX_generate_interactive_dashboards_aggregated.py`, and the `trn_plots/` and
`ostram_plots/` script families. Their training outcome is covered by the
profile report and the interconnector comparison routes above.
