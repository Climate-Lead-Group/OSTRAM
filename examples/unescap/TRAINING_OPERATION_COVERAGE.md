# UNESCAP training-operation coverage

This ledger records the release-readiness audit against the read-only legacy source at
`OSTRAM_training_source`. The canonical interface is always `python -m ostram`; no
legacy Python generator is copied into this example.

| Legacy exercise operation or generator | Canonical package / CLI operation | Expected workspace artifact or report | Status |
|---|---|---|---|
| Initial t1 confection and manual folder reset | `example prepare unescap`; guarded replacement with `--reset` | stamped `workspace/profiles/unescap/` bundle | covered; repeat and reset protection tested |
| Ad-hoc resource/path inspection | `--profile unescap inspect-resources` | JSON containing resolved manifest, prepared authorities, and stage paths | covered; caller-CWD independent |
| `Z_generate_country_template.py` | `--profile unescap country template` | `preparation/country_templates/<ISO3>/` | covered; generated helper remains in ignored workspace only |
| generated `templates/<ISO3>/merge_into_inputs.py` | `--profile unescap country merge <ISO3>` | merged mutable OSeMOSYS CSV authorities plus timestamped backups | covered by reusable package command |
| `Z_validate_country_data.py` | `--profile unescap country validate <ISO3>` | terminal validation summary | covered by reusable package command |
| `populate_v18_new_country.py` | `--profile unescap scenario sync-country --country <ISO3>` | prepared-workspace `OSTRAM_Scenario_Inputs.xlsx` | covered; dry-run, schema checks, atomic replacement, and idempotence |
| manual scenario-country row synchronization | same `scenario sync-country` route | six mapped scenario sheets synchronized from canonical post-A2 BAU A-O workbook | covered; deterministic BAU selection |
| `set_interconnector_direction.py` and per-direction YAML copies | edit the profile scenario YAML, then `--profile unescap run --scenarios B_Optimised_VRE` | materialized B scenario A-O workbook and compiled/solved result | covered by registered A3 rule; no direct-script command |
| `A3_process.py` / scenario-specific direct scripts | `--profile unescap run --scenarios <name> --compile-only` | compiled B1 CSV domain and final governed `.txt` under the profile workspace | covered; solver-free boundary is explicit |
| direct CBC invocation / legacy B2 solve wrapper | `--profile unescap run --scenarios B_Optimised_VRE --skip-pull` | CBC `.sol`, otoole CSV outputs, and combined result CSV under `execution/` | covered; CBC is discovered in the active environment on Windows and POSIX |
| `ostram_training_dashboard.py` | `example report unescap` | `reports/unescap.html` | covered with synthetic and real result routes |
| manual copying of combined result files | `example report unescap --capture <label>` | immutable labelled CSV under `reports/snapshots/` | covered; labels are validated and overwrite is refused |
| `generate_direction_comparison.py` | `example report unescap --compare forward,reverse,bidirectional` | `reports/unescap-interconnector-comparison.html` | covered by the canonical self-contained dashboard using only the selected captures |
| standalone training-dashboard launch | the same `example report` route | self-contained HTML with embedded `ostram-profile-report-v1` data | covered; no web server or caller-CWD dependency |

## Exercise command audit

The commands in `exercises/training.html`, `exercises/add-country.html`,
`exercises/add-interconnector.html`, and `README.md` were checked against the current
parser help. Retired `--rebuild`, invented `--force`, `country populate-workbook`, and
`scenario list` spellings are rejected by regression tests and are absent from the live
instructions. The workbook synchronization command deliberately targets the ignored
prepared copy, never the committed seed.

## Optional retired plotting families

The live exercises do not require these legacy plotting families, so they were not
restored: `ostram_trn_plotter.py`, `Z_AUX_interconnections_dashboard.py`,
`Z_AUX_generate_transmission_maps.py`, `Z_AUX_generate_RES_diagram.py`,
`Z_AUX_generate_interactive_dashboards_aggregated.py`, and the `trn_plots/` and
`ostram_plots/` script families. Their required training outcome is covered by the
profile report and interconnector comparison routes above.
