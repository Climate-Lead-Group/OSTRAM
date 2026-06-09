# BAU scenario configs

The BAU scenario currently chains only `add_max_cap_investment_lid_rule.py`,
which carries its configuration in-script (no YAML required).

If you add YAML-driven scripts to the BAU chain, place their config files here
using the script's `YAML_FILE_NAME` (e.g. `retirement_schedule.yaml`,
`bau_calibration.yaml`, `set_vre_targets.yaml`, `relax_interconnectors.yaml`).

Resolution order at runtime:
  1. `rules_scripts/configs/<scenario>/<YAML_FILE_NAME>`   (this folder)
  2. `rules_scripts/<YAML_FILE_NAME>`                       (default fallback)
