# Sensitivity patch materialization

`apply_patches.py` is the retained interim post-A3 materializer for derived
scenario patches. It rebuilds a target from a fresh source scenario, applies
the shared VRE ceiling from `reference/vre_ceilings_base.json`, then applies
the scenario's tracked `patches.json`.

The maintained inputs in this folder are:

- `apply_patches.py`
- `reference/vre_ceilings_base.json`
- `reference/vre_ceilings.csv`
- `reference/vre_ceiling_provenance.md`
- `reference/interconnector_direction_references.md`

Scenario-specific patch files remain under
`A3_process/rules_scripts/configs/<scenario>/patches.json`.

This helper does not run A3, B1, B2, or a solver. Canonical derived-scenario
selection, dependency handling, base selection, and direction overlays are
the responsibility of the Stage 2 scenario-automation work.
