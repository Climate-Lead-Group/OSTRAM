# ostram/pipeline/scenarios — stage A3: registry, materialization and root transformation

A3 turns the post-A2 snapshot of each root scenario into its final `A-O_*` workbooks and derives
the remaining decision scenarios from their registered roots. `registry.py` validates the
contract in `config/profiles/full.yaml::scenario_registry`; `materializer.py` drives an exact
scenario selection (roots through `transform.py`, then derived scenarios through
`apply_patches.py` and the direction overlay); `transform.py` and `orchestrator.py` run the
ordered stages of one root, calling `transformations/` and `rules/`. Stage contract:
`docs/pipeline.md`, "Build selected scenarios (A3)" and "Scenario selection". All outputs go to
`<workspace>/preparation/A1_Outputs/A1_Outputs_<scenario>/`, with disposable run directories under
`<workspace>/scenarios/`.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `registry.py` | Loads and validates the scenario registry (roots, derived scenarios, decision order, result and restriction dependencies), resolves an exact canonical selection and its required roots, resolves declared result dependencies (`OSTRAM_A_CALIBRATED_BAU_RESULT` or `<execution-workspace>/Executables/A_Calibrated_BAU_0`) | `config/profiles/full.yaml::scenario_registry` (and `scenario_inputs` when validating files), the dependency environment variables | Only the four `A1_Outputs_<root>/` directories (`ensure_root_output_directories`) | Imported only (`materializer.py`, `ostram/pipeline/orchestration.py`, A1, A2, `rules/apply_base_year_pin.py`) |
| `materializer.py` | Materializes the selection: validates the workbook `Control` sheet against the registry, requires every root's post-A2 snapshot, resolves result dependencies, runs `transform.py` per root (wiring exported restriction states between roots), then builds each derived scenario with `apply_patches.build_scenario` and applies its direction overlay, and writes provenance | The registry, the scenario workbook's `Control` sheet, `_post_a2_snapshot_<root>`, `OSTRAM_A_CALIBRATED_BAU_RESULT` / `--a-result-seed` | `A1_Outputs/A1_Outputs_<scenario>/` (via `transform.py` and `apply_patches.py`), the AR workbooks of direction scenarios, `A1_Outputs/_scenario_materialization.json`, ephemeral `<scenarios workspace>/_restriction_states/` | `python -m ostram run` (stage A3) as `python -B -m ostram.pipeline.scenarios.materializer [--scenarios a,b] [--a-result-seed PATH]` |
| `transform.py` | Single-root A3 runner: builds the run directory, launches every transformation and rule step as a `python -B -m` subprocess in the documented order, applies inherited restrictions, syncs the timeslice fabric to the A1 CSVs and the compilation YAML, persists run restrictions to a disposable copy, delivers the four workbooks and finalises support `BAU` | `ProjectPaths.scenario_workbook` (`Control`: `rules_script`, `inherit_restrictions_from`), the timeslice workbook, `config/scenarios/technology_types.csv`, `config/scenarios/<scenario>/<rule>.yaml`, `config/scenarios/rules/*.csv`, `Config_country_codes.yaml`, `OSTRAM_TEMPLATE_PATH`, profile policies | `<workspace>/scenarios/_run_<ts>/{stage1,stage1b,stage2,stage3,stage5}`, `A1_Outputs/A1_Outputs_<scenario>/A-O_*.xlsx` (or `--output-dir`), `_scenario_run_state_<scenario>.xlsx`, optional `--run-state-out`, and through `sync_og_to_ts20` the `og_csvs_inputs/` CSVs and `<workspace>/compilation/Config_MOMF_T1_A.yaml` | `python -m ostram transform --scenario <root> [--soasia F] [--rules-script a.py,b.py] [--inherit-from S] [--input-dir D] [--output-dir D] [--keep-workdir] [--run-state-out F] [--restrictions-source S=PATH]`; also launched by `materializer.py` |
| `orchestrator.py` | Behaviour-preserving planning and stage sequencing for one root: plan resolution, snapshot restoration, the fixed Stage 0-6 dispatch, delivery and cleanup, with the real work injected through `A3Dependencies` | `A1_Outputs/_post_a2_snapshot_<scenario>`, the input directory, the scenario workbook | Wipes and recreates `A1_Outputs_<scenario>/` from the snapshot, the run directory, `_materialized_<scenario>.xlsx`, `_scenario_run_state_<scenario>.xlsx` | Imported only (`transform.py`) |
| `apply_patches.py` | Builds a derived scenario's `A1_Outputs_<scenario>/` as a fresh copy of its registered `base_scenario` output, applies the shared `vre_ceilings` layer and the scenario's `edits` (`multiply`, `set_flat`, `set_to_residual`, `clamp_to_residual`, `set_to_residual_factor_floor`, explicit `values`), then re-applies the audited base-year pin and writes a change log | `config/profiles/full.yaml::scenario_inputs[scenario]` and `::vre_ceilings`, `A1_Outputs_<base>/A-O_Parametrization.xlsx`, the minimum-boundary authority when an edit asks for it, `config/scenarios/rules/pwr_min_2023_2026_pin.csv` | `A1_Outputs/A1_Outputs_<scenario>/` (replaced), `apply_patches_CHANGES_<stamp>.json`, `A1_Outputs_<scenario>_PREPATCH_<stamp>` backup unless skipped | Imported by `materializer.py`; CLI `python -m ostram.pipeline.scenarios.apply_patches --scenario NAME [--source-scenario BASE] [--skip-backup] [--restore] [--self-test]` |

## How to run

From the repo root in `OSTRAM-env`, after A1/A2 have produced the post-A2 snapshots:

```bash
python -m ostram transform --scenario A_Calibrated_BAU                                   # one root
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU             # A3 via materializer, then B1
python -m ostram run --skip-pull --scenarios <16 others> --a-result-seed <A result dir>  # C needs a solved A
```

Omitting `--scenarios` selects support `BAU` plus every decision scenario (`registry.select`).

## Order / dependencies

1. `materializer.py`: `registry.select` → `validate_control_roots` → `required_roots` (base roots
   plus transitive restriction prerequisites, in `Control` order) → snapshot check →
   `result_dependencies` (fails closed for `C_Target_VRE` without a solved A result).
2. Per root, `transform.py` through `orchestrator.execute_plan`: restore the input dir from
   `_post_a2_snapshot_<root>` → Stage 0 materialize the template → Stages 1, 1b, 2/2.5, 3
   (`transformations/`, see its README) → Stage 4 consolidation → Stage 4.5 inherited
   `Restrictions` → Stage 5 the `Control` `rules_script` chain, left to right (`rules/`) → WS-3
   interconnector costs, internal transmission, internal losses → WS-4 base-year pin (A, B, C
   roots only) → `sync_og_to_ts20` → persist restrictions → deliver → `finalize_support_bau`
   (`BAU` only).
3. Per derived scenario, in canonical order: `apply_patches.build_scenario` → direction overlay
   (`rules/set_interconnector_direction.run`) for `B_Opt_DirBidir` and `B_Opt_DirContractual`.

## Gotchas

- `C_Target_VRE` requires a completed `A_Calibrated_BAU` result: `OSTRAM_A_CALIBRATED_BAU_RESULT`,
  `--a-result-seed`, or the default `Executables/A_Calibrated_BAU_0` under the execution workspace
  (`registry.result_dependencies`). Solve A first (`docs/pipeline.md`).
- The `Control` sheet of the scenario workbook must list exactly the registry roots, all active;
  its `rules_script` column defines the Stage 5 chain per root. `BAU` has no
  `config/scenarios/BAU/` folder, so only the lid rule (which tolerates a missing YAML) can run
  for it; other YAML-driven rules raise `FileNotFoundError`.
- `A1_Outputs_<root>/` is wiped and rebuilt from the snapshot on every run; nothing is incremental.
  `apply_patches.py` likewise replaces the derived scenario folder from a fresh copy of its base.
- `transform.py` requires `OSTRAM_TEMPLATE_PATH` (set by Stage 0) or a profile
  `interconnector_authority` for Stage 3; without either it raises.
- Stage 4.5 overwrites the Stage 1b `9999` placeholders in the same `TotalAnnualMaxCapacityInvestment`
  cells; this is intended (`transform.py`).
- `sync_og_to_ts20` writes outside the disposable run directory (A1 CSVs and the compilation YAML in
  the workspace) on every root run.
- Stage 6 persists only `*_PRE_*_CHANGES.json` logs written by rules that ran with backups; the
  WS-3/WS-4 stages and the pin run with `--skip-backup` and emit none. Persistence targets the
  disposable `_scenario_run_state_<scenario>.xlsx`, never the maintained workbook.
- `config/profiles/full.yaml::scenario_inputs` contains entries (`B_Opt_LinkFreeze`,
  `B_Opt_SolarHi10`, `B_Opt_TradeCap30`, `B_Opt_TradeCap50`) that are not registry-derived; the
  materializer never builds them, only the `apply_patches.py --scenario` CLI can.
- Docstrings in `transform.py` and the rules still mention `rules_scripts/configs/<scenario>/`;
  the code reads `config/scenarios/<scenario>/`. The `transform.py` header also lists two
  `add_max` copies that are now one merged call.
- `materializer.py` defines `A3_ENTRYPOINT` / `A3_PROCESS` constants that nothing else appears to
  use.
