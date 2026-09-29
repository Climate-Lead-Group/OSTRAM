# ostram — the OSTRAM package: CLI, path resolution, profiles, terminal and the `run` orchestrator

`python -m ostram` is the only supported entry point (`README.md`, "Canonical interface"). This
folder holds the dispatcher and the base libraries every stage imports: path and workspace
resolution (`paths.py`), fail-closed profile manifests (`profiles.py`, `profile_workspace.py`),
terminal-safe progress and logging (`terminal.py`), the example-profile commands (`examples.py`),
the `run` orchestrator (`pipeline/orchestration.py`) and profile-domain validation
(`validation/profile.py`). The stage packages live under `pipeline/` (`preparation/`,
`scenarios/`, `compilation/`, `execution/`), each with its own README; `reporting/` holds the
result reports; `resources/` is read-only package data. Path contract: `docs/lineage.md`.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `__main__.py` | Parses the global prefix (`--project-root`, `--workspace`, `--profile`), activates the profile, exports `OSTRAM_PROJECT_ROOT` / `OSTRAM_WORKSPACE` and the profile environment for the duration of the command, and routes `run`, `transform`, `compile-inputs`, `example`, `country`, `scenario` and `inspect-resources` to their modules | The profile registry and manifest, the prepared-profile stamp; `inspect-resources` reads the registry, workbooks, configs and model header | Nothing on disk (environment variables and `sys.argv` are set and restored) | `python -m ostram [--project-root P] [--workspace W] [--profile ID] COMMAND ...` |
| `paths.py` | Resolves the project bundle and the mutable workspace (CLI > environment > editable checkout), exposes every input, config, model, resource and workspace path as `ProjectPaths` properties, applies profile authority overrides from `OSTRAM_PROFILE_AUTHORITIES`, and bounds workbook paths to the Windows path budget | `OSTRAM_PROJECT_ROOT`, `OSTRAM_WORKSPACE`, `OSTRAM_PROFILE_AUTHORITIES`; the bundle markers `ostram/__init__.py`, `environment.yaml`, `inputs/`, `config/`, `model/` | Only `stage_workspace(..., create=True)` creates directories under the workspace | Imported by almost every module (`resolve_paths()`, `ProjectPaths`) |
| `profiles.py` | Loads profile registries and manifests (`ostram-profile-registry-v1`, `ostram-profile-v1`), resolves namespaced authority references (`project:`, `profile:`, `package:`), encodes the profile environment and exposes `profile_policy()` | `config/profiles/registry.json` and the manifests it names; `OSTRAM_PROFILE*` environment variables | Nothing | Imported only (`__main__.py`, `examples.py`, stage modules) |
| `profile_workspace.py` | Prepares a mutable profile workspace atomically (`<workspace>/profiles/<id>/`, staging dir, stamp written last) and validates or refreshes it on every later command | The manifest's source authorities, the stamp `.ostram-profile.json`, hashes of mutable authorities | `<workspace>/profiles/<id>/authorities/...`, `.<id>.preparing`, `.<id>.reset-backup`, the stamp | Imported only (`__main__.py`, `examples.py`) |
| `terminal.py` | Live stage status on stderr (interactive) or append-only status lines (redirected), the UTF-8 run log, subprocess output capture with encoding-safe printing and interrupt handling | Stream encodings; no project files | `<workspace>/logs/<run-id>/run.log` | Imported only (`pipeline/orchestration.py`, `__main__.py`, `pipeline/execution/runner.py`) |
| `examples.py` | `example prepare <profile>` builds the profile workspace; `example report <profile>` captures result snapshots and renders the training dashboard, optionally comparing captures | `OSTRAM_PROFILE*`, result CSVs matching the manifest's `reporting.result_glob` (else `*Combined_Inputs_Outputs.csv` under the execution workspace), `<workspace>/reports/snapshots/*.csv` | `<workspace>/reports/snapshots/<label>.csv`, `<workspace>/reports/<profile>.html` or `<profile>-interconnector-comparison.html` | `python -m ostram example prepare <profile> [--reset]`; `python -m ostram example report <profile> [--capture LABEL | --compare L1,L2]` |
| `pipeline/orchestration.py` | The `run` route: checks Conda and the environment, optional DVC init/pull, A1+A2 when a required post-A2 snapshot is missing, then A3, B1 and B2 for the selected scenarios, each as `python -B -m <module>` with cwd under the workspace | `environment.yaml`, `dvc.yaml`, the scenario registry, `<workspace>/preparation/A1_Outputs/_post_a2_snapshot_*`, `CONDA_EXE` | The Conda environment, DVC state, the run log; every other output is written by the stage modules | `python -m ostram run [--env-name N] [--env-file F] [--dvc-file F] [--skip-pull] [--skip-a3] [--skip-b1] [--skip-b2] [--scenarios a,b] [--a-result-seed PATH] [--compile-only] [--verbose]` |
| `validation/profile.py` | Profile-domain validation: the seed `TECHNOLOGY` / `FUEL` sets must match the manifest's `domain_contract`, and the compiled sets must equal the seed plus the declared `generated_delta` | `<osemosys_inputs>/TECHNOLOGY.csv` and `FUEL.csv`, the compiled `TECHNOLOGY.csv` / `FUEL.csv`, `OSTRAM_PROFILE_MANIFEST` | Nothing (prints a skip notice) | Imported only (`profile_workspace.py` at `example prepare`, `pipeline/execution/orchestrator.py` before the solve) |

## How to run

From the repo root in `OSTRAM-env` (see `AGENTS.md`, "Canonical commands"):

```bash
python -m ostram --help
python -m ostram inspect-resources
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU
python -m ostram --profile unescap example prepare unescap
```

## Order / dependencies

`__main__.py` resolves paths and the profile before importing any route. `run`
(`pipeline/orchestration.py`) sequences A1 → A2 → A3 (`pipeline/scenarios/materializer.py`) →
B1 (`pipeline/compilation/runner.py`) → B2 (`pipeline/execution/runner.py`). Non-`full` profiles
need `example prepare` first; `validation/profile.py` gates the seed at prepare time and the
compiled domain at B2.

## Gotchas

- Global options must precede the command; `--profile` after the command is passed to the route's
  own parser (`tests/regression/test_profiles.py`).
- `run` requires Conda and creates or updates the environment even though the child stages use the
  active interpreter; it also pins `PYTHONHASHSEED=0` and puts the project root first on the
  child `PYTHONPATH`.
- `--skip-pull` skips DVC init and pull; without it, `run` initialises DVC metadata and pulls only
  when a remote is configured (`docs/installation.md`).
- `paths.py` fails closed when `OSTRAM_PROFILE_AUTHORITIES` is present but lacks a role; a
  non-editable install without a colocated bundle needs `--project-root`. A workbook path of 240
  or more UTF-16 units is compacted to `<stage>_<sha256[:16]>.xlsx`.
- A leftover `.<id>.preparing` or `.<id>.reset-backup` directory makes `profile_workspace.py` refuse
  to run; an existing prepared workspace needs `--reset`.
- `terminal.py` redacts arguments containing `token`, `password`, `secret` or `license` in the
  logged command.
- `examples.py` refuses to overwrite an existing snapshot label and requires the active profile to
  equal the requested one.
- `validation/profile.py` skips strict compiled-domain validation, with a printed notice, when the
  seed domain has grown beyond the contract (after `country merge`).
