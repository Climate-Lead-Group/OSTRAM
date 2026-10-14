# Changes

This repository has no GitHub Releases. This file lists the changes by
release and by pull request.

## Unreleased

- PR #43. Each A3 rule stops on a configuration key that its code does not
  read, and names the file and the key. A scenario edit stops when it names a
  year that is not a year column. B1 stops with a message when neither
  ruamel.yaml nor PyYAML is installed. B2 accepts the previous names of the
  reserve-margin adjust keys. The `relax_interconnectors` mode `unconstrained`
  became `cap_9999`, with the old name still accepted, and the rule runs from
  the package. When the source of a training authority changes, an edited
  workspace copy moves to `authorities_previous` before the source is copied
  in. A repeated `example prepare` or capture prints its message without a
  traceback. Unread configuration keys, comments that cited past runs, em
  dashes in configuration comments and figures without a source were removed.
  All registered scenarios compile to identical model files.
- PR #42, commit `0ef3552`, 2026-10-13. The installation page uses Miniforge
  and PowerShell, tells the reader to stop when the environment cannot be
  created and checks the environment prefix. The isolated reproduction
  resolves its parent folder from the PowerShell location. The training notes
  describe what happens to the workspace after a change to the profile or its
  inputs. `CONTRIBUTING.md` gained the section After changing inputs.
- PR #41, commit `41d72a6`, 2026-10-12. Stored hashes and counts were removed
  from the base-year pin, the authority loaders, the training profile and the
  test comparator; the per-row checks stay. A changed `profile.yaml` keeps the
  prepared workspace when its mutable authority roles are the same and moves
  it aside when they differ. The standalone scenario edit command prints its
  stop message without a traceback. Three timeslice scripts lost their machine
  paths. Stage names replaced the WS-3 and WS-4 labels, "patch" wording and em
  dashes left the code strings, and docstrings were corrected or added. All
  registered scenarios compile to identical model files.
- Documentation for a stranger: the glossary, the file map, the stage
  walk-through on the pipeline page, the scenario rules page, the scenario
  edits reference, the `full.yaml` section map, the generated scenario
  catalogue with its test, the pages on adding a scenario and on adapting
  OSTRAM to another system, `CONTRIBUTING.md`, and corrections to the
  existing pages, `README.md`, `AGENTS.md` and the notes of the training
  profile. Five scenario descriptions in `config/profiles/full.yaml` and the
  README sheets of two input workbooks were reworded. No pipeline code and
  no model input changed.

## October 2026 cleanup

- PR #39, commit `935f62a`, 2026-10-11. `apply_patches.py` became
  `scenario_edits.py`, `execution/patches/` became `model_file_adjustments/`,
  and the "patch" names in code and configuration became "scenario edits",
  "model-file adjustments" and "rules". The generator settings
  `include_fuel`, `cap_fraction` and `export_factor` and the registry field
  `patches:` were removed from `full.yaml`. A scenario edit that matches
  nothing now stops the run with a message. The conda executable is found in
  a standard Windows install. Tests were added for the scenario edit check,
  for the list of files under `config/` and for the conda lookup. All 18
  scenarios compile to identical model files.
- PR #38, commit `89eee70`, 2026-10-11. Dead scenario configuration and
  superseded working notes were removed.
- PR #37, commit `003fdbe`, 2026-10-01. `AGENTS.md` was added.

## September 2026 release

- Tag `v2026.09-release`, commit `3a428dc`, 2026-09-19. The accepted
  17-scenario cost release with native country fuel accounting. The scenario
  workbook and the monetary CSVs carry the revised costs in constant 2023 US
  dollars. The transmission freeze and gross import cap scenarios were
  registered, and the scenario edits were consolidated into
  `config/profiles/full.yaml`.

## August 2026

- 2026-08-13. `ostram/reporting/summary_workbook.py` was added: a workbook
  writer over the training dashboard metrics, run as
  `python -m ostram.reporting.summary_workbook`.
