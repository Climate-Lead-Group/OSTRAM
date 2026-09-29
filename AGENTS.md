# AGENTS.md

Instructions for AI coding agents working on OSTRAM. People should start with `README.md`.

## Project

OSTRAM (OSeMOSYS Transmission Model) builds on OSeMOSYS and OSeMOSYS Global. It adds detail on internal transmission and interconnectors, and supports "what if" comparisons between scenarios. The South Asia model, the default `full` profile, is the reference application. `examples/unescap` is a separate, reduced two-region training profile.

The pipeline runs in this order:

1. A1 `ostram.pipeline.preparation.base_inputs`: source CSVs and workbook templates become initial A-O workbooks for every root scenario.
2. A2 `ostram.pipeline.preparation.transmission`: adds transmission.
3. A3 `ostram.pipeline.scenarios.materializer`: builds each root with its rule scripts, then each derived scenario from its root.
4. B1 `ostram.pipeline.compilation.runner`: A-O workbooks become OSeMOSYS CSVs.
5. B2 `ostram.pipeline.execution.runner`: CSVs become the patched solver datafile. B2 then solves and collects results: `glpsol --wlp` writes the LP, and CPLEX solves it (the `unescap` profile uses CBC). `--compile-only` stops B2 after the datafile.

## Layout

- `ostram/`: the package.
- `config/profiles/full.yaml`: scenario registry, derived-scenario patches, VRE ceilings and run policies.
- `config/scenarios/<root>/`: rule YAMLs for the roots `A_Calibrated_BAU`, `B_Optimised_VRE` and `C_Target_VRE`. The fourth root, support `BAU`, has no folder; its rule script runs on its built-in defaults.
- `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx`: the scenario workbook. Its Control sheet lists each root's rule scripts in run order.
- `tests/`: tests. `docs/`: Sphinx documentation.
- `workspace/`: generated outputs, git-ignored. Never commit it.

## Scenarios

- A root runs the A3 rule scripts named in the Control sheet. Each script reads its YAML from `config/scenarios/<root>/`.
- A derived scenario copies its root's finished A3 output, applies the VRE ceiling layer, applies its ordered edits from `full.yaml` under `scenario_inputs/<name>`, then re-applies the base-year pin. Its parent is always a root. `B_Opt_DirBidir` and `B_Opt_DirContractual` then also get their direction overlay.
- `C_Target_VRE` and `C_Target_VRE_Clipped` need a solved `A_Calibrated_BAU` result, passed with `--a-result-seed`.

## Environment (Windows)

- Pass an existing conda env with `--env-name <env>`, after `run`, and run `python -m ostram` with that env's Python: every stage runs under the calling interpreter. Without the flag, the CLI checks `OSTRAM-env`, the `name:` in `environment.yaml`. Whichever env it checks, it runs `conda env create` if that env is missing and installs any missing dependency into it. `--skip-pull` skips only DVC.
- The CLI takes conda from `CONDA_EXE` if that names an existing file, otherwise `conda.exe` or `conda` on `PATH`, skipping `.bat` and `.cmd` matches such as those in `condabin`. Set `CONDA_EXE`, or add `<anaconda>\Scripts` to `PATH`.
- Before pipeline commands, set `PYTHONUTF8=1`, as `docs/installation.md` and CI do.
- If the env holds another editable install of `ostram`, set `PYTHONPATH` to the clone root and run from the clone root. Check with `python -c "import ostram; print(ostram.__file__)"`.
- Unset every `OSTRAM_*` environment variable unless the task sets one on purpose.
- Clone to a short path. Some output paths come close to the Windows path limit.

## Commands

- Build model inputs without solving: `python -m ostram run --skip-pull --compile-only --env-name <env> --scenarios <name1,name2> [--a-result-seed <dir>]`. `--scenarios` takes one comma-separated list.
- Rebuild one derived scenario from an existing root output: `python -B -m ostram.pipeline.scenarios.apply_patches --scenario <name> --source-scenario <root>`, then `python -m ostram compile-inputs --scenarios <name>`. Direction variants also need the direction overlay (`set_interconnector_direction.run`).
- Tests, no solver: `python -m unittest discover -s tests -p 'test_*.py'`
- Documentation: `python -m sphinx -b html docs <out-dir>`, with the packages in `docs/requirements.txt`. Put `<out-dir>` outside the clone; `docs/_build` is not git-ignored.

## Rules

- Do not run the solver unless the task asks for it. Use `--compile-only`.
- Changes under `config/`, `inputs/` or `ostram/pipeline/` need the compile gate below before they are pushed.
- Rule order in A3 is load-bearing. Retirement runs before the build cap, because the build cap reads the ResidualCapacity that retirement writes.
- In the LP, `TotalAnnualMaxCapacityInvestment = 0` forbids new capacity. Only -1, the default for a row that B1 does not write, leaves it unconstrained. B1 does not write rows whose `Projection.Mode` is `Zero`, `EMPTY` or blank. Some explicit zeros are rewritten before the LP: in A3 by Stage 1b `--fill-zeros` and the build-cap (lid) rule, and in B2 for `PWRBCK*`, `PWRPET*`, `PWROIL*` and `PWRNGS*`. Before relying on a 0, check the final `Pre_processed_*` datafile in `Executables/<S>_0/`.
- To freeze capacity at its existing level, set `TotalAnnualMaxCapacity` equal to `ResidualCapacity`, as `cap_trn_to_residual` does.
- Scripts you write read their inputs and write new files beside them. They never modify files in `inputs/` or `config/` in place. Deliberate edits to those files are commits and go through the compile gate.
- Do not add machine-specific paths such as `C:\Users\...` to tracked files.

## Git

- Branch from `origin/main`. Never commit to `main`. Changes reach `main` through pull requests.
- Agents never merge pull requests; a person does.
- One logical change per commit. Before committing, `git status --porcelain` shows only the intended paths.
- Never force-push, rewrite pushed commits, or delete branches or tags.
- Commit subjects: `type: summary`, imperative, under 72 characters. Types: `feat`, `fix`, `docs`, `chore`, `refactor`, `test`.

## Verification gates

- Every change: the solver-free tests pass.
- Docstring-only commits: with docstrings stripped, every module's AST is unchanged.
- Config, input or pipeline changes: build all 18 registered scenarios (support `BAU` plus the 17 decision scenarios) with `--compile-only`, on `main` and on the branch, in separate clones. Compare `workspace/compilation/A2_Output_Params/<S>/`, `workspace/execution/A2_Outputs_Params_otoole/<S>/` and `workspace/execution/Executables/<S>_0/` byte for byte. Warnings files may differ only in clone paths and timestamps.
- Documentation changes: the Sphinx build adds no new warnings.
