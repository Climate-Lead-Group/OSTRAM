# Contributing

`AGENTS.md` holds the same rules for AI coding agents, with the pipeline
summary they need.

## Branches and pull requests

- Branch from `origin/main`. Never commit to `main`.
- Open a draft pull request. A person merges it; agents never merge.
- Never force-push, rewrite pushed commits, or delete branches or tags.

## Commits

- One logical change per commit. Moves and renames go in their own commits.
- Subjects: `type: summary`, imperative, under 72 characters. Types: `feat`,
  `fix`, `docs`, `chore`, `refactor`, `test`.
- Before committing, `git status --porcelain` shows only the intended paths.

## Tests

- Regression checks are the solver-free checks below and, for input changes,
  the steps in [After changing inputs](#after-changing-inputs).
- Every change: the solver-free checks pass. This is the one list of them.
  Other pages link here. They run no solver and build no matrix. Run them from
  the repository root, in the activated environment:

  ```powershell
  python -B -m compileall -q ostram tests
  python -B -m unittest discover -s tests -p "test_*.py"
  python -B -m tests.validation.test_scenarios_lite
  python -m ostram inspect-resources
  git diff --check
  ```

  `tests.validation.test_scenarios_lite` validates the seven compact scenarios
  without a matrix builder or a solver. `git diff --check` needs a Git
  checkout. Five tests in `tests/regression/test_materialization_paths.py` create
  folders named `tmp*` in the folder that holds the clone, and remove them
  when each test ends. That folder must be writable.

- Changes under `config/`, `inputs/` or `ostram/pipeline/`: build every
  registered scenario with `--compile-only` on `main` and on the branch, in
  separate clones, and compare `workspace/compilation/A2_Output_Params/<S>/`,
  `workspace/execution/A2_Outputs_Params_otoole/<S>/` and
  `workspace/execution/Executables/<S>_0/` byte for byte. The steps are in
  [After changing inputs](#after-changing-inputs).
- Documentation changes: the Sphinx build adds no warning. Build with
  `python -m sphinx -b html docs <out-dir>` and the packages in
  `docs/requirements.txt`, with `<out-dir>` outside the clone.
- Docstring-only changes: with docstrings stripped, every module's AST is
  unchanged.
- Scenario catalogue: the page lists the delivered scenarios, the rows of its
  origin table. The generator is safe to run at any time. It rewrites only the
  generated block and leaves out a registered scenario that has no row in the
  origin table. A new scenario appears in the catalogue only after its row is
  added to that table. After a delivered scenario's file changes, run
  `python tools/generate_scenario_catalogue.py` and commit the page.

## After changing inputs

This applies to changes under `config/`, `inputs/` or `ostram/pipeline/`.

1. Make two clones at short paths, such as folders directly under the root of a
   drive, with `main` checked out in one and the branch in the other.
2. Use one conda environment. It holds an editable install of one clone, so
   build each clone from its own root with `PYTHONPATH` set to that root.
3. In each clone, build every registered scenario. Omit `--scenarios`. The two
   C scenarios need a solved `A_Calibrated_BAU` result: pass the same folder
   to both clones. `B_Opt_TxCap150` and `B_Opt_TradeCap15` need a solved
   `B_Optimised_VRE` result and `A_Calibrated_BAU_Clipped_TxFreeze2026` a solved
   `A_Calibrated_BAU_Clipped` result: set the two environment variables below
   to the same folders in both clones. The
   [portfolio sequence](docs/pipeline.md#17-scenario-portfolio) shows
   how to make them.

   ```powershell
   $env:PYTHONUTF8 = '1'; $env:PYTHONPATH = (Get-Location).Path
   $env:OSTRAM_B_OPTIMISED_VRE_RESULT = '<folder of B_Optimised_VRE_0>'
   $env:OSTRAM_A_CALIBRATED_BAU_CLIPPED_RESULT = '<folder of A_Calibrated_BAU_Clipped_0>'
   python -m ostram run --skip-pull --compile-only --env-name <env> --a-result-seed <folder of A_Calibrated_BAU_0>
   ```

4. Compare the three folders listed under Tests for the two clones. Empty
   output means they are identical. `<=` marks a line from the `main` clone and
   `=>` a line from the branch clone. A path listed on both sides differs in
   content; a path listed once exists in one clone only.

   ```powershell
   function Get-TreeHash($root) {
       $root = (Resolve-Path -LiteralPath $root).Path
       Get-ChildItem -LiteralPath $root -Recurse -File | ForEach-Object {
           [pscustomobject]@{ Path = $_.FullName.Substring($root.Length); Hash = (Get-FileHash -LiteralPath $_.FullName).Hash }
       }
   }
   foreach ($part in 'compilation\A2_Output_Params', 'execution\A2_Outputs_Params_otoole', 'execution\Executables') {
       Compare-Object (Get-TreeHash "<main clone>\workspace\$part") (Get-TreeHash "<branch clone>\workspace\$part") -Property Path, Hash
   }
   ```

5. Files ending in `.warnings.txt` may differ in clone paths and timestamps
   only. Any other difference changes the model for the scenario named by the
   first folder of the path. Check that it is a scenario the change affects.

## Writing

- Short declarative sentences, active voice, UK spelling, exact numbers, no
  em dashes.
- Each term has one meaning, defined once in the [glossary](docs/glossary.md)
  and used the same way everywhere. Code names appear in code spans.
- 9999 is a finite cap.
- No cost figures from model results, no hash or checksum values and no
  machine paths in tracked files.
- Every number traces to a source in the repository.

## Scripts and data

- Scripts read their inputs and write new files beside them. They never
  modify files in `inputs/` or `config/` in place. Deliberate edits to those
  files are commits and go through the compile gate above.
- No hash or checksum files, no pins of file bytes in tests, and no
  re-pinning. Checks compare outputs live.
