# tests — solver-free test suites and the documentation checker

`regression/` holds the unittest-style regression suites (CLI, path resolution, A3, B1, B2,
profiles, reporting, accepted baseline), `validation/` the scenario-workbook lite checks and the
strip-storage CLI gate, and this folder the documentation checker. No test invokes a solver or
builds a matrix; the maintained validation commands are listed in `README.md`, "Safe validation",
and `docs/regression.md`. Each sub-folder has its own README with the behaviour every file guards.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `check_docs.py` | Checks that every code folder has a README listing its scripts, every module has a real docstring and every required public function is documented; `--compare-ast` proves a docstring-only diff; `--selftest` runs its fixture tests | `git ls-files` under `CODE_ROOTS` (`ostram`, `tests`, `timeslice_module/scripts`), the READMEs, the `.py` sources | Nothing (temp dir during `--selftest`) | `python tests/check_docs.py [--summary] [--paths P ...] [--compare-ast REF] [--selftest]` |

## How to run

From the repo root (any interpreter with git on `PATH`; the checker is stdlib-only):

```bash
python tests/check_docs.py                    # must exit 0
python tests/check_docs.py --compare-ast main # docstring-only commits
PYTHONPATH=tests python -m pytest tests -q    # the suites (see AGENTS.md for the recorded counts)
python -B -m unittest discover -s tests -p "test_*.py"
```

## Gotchas

- `check_docs.py` only sees files known to git; run it after `git add`, or a new `.py` is silently
  skipped.
- `tests/regression/test_support_bau_transmission.py` imports `regression.test_a3_orchestration`,
  which needs `tests/` on `sys.path` (hence `PYTHONPATH=tests`), while
  `test_restriction_prerequisites.py` imports `tests.regression...`, which needs the repo root.
  Running from the repo root with `PYTHONPATH=tests` satisfies both.
- Several tests write temporary directories inside the tree: under `tests/regression/`
  (`test_b1_runner.py`, `test_run_orchestration.py`, `test_run_terminal.py`), in the parent of the
  repo (`test_materialization_paths.py`) and under the gitignored `workspace/`
  (`test_paths.py`, which also leaves `workspace/stage11-lazy-test/` behind). See
  `regression/README.md`.
- Pre-existing failures on Windows with pandas 3 are listed in `AGENTS.md`, "Known gotchas".
