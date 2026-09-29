# tests/validation — scenario workbook lite checks and the strip-storage CLI gate

Two standalone check scripts. `test_scenarios_lite.py` exercises the multi-scenario workbook
reader on temporary copies of the maintained workbook without running the pipeline;
`test_strip_storage_cli.py` runs the storage-stripping patch as a subprocess on a preprocessed
datafile and verifies the result structurally. Both are also collected by pytest.

## Scripts

| Script | What it does | Reads | Writes | Run |
|---|---|---|---|---|
| `test_scenarios_lite.py` | Guards `scenario_workbooks`: bad-config errors (unknown inherit reference, duplicate scenario, empty restriction source), inheritance-cycle detection, topological order, and the `BAU` materialization contract (v18 runtime sheets and headers) | `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx` (copied to temp dirs) | Temp dirs only | `python -B -m tests.validation.test_scenarios_lite` (exit 1 on failure) or `python -m pytest tests/validation/test_scenarios_lite.py -q` |
| `test_strip_storage_cli.py` | Runs `ostram.pipeline.execution.patches.strip_storage` on a datafile and checks that the stripped storage set, storage and link parameters, auxiliary sets, zeroed `TotalAnnualMaxCapacity` rows, minimum parameters and bare `PWRSDS`/`PWRLDS` mentions are all clean | The input datafile (positional), `patches/strip_storage.py` constants and parsers | The `-o` output datafile written by the subprocess | `python tests/validation/test_strip_storage_cli.py INPUT -o OUTPUT --mode {tech,class,all} [--targets T ...]` (exit 0 pass, 1 fail) |

## How to run

From the repo root in `OSTRAM-env`:

```bash
python -B -m tests.validation.test_scenarios_lite
python tests/validation/test_strip_storage_cli.py <workspace>/execution/Executables/<S>_0/Pre_processed_<S>_0.txt -o /tmp/out.txt --mode all
```

## Gotchas

- The `test_scenarios_lite.py` functions return `(ok, message)` tuples, so pytest reports
  `PytestReturnNotNoneWarning` for each; its docstring says the expensive checks (a full A3 run)
  are exercised separately with `python -m ostram run`.
- `test_scenarios_lite.py` assumes the maintained workbook has a seeded `BAU` scenario.
- `test_strip_storage_cli.py` rejects `--strip-script` overrides ("not supported by the
  package-layout gate") and needs the input file and the package patch module to exist; its
  docstring examples still use pre-refactor file names. Under pytest it collects no test functions.
