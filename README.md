# OSTRAM

OSTRAM, OSeMOSYS for Transmission – Framework to represent power systems
connectivity, is a framework for power-system scenario studies. It builds on
[OSeMOSYS](https://osemosys.readthedocs.io/) and OSeMOSYS Global, adds detail
on internal transmission and on the interconnectors between countries, and
supports "what if" comparisons between scenarios. The South Asia model is its
reference application: Bangladesh, Bhutan, India in five regions, Nepal, Sri
Lanka and the Maldives, from 2023 to 2050 in 20 timeslices. A reduced
two-region training model lives in [`examples/unescap/`](examples/unescap/README.md).

## Install

Follow the [installation](docs/installation.md) page. In short: clone to a
short path, create the conda environment from `environment.yaml`, install the
package in editable mode and set `PYTHONUTF8=1`.

```powershell
git clone https://github.com/Climate-Lead-Group/OSTRAM.git OSTRAM
cd OSTRAM
conda env create -n ostram -f environment.yaml
conda activate ostram
python -m pip install -e .
$env:PYTHONUTF8 = '1'
python -m ostram inspect-resources
```

GLPK and CBC come with the environment. CPLEX and Gurobi are external solver
choices.

## Commands

Every command starts with `python -m ostram`. Pass the name of the environment
created above to `run`, so that the command checks that environment and
creates none.

```powershell
python -m ostram --help
python -m ostram inspect-resources
python -m ostram run --skip-pull --compile-only --env-name ostram --scenarios A_Calibrated_BAU
python -m ostram run --skip-pull --env-name ostram --scenarios A_Calibrated_BAU --verbose
python -m ostram transform --scenario A_Calibrated_BAU
python -m ostram compile-inputs --scenarios "A_Calibrated_BAU,B_Optimised_VRE"
```

`run` prepares the base model (A1), adds the transmission network (A2),
builds the selected scenarios (A3), compiles the model inputs (B1), and runs
the model and collects the results (B2). `--compile-only` stops after the
solver datafile and runs no solver. `--scenarios` takes one comma-separated
list. `C_Target_VRE` needs a solved `A_Calibrated_BAU` result, passed with
`--a-result-seed`. The [pipeline](docs/pipeline.md) page walks through the
stages, and the [scenario catalogue](docs/reference/scenarios/scenario_catalogue.md)
lists the 18 registered scenarios: the support scenario `BAU` and the 17
decision scenarios.

The global options `--project-root` and `--workspace` go before the command.
They override `OSTRAM_PROJECT_ROOT` and `OSTRAM_WORKSPACE`. Without them, the
clone is the project root and `workspace/` under it holds the generated
files. Every run writes `<workspace>/logs/<run-id>/run.log`.

## Layout

```text
ostram/       the Python package
inputs/       model inputs
config/       configuration, including config/profiles/full.yaml
model/        the OSeMOSYS model file
docs/         documentation
examples/     the two-region training profile
tests/        solver-free tests
workspace/    generated files, ignored by Git
```

The [file map](docs/file-map.md) describes each folder and key file.

## Checks

These commands run no solver and build no matrix.

```powershell
python -B -m compileall -q ostram tests
python -B -m unittest discover -s tests -p "test_*.py"
python -B -m tests.validation.test_scenarios_lite
python -m ostram inspect-resources
git diff --check
```

## Documentation

The documentation is in [`docs/`](docs/index.rst): [installation](docs/installation.md),
[quickstart](docs/quickstart.md), [pipeline](docs/pipeline.md),
[scenario rules](docs/rules.md), [scenario edits](docs/scenario-edits.md),
[scenario catalogue](docs/reference/scenarios/scenario_catalogue.md),
[adding a scenario](docs/adding-a-scenario.md),
[adapting OSTRAM to another system](docs/adapting.md),
[configuration](docs/configuration.md), [file map](docs/file-map.md) and
[glossary](docs/glossary.md).

The sources, conversions and assumptions behind the inputs are documented in
the scenario workbook `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx`: in its
README sheet, in its Source, Notes and References cells and in its cell
notes. `AGENTS.md` holds the rules for AI coding agents and `CONTRIBUTING.md`
the rules for changes.
