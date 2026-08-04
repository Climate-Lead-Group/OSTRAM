# OSTRAM

OSTRAM prepares, transforms, compiles, and executes OSeMOSYS scenarios. The
repository is an installable Python project: maintained inputs, configuration,
model text, package resources, and mutable outputs have separate locations.

## Install

Create the Conda environment and install the checkout in editable mode:

```powershell
conda env create -f environment.yaml
conda activate OSTRAM-env
python -m pip install -e .
```

The package metadata declares the Python dependencies. GLPK, CBC, CPLEX, and
Gurobi remain external solver choices; see [installation](docs/installation.md).

## Canonical interface

All supported commands begin with `python -m ostram`:

```powershell
python -m ostram --help
python -m ostram inspect-resources
python -m ostram run --skip-pull --compile-only
python -m ostram run --verbose
python -m ostram transform --scenario BAU
python -m ostram compile-inputs --scenarios "BAU,B_Optimised_VRE"
```

Global path options precede the command:

```powershell
python -m ostram `
  --project-root C:\path\to\OSTRAM `
  --workspace "D:\OSTRAM work\run α" `
  inspect-resources
```

`--project-root` overrides `OSTRAM_PROJECT_ROOT`; `--workspace` overrides
`OSTRAM_WORKSPACE`. Without either, an editable checkout supplies the project
root and `<project-root>/workspace` is selected lazily. Caller CWD is never a
resource root.

An actual run reports live progress through preparing the base model (A1),
adding the transmission network (A2), building scenarios (A3), compiling model
inputs (B1), and running the model/collecting results (B2). Redirected output is
plain and append-only. `--verbose` streams complete child output and command
diagnostics. Every mode retains a detailed UTF-8 log at
`<workspace>/logs/<run-id>/run.log`; help and resource inspection create no log.

## Layout

```text
ostram/       Python package and read-only package resources
inputs/       authoritative model and scenario inputs
config/       preparation, scenario, compilation, and execution configuration
model/        maintained OSeMOSYS model source
workspace/    ignored mutable runtime state, created only when needed
tests/        solver-free regression and validation suites
```

Important authorities include:

- `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx`
- `inputs/scenarios/OSTRAM_Timeslice_Inputs.xlsx`
- `config/scenarios/registry.json`
- `config/compilation/Config_MOMF_T1_A.yaml`
- `config/execution/Config_MOMF_T1_AB.yaml`
- `model/osemosys_fast_preprocessed.txt`
- `ostram/resources/compilation/conversion_format.yaml`

Code reads these locations through `ostram.paths`; generated state belongs
under the selected workspace. Package resources are opened read-only through
the installed package.

## Safe validation

The compact checks below do not invoke a solver or build a matrix:

```powershell
python -B -m compileall -q ostram tests
python -B -m unittest discover -s tests -p "test_*.py"
python -B -m tests.validation.test_scenarios_lite
python -m ostram inspect-resources
git diff --check
```

See the [quickstart](docs/quickstart.md), [pipeline](docs/pipeline.md),
[configuration](docs/configuration.md), and [lineage](docs/lineage.md) for the
maintained operating contract.
