# Installation

## Supported environment

OSTRAM requires Python 3.10 or newer. The maintained Conda definition is
`environment.yaml`; Python package requirements are also declared in
`pyproject.toml`.

```powershell
conda env create -f environment.yaml
conda activate OSTRAM-env
python -m pip install -e .
```

For an existing environment:

```powershell
conda env update -f environment.yaml --prune
python -m pip install -e .
```

The editable install is important: it makes `ostram` and its package data
import-addressable from any current directory while project authorities remain
in the checkout.

## Verify the installation

```powershell
python -m ostram --help
python -m ostram inspect-resources
python -B -m unittest discover -s tests -p "test_*.py"
```

From outside the checkout, supply the bundle explicitly:

```powershell
python -m ostram --project-root C:\path\to\OSTRAM inspect-resources
```

Use `--workspace` or `OSTRAM_WORKSPACE` to place mutable state elsewhere.
Explicit command-line values override environment variables. Inspection is
read-only and does not create the workspace.

## Solvers

`environment.yaml` installs GLPK and CBC. CPLEX and Gurobi require their vendor
installations and licenses. Select the solver in
`config/execution/Config_MOMF_T1_AB.yaml`.

Verify an external solver before a production run:

```powershell
glpsol --version
cbc -stop
```

Use the compile-only route when validating without a solver:

```powershell
python -m ostram run --skip-pull --compile-only
```

## DVC

DVC is a package dependency and `dvc.yaml` uses only canonical package commands
and workspace paths. The OSTRAM runner initializes local DVC metadata when
needed and pulls only when a remote is configured. It does not use caller CWD
to locate project data.

```powershell
dvc remote list
python -m ostram run --skip-pull --compile-only
```

Generated DVC cache/tmp state and the central workspace are ignored. Maintained
inputs, configuration, model files, and package resources remain tracked and
must not be replaced with generated copies.
