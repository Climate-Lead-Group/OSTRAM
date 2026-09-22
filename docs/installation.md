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

### Isolated reproduction environment

Keep the checkout, environment, package cache, logs and outputs separate from
production. From a new checkout, the following creates the documented
environment under an explicit local directory instead of updating an existing
environment. Set `CONDA_EXE` to the installed Conda executable if it is not
already available through an activated Conda prompt. Use a dedicated PowerShell
session. The recorded clean-clone execution cleared inherited OSTRAM/Python
path overrides and fixed Python's hash seed before launching the runner:

```powershell
Get-ChildItem Env: | Where-Object Name -Like 'OSTRAM_*' | Remove-Item
Remove-Item Env:PYTHONPATH, Env:PYTHONHOME -ErrorAction SilentlyContinue
$env:PYTHONUTF8 = '1'
$env:PYTHONHASHSEED = '0'
$env:PYTHONDONTWRITEBYTECODE = '1'
$env:PYTHONUNBUFFERED = '1'
$runRoot = [IO.Path]::GetFullPath('..') # parent of this isolated checkout
$env:CONDA_ENVS_PATH = Join-Path $runRoot 'conda-envs'
$env:CONDA_PKGS_DIRS = Join-Path $runRoot 'conda-pkgs'
conda env create -n ostram-reproduce -f environment.yaml -y
if ($LASTEXITCODE -ne 0) { throw 'Conda installation failed' }
$python = Join-Path $env:CONDA_ENVS_PATH 'ostram-reproduce\python.exe'
& $python -m pip install -e .
if ($LASTEXITCODE -ne 0) { throw 'Editable installation failed' }
& $python -m pip check
& $python -m pip freeze > (Join-Path $runRoot 'pip-freeze.txt')
& $python -m ostram inspect-resources
```

Use `& $python` instead of `python`, and add `--env-name ostram-reproduce`
to the `run` commands in the [portfolio sequence](pipeline.md#accepted-17-scenario-portfolio).
Record `conda list`, the Python version, `glpsol --version` and the selected
vendor solver version with each reproduction. Solver binaries and vendor
licenses are external runtime dependencies; an old prepared workspace is not
an installation dependency.

The 21 September 2026 clean-clone execution installed Python 3.10.21,
NumPy 2.2.6, pandas 2.3.3, openpyxl 3.1.5, otoole 1.1.5, PyYAML 6.0.3,
ruamel.yaml 0.19.1, XlsxWriter 3.2.9 and DVC 3.67.1. It used Conda GLPK 5.0
for matrix generation and the existing licensed CPLEX 22.1.2.0 installation,
with four threads, random seed 12345 and deterministic parallel mode.
Keep the complete environment/package export with run evidence; unbounded
dependency resolution on a later date is not an exact environment replay.

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
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU
```

## DVC

DVC is a package dependency and `dvc.yaml` uses only canonical package commands
and workspace paths. Without `--skip-pull`, the OSTRAM runner initializes local
DVC metadata when needed and pulls only when a remote is configured. With
`--skip-pull`, it performs neither DVC initialization nor a pull. It does not
use caller CWD to locate project data.

```powershell
dvc remote list
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU
```

Generated DVC cache/tmp state and the central workspace are ignored. Maintained
inputs, configuration, model files, and package resources remain tracked and
must not be replaced with generated copies.
