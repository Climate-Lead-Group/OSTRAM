# Installation

## Supported environment

OSTRAM requires Python 3.10 or newer, conda, Git and a solver. The conda
environment is defined in `environment.yaml`, and the Python package
requirements are also declared in `pyproject.toml`. The steps below are for
Windows PowerShell. Terms are in the [glossary](glossary.md).

## Install

1. Install Git and Miniforge, a conda installer that uses only the conda-forge
   channel. Download Miniforge from <https://conda-forge.org/download/> and
   keep the default options. Open the Miniforge Prompt from the Start menu, run
   `conda init powershell` once, and close it.

   If PowerShell later reports that running scripts is disabled, run
   `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once in PowerShell.

   Open a new PowerShell window. `conda --version` and `conda env list` work
   there. Run every command on this page in that window.

   If you already use Miniconda or Anaconda and `conda env create` stops with a
   message about Anaconda's terms of service, install Miniforge, or accept
   those terms if your organisation's licence covers it.

2. Clone into a short path, such as a folder directly under the root of a
   drive. Some output paths of the scenario
   `A_Calibrated_BAU_Clipped_TxFreeze2026` come close to the Windows path
   limit, and a long clone path makes them fail.

   ```powershell
   git clone https://github.com/Climate-Lead-Group/OSTRAM.git OSTRAM
   cd OSTRAM
   ```

3. Create the environment and install the package in editable mode. The
   editable install makes `ostram` importable from any folder while the
   inputs and the configuration stay in the clone.

   ```powershell
   conda env create -n ostram -f environment.yaml
   ```

   If this command fails, stop. Every later command needs the environment.

   ```powershell
   conda activate ostram
   python -c "import sys; print(sys.prefix)"
   python -m pip install -e .
   python -m pip check
   ```

   The printed path ends with `envs\ostram`.

4. Set the encoding and check the installation. Set `PYTHONUTF8` again in
   every new window before a pipeline command.

   ```powershell
   $env:PYTHONUTF8 = '1'
   python -m ostram --help
   python -m ostram inspect-resources
   glpsol --version
   cbc -stop
   ```

5. Run the solver-free tests:

   ```powershell
   python -B -m unittest discover -s tests -p "test_*.py"
   ```

Pass `--env-name ostram` to every `python -m ostram run`. Without it, the
command checks the environment named in `environment.yaml` and creates that
environment when it is missing. Every stage runs under the Python that
called the command, so run the commands from the activated environment.

For an existing environment, update it with:

```powershell
conda env update -n ostram -f environment.yaml --prune
python -m pip install -e .
```

## How the command finds conda

`python -m ostram run` calls conda to check the environment. It takes conda
from `CONDA_EXE` when that variable names an existing file; conda sets it in
every activated prompt. Otherwise it takes the first of `conda.exe`, `conda`
and `conda.bat` found on `PATH`. A `.bat` or `.cmd` hit, such as
`condabin\conda.bat` of a standard install, resolves to `Scripts\conda.exe`
of the same installation when that file exists. With no hit it runs the bare
name `conda`.

## Isolated reproduction environment

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
$runRoot = (Resolve-Path '..').Path # parent of this isolated checkout
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

Use `& $python` instead of `python`, and pass `--env-name ostram-reproduce`
to the `run` commands in the [portfolio sequence](pipeline.md#accepted-17-scenario-portfolio).
Record `conda list`, the Python version, `glpsol --version` and the selected
vendor solver version with each reproduction. Solver binaries and vendor
licences are external runtime dependencies; an old prepared workspace is not
an installation dependency.

Keep the complete environment/package export with run evidence; unbounded
dependency resolution on a later date is not an exact environment replay.

## Verify the installation

```powershell
python -m ostram --help
python -m ostram inspect-resources
python -B -m unittest discover -s tests -p "test_*.py"
python -m ostram run --skip-pull --compile-only --env-name ostram --scenarios A_Calibrated_BAU
```

The last command prepares the base model, builds `A_Calibrated_BAU`, compiles
its inputs and stops after the solver datafile. It runs no solver.

From outside the checkout, name the project root explicitly:

```powershell
python -m ostram --project-root <path to the clone> inspect-resources
```

Use `--workspace` or `OSTRAM_WORKSPACE` to place mutable state elsewhere.
Explicit command-line values override environment variables. Inspection is
read-only and does not create the workspace. The support `BAU` command is on
the [pipeline](pipeline.md#support-bau) page.

## Solvers

`environment.yaml` installs GLPK and CBC. CPLEX and Gurobi require their vendor
installations and licences. Select the solver in
`config/execution/Config_MOMF_T1_AB.yaml`. The tracked full profile selects
CPLEX, and the training profile selects CBC. For CPLEX, add the folder that
holds `cplex.exe` to `PATH` in the session.

Verify a solver before a production run:

```powershell
glpsol --version
cbc -stop
cplex -c quit
```

Use the compile-only route when validating without a solver:

```powershell
python -m ostram run --skip-pull --compile-only --env-name ostram --scenarios A_Calibrated_BAU
```

## DVC

DVC is a package dependency and `dvc.yaml` uses only canonical package commands
and workspace paths. Without `--skip-pull`, the OSTRAM runner initializes local
DVC metadata when needed and pulls only when a remote is configured. With
`--skip-pull`, it performs neither DVC initialization nor a pull. It does not
use caller CWD to locate project data.

```powershell
dvc remote list
python -m ostram run --skip-pull --compile-only --env-name ostram --scenarios A_Calibrated_BAU
```

Generated DVC cache/tmp state and the central workspace are ignored. Maintained
inputs, configuration, model files, and package resources remain tracked and
must not be replaced with generated copies.
