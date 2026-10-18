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
   conda env create -f environment.yaml
   ```

   The environment is named `OSTRAM-env`, the `name` in `environment.yaml`.
   The channel list comes from `environment.yaml`, which names conda-forge, and
   from your own conda settings. Conda prints `Channels: conda-forge, defaults`
   when your settings add `defaults`. Conda 24.11 does not accept
   `--override-channels` with `conda env create -f`. To see the list, run
   `conda config --show channels`.

   If this command fails, stop. Every later command needs the environment.

   ```powershell
   conda activate OSTRAM-env
   python -c "import sys; print(sys.prefix)"
   python -m pip install -e .
   python -m pip check
   ```

   The printed path ends with `envs\OSTRAM-env`.

   An environment is broken when its folder has no `python.exe`. `conda env
   list` shows the folder of each environment. Remove a broken `OSTRAM-env` and
   create it again:

   ```powershell
   conda env remove -n OSTRAM-env
   conda env create -f environment.yaml
   ```

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

`python -m ostram run` checks the environment named in `environment.yaml`,
`OSTRAM-env`, and creates it when it is missing. Pass `--env-name <name>` to
every `run` only when you created the environment under another name. Every
stage runs under the Python that called the command, so run the commands from
the activated environment.

For an existing environment, update it with:

```powershell
conda env update -n OSTRAM-env -f environment.yaml --prune
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

When `run` finds no conda, it stops at its first step with `Requirement
'conda' not found in PATH`. There are two ways to fix this:

- Open the Miniforge Prompt or the Anaconda Prompt from the Start menu and run
  the commands there. Or run `conda init powershell` once in that prompt, and
  open a new PowerShell window.
- As a fallback, set `CONDA_EXE` to the full path of `conda.exe` in the
  window, before `run`:

  ```powershell
  $env:CONDA_EXE = '<conda folder>\Scripts\conda.exe'
  ```

`CONDA_EXE` is all that `run` needs from conda, because every conda call of
`run` uses that file. Without `conda activate`, call the environment's Python
by its full path, `<environment folder>\python.exe -m ostram run ...`, as every
stage runs under the Python that called the command. `conda env list` shows
the environment folder.

## Verify the installation

```powershell
python -m ostram --help
python -m ostram inspect-resources
python -B -m unittest discover -s tests -p "test_*.py"
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU
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
`config/execution/Config_MOMF_T1_AB.yaml`. The tracked `south_asia` setup selects
CPLEX, and the training setup selects CBC. For CPLEX, add the folder that
holds `cplex.exe` to `PATH` in the session.

Verify a solver before a production run:

```powershell
glpsol --version
cbc -stop
cplex -c quit
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
