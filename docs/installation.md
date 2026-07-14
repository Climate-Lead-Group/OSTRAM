# Installation

This guide covers the full setup process for running OSTRAM on a Windows machine.

## System Requirements

- **Operating System:** Windows 10 or higher
- **Git:** [Git for Windows](https://gitforwindows.org/) (version 2.40+)
- **Python distribution:** [Miniconda](https://docs.conda.io/en/latest/miniconda.html) or [Anaconda](https://www.anaconda.com/download)
- **Solver:** At least one of the following LP/MIP solvers:
  - [GLPK](https://www.gnu.org/software/glpk/) (open source, included in conda-forge)
  - [CBC](https://github.com/coin-or/Cbc) (open source)
  - [CPLEX](https://www.ibm.com/products/ilog-cplex-optimization-studio) (commercial, IBM)
  - [Gurobi](https://www.gurobi.com/) (commercial, free academic license)

## Clone the Repository

```bash
git clone https://github.com/Climate-Lead-Group/OSTRAM.git
cd OSTRAM
```

## Conda Environment

OSTRAM uses a Conda environment defined in `environment.yaml`. The `run.py` launcher automatically creates it if it does not exist, but you can also create it manually:

```bash
conda env create -f environment.yaml
conda activate OSTRAM-env
```

### Environment Dependencies

The environment installs the following packages:

| Package | Version | Purpose |
|---------|---------|---------|
| Python | 3.10 | Runtime |
| pandas | >= 2.1 | Data manipulation |
| numpy | >= 1.26 | Numerical computation |
| openpyxl | >= 3.1 | Excel file reading |
| xlsxwriter | >= 3.2.4 | Excel file writing |
| pyyaml | >= 6.0 | YAML configuration parsing |
| git | >= 2.40 | Version control |
| dvc | latest | Data Version Control pipeline |
| otoole | >= 1.1.1 | OSeMOSYS data format conversion |

## Solver Setup

### GLPK (simplest option)

GLPK can be installed directly via conda:

```bash
conda activate OSTRAM-env
conda install -c conda-forge glpk
```

### CBC

CBC can also be installed via conda:

```bash
conda activate OSTRAM-env
conda install -c conda-forge coin-or-cbc
```

### CPLEX

1. Install IBM ILOG CPLEX Optimization Studio from the [IBM website](https://www.ibm.com/products/ilog-cplex-optimization-studio).
2. Ensure the `cplex` binary is available on your system `PATH`.
3. Set `solver: 'cplex'` in `t1_confection/Config_MOMF_T1_AB.yaml`.

### Gurobi

1. Install Gurobi from the [Gurobi website](https://www.gurobi.com/downloads/).
2. Activate your license (`grbgetkey <license-key>`).
3. Ensure the `gurobi` binary is available on your system `PATH`.
4. Set `solver: 'gurobi'` in `t1_confection/Config_MOMF_T1_AB.yaml`.

## Verify Installation

After setup, verify that the environment works:

```bash
conda activate OSTRAM-env
python -c "import pandas; import numpy; import openpyxl; import yaml; print('All dependencies OK')"
```

To verify your solver:

```bash
# For GLPK
glpsol --version

# For CBC
cbc --version

# For CPLEX
cplex -c "quit"

# For Gurobi
gurobi_cl --version
```

## DVC Setup (Optional)

OSTRAM uses [DVC](https://dvc.org/) (Data Version Control) to manage the pipeline. The `run.py` script handles DVC commands automatically, but if you want to run stages manually:

```bash
conda activate OSTRAM-env
dvc repro
```

This will execute all pipeline stages that have outdated outputs based on dependency tracking.
