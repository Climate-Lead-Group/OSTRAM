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
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU
python -m ostram run --skip-pull --scenarios A_Calibrated_BAU --verbose
python -m ostram transform --scenario A_Calibrated_BAU
python -m ostram compile-inputs --scenarios "A_Calibrated_BAU,B_Optimised_VRE"
```

Global path options precede the command:

```powershell
python -m ostram `
  --project-root C:\path\to\OSTRAM `
  --workspace "D:\OSTRAM work\run α" `
  inspect-resources
```

For all 17 accepted decision scenarios, follow the
[dependency-ordered portfolio commands](docs/pipeline.md#accepted-17-scenario-portfolio).
`BAU` is a support scenario, distinct from `A_Calibrated_BAU`; it is not one of
the accepted 17 outputs. Omitting `--scenarios` includes support BAU and C,
whose materialization requires a completed A result. It is not a fresh-portfolio
shortcut. A1/A2/A3/B1/B2 are pipeline stages, not scenario names.

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
- `config/profiles/full.yaml::scenario_registry`
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

The adopted sources, conversions, assumptions, historical calibration and
remaining evidence limitations are documented inside the existing
[scenario workbook](inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx), in its README,
Source/Notes/References cells and Excel cell notes. The annual values remain
the operational inputs; the added explanations are not another input authority.
The September 2026 cost update and reconciled reporting convention are also
described in [lineage](docs/lineage.md#september-2026-cost-release). The external release
review at `D4_SCAN_STATE/RELEASE_20260919/RELEASE_REVIEW.md` records the selected
17 cases, verification gates, evidence archive, and publication state.
For a compile-only check using the selected A result for C, use the normal
`run --skip-pull --compile-only --a-result-seed <selected-A-run-directory>`
command with an isolated `--workspace`. A fresh full solve must solve A before
materialising C through its declared dependency. The transmission-freeze
case retains its fixed historical reference; legacy and strict import caps
remain separate definitions.

A fresh clone includes maintained inputs, but not the local `D4_SCAN_STATE`
release archive, selected solver outputs or original country source packages.
Those locally retained records have not been published as downloadable release
assets. Exact historical replay requires those artifacts; a new full run obtains
C's dependency by solving A first. The transmission-freeze coefficients are
already embedded in the registered YAML and require no external M0 run at
runtime. Source URLs identify publications but do not guarantee continuing
online availability.

The calibration CSV is byte-hash guarded. Its scoped `.gitattributes` rule
preserves the released CRLF bytes on fresh checkouts; do not normalize that
file independently or change its accepted hash to bypass the guard.
