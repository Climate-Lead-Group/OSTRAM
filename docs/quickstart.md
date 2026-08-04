# Quickstart

## 1. Install the editable package

From the project root:

```powershell
conda env create -f environment.yaml
conda activate OSTRAM-env
python -m pip install -e .
python -m ostram inspect-resources
```

The last command reads representative input, configuration, model, and package
resources without creating a workspace.

## 2. Choose project and workspace paths

An editable checkout is discovered automatically. For an installed package or
an external working directory, provide the project bundle explicitly:

```powershell
python -m ostram `
  --project-root C:\path\to\OSTRAM `
  --workspace "D:\OSTRAM work\run α" `
  inspect-resources
```

The equivalent environment variables are `OSTRAM_PROJECT_ROOT` and
`OSTRAM_WORKSPACE`. Command-line values take precedence. All resolved paths are
absolute, and the caller's current directory is never searched for resources.

### Supported Windows locations

Use a local, nonsynchronised location for both the project and workspace.
Unicode and spaces are supported. Before a governed run, the final-output
preflight must predict every final path at no more than 240 UTF-16 code units.

Deeply nested roots, network-mounted roots, and cloud-synchronised roots,
including OneDrive, are unsupported. If the preflight fails, choose a shorter
local root. The application must not silently truncate final governed output
names.

## 3. Run a solver-free preparation gate

```powershell
python -m ostram run --skip-pull --compile-only
```

`--compile-only` reaches the final preprocessed text inputs and stops before
matrix creation, every solver adapter, cleanup, and result post-processing.

Useful focused commands are:

```powershell
python -m ostram transform --scenario BAU
python -m ostram compile-inputs --scenarios "BAU,B_Optimised_VRE"
```

## 4. Run the configured workflow

After selecting and configuring a solver in
`config/execution/Config_MOMF_T1_AB.yaml`:

```powershell
python -m ostram run
```

The runner conditionally prepares A1/A2 snapshots, materializes the exact
registry selection, compiles it, and executes B2. Use `--scenarios` for an
explicit comma-separated selection and the documented `--skip-*` flags when a
stage is already satisfied.

## 5. Know where files belong

- Maintained scenario workbooks: `inputs/scenarios/`
- OSeMOSYS CSV authorities: `inputs/osemosys_global/`
- Scenario policy: `config/scenarios/`
- Compilation/execution configuration: `config/compilation/` and
  `config/execution/`
- Read-only package resources: `ostram/resources/`
- Generated workbooks, compiled parameters, executables, and results:
  `<workspace>/preparation`, `<workspace>/scenarios`,
  `<workspace>/compilation`, and `<workspace>/execution`

`workspace/` is ignored and created lazily. Do not copy maintained sources into
it or treat generated state as source authority.

## 6. Run solver-free validation

```powershell
python -B -m compileall -q ostram tests
python -B -m unittest discover -s tests -p "test_*.py"
python -B -m tests.validation.test_scenarios_lite
```

The scenario-lite command validates all seven compact scenarios without
calling a matrix builder or solver.
