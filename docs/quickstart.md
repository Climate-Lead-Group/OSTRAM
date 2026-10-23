# Quickstart

## 1. Install the editable package

Follow the [installation](installation.md) page, then check from the project
root:

```powershell
python -m ostram inspect-resources
```

The last command reads representative input, configuration, model, and package
resources without creating a workspace.

## 2. Choose project and workspace paths

An editable checkout is discovered automatically. For an installed package or
an external working directory, provide the project bundle explicitly:

```powershell
python -m ostram `
  --project-root <path to the clone> `
  --workspace "<path to a workspace>" `
  inspect-resources
```

The equivalent environment variables are `OSTRAM_PROJECT_ROOT` and
`OSTRAM_WORKSPACE`. Command-line values take precedence. All resolved paths are
absolute, and the caller's current directory is never searched for resources.

### Supported Windows locations

Use a local, nonsynchronised location for both the project and workspace.
Unicode and spaces are supported. Before a run, the final-output
preflight must predict every final path at no more than 240 UTF-16 code units.

Deeply nested roots, network-mounted roots, and cloud-synchronised roots,
including OneDrive, are unsupported. If the preflight fails, choose a shorter
local root. The application must not silently truncate final output
names.

## 3. Run a solver-free preparation gate

```powershell
python -m ostram run --skip-pull --compile-only --scenarios A_Calibrated_BAU
```

`--compile-only` reaches the final preprocessed text inputs and stops before
matrix creation, every solver adapter, cleanup, and result post-processing.

The run takes minutes. In one run on one machine, compiling two derived
scenarios took about 20 minutes: A1 4 minutes, A2 45 seconds, A3 8 minutes, B1
5.5 minutes and B2 25 seconds. A3 and B1 are the longest stages. Other machines
differ. Progress goes to stderr. When output is redirected in Windows
PowerShell 5.1, the normal progress lines appear as `NativeCommandError`. The
exit code decides success: `$LASTEXITCODE` is 0 after a successful run, which
ends with `OSTRAM compile-only run completed successfully`.

Useful focused commands are:

```powershell
python -m ostram transform --scenario A_Calibrated_BAU
python -m ostram compile-inputs --scenarios "A_Calibrated_BAU,B_Optimised_VRE"
```

## 4. Run the configured workflow

After selecting and configuring a solver in
`config/execution/Config_MOMF_T1_AB.yaml`:

```powershell
python -m ostram run --skip-pull --scenarios A_Calibrated_BAU
```

The runner conditionally prepares A1/A2 snapshots, materializes the exact
registry selection, compiles it, and executes B2. Use `--scenarios` for an
explicit comma-separated selection and the documented `--skip-*` flags when a
stage is already satisfied.

For all 23 decision cases, use the
[portfolio sequence](pipeline.md#23-scenario-portfolio): solve A, then
materialize and solve the remaining decision cases using that new A result.
Support `BAU` has a separate [standalone execution command](pipeline.md#support-bau)
and permits existing transmission plus committed additions. It is not the
A scenario. The unfiltered default also
includes C, which cannot be materialized before its A result exists.

The interactive display is compact and updates in place on stderr. Redirected
output uses append-only status lines. Add `--verbose` to disable the moving
display and stream complete child output. The detailed UTF-8 log is always at
`<workspace>/logs/<run-id>/run.log` for an actual run.

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

The solver-free checks are listed under [Tests](https://github.com/Climate-Lead-Group/OSTRAM/blob/main/CONTRIBUTING.md#tests) in `CONTRIBUTING.md`.
They call no matrix builder or solver.
