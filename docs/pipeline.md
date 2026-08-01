# Pipeline

The canonical workflow is:

```text
A1 base-input preparation
  -> A2 transmission enrichment and root snapshots
  -> A3 root transformation and derived-scenario materialization
  -> B1 OSeMOSYS input compilation
  -> B2 preprocessing, optional matrix/solve, and results
```

Run it with:

```powershell
python -m ostram run [options]
```

## A1 and A2: preparation

A1 reads the maintained CSV authorities in `inputs/osemosys_global/`, the
preparation configuration in `config/preparation/`, and workbook assets in
`inputs/preparation/`. A2 adds the governed transmission content. Their mutable
products and post-A2 snapshots live under `<workspace>/preparation/`.

The full runner treats A1+A2 as a pair. If every selected root already has its
post-A2 snapshot, the pair is skipped; otherwise the root preparation is
rebuilt before A3.

## A3: scenario materialization

The maintained authorities are:

- `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx`
- `inputs/scenarios/OSTRAM_Timeslice_Inputs.xlsx`
- `config/scenarios/registry.json`
- scenario-specific rule YAML/JSON files in `config/scenarios/<scenario>/`

The registry declares four roots and the accepted derived scenarios. A root is
transformed through the ordered workbook stages. A derived scenario begins
from its declared root output and then receives only its registered patch and
optional direction overlay.

Focused root transformation uses:

```powershell
python -m ostram transform --scenario BAU
```

Temporary stage directories are created under `<workspace>/scenarios/` and are
removed unless `--keep-workdir` is set. Python implementation files are never
copied into those directories; subprocess stages execute installed package
modules with explicit workspace and project paths.

## B1: compilation

B1 reads materialized scenario workbooks, copies the maintained compiler YAML
into `<workspace>/compilation/`, updates the selected scenario in that runtime
copy, and writes compiled parameter CSVs under
`<workspace>/compilation/A2_Output_Params/`.

```powershell
python -m ostram compile-inputs
python -m ostram compile-inputs --scenarios "BAU,B_Optimised_VRE"
```

The maintained `config/compilation/Config_MOMF_T1_A.yaml` is not mutated.

## B2: execution boundary

B2 reads:

- `config/execution/Config_MOMF_T1_AB.yaml`
- `model/osemosys_fast_preprocessed.txt`
- `ostram/resources/compilation/conversion_format.yaml`
- compiled scenario CSVs from the selected workspace

It writes preprocessed text, optional solver artifacts, and result tables only
under `<workspace>/execution/` (apart from the documented root datafile export
when enabled).

The solver-free boundary is:

```powershell
python -m ostram run --skip-pull --compile-only
```

This route stops before matrix creation, solver adapters, cleanup, and result
post-processing. A normal `python -m ostram run` follows the solver and policy
settings in the execution YAML.

## Scenario selection

`--scenarios` accepts a comma-separated exact selection. The registry validates
names, preserves canonical order, calculates required roots, and resolves
declared result dependencies. `C_Target_VRE` requires the accepted A-result
seed boundary; no result is discovered through caller CWD.

## Path and process guarantees

- Project and workspace selection follow CLI > environment > editable checkout.
- Every resource and runtime path is absolute after resolution.
- Resource inspection never creates the workspace.
- Production code has no `sys.path` manipulation or global `chdir`.
- Python subprocess stages use `python -B -m <package.module>`.
- Commands use token lists, never `shell=True`.
- Package resources are opened read-only from the installed package.

See [configuration](configuration.md) for parameter details and
[lineage](lineage.md) for source ownership.
