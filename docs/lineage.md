# Runtime and source lineage

OSTRAM has one maintained runtime layout and one public command family. Python
code lives under `ostram/`; all supported invocations begin with
`python -m ostram`.

## Path ownership

| Content | Maintained location | Runtime treatment |
|---|---|---|
| Global OSeMOSYS sets and parameters | `inputs/osemosys_global/` | Read-only authority |
| Preparation sources and workbook templates | `inputs/preparation/` | Read-only authority |
| Scenario and timeslice workbooks | `inputs/scenarios/` | Read-only authority; materialized copies enter the workspace |
| Scenario registry and rule configuration | `config/scenarios/` | Read-only policy authority |
| Compiler configuration | `config/compilation/` | Copied to the compilation workspace before controlled mutation |
| Execution configuration | `config/execution/` | Read-only configuration |
| Maintained model | `model/osemosys_fast_preprocessed.txt` | Read-only model authority |
| Conversion schema and CSV templates | `ostram/resources/compilation/` | Import-addressable package data, opened read-only |
| Generated state | `<workspace>/` | Mutable, ignored, and lazily created |

The path resolver applies this precedence:

1. `--project-root` / `--workspace`
2. `OSTRAM_PROJECT_ROOT` / `OSTRAM_WORKSPACE`
3. validated editable-checkout project root / `<project-root>/workspace`

Caller CWD is never a resource-selection input.

## Pipeline lineage

```text
inputs + config + package resources
        |
        v
preparation workspace (A1/A2 snapshots)
        |
        v
scenario workspace (root transformations and derived overlays)
        |
        v
compilation workspace (one CSV family per selected scenario)
        |
        v
execution workspace (preprocessed text, optional matrix/solve, results)
```

`config/scenarios/registry.json` defines the four roots, the accepted derived
scenarios, selection order, patch layers, direction overlays, and explicit
result dependencies. Derived scenarios reuse their declared root source and do
not become competing workbook authorities.

The A-result seed used by `C_Target_VRE` remains an explicit dependency. Supply
it through the documented option/environment boundary; it is never inferred
from caller CWD.

## Interfaces and safety boundaries

- `python -m ostram run` owns full orchestration.
- `python -m ostram transform` owns focused root transformation.
- `python -m ostram compile-inputs` owns focused B1 compilation.
- `python -m ostram inspect-resources` proves read-only resource access.

Subprocesses use argument lists and installed module names. Production code
does not manipulate `sys.path`, change the global process directory, use
`shell=True`, or launch Python files by pathname.

The solver-free boundary is `python -m ostram run --compile-only`: it may
prepare final text inputs but must stop before matrix creation, every solver
adapter, cleanup, and result conversion. Full execution proceeds only when the
user deliberately runs the configured solver route.

## Preservation contract

Project-layout changes do not alter workbook bytes, numerical assumptions,
scenario policy, source authority, or solver policy. Mutable workspace files
are disposable products; the tracked `inputs/`, `config/`, `model/`, and
`ostram/resources/` trees remain the sources of truth.
