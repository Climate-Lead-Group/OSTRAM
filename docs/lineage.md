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

`config/profiles/full.yaml::scenario_registry` defines the four roots, the accepted derived
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

## Cost conventions and provenance

Prices are constant 2023 USD/GJ (numerically MUSD/PJ); capital and fixed O&M use
MUSD/GW and MUSD/GW-year. `VariableCost` country MIN rows provide fuel prices,
including the `B_Opt_IndiaCostsFuel` substitution. The compiler reads
`Config_MOMF_T1_A.yaml::fuel_cost_allocation` (implemented in
`ostram/pipeline/compilation/transforms/effects.py`, checked by
`tests/regression/test_a1_b1_transforms.py`): for each fuel and year the positive
minimum mapped national price is charged upstream, and each generator retains
its non-fuel VOM plus the national-price difference multiplied by its unchanged
fuel `InputActivityRatio`. Country supply permissions, efficiencies, physical
routes and constraints are unchanged by this allocation; `COAINT` names an
accounting route, not procurement origin. Monetary values are rounded to four
decimals.

The sources, conversions, proxies and caveats behind every monetary and physical
input are documented inside `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx`
(README sheet, Source/Notes/References cells and cell notes). The annual values
are the operational inputs; those notes are not a separate input authority.
Release evidence for the September 2026 cost update (accepted manifest, selected
17-case index, evidence workbook and fuel-price register) is retained outside
Git beside the checkout, under `D4_SCAN_STATE/`; it is not needed at runtime and
is not shipped with a clone. In that release `B_Opt_TradeCap15` was accepted from
a numerical-emphasis recovery of the same LP (feasibility/optimality tolerances
`1e-6`, no coefficient change), now reproduced by
`cplex_numerical_emphasis_scenarios` in `config/execution/Config_MOMF_T1_AB.yaml`.

Full-profile result concatenation reconciles omitted default VOM and duplicate
identical exporter rows against the exact compiled coefficients and exporter CSV
(`reconcile_variable_cost_exports: true`). Original `Outputs/*.csv` remain
unchanged; concatenated cost series and summary workbooks use the reconciled
values, and an adjacent `*_cost_reconciliation.json` records raw total,
adjustment, source hashes and the reconciled total. Raw fixed O&M already
includes residual stock and is retained.

## Preservation contract

Project-layout changes do not alter workbook bytes, numerical assumptions,
scenario policy, source authority, or solver policy. Mutable workspace files
are disposable products; the tracked `inputs/`, `config/`, `model/`, and
`ostram/resources/` trees remain the sources of truth.
