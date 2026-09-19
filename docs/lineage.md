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

## Preservation contract

### September 2026 cost release

The accepted cost candidate is identified by manifest SHA256
`882c38ae33178f931a25dab78567b94f8becaed24a31c5d8bd81d6bb265badf4`.
This is a manifest identity, not the scenario workbook hash. Release verification
and the selected 17-case index live outside Git in
`D4_SCAN_STATE/RELEASE_20260919/RELEASE_REVIEW.md` alongside this checkout.
That review determines whether production transfer and publication passed.

The scenario workbook and existing monetary CSVs retain the complete revised
monetary set and the eight accepted correction families. Prices are constant
2023 USD/GJ (numerically MUSD/PJ); capital and fixed O&M retain their documented
MUSD/GW and MUSD/GW-year conventions. `VariableCost` country MIN rows provide
fuel prices, including the existing IndiaCostsFuel substitution. The compiler
reads `Config_MOMF_T1_A.yaml::fuel_cost_allocation`; its implementation is in
the existing compilation `transforms/effects.py` and its focused checks in
`test_a1_b1_transforms.py`. No additional input authority or scenario patch is
required. CSV and workbook redundancy remains part of the existing workflow.

For each fuel/year the positive minimum mapped national price is charged
upstream. Each generator retains its non-fuel VOM plus the national-price
difference multiplied by its unchanged fuel InputActivityRatio. Effective
fuel expenditure is the upstream charge plus that consumer adjustment.
Country supply permissions, efficiencies, physical routes and constraints
remain unchanged. COAINT names an accounting route, not procurement origin.
The compiler retains four-decimal monetary rounding.

India coal uses the NTPC utility-cost proxy and 1% annual real escalation;
extension through 2050 is a study assumption. Bangladesh gas uses the same
FY2023/24 procurement proxy (5.422227767430297 USD2023/GJ) for 2023 and 2024,
with IEPMP ratios thereafter; its boundary excludes midstream. Sri Lanka
coal's 2023 value is 7.96 USD/GJ with the invoice-finalisation caveat. Maldives
diesel proxies Sri Lanka. Approved international/CIF multipliers apply once
(1.3, fuel oil 1.0); national observations receive no invented delivery charge.
OTH/COG retain oil-/gas-equivalent proxies. Unity mapping to the inherited
fleet's incompletely documented heat basis is a disclosed compatibility
assumption, not a verified HHV/LHV conversion.

The final evidence workbook is
`D4_SCAN_STATE/COST_UPDATE_20260919/OSTRAM_Fuel_Prices_Cost_Update.xlsx`.
Its retained non-fuel/source sheets, `fuel_price_register.csv`, and the release
archive index identify source documents, units, dates, boundaries and proxies.
Historical comparisons use H2_r4 with selected execution under H4; the newer
evidence-review candidate must not replace this full monetary set.

Full-profile result concatenation reconciles omitted default VOM and duplicate
identical exporter rows against the exact compiled coefficients and exporter
CSV. Original `Outputs/*.csv` remain unchanged. Concatenated cost series and
normal summary workbooks use the reconciled values; an adjacent
`*_cost_reconciliation.json` records raw total, adjustment, source hashes and
the reconciled total. Raw fixed O&M already includes residual stock and is
retained. The full-profile policy is `reconcile_variable_cost_exports: true`.
The selected analysis summaries, rather than intermediate `result_*.json`
files, are the historical result authorities.

Project-layout changes do not alter workbook bytes, numerical assumptions,
scenario policy, source authority, or solver policy. Mutable workspace files
are disposable products; the tracked `inputs/`, `config/`, `model/`, and
`ostram/resources/` trees remain the sources of truth.
