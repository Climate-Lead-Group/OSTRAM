# UNESCAP South Asia training setup

A complete, self-contained OSTRAM example: a reduced two-region power-system model
built for the UNESCAP training sessions. Bangladesh (`BGDXX`) and India East (`INDEA`),
joined by the real `TRNBGDXXINDEA` cross-border corridor, over 2023–2050 in 20 timeslices.

Everything the setup needs is in this directory — input CSVs, the scenario workbook,
preparation/compilation/execution configuration and the scenario registry. Nothing here falls back to the project-root configuration of
the same name.

The setup engine and these assets are integrated. All commands below resolve the
setup's explicit authorities; none implicitly borrow a same-named South Asia input.

## The model

| | |
|---|---|
| Regions | `BGDXX`, `INDEA` |
| OSeMOSYS `REGION` | `GLOBAL` (single region) |
| Interconnector | `TRNBGDXXINDEA` |
| Years | 2023–2050 (28) |
| Timeslices | 20 (4 seasons × 1 day type × 5 daily brackets) |
| Seed technologies / fuels | 89 / 43 (pre-preparation authorities) |
| Prepared/compiled technologies / fuels | 90 / 49 |
| Storage | 4 — `LDSBGDXX01`, `LDSINDEA01`, `SDSBGDXX01`, `SDSINDEA01` |
| Solver | CBC |
| Scenario roots | `BAU`, `A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE` |

The two rows above are stage-specific contracts, not competing counts. Preparation
performs a count-preserving reconciliation of legacy technology identities. Compilation
then adds exactly one technology, `PWRSHPINDEA`, and six ELC dispatch fuels:
`ELCBGDXX00`, `ELCBGDXX03`, `ELCBGDXX04`, `ELCINDEA00`, `ELCINDEA03`, and
`ELCINDEA04`. `setup.yaml` records the reconciliation and the generated delta, and the
runtime rejects any other compiled addition or removal before crossing the matrix/solver
boundary.

`C_Target_VRE` declares a completed-result dependency on `A_Calibrated_BAU`: it scales
its NDC-derived floors against solved CalBAU generation, so A must be solved first. The
dependency is declared in [`config/scenarios/registry.json`](config/scenarios/registry.json),
not left to the operator to remember.

`A_Calibrated_BAU` declares a restriction-materialization dependency on `BAU`
(`{"type": "restrictions", "scenario": "BAU"}` in the same registry): this workbook
ships a header-only Restrictions sheet, and A inherits the rows that BAU's rules
generate. Selecting A therefore materializes BAU first in the same A3 run and hands
BAU's generated Restrictions to A through an exported disposable run state; the
downstream compile/solve selection stays exactly what was requested, and requesting
only `B_Optimised_VRE` never adds BAU.

## Layout

```
examples/unescap_training/
  setup.yaml                     the setup's authorities and resolution policy
  README.md                        this file
  inputs/
    osemosys_global/*.csv          64 OSeMOSYS Global parameter and set files
    scenarios/
      OSTRAM_Scenario_Inputs.xlsx  the scenario workbook (20 sheets, incl. Restrictions)
  config/
    preparation/
      Config_country_codes.yaml    countries, technology mappings, template generation
      Config_region_consolidation.yaml   disabled for this setup
    compilation/
      Config_MOMF_T1_A.yaml        timeslice fabric, storage set, A-stage authorities
    execution/
      Config_MOMF_T1_AB.yaml       solver selection and execution-stage switches
    scenarios/
      registry.json                ostram-scenario-registry-v1 — four roots, no derived
      technology_types.csv         technology → category map, trimmed to this setup
      ao_extension_decisions.csv   AO extension decision record
      A_Calibrated_BAU/*.yaml
      B_Optimised_VRE/*.yaml
      C_Target_VRE/*.yaml
```

Nothing generated lives here. Preparation output, compiled parameters, solver inputs,
`.sol` files, `Executables/`, `Outputs/`, dashboards, and plots are all written under the
run workspace (`--workspace` / `OSTRAM_WORKSPACE`), never inside `examples/unescap_training/`.

## Commands

Install a clean checkout with the [installation](../../docs/installation.md) page.
The commands below use its `ostram` environment and a new workspace.
The setup is named, not located, so no command depends on your working directory.

```
python -m ostram --workspace workspace/unescap_training example prepare unescap_training
python -m ostram --workspace workspace/unescap_training --setup unescap_training run --env-name ostram --scenarios "A_Calibrated_BAU,B_Optimised_VRE" --compile-only --skip-pull
python -m ostram --workspace workspace/unescap_training --setup unescap_training run --env-name ostram --scenarios "A_Calibrated_BAU,B_Optimised_VRE" --skip-pull
python -m ostram --workspace workspace/unescap_training --setup unescap_training run --env-name ostram --scenarios "C_Target_VRE" --skip-pull
python -m ostram --workspace workspace/unescap_training example report unescap_training --capture baseline
python -m ostram --workspace workspace/unescap_training example report unescap_training
```

The first run stops before matrix generation and solving. The next two runs compile,
solve with CBC, export results and reconcile costs. C consumes the newly completed A
result in the same workspace. Select A/B/C explicitly: support `BAU` supplies A's
restriction materialization and is not a requested decision solve. The backstop is
enabled by this setup; include its exported use when interpreting a feasible result.

Setup-aware country commands:

```
python -m ostram --setup unescap_training country template
python -m ostram --setup unescap_training country merge MMR
python -m ostram --setup unescap_training country validate MMR
python -m ostram --setup unescap_training scenario sync-country --country MMR
```

Setup-aware scenario commands:

```
python -m ostram --setup unescap_training scenario sync-country --help
python -m ostram --setup unescap_training transform --scenario B_Optimised_VRE
python -m ostram --setup unescap_training compile-inputs
python -m ostram --setup unescap_training inspect-resources
```

`setup.yaml` sets `metadata.requires_prepare: true`. The setup ships input CSVs and a
scenario workbook, not a built model, so `example prepare unescap_training` must succeed before any
scenario can be compiled or solved.

## Path resolution

`setup.yaml` names every resource under exactly one authority:

- **`setup:`** — everything in this directory: inputs, configs and registry.
- **`project:`** — repository-level resources borrowed as-is: the maintained OSeMOSYS
  model at `model/osemosys_fast_preprocessed.txt`, and `inputs/execution/`.
- **`package:`** — resources shipped inside the installed `ostram` package:
  `resources/compilation/` (otoole conversion format, parameter templates) and
  `resources/preparation/`.

`resolution.implicit_file_fallback: false`. If a setup-owned file is missing, resolution
fails and names it; it does not quietly reach for the project-root file of the same name.
The config files use `${...}` tokens for every path — none of them contain a physical
legacy path such as `./A1_Outputs` or `../Executables/A_Calibrated_BAU_0`.

## Known discrepancies

The workbook's exact **2.496 GW** `TRNBGDXXINDEA` residual
capacity is also the 2023 anchor in every executable scenario schedule; no executable
authority rounds it. In the active relaxed scenario the resulting
`TotalAnnualMaxCapacity` is exactly `9999.0`, a finite cap of up to 9999 GW in total. It
does not alter the `ResidualCapacity` series.
