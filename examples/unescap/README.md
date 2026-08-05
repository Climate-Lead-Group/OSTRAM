# UNESCAP South Asia training profile

A complete, self-contained OSTRAM example: a reduced two-region power-system model
built for the UNESCAP training sessions. Bangladesh (`BGDXX`) and India East (`INDEA`),
joined by the real `TRNBGDXXINDEA` cross-border corridor, over 2023–2050 in 20 timeslices.

Everything the profile needs is in this directory — input CSVs, the scenario workbook,
preparation/compilation/execution configuration, the scenario registry, the exercises,
and the provenance record. Nothing here falls back to the project-root configuration of
the same name.

The profile engine and these assets are integrated. All commands below resolve the
profile's explicit authorities; none implicitly borrow a same-named full-model input.

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
`ELCINDEA04`. `profile.yaml` records membership hashes for both stages, and the runtime
rejects any other addition or removal before crossing the matrix/solver boundary.

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
examples/unescap/
  profile.yaml                     the profile's authorities and resolution policy
  README.md                        this file
  AUTHORING_AND_ACCEPTANCE.md      Natalia's Windows authoring + Actions acceptance start
  TRAINING_OPERATION_COVERAGE.md   legacy-to-canonical operation ledger
  exercises/
    training.html                  main guide — 4 scenario exercises + 5 interconnector sub-exercises
    add-country.html               Exercise A — add Myanmar to the model
    add-interconnector.html        Exercise B — add a BGD↔MMR corridor
  docs/
    git-setup.html                 version control and environment setup
  inputs/
    osemosys_global/*.csv          64 OSeMOSYS Global parameter and set files
    scenarios/
      OSTRAM_Scenario_Inputs.xlsx  the scenario workbook (20 sheets, incl. Restrictions)
  config/
    preparation/
      Config_country_codes.yaml    countries, technology mappings, template generation
      Config_region_consolidation.yaml   disabled for this profile
    compilation/
      Config_MOMF_T1_A.yaml        timeslice fabric, storage set, A-stage authorities
    execution/
      Config_MOMF_T1_AB.yaml       solver selection and execution-stage switches
    scenarios/
      registry.json                ostram-scenario-registry-v1 — four roots, no derived
      technology_types.csv         technology → category map, trimmed to this profile
      ao_extension_decisions.csv   AO extension decision record (see provenance)
      A_Calibrated_BAU/*.yaml
      B_Optimised_VRE/*.yaml
      C_Target_VRE/*.yaml
  references/
    interconnector-direction-results.html   measured forward/reverse/bidirectional comparison
    provenance.md                  where every file came from, and what changed
    release-readiness.md           Windows CBC and cross-platform acceptance evidence
```

Nothing generated lives here. Preparation output, compiled parameters, solver inputs,
`.sol` files, `Executables/`, `Outputs/`, dashboards, and plots are all written under the
run workspace (`--workspace` / `OSTRAM_WORKSPACE`), never inside `examples/unescap/`.

Training authors should start with
[`AUTHORING_AND_ACCEPTANCE.md`](AUTHORING_AND_ACCEPTANCE.md). The audited migration map
is [`TRAINING_OPERATION_COVERAGE.md`](TRAINING_OPERATION_COVERAGE.md).

## Commands

The profile is named, not located, so no command depends on your working directory.

```
python -m ostram example prepare unescap                       # required first — see below
python -m ostram --profile unescap run
python -m ostram --profile unescap run --scenarios "A_Calibrated_BAU,B_Optimised_VRE"
python -m ostram --profile unescap run --scenarios "C_Target_VRE"
python -m ostram example report unescap
python -m ostram example report unescap --capture baseline
```

Profile-aware country commands, used by [Exercise A](exercises/add-country.html):

```
python -m ostram --profile unescap country template
python -m ostram --profile unescap country merge MMR
python -m ostram --profile unescap country validate MMR
python -m ostram --profile unescap scenario sync-country --country MMR
```

Profile-aware scenario commands:

```
python -m ostram --profile unescap scenario sync-country --help
python -m ostram --profile unescap transform --scenario B_Optimised_VRE
python -m ostram --profile unescap compile-inputs
python -m ostram --profile unescap inspect-resources
```

`profile.yaml` sets `metadata.requires_prepare: true`. The profile ships input CSVs and a
scenario workbook, not a built model, so `example prepare unescap` must succeed before any
scenario can be compiled or solved.

## Path resolution

`profile.yaml` names every resource under exactly one authority:

- **`profile:`** — everything in this directory: inputs, configs, registry, docs.
- **`project:`** — repository-level resources borrowed as-is: the maintained OSeMOSYS
  model at `model/osemosys_fast_preprocessed.txt`, and `inputs/execution/`.
- **`package:`** — resources shipped inside the installed `ostram` package:
  `resources/compilation/` (otoole conversion format, parameter templates) and
  `resources/preparation/`.

`resolution.implicit_file_fallback: false`. If a profile-owned file is missing, resolution
fails and names it; it does not quietly reach for the project-root file of the same name.
The config files use `${...}` tokens for every path — none of them contain a physical
legacy path such as `./A1_Outputs` or `../Executables/A_Calibrated_BAU_0`.

## Provenance and known discrepancies

[`references/provenance.md`](references/provenance.md) records the source commit, the
source path of every migrated file, the SHA-256 of the scenario workbook, and every
transformation applied. The workbook's exact **2.496 GW** `TRNBGDXXINDEA` residual
capacity is also the 2023 anchor in every executable scenario schedule; prose may describe
it as approximately 2.5 GW, but no executable authority rounds it. Choosing 2.500 GW is
an exercise edit, not the shipped seed. In the active relaxed scenario the resulting
`TotalAnnualMaxCapacity` is intentionally unbound at exactly `9999.0`; that does not alter
the `ResidualCapacity` series.
