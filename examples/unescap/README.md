# UNESCAP South Asia training profile

A complete, self-contained OSTRAM example: a reduced two-region power-system model
built for the UNESCAP training sessions. Bangladesh (`BGDXX`) and India East (`INDEA`),
joined by the real `TRNBGDXXINDEA` cross-border corridor, over 2023–2050 in 20 timeslices.

Everything the profile needs is in this directory — input CSVs, the scenario workbook,
preparation/compilation/execution configuration, the scenario registry, the exercises,
and the provenance record. Nothing here falls back to the project-root configuration of
the same name.

> **This branch carries the assets only.**
> The profile engine that reads `profile.yaml` and resolves `${profile.…}` /
> `${project.…}` / `${package.…}` / `${workspace.…}` tokens lives on the parallel
> profile-engine branch. Until that branch is merged, none of the commands below run:
> there is no `--profile` option and no `example` subcommand yet. Treat the commands as
> the interface these assets are written against.

## The model

| | |
|---|---|
| Regions | `BGDXX`, `INDEA` |
| OSeMOSYS `REGION` | `GLOBAL` (single region) |
| Interconnector | `TRNBGDXXINDEA` |
| Years | 2023–2050 (28) |
| Timeslices | 20 (4 seasons × 1 day type × 5 daily brackets) |
| Technologies | 89 |
| Fuels | 43 |
| Storage | 4 — `LDSBGDXX01`, `LDSINDEA01`, `SDSBGDXX01`, `SDSINDEA01` |
| Solver | CBC |
| Scenario roots | `BAU`, `A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE` |

`C_Target_VRE` declares a completed-result dependency on `A_Calibrated_BAU`: it scales
its NDC-derived floors against solved CalBAU generation, so A must be solved first. The
dependency is declared in [`config/scenarios/registry.json`](config/scenarios/registry.json),
not left to the operator to remember.

## Layout

```
examples/unescap/
  profile.yaml                     the profile's authorities and resolution policy
  README.md                        this file
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
```

Nothing generated lives here. Preparation output, compiled parameters, solver inputs,
`.sol` files, `Executables/`, `Outputs/`, dashboards, and plots are all written under the
run workspace (`--workspace` / `OSTRAM_WORKSPACE`), never inside `examples/unescap/`.

## Commands

The profile is named, not located, so no command depends on your working directory.

```
python -m ostram example prepare unescap                       # required first — see below
python -m ostram --profile unescap run
python -m ostram --profile unescap run --scenarios "A_Calibrated_BAU,B_Optimised_VRE"
python -m ostram --profile unescap run --scenarios "C_Target_VRE"
python -m ostram example report unescap
python -m ostram example report unescap --label baseline
```

Profile-aware country commands, used by [Exercise A](exercises/add-country.html):

```
python -m ostram --profile unescap country template
python -m ostram --profile unescap country merge MMR
python -m ostram --profile unescap country validate MMR
python -m ostram --profile unescap country populate-workbook MMR
```

Profile-aware scenario commands:

```
python -m ostram --profile unescap scenario list
python -m ostram --profile unescap transform --scenario B_Optimised_VRE
python -m ostram --profile unescap compile-inputs
python -m ostram --profile unescap inspect-resources
```

`profile.yaml` sets `runtime.requires_prepare: true`. The profile ships input CSVs and a
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
transformation applied. It also documents the **2.496 vs 2.5 GW** discrepancy in the
`TRNBGDXXINDEA` residual capacity — the workbook says 2.496, the scenario YAMLs and the
exercise text say 2.5. That discrepancy is documented, not resolved.
