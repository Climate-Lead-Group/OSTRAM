# Provenance — integrated UNESCAP profile

This record describes the runnable UNESCAP example after integration with the OSTRAM
profile engine. Generated workspaces remain outside `examples/unescap/`.

## Integration lineage

| Item | Commit |
|---|---|
| Integration base | `8636ccccc324dacdd7bb7137fcbfe31d02c5e67d` |
| Profile engine component | `49bf4c6a4ade7a9f0e88ce7625964ab071ea123d` (cherry-picked here as `c6293c2`) |
| Training operations component | `1f323b17d55574e465fcd9263843847491e217d3` (cherry-picked here as `15573ba`) |
| UNESCAP seed/assets component | `cecf7f776d9890f5897fa83d66d083c7f7f55716` (cherry-picked here as `ef66110`) |
| Asset source repository | `OSTRAM_training_source`, read-only source commit `6e00e8b00144b6859344f54022df34417c075ae9` |

The integrated manifest uses only explicit `profile:`, `project:`, and `package:`
authorities. Missing profile assets fail closed; a same-named full-model resource is never
used as an implicit fallback. Shared maintained-model, execution-support, and package
resources are declared explicitly.

## Seed inputs and stage-specific domain authority

The 64 OSeMOSYS Global CSVs came from `t1_confection/OG_csvs_inputs/*.csv`. Git's
repository-wide text normalization changes CRLF to LF where applicable but does not change
CSV members or values.

The authoritative counts are stage-specific:

| Stage | Technologies | Fuels | Meaning |
|---|---:|---:|---|
| Seed | 89 | 43 | `inputs/osemosys_global/{TECHNOLOGY,FUEL}.csv`, before preparation |
| Projected seed | 89 | 43 | count-preserving preparation reconciliation |
| Prepared/compiled | 90 | 49 | valid B2 CSV domain before matrix/solver work |

Membership hashes use `sha256-sorted-values-v1`: sorted `VALUE` members, each encoded as
UTF-8 and terminated by LF.

| Stage / set | SHA-256 |
|---|---|
| Seed TECHNOLOGY | `a866b63a168d413cec240b31598bbabafb1f8da99e202caf05d4e1bf5373c66b` |
| Seed FUEL | `c45a6db17552f4a3527bfa7fc8d014fee1e46457a6fabd81b77f4661b04f367c` |
| Projected-seed TECHNOLOGY | `b198cdaa23bcff3157668ead9a0578fd18f63a08127529cb344205d8c1a67f35` |
| Projected-seed FUEL | `c45a6db17552f4a3527bfa7fc8d014fee1e46457a6fabd81b77f4661b04f367c` |
| Compiled TECHNOLOGY | `6266dd062c5d593be8cb62f5169eb2759f4a7301200e26d147bb4f934ebec4d3` |
| Compiled FUEL | `e924cb09c945ad23e1b3c7bb2f760dce245eafe4fbeca065207fd7868d8501e8` |

Preparation first removes four duplicate legacy `*00` technology identities, normalizes
the declared `PWR...01` identities, removes ten matrix-filtered identities, and adds the
fourteen A1/A2 canonical dispatch/transmission identities listed in `profile.yaml`. That
projection remains exactly 89 technologies. The only net compiled technology addition is:

- `PWRSHPINDEA`

The only compiled fuel additions are:

- `ELCBGDXX00`, `ELCBGDXX03`, `ELCBGDXX04`
- `ELCINDEA00`, `ELCINDEA03`, `ELCINDEA04`

The engine validates both count and membership at seed preparation, then validates the
exact projected-seed-to-compiled delta immediately after B2 CSV materialization and before
text preprocessing, matrix creation, or solver dispatch. Same-count substitutions,
undeclared additions, and removals all fail closed.

## Scenario workbook transformation

Source:
`t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx` at source commit
`6e00e8b00144b6859344f54022df34417c075ae9`.

| Property | Source | Integrated seed |
|---|---:|---:|
| Size | 257,429 bytes | 162,918 bytes |
| SHA-256 | `f9e3ed1fc720c3c2906cd3ffc18e0023ea18df7f977260580c6d7f37a82019a8` | `4cf3323f28fe3da2c552433a9083ee4d4ee03d7c249a14d1925fa4afb3b41253` |
| `Restrictions` data rows | 2,537 generated rows | 0; header retained |

The cleanup was archive-preserving rather than a workbook resave. The source and result
both contain 59 ZIP members in the same order. Only `xl/worksheets/sheet3.xml` changed;
all other member bytes and archive metadata are identical. All 20 sheets, the
`Restrictions` header, `Control` contents, formulas, formatting, validations, defined
names, and unrelated worksheet XML were preserved.

## Reduced mappings and preparation authorities

`config/scenarios/technology_types.csv` is a deliberate reduced projection of the
386-row full source. It has 92 physical lines and 90 nonblank, unique mappings. It covers
the complete compiled 90-technology taxonomy, including `PWRSHPINDEA`, the A2-generated
dispatch/transmission technologies, four storage technologies, and the single physical
interconnector `TRNBGDXXINDEA`. Irrelevant full-model rows were not restored.

`config/scenarios/ao_extension_decisions.csv` contains only the live reduced decisions:
`PWRNGSBGDXX` and `PWRSHPINDEA`. The reduced base-year pin authority contains 510 rules
over 50 technologies and has SHA-256
`984c3885f7bcee992d634c602402e3c6183a2f3bc8a1d8a0620ae214c1a1d872`.

The profile also carries the required preparation sources absent from the initial asset
copy: `Tech_Country_Matrix.xlsx`, the reduced timeslice workbook, country centerpoints,
and secondary-technology inputs. Exercise A's ELC dispatch nodes are derived from the
active profile country list; the shipped BGD/INDEA list yields `ELCBGDXX01` and
`ELCINDEA01`, and adding a country extends the set through configuration rather than a
source-code list.

## Interconnector authority

`OSTRAM_Scenario_Inputs.xlsx`, sheet `Interconnector_Params`, is authoritative for
`TRNBGDXXINDEA` residual capacity:

- exactly `2.496` GW for 2023–2028;
- `2.996` GW in 2029;
- `3.746` GW in 2030–2032;
- `4.496` GW in 2033–2050.

Every shipped executable scenario schedule begins at exactly `2.496`; none rounds the
seed to `2.50`. Training prose that says “about 2.5 GW” is descriptive. An exact 2.500-GW
value, including the historical embedded direction-results snapshot, is an exercise edit
or rounded reference result, not the baseline authority.

The active `B_Optimised_VRE` relaxation intentionally lifts
`TotalAnnualMaxCapacity(TRNBGDXXINDEA, 2023..2050)` to exactly `9999.0`. This is the lid
rule's explicit unbinding value; it neither overwrites nor rounds `ResidualCapacity`.

## Runtime integration decisions

### Bangladesh gas emissions and training A3 checks

Before the training reproduction repair (commit
`6ff8fc91c89b81a5e44cbf6798f0ec30a0fbc77a`) the profile-owned
`EmissionActivityRatio.csv` had **no** `PWRNGSBGDXX` rows (nor legacy Bangladesh
CCG/OCG emissions to reconcile). A3 creates the Bangladesh gas technology from
the maintained extension decision, but that does not create its emissions rows.
The shared B1 combustion check therefore correctly rejects mode 1 in 2023.
This is independent of the valid `URN` nuclear mappings.

Adding the missing rows alone is insufficient: the training technology-country
matrix deliberately excludes Bangladesh gas from A1, before A3 introduces it.
A1 consequently discards the new emissions too. The profile now opts into
`preserve_declared_extension_emissions`: A1 retains only emissions whose exact
technology is an `Include=Y` A3 extension and is present in the active taxonomy.
The other parameters still follow the matrix, so the late extension's activity
ratios, costs and capacity data are not replaced with legacy seed values.
The full profile retains the existing filter semantics by default.

The correction adds exactly 56 rows: `GLOBAL / PWRNGSBGDXX / CO2BGD`, modes 1
and 2, years 2023–2050. The prepared input activity ratio is `2.33765`; B1's
existing four-decimal conversion yields `2.3376` PJ gas per PJ activity. Each
mode produces `ELCBGDXX01` at output ratio 1, so its activity is PJ of plant
electricity output. The compiler's approved fuel factor is `0.0561` Mt CO₂ per
PJ gas. Thus both modes require `0.0561 × 2.3376 = 0.13113936` **Mt CO₂ per PJ
activity**. Mode 1 consumes `GASBGD`; mode 2 consumes `GASINT`. Both combust in
Bangladesh and emit `CO2BGD`. No emissions are added at gas mining/import, and
no full-model rows, technologies, fuel mappings, efficiencies or capacities
are imported. The compiler's exact factor and output-basis checks remain active.

The A3 extension and manual-fix checks use the active profile's explicit
technology taxonomy to determine applicability: out-of-domain substitutions (for
example `PWRGEOINDNO`, `PWRNGSMDVXX`, `PWROILMDVXX`, `PWROILNPLXX`) are logged,
missing or duplicate in-domain rows still fail, small-hydro validation requires
exactly the declared set (`PWRSHPINDEA` in training, all five Indian codes in
`full`), and both stages exit unsuccessfully when an applicable check fails.

CBC marked storage activities with `**` bound violations (largest about `1.27e-6`)
on the first A solve. The training execution configuration therefore requests
primal tolerance `1e-10` for A/B and, through `cbc_primal_tolerance_by_scenario`,
`1e-12` for C (the tighter value caused numerical cycling on A), with dual
tolerance `1e-8` and seed 12345. The shared CBC adapter keeps its `1e-8` primal
default for other profiles and rejects material saved `**` violations even when
the header says optimal; negative postsolve values within `1e-8` of zero remain
within tolerance. No model coefficient is changed to repair numerical feasibility.

Training also enables the `reconcile_variable_cost_exports` policy: the first A
export omitted about 60.775 MUSD of compiled default variable costs, which
reconciliation accounts for in the combined cost tables while retaining the raw
CSVs. When comparing reported cost to the CBC objective, include the constant
recorded in GLPK's LP comment (26,631.6286009204 MUSD for A), which GLPK omits
from the solver objective. Exported values have about eight significant digits
and are checked at `1e-7` scaled row/bound residual.

- `lid_rule_new_semantics: true` applies only to UNESCAP; `full` and unqualified behavior
  retain the historical semantics.
- `storage_delay_active` reads the declared maintained model and writes a scenario-local
  patched model under `Executables/<scenario>_0/`; it never resolves a removed legacy
  patched-model path or modifies the maintained model.
- Preparation is stamped atomically. Repeated preparation requires `--reset`, refuses
  foreign/unstamped workspaces, and profile workspaces are isolated by profile id.
- Reporting/capture and resource inspection resolve the activated profile bundle.

## Accepted training results (2026-09-22)

Final A and B completed native compilation, CBC solving, export and cost
reconciliation from an empty workspace; C then completed the same chain using that
workspace's freshly generated A result. All three CBC headers declare optimal
solutions (Windows, Python 3.10.21, CBC 2.10.13, GLPK 5.0, otoole 1.1.5; CBC primal
tolerance `1e-10` for A/B and `1e-12` for C, dual tolerance `1e-8`, seed 12345).

| Scenario | CBC objective (MUSD) | Reconciled total (MUSD) | Backstop generation, 2023–2050 (PJ) |
|---|---:|---:|---:|
| A_Calibrated_BAU | 586,430.07677182 | 613,061.705301 | 0.000000 |
| B_Optimised_VRE | 555,111.93041688 | 581,743.559485 | 0.000000 |
| C_Target_VRE | 540,925.25041220 | 567,556.879320 | 0.000000 |

The reconciled total includes GLPK's omitted objective constant; it is compared with
that constant plus the CBC objective, and every cost difference is within `1e-7`
relative tolerance. All three cases have zero backstop generation and capacity while
the profile's backstop remains enabled.
