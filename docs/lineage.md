# Runtime, scenario, and data lineage

This is the central guide to the current OSTRAM product tree. It separates
maintained authority from generated runtime state, explains how a scenario is
constructed, and follows each important input class to the compiled OSeMOSYS
datafile and eventual results.

The layout is intentionally interim. Pull Request A retains root `run.py`, the
repository-local `ostram` command module, and the complete
`t1_confection/` runtime shell. A later Pull Request B may move Python modules
and remove caller-working-directory assumptions. Until then, run documented
commands from the repository root and treat the external Stage 6
path-assumption ledger as migration evidence, not as runtime configuration.

## Install and prepare

OSTRAM is currently a Windows, repository-local application. Install Git,
Miniconda or Anaconda, and the environment in `environment.yaml`; detailed
steps are in [Installation](installation.md). A production optimization also
requires a solver supported by `t1_confection/Config_MOMF_T1_AB.yaml`.

From an Anaconda Prompt:

```powershell
conda env create -f environment.yaml
conda activate OSTRAM-env
python -m ostram --help
```

The command does not install a bare `ostram` executable. Python dependencies
are fixed by the environment file, while exact numerical reproducibility also
depends on the selected solver, its version, and its settings.

## Public commands

These interfaces remain supported in Pull Request A:

| Purpose | Preferred command | Retained direct command |
| --- | --- | --- |
| Full conditional A1/A2, then A3/B1/B2 | `python -m ostram run [args]` | `python run.py [args]` |
| One root A3 transformation | `python -m ostram transform [args]` | `python t1_confection/A3_process.py [args]` |
| Selected-scenario B1 compilation | `python -m ostram compile-inputs [args]` | `python t1_confection/B1_Run_Compiler.py [args]` |
| Exact root/derived materialization | — | `python t1_confection/scenario_materializer.py [args]` |
| B2 preparation, optional solve, and results | — | `python t1_confection/B2_Executing_OG_Model.py [args]` |
| A0 technology-country workbook generation | — | `python t1_confection/A0_generate_tech_country_matrix.py` |
| Optional secondary-technology editing | — | `python t1_confection/D1_generate_editor_template.py`, then `python t1_confection/D2_update_secondary_techs.py` |

`run` and the direct B2 command are production commands: with the live
configuration they can install dependencies, change generated state, create a
matrix, and execute a solver. `--help` only proves parser reachability.
There are no `prepare-model` or `solve` subcommands because B2 owns one
configuration-driven preparation/solve/result boundary.

For a solver-free B2 boundary after A1, A2, materialization, and B1:

```powershell
python t1_confection/B2_Executing_OG_Model.py `
  --scenarios "A_Calibrated_BAU,B_Optimised_VRE" `
  --compile-only
```

`--compile-only` compiles the final text input and returns before matrix
creation, solver execution, cleanup, solution conversion, and result
concatenation. It is a validation boundary, not a production optimization.

## Scenario authority

### Four maintained roots

`t1_confection/A3_process/OSTRAM_Scenario_Inputs.xlsx::Control` contains
exactly four active roots:

1. `BAU`
2. `A_Calibrated_BAU`
3. `B_Optimised_VRE`
4. `C_Target_VRE`

The workbook also owns scenario parameters, inherited Restrictions, the
`Interconnector_Params` table, the complete maintained RNWBIO
`VariableCost` rows, and the 11 human `AO_Extension_Decisions`. A3 works on a
disposable scenario-state copy; generated Restrictions are never persisted to
the authoritative workbook.

`t1_confection/A3_process/OSTRAM_Timeslice_Inputs.xlsx` is the second
maintained A3 workbook. Its 23 sheets supply the 20-timeslice fabric. A3 stages
a copy, merges it into the scenario workbooks, and synchronizes the temporal
CSV/YAML representation without mutating this authority.

### Derived scenarios

`t1_confection/scenario_registry.json` is the canonical registry. A derived
scenario selects one of the three decision roots, then applies only its
declared patch and optional direction overlay:

| Derived scenario | Base | Additional declaration |
| --- | --- | --- |
| `A_Calibrated_BAU_Clipped` | `A_Calibrated_BAU` | `patches.json` |
| `B_Opt_Clipped` | `B_Optimised_VRE` | `patches.json` |
| `B_Opt_DirBidir` | `B_Optimised_VRE` | patch plus bidirectional overlay from 2027 |
| `B_Opt_DirContractual` | `B_Optimised_VRE` | patch plus contractual overlay from 2027 |
| `B_Opt_IndiaCosts` | `B_Optimised_VRE` | `patches.json` |
| `B_Opt_IndiaCostsFuel` | `B_Optimised_VRE` | `patches.json` |
| `B_Opt_SolarCapex130` | `B_Optimised_VRE` | `patches.json` |
| `B_Opt_SolarCapexHi` | `B_Optimised_VRE` | `patches.json` |
| `B_Opt_SolarCapexSpike` | `B_Optimised_VRE` | `patches.json` |
| `B_Opt_TradeCap15` | `B_Optimised_VRE` | `patches.json` |
| `B_Opt_TxCap150` | `B_Optimised_VRE` | `patches.json` |
| `C_Target_VRE_Clipped` | `C_Target_VRE` | `patches.json` |

Patch files live under
`t1_confection/A3_process/rules_scripts/configs/<scenario>/`. The
materializer rejects unknown names, resolves the required base, copies its
materialized workbooks, applies the patch atomically, applies a declared
direction overlay, and records provenance. Derived names are not extra Control
rows and must not be maintained as hand-copied A1/A2 folders.

### C depends on an A result

`C_Target_VRE` is not a single-pass cold-start scenario. Its VRE-target rule
uses the accepted result from `A_Calibrated_BAU` to derive its maintained
target transformation. Production therefore requires two passes:

```powershell
# Pass 1: generate and solve A, producing its real result.
python -m ostram run --skip-pull --scenarios "A_Calibrated_BAU"

# Inspect the successful A result, then materialize/compile/solve C.
python -m ostram run --skip-pull --scenarios "C_Target_VRE"
```

The first pass must complete far enough to create
`t1_confection/Executables/A_Calibrated_BAU_0/Pre_processed_A_Calibrated_BAU_0_StorageDelayN5_OpenBCK_RMCarefulXLSX_output.csv`.
The second pass consumes that result while materializing C. The frozen
external A-result seed exists only for authenticated, solver-free acceptance
runs; it is not a production data source.

## Runtime order and ownership

The ordinary order is:

| Stage | Reads maintained authority | Writes generated state | Important overwrite behavior |
| --- | --- | --- | --- |
| A0, manual | country YAML | `Tech_Country_Matrix.xlsx` | Regenerates the workbook; user decisions in it must be reviewed before replacement. |
| A1, conditional | 64 `OG_csvs_inputs/*.csv`, matrix, country/region YAML, `Miscellaneous/` templates | `A1_Outputs/A1_Outputs_BAU/`, four `A2_Extra_Inputs/*.xlsx` | Normalizes selected temporal/model CSVs back into `OG_csvs_inputs/` and updates compiler temporal settings. |
| A2, conditional | BAU A1 workbooks, country YAML, extra inputs | transmission-enriched BAU and `_post_a2_snapshot_BAU/` | Replaces the post-A2 snapshot after successful assembly. |
| A3 | post-A2 snapshot, two maintained A3 workbooks, registry and rule data | one `A1_Outputs/A1_Outputs_<scenario>/` set per selected identity | Restores the snapshot first, then applies template, AO, interconnector, pin, root-rule, patch, direction, and dependency behavior in declared order. |
| B1 | selected A3 workbooks, `Config_MOMF_T1_A.yaml`, conversion schema/templates | `A2_Output_Params/<scenario>/` and optionally `A2_Outputs_Params_otoole/<scenario>/` | Recompilation replaces generated parameter CSVs for the selected scenario. |
| B2 | selected parameter CSVs, `Config_MOMF_T1_AB.yaml`, model and patch inputs | `Executables/<scenario>_0/`, root datafile, and—unless compile-only—`Outputs/` and combined result CSVs | Patch order can replace the intermediate datafile; solve/result routes are skipped only by the explicit compile-only gate. |

A0 is not called by `run.py`. A1 and A2 run only when the post-A2 snapshot is
absent. A3, B1, and B2 otherwise receive the same selected scenario set.
Optional D1/D2 editing sits between A3 and B1 and is outside `run.py`.

## Authoritative input-to-result lineage

The rows below are exhaustive by tracked path family. Each family names its
maintained owner, reader, generated intermediates, later overwrite point,
scenario scope, compiled destination, and result domain.

| Input class and maintained owner | Reader / writer | Generated intermediate and later overwrite | Scenario scope | Final compiled destination and eventual influence |
| --- | --- | --- | --- | --- |
| Core OSeMOSYS sets and parameters: all 64 tracked files in `t1_confection/OG_csvs_inputs/*.csv` | A1 loads the complete directory; A1 and A3 temporal sync are the only pipeline writers back to this family | A1 model workbooks; A2/A3 may transform the workbook representation; B1 emits one CSV per set/parameter | All roots and every derived scenario through its base | Same-named blocks in the per-scenario `Pre_processed_*.txt`; capacity, activity, demand, trade, storage, emissions, and cost results according to the parameter |
| Technology/country inclusion and aggregation: `Tech_Country_Matrix.xlsx::{Matrix,NGS_Unification,Aggregation_Rules,Tech_Reference,Country_Reference}` | A0 generates; user maintains decisions; A1 reads | A1 filters/unifies/consolidates before workbook creation; later scenario rules operate on the surviving technologies | All scenarios | `TECHNOLOGY` and all technology-indexed parameter blocks; therefore feasible build/activity and their costs/emissions |
| Country and transmission configuration: `Config_country_codes.yaml` and `Config_region_consolidation.yaml` | A0, A1, A2, A3 interconnector stages, and auxiliary country tools | A1 region/technology transformations; A2 transmission rows; A3 internal transmission/loss rows; later scenario rules may cap declared links | All scenarios, with scenario overlays where declared | Region, technology, activity-ratio, capacity, cost, life, and limit blocks; trade flows, capacity, activity, cost, and emissions |
| A1 compiler structure: `Config_MOMF_T1_A.yaml` plus the staged `A3_process/Config_MOMF_T1_A.yaml` | A1 updates temporal values; A3 stages/synchronizes; B1 reads | B1 parameter CSVs and otoole CSVs, overwritten on recompilation | Selected scenarios | Shapes every compiled set/parameter block and temporal index; therefore all results |
| B2 execution and patch configuration: `Config_MOMF_T1_AB.yaml` | B2 reads; no earlier stage writes it | otoole datafile, storage-delay/strip/backstop/reserve-margin patched datafile; later patch replaces the prior intermediate | Selected scenarios | `Executables/<scenario>_0/Pre_processed_<scenario>_0_<suffix>.txt`; selects solver/result route and can affect storage, firm capacity, costs, activity, and feasibility |
| A1 workbook templates: `t1_confection/Miscellaneous/A-O_*.xlsx`, `A-Xtra_Emissions.xlsx`, and `A-Xtra_Storage.xlsx` | A1 reads and populates | A1 BAU workbooks and `A2_Extra_Inputs` outputs; A2/A3 later transform scenario copies | All scenarios through the BAU snapshot | Demand, ratios, technology parameters, emissions, and storage blocks; corresponding activity, capacity, cost, storage, and emissions results |
| A1 extra input outputs: the four tracked files `A2_Extra_Inputs/A-Xtra_{Battery_Replacement,Emissions,Projections,Storage}.xlsx` | A1 is the governed writer; A2 consumes them | A2 BAU workbook/snapshot; overwritten when A1 is deliberately rerun | All scenarios through the post-A2 snapshot | Projection, storage, emissions, and replacement-related parameter blocks; corresponding capacity/activity/cost/emission results |
| Scenario root authority: `OSTRAM_Scenario_Inputs.xlsx::{Control,Restrictions,AO_Extension_Decisions,Interconnector_Params,VariableCost,...}` | scenario registry helpers, A3, AO overlay, and interconnector-cost stage read; only a disposable scenario-state copy receives generated Restrictions | Root workbooks; AO rows are regenerated then 11 decisions overlaid; explicit interconnector authority is applied later than generic AO propagation | Four roots; derived scenarios inherit exactly their registry base | Scenario-specific parameter blocks including RNWBIO `VariableCost`; objective cost, generation/activity, capacity, trade, and emissions |
| Timeslice authority: all 23 sheets in `OSTRAM_Timeslice_Inputs.xlsx` | A3 timeslice merge reads a staged copy | 20-timeslice workbook fabric; A3 sync then overwrites temporal CSV/YAML representations for the materialized run | Every materialized scenario | `TIMESLICE`, `YearSplit`, `DaySplit`, conversions, profiles, and time-indexed parameters; all time-resolved activity/capacity-factor/storage results |
| Scenario selection and derivation: `scenario_registry.json` and `rules_scripts/configs/<scenario>/*.{yaml,json}` | registry/materializer select bases; A3 rule scripts and patch engine apply declarations | Root snapshots, then derived copies; a derived patch/direction overlay is deliberately last within its declared responsibility | Exact registered identity only | Any parameter named by the root rule or patch; corresponding result measure |
| Technology taxonomy: `A3_process/TECH_TYPES.csv` | AO, retirement, VRE, capacity-floor, lid, direction, and relaxation rules | Rule-specific workbook edits; later explicit owner can overwrite only its own declared cells | Root rules and derived copies | Technology-indexed capacity, activity, cost, VRE, storage, and interconnector blocks; matching results |
| Maintained interconnector authority: scenario workbook `Interconnector_Params`, `Config_country_codes.yaml`, `rules_scripts/internal_tx_residuals.csv`, and the two registered direction-overlay YAMLs | A3 cost, internal-transmission, loss, relaxation, and direction stages | A-O parametrization rows; explicit cost/residual/minimum and direction stages supersede generic generated rows | Roots plus declared B direction variants | `ResidualCapacity`, `CapitalCost`, `OperationalLife`, activity ratios, annual capacity/investment limits; trade flow, investment, cost, and emissions |
| Base-year PWR/MIN pin: `rules_scripts/pwr_min_2023_2026_pin.csv` | `apply_base_year_pin.py` | Late-A3 allowlisted capacity cells; it does not rewrite the authority CSV | Root-gated scenarios and their derived copies | PWR/MIN capacity limit blocks for 2023–2026; feasible capacity/activity and cost |
| B1 conversion contract: `Miscellaneous/conversion_format.yaml` and every header template in `Miscellaneous/templates/*.csv` | B1/preprocess utilities and optional D2 default lookup | otoole-format parameter CSVs; regenerated on every B1 run | Selected scenarios | Ordering/schema/defaults of all compiled blocks; indirect influence through correct model interpretation |
| Model and B2 repair inputs: `osemosys_fast_preprocessed.txt` and `firm_capacity_fallbacks_by_cr.xlsx::fallbacks` | B2 and reserve-margin patcher | model matrix when enabled; careful reserve-margin-patched text datafile | Selected scenarios under configured toggles | Solver equations plus reserve capacity-credit/cap fallback values; feasibility, capacity, activity, and total cost |
| Optional D1/D2 authorities: installed-capacity workbook, generation workbook, PET/OIL shares, power-generation shares, and user-edited `Secondary_Techs_Editor.xlsx` | D1 generates the editor; user edits; D2 reads all selected sources and writes materialized scenario workbooks | A3 scenario workbooks are edited after materialization; a later A3 rematerialization overwrites those manual edits | Scenarios selected in the editor | Residual capacity, demand, activity limits, shares, and secondary-technology parameters; associated capacity/activity/cost/emission results |
| Auxiliary new-country template: tracked `templates/MDV/*.csv` and `centerpoint.csv` | the template merge helper reads; a user explicitly merges into `OG_csvs_inputs` | OG inputs after reviewed merge; then ordinary A1 lineage applies | Only a user-created country and later scenarios | Same-named OSeMOSYS blocks and their result domains |
| Sensitivity-construction references: `sensitivity_expansion/reference/vre_ceilings.csv` and `vre_ceilings_base.json` | sensitivity generation tooling only | proposed patch declarations; current accepted identities still come only from `scenario_registry.json` | No implicit production scenario | No compiled effect until a reviewed registry patch explicitly references the generated declaration |

The named families intentionally distinguish a maintained writer from a later
overwrite. In particular, generated A1/A2 trees are not source authority,
derived folders are not scenario definitions, and the external comparator is
not read by production.

## Generated versus tracked

Tracked content includes code, the configuration and authority families above,
documentation, and maintained tests. The following runtime roots are ignored
and must be recreated from tracked inputs:

- `t1_confection/A1_Outputs/`
- `t1_confection/A2_Output_Params/`
- `t1_confection/A2_Outputs_Params_otoole/`
- `t1_confection/Executables/`
- `t1_confection/Outputs/`

Generated editor workbooks, root datafiles, matrices, solutions, result CSVs,
logs, caches, backups, and date-stamped outputs are also runtime products.
Their presence in a working directory never makes them authority. The tracked
`dvc.yaml` and `dvc.lock` describe older partial data-versioning state;
`run.py` may call `dvc pull`, but does not use `dvc repro` as the production
orchestrator.

## Validation evidence is not production input

Housekeeping acceptance regenerates a fresh tracked-only candidate and compares
its 15 compiled `.txt` files with the authenticated external
`STAGE_2_GOVERNED_COMPARATOR_MANIFEST.csv`. The maintained
`tests/regression/accepted_baseline.py` utility checks exact scenario order,
authority class, SHA-256, byte size, and line count. It neither normalizes
files nor runs a solver.

Production does not read the comparator, a reference repository, an evidence
report, or the frozen A seed. These are validation evidence only. See
[Regression and cleanup acceptance](regression.md) for the guarded command.

## End-to-end lineage

For any ordinary parameter, the lineage is:

```text
maintained CSV/workbook/YAML/configuration
  -> A1 BAU workbook
  -> A2 transmission-enriched post-A2 snapshot
  -> A3 root rule / declared derived patch and dependency
  -> B1 per-parameter CSV
  -> B2 per-scenario Pre_processed_*.txt
  -> optional matrix and solver
  -> per-scenario output CSV
  -> combined Inputs / Outputs / Combined_Inputs_Outputs result files
```

Generated intermediates are inspectable but replaceable. To change a result
reproducibly, change the maintained owner, rematerialize from the correct base,
and preserve the declared overwrite order.
