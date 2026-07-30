# A1/A2-to-B1 transformation contract

> **Historical transformation contract:** The branch-scoped no-solver evidence and
> claims below are preserved as recorded. The portable
> [final 15-scenario baseline](../tests/regression/reports/accepted_compiled_solver_baseline_15.json)
> is historical source-bound evidence. Current acceptance authority for
> derived identities is the authenticated external Run 3
> `STAGE_2_GOVERNED_COMPARATOR_MANIFEST.csv`, generated from current maintained
> roots plus declared rules.
>
This document freezes the transformation boundary exercised by
`refactor/a1-b1-transforms`. It supplements the
[core workflow characterization](core-workflow-characterization.md) and the
[offline regression policy](regression.md).

This is a behavior-preserving Tier 3 structural refactor. Its purpose is to make
the established transformation plan, pure table operations, I/O effects,
validation, and delivery independently testable. It does not authorize a model
calculation, data, scenario, path, filename, writer-order, message, exit, or failure
semantic change. Compiled-input equality is required pre-solver evidence; it is not
a solver-backed behavioral or numerical equivalence claim.

The governing branch sequence recorded at this checkpoint was:

1. `refactor/a1-b1-transforms` -- structural isolation only;
2. `fix/interconnector-v18-source-of-truth` -- the merged residual-capacity
   authority correction;
3. `refactor/retire-naty-workbook` -- the behavior-preserving authority
   migration recorded here; and
4. `validation/solver-baseline-15` -- separately scoped work outside this
   branch and outside the evidence recorded here; it was subsequently completed
   and superseded as current authority by the accepted baseline linked above.

Work assigned to a later branch must not be prepared silently on this branch.
The separately authorized `fix/v18-pwr-min-2023-2026-pin` branch is a narrow
exception to the historical behavior-preserving scope: it changes only the
frozen non-Maldives 2023--2026 PWR/MIN allowlist through a root-gated late-A3
stage. Historical validation records below remain scoped to their named commits.

## Frozen predecessor chain

The outer workflow remains
`run.py` -> A1 -> A2 -> per-active-scenario A3 -> B1 runner -> B1 compiler -> B2.
The snapshot gate, skip flags, scenario propagation, child working directories,
environment inheritance, messages, exit behavior, and failure propagation described
in the core workflow characterization remain unchanged. The refactor does not move
or replace a public entrypoint.

### A1: source preprocessing and scenario artifacts

[`A1_Pre_processing_OG_csvs.py`](../t1_confection/A1_Pre_processing_OG_csvs.py)
retains its established sequence:

1. read and sort the original model CSV inputs;
2. normalize the temporal inputs;
3. apply the `JAM`-to-`BRB` handling;
4. filter the configured year range;
5. build and apply the technology matrix filtering;
6. unify the NGS representation;
7. consolidate configured regions;
8. clean and merge the PWR data;
9. discover scenario definitions;
10. update demand, parametrization, emissions, base-year, projection,
    configuration, and storage artifacts in their predecessor order; and
11. rewrite the original-format CSV outputs.

The predecessor's exception boundaries are part of the contract. Workbook-update
errors that were reported and continued remain report-and-continue; the uncaught
base-year update route remains fail-fast. This branch does not make A1 transactional
or change its model calculations.

### A2: scenario assembly

[`A2_AddTx.py`](../t1_confection/A2_AddTx.py) continues to discover scenario
directories in sorted order. Within each scenario it applies base-year,
projections, parametrization, and demand work in that order. Its established BAU
snapshot replacement remains remove-then-copy after scenario processing. Argument
parsing placement, path interpretation, partial-state behavior, and native failures
remain unchanged even where they are awkward.

### A3: ordered scenario materialization

[`A3_process.py`](../t1_confection/A3_process.py) and
[`a3_orchestrator.py`](../t1_confection/a3_orchestrator.py) retain the characterized
stage sequence:

`stage_1_scripts_1_to_5` -> `stage_1b` ->
`stage_2_and_2_5` -> `stage_3_fix_2` -> `stage_4_consolidate` ->
`stage_4_5_apply_inherited_restrictions` -> `stage_5_rules_scripts` ->
`stage_ws3_interconnector_costs` -> `stage_ws3_internal_transmission` ->
`stage_ws3_internal_tx_losses` -> `stage_ws4_pwr_min_pin` ->
`stage_6_sync_og_to_ts20` ->
`stage_6_persist_restrictions` -> `deliver_outputs`.

Normal A3 execution still contains only the four active Control definitions:
`BAU`, `A_Calibrated_BAU`, `B_Optimised_VRE`, and `C_Target_VRE`. Derived decision
scenarios remain later, ordered scenario-specific patches; this refactor does not
promote them into the active A3 loop or change patch precedence. A3 helper failure
remains fail-fast, while work-directory and `OSTRAM_TEMPLATE_PATH` cleanup remain
success-path-only.

The late `stage_ws4_pwr_min_pin` dispatches only for the exact
`A_Calibrated_BAU`, `B_Optimised_VRE`, and `C_Target_VRE` roots. It applies the
version-controlled `pwr_min_2023_2026_pin.csv` allowlist to existing workbook
rows before Stage 6. The transformer validates the allowlist digest and contract,
writes only its explicit 2023--2026 PWR/MIN cells, never reads solver output, and
does not emit a generic `*_CHANGES.json`. Plain `BAU` is untouched; derived
scenarios inherit their corrected root before their bounded sensitivity patches.

The materialized v18 `Interconnector_Params` sheet is the sole runtime authority
for three distinct interconnector families: final `ResidualCapacity`, additive
`TotalAnnualMinCapacityInvestment` FUTURE contributions, and the separate
`MinimumInvestmentClampBoundary` compatibility family. The additive family
contains only the contribution merged into any pre-existing minimum; it is not a
copy of the final minimum row. The clamp family is used only by the two LinkFreeze
minimum clamps and is never interpreted as installed residual capacity or a FUTURE
addition. The former NATY workbook is not staged, passed on the command line, or
available as a runtime fallback.

### B1: selection, compilation, and CSV delivery

[`B1_Run_Compiler.py`](../t1_confection/B1_Run_Compiler.py) remains the public CLI
and [`b1_runner.py`](../t1_confection/b1_runner.py) retains scenario discovery,
filtering, config backup/update/restoration, and compiler dispatch. Eligible
scenario directories are discovered in sorted order. A valid requested filter is
resolved in discovery order and collapses requested duplicates. The child command
remains the current interpreter plus the absolute `B1_Compiler.py` path, with
`t1_confection/` as its working directory and the complete environment inherited.

[`B1_Compiler.py`](../t1_confection/B1_Compiler.py) remains a directly executable,
top-level compiler. Importing it is still not a supported way to obtain helper
functions because import executes compiler work. Its established transformation
groups remain ordered as follows:

1. activity-ratio tables;
2. demand tables and demand-timeslice selection;
3. parametrization and optional transport tables;
4. `YearSplit`;
5. `DaySplit`;
6. emissions, including the original-module route when configured;
7. other-technology/NDP accumulation;
8. storage tables;
9. technology-to-storage and storage-to-technology conversion modes;
10. system parameters; and
11. intermediate workbook and final CSV delivery.

Configuration and all relative input paths retain current-working-directory
semantics. `Config_MOMF_T1_A.yaml` is read by its relative name. A1 scenario
workbooks retain the exact nested formula
`os.path.join(A1_outputs, A1_outputs + "_" + Main_Scenario + suffix)`; A2 extra
inputs retain string-concatenation semantics; the original emissions source remains
`OG_csvs_inputs/EMISSION.csv`. Reads remain lazy at their predecessor positions so
a missing later source fails only when that source is reached.

The fixed delivery order is:

1. completed demand workbook (`Print_Dem_Completed`);
2. completed parametrization workbook (`Print_Paramet_Completed`);
3. completed natural parametrization workbook
   (`Print_Paramet_Natural_Completed`);
4. completed projections workbook (`Print_Proj_Completed`);
5. the configured structure workbook (`Print_A2_Struct_List`, currently the
   `A2_Structure_Lists.xlsx` contract);
6. main-scenario parameter CSVs in parameter-mapping insertion order;
7. main-scenario set CSVs in sorted set-name order; and
8. each configured additional scenario in configured order, including duplicates,
   with parameter and set CSVs sorted by name.

Parameter and set filenames remain `<name>.csv`. The main output directory remains
`os.path.join(A2_output_main_scen, Main_Scenario)` and each additional output
directory remains `os.path.join(A2_output, scenario)`. Existing files are not swept
before delivery. A write error can therefore leave already-written workbooks or
CSVs and can coexist with stale files; this branch characterizes that behavior and
does not repair it.

### B2: final compiler delivery and the solver boundary

[`B2_Executing_OG_Model.py`](../t1_confection/B2_Executing_OG_Model.py) and
[`b2_orchestrator.py`](../t1_confection/b2_orchestrator.py) retain the compile path
from B1 CSV directories through `process_scenario_folder`, otoole conversion, and
the ordered preprocessing/patch chain documented in the core workflow
characterization. The canonical final filename suffix remains
`Pre_processed_<scenario>_0_StorageDelayN5_OpenBCK_RMCarefulXLSX.txt`.

B2 is also the matrix and solver boundary. Compile-only validation must prove that
`execute_model`, `create_matrix`, `concat_otoole_csv`, and
`concat_scenarios_csv` are each the boolean `False` before invocation and must
install fail-closed barriers around matrix and solver process routes. This contract
does not authorize a matrix build or solver run.

## Extracted production boundaries

The new import-safe package is
[`t1_confection/a1_b1_transforms/`](../t1_confection/a1_b1_transforms/). Importing
any module in it must not read configuration, open a workbook, create a writer,
change the current working directory, launch a process, or run any pipeline stage.
The package does not make `B1_Compiler.py` itself import-safe and does not add a new
public CLI.

| Module | Responsibility | Preserved dependency boundary |
|---|---|---|
| [`planning.py`](../t1_confection/a1_b1_transforms/planning.py) | Derive the inclusive year vector, setup mapping, and exact lazy path formulas from the loaded configuration. | No filesystem access. The configured `Timeslices` list is still sorted in place; paths remain unnormalized strings. |
| [`tables.py`](../t1_confection/a1_b1_transforms/tables.py) | Normalize year-like column labels, expand system-parameter rows, and construct padded structure/set tables. | Pure pandas/numpy inputs and returned tables; no workbook or CSV effects. |
| [`effects.py`](../t1_confection/a1_b1_transforms/effects.py) | Read YAML, CSV, workbook, and pickle inputs and write completed/structure workbooks. | Openers, readers, loaders, writer factories, and frame writers are injectable while default calls retain predecessor semantics. |
| [`validation.py`](../t1_confection/a1_b1_transforms/validation.py) | Apply the demand, capacity-factor, `YearSplit`, and `DaySplit` configuration checks. | Reporting and stopping are injectable; default text and bare `sys.exit()` behavior remain exact. |
| [`delivery.py`](../t1_confection/a1_b1_transforms/delivery.py) | Clean the delivery mappings and emit main/additional parameter and set CSVs in the frozen order. | Directory creation, frame construction, and CSV writing are injectable; no implicit cleanup is added. |
| [`__init__.py`](../t1_confection/a1_b1_transforms/__init__.py) | Declare the package without importing or executing compiler work. | Import safety only; no orchestration. |

Each of these paths is operational production code and must be registered as an
exact protected path. Protection is additive: introducing the package must not
weaken the protected manifest, exclude an existing file, or relax the EOL policy.

## Frozen data and failure semantics

The extraction must retain all of the following predecessor details:

- A year-like integer, integer-valued float, or digit-only stripped string column
  is renamed to its plain decimal string. If no column qualifies, the same DataFrame
  object is returned. Normalization collisions remain duplicate column labels.
- System-parameter expansion follows input index order and then inclusive year
  order. Missing values are omitted; duplicate rows and keys are retained; values
  use Python `float` conversion and four-decimal rounding. Missing year/parameter
  columns and invalid values retain native `KeyError`/`ValueError` timing.
- Structure lists preserve source order, values, and duplicates. They are padded to
  the longest set with empty strings; no deduplication or model-driven reordering is
  introduced.
- Demand and workbook frames retain predecessor dtype conversions and missing-value
  behavior. Rounding occurs at the existing writer boundary, not earlier in the
  computational path.
- The demand, capacity-factor, `YearSplit`, and `DaySplit` checks retain their exact
  conditions, misspelled user-visible messages, and bare `sys.exit()` result. The
  known one-daily-time-bracket branch is not repaired.
- Workbook writers are opened before the predecessor conversions, are explicitly
  closed only after successful writes, and are not wrapped in new failure cleanup.
  Sheet mappings retain insertion order.
- Parameter cleanup removes missing mapping keys. The additional/NDP delivery
  mapping is intentionally rebuilt from the cleaned main mapping rather than the
  previously accumulated NDP mapping. This accepted historical quirk must not be
  corrected during isolation.
- Main parameter CSVs retain insertion order. Main set CSVs, and all additional
  parameter/set CSVs, retain their established sorted order. Additional scenario
  replacement changes only the `Scenario` value mapping already changed by the
  predecessor.
- Config backup failure still propagates before B1 enters its restoration scope.
  Config-update failure still skips one scenario; a nonzero compiler return still
  reports and continues; compiler-launch exceptions and interrupts still restore
  then propagate. Restore warnings and partial-state hazards remain as characterized.
- Missing config keys, missing files/sheets/columns, malformed data, writer failures,
  and unexpected exceptions retain their native failure type, message route, exit
  behavior, and partial output timing unless a specific predecessor validation check
  handles them.
- Public entrypoints, callable B1 runner helpers, direct-script imports, stdout/stderr
  messages, generated filenames, current working directory, and inherited
  environment remain unchanged.

Focused characterization must compare the predecessor and candidate traces for
these rules before production edits. The test doubles must prove exact source/path
selection and writer events, including the established fixture's complete ordered
delivery trace, without running A1, A2, A3, B1, B2, the compiler, a matrix command,
or a solver process.

## Scenario and evidence scopes

The three scopes are deliberately different:

| Scope | Required count | Membership rule |
|---|---:|---|
| Scenario registry | 16/16 | BAU plus the frozen accepted decision set in `t1_confection/scenario_registry.json`; only the four roots are prepared by A1/A2. |
| Static cleanup acceptance | 16/16 | Plain `BAU` plus the 15 decision-relevant scenarios. |
| Compiled-input equivalence | 15/15 | The decision scenarios below; plain `BAU` is excluded. |

The compiled-input scenarios are `A_Calibrated_BAU`,
`A_Calibrated_BAU_Clipped`, `B_Optimised_VRE`, `B_Opt_Clipped`,
`B_Opt_DirBidir`, `B_Opt_DirContractual`, `B_Opt_IndiaCosts`,
`B_Opt_IndiaCostsFuel`, `B_Opt_SolarCapex130`, `B_Opt_SolarCapexHi`,
`B_Opt_SolarCapexSpike`, `B_Opt_TradeCap15`, `B_Opt_TxCap150`,
`C_Target_VRE`, and `C_Target_VRE_Clipped`. `A_Calibrated_BAU` is the decision
baseline.

Plain `BAU` remains protected, discovered, and part of static acceptance, but it is
non-decision support evidence. The four superseded definitions --
`B_Opt_LinkFreeze`, `B_Opt_SolarHi10`, `B_Opt_TradeCap30`, and
`B_Opt_TradeCap50` -- remain protected and preservation-visible; their exclusion
from the 16- and 15-scenario scopes does not permit deletion, regeneration, or
promotion.

## A-derived validation lineage

The accepted A-derived artifacts have a historical A-only dependency on the
external scratch helper `relax_activity_band.py` (2,807 bytes; SHA-256
`df861d50076620adc605276d6276d5624f12fb487efa60b5c776da3f843f8185`). That file is
lineage evidence, not production source. It must not be copied into the repository,
imported by the compiler, or executed as part of the normal transformation path.
The tracked
[`apply_base_year_pin.py`](../t1_confection/A3_process/rules_scripts/apply_base_year_pin.py)
is now the static production transformer. Its audited allowlist is mechanically
projected from the frozen corrected evidence, and its embedded digests bind both
that canonical source (`canonical_source_rules.csv` SHA-256
`9c28f9d43c3037daa668554a94061e829d0974662a746efa48d4a2dc341b9ca6`)
and the 1,956-row production projection (`pwr_min_2023_2026_pin.csv` SHA-256
`cdcb0aeb570486b40ab96be68f6db031af54afa3ac02e4832a456522ca73a17c`).
It creates no row or column and fails closed on an invalid rule, asset, scenario,
or workbook match. The former solve-derived `--band 0.002` recipe is historical
lineage only; it is not part of the production interface or derivation.

When changed production paths begin at B1, validation may consume the accepted,
frozen A-derived artifacts and hashes without rerunning A3. If those artifacts or
their frozen recipe cannot be identified exactly, validation must stop rather than
reopen numerical provenance or substitute the scratch helper. Any proposal that
would make the external helper affect normal production behavior belongs on a
separate, explicitly authorized branch.

## Interconnector authority migration

PR #22 established the v18 residual-capacity authority. This follow-up migration
moves the remaining additive minimum/FUTURE contribution and LinkFreeze
minimum-only compatibility boundary into the same production workbook without
changing their effective values, row participation, additive merge, or writer
order. The following unaffected production scripts remain protected:

| File | Frozen SHA-256 |
|---|---|
| [`cap_trn_to_residual.py`](../t1_confection/A3_process/cap_trn_to_residual.py) | `f9f876d1e58cc8dd1339aea703477fe7a85bef776ad227ab99d972d25f7c6a36` |
| [`relax_interconnectors.py`](../t1_confection/A3_process/rules_scripts/relax_interconnectors.py) | `e496d54157459e7da2eb460d0cc76264eeee26a386e6a8c811cad1285424fbb7` |

The following byte-identity requirements apply to that earlier
behavior-preserving migration. The later PWR/MIN correction instead requires the
candidate/current compiled-input difference set to equal exactly its frozen
target set.

This migration must not change TRN costs, operational life,
`ResidualCapacity`, `FUTURE` classification, `TotalAnnualMaxCapacity`,
`TotalAnnualMaxCapacityInvestment`, or `relax_interconnectors` behavior.
Historical citation workbooks are not runtime authorities. The retired NATY
workbook has no production, configuration, or test fallback.

Any candidate compiled-input difference is a blocker here. It must not be
allowlisted or normalized away. Validation on this branch is compile-only and
no-solver; it establishes solver-consumed-input identity and does not claim
solver-backed behavioral or numerical equivalence.

## Required disposable validation evidence

Validation must use the specified OSTRAM interpreter and only clean disposable
checkouts bound to the exact candidate commit. The primary repository and the
read-only accepted reference checkout must not run a generating command. Raw
generated artifacts stay outside Git.

The minimum required evidence is:

1. focused predecessor/candidate transformation tests covering exact transformation
   and writer order; scenario discovery/filter/order; paths and filenames;
   row/key/year/technology selection; dtype, missing-value, and duplicate behavior;
   configuration/environment handling; expected and unexpected failures;
   restoration/cleanup; entrypoint compatibility; and import safety;
2. proof from doubles and process guards that unit tests invoke no real pipeline,
   compiler, B2, matrix, or solver process;
3. the existing A3, B1, B2, and `run.py` orchestration tests plus the full safe
   regression suite;
4. preservation discovery at 20/20, cleanup discovery and acceptance at 16/16, the
   cleanup gate, protected-tree verification, and strict baseline self-comparison;
5. AST parsing of every changed Python file, Markdown/link/path checks for changed
   documentation, EOL verification, and `git diff --check`;
6. an inventory/hash comparison proving that the smallest candidate chain did not
   drift any unintended intermediate artifact or fail to restore configuration;
7. B1 regeneration of the established 15 scenarios followed by B2 compile-only with
   `PYTHONIOENCODING=utf-8`, the four execution/concatenation booleans independently
   confirmed as `False`, fail-closed process barriers, and concurrent monitoring;
8. exactly 15 canonical final files, no missing or extra file, 15/15 byte-exact and
   15/15 normalized-exact comparison with the supplied read-only accepted files,
   zero parameter or set-definition drift, and zero duplicate parameter keys or set
   memberships; and
9. a scan proving that no solver, matrix, solution, result, or solver-log artifact
   was created, followed by clean-status checks of both primary repositories.

The disposable command audit must reject `run.py`, DVC, every batch file, and every
solver. B2 compile-only is permitted only through the inspected guarded path above;
neither `execute_model` nor `create_matrix` may become true. Validation must record
the disposable location, exact candidate commit, accepted-reference identity,
scenario list, config backup/restoration evidence, comparison rule, and process and
artifact monitoring result.

At this checkpoint, evidence not yet completed had to be described as pending. Even a
fully byte-exact 15-scenario result established solver-consumed-input equivalence
only. Solver-backed behavioral and numerical equivalence remained outside this
branch's authorized claim boundary.

## Validation record for `0cc2c68`

This historical record predates and does not validate the late-WS4 static
PWR/MIN pin.

The contract above was exercised at candidate commit
`0cc2c68234df903beadf037082301e2211557e61` in the disposable clean clone
`C:\Users\luisfernando\AppData\Local\Temp\OSTRAM_a1_b1_validate_0cc2c68_20260717`.
The accepted files were read only from
`C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_mainredo\t1_confection\Executables`.
No generating command ran in either primary checkout.

The smallest changed production chain was B1 followed by guarded B2 compile-only.
A3 was not rerun because the changed production path begins at B1 and the accepted
A-derived artifacts already carry the frozen derivation. The external
`relax_activity_band.py` scratch helper was neither copied nor executed.

The B1 invocation used the prescribed interpreter, `PYTHONHASHSEED=0`, and this
explicit 15-scenario filter:

```powershell
& $Py -u .\B1_Run_Compiler.py --scenarios `
  'A_Calibrated_BAU,B_Optimised_VRE,C_Target_VRE,A_Calibrated_BAU_Clipped,B_Opt_Clipped,C_Target_VRE_Clipped,B_Opt_SolarCapexHi,B_Opt_SolarCapex130,B_Opt_SolarCapexSpike,B_Opt_TradeCap15,B_Opt_TxCap150,B_Opt_IndiaCosts,B_Opt_IndiaCostsFuel,B_Opt_DirBidir,B_Opt_DirContractual'
```

B1 completed all 15 scenarios in discovered directory order and restored
`Config_MOMF_T1_A.yaml` to SHA-256
`8b50cfdc2f55a788b69676d9da751ffa910744626ec3d7bf776adc2148d89137`.
The regenerated A1 workbooks were normalized-exact to the pre-run inventory; both
the A2 and otoole CSV trees had empty Git-normalized diffs. The generated structure
workbook was normalized-exact to the accepted structure workbook. All 105 selected
compile intermediates were byte-exact to the accepted checkout.

Before B2, a separate configuration read confirmed that `execute_model`,
`create_matrix`, `concat_otoole_csv`, and `concat_scenarios_csv` were each the
boolean `False`; `parallel` was also `False`, while `A2_otoole_outputs` and
`write_txt_model` remained `True`. `storage_delay_model_output` was redirected to
the disposable `.validation_runtime` directory. B2 ran through its public `main`
path with `PYTHONIOENCODING=utf-8`, an injected `main_executer` barrier, a denied
multiprocessing boundary, a denied matrix runner, and a subprocess filter rejecting
solver executables and matrix/solution file arguments. It executed 91 permitted
otoole/Python/environment subprocess commands and invoked no execution barrier.
`Config_MOMF_T1_AB.yaml` was then restored byte-for-byte to SHA-256
`61e197cf175a63c1445c42fb9f71aea311d9a6db5d1642db9bf6a8feefe5bdbb`.

The resulting canonical comparison was 15 expected, 15 present, 15 byte-exact,
15 normalized-exact, with no missing or extra canonical file. An independent parser
found zero parameter-definition drift, zero set-definition drift, zero duplicate
parameter declarations or keys, and zero duplicate set declarations or memberships.
Monitoring observed no solver process. The final scan found no matrix, solution,
result, or new/modified solver-log artifact; the three existing zero-byte historical
logs were tracked baseline files and remained unchanged.

The safe regression result was 169 passing tests with three optional-dependency
skips. Preservation discovery passed 20/20, cleanup discovery and acceptance passed
16/16, and the cleanup gate, protected-tree verification, and strict baseline
self-comparison all passed. The protected-tree transition is 2,354 files,
266,969,431 bytes, aggregate raw SHA-256
`1b42974e549d181484a161e157dbccd957f666e189fb9cb3a24726e858e5753d`.

This evidence proves exact solver-consumed-input preservation for the authorized
15-scenario scope. It is not a solver-backed result, and it neither implements nor
prepares the parked interconnector correction.
