# Provenance — UNESCAP example profile

Every file under `examples/unescap/` and where it came from, what was changed on the way,
and what was deliberately left behind.

## Source

| | |
|---|---|
| Source repository | `OSTRAM_training_source` (read-only clone; never modified) |
| Source commit | `6e00e8b00144b6859344f54022df34417c075ae9` |
| Target branch | `feat/unescap-example-assets` |
| Base commit | `8636ccccc324dacdd7bb7137fcbfe31d02c5e67d` |
| Files added | 95, all under `examples/unescap/` |

This branch carries **assets only**. The profile engine that reads `profile.yaml` and
resolves `${profile.…}` / `${project.…}` / `${package.…}` / `${workspace.…}` is on the
parallel profile-engine branch. Nothing here is runnable until that branch lands, and the
exercise pages say so on the page.

---

## 1. Copied verbatim

### Model input CSVs — 64 files, byte-identical

| | |
|---|---|
| Source | `t1_confection/OG_csvs_inputs/*.csv` |
| Destination | `inputs/osemosys_global/*.csv` |
| Transformation | none — filenames, rows, columns and values preserved |

The files were copied byte-for-byte and verified byte-identical to the source in the
working tree. One thing does change at commit time: the source CSVs use CRLF line endings,
and the repository's `.gitattributes` sets `* text=auto eol=lf`, so git normalises them to
LF in the index. That is the repository's existing standard — the tracked CSVs in the
project-root `inputs/osemosys_global/` are LF too — and it is applied here rather than
carving out an exception, since `.gitattributes` is outside this branch's ownership
boundary. No filename, row, column, or value is affected.

Verified set sizes:

| Set | Members |
|---|---|
| `TECHNOLOGY` | 89 |
| `FUEL` | 43 |
| `STORAGE` | 4 (`LDSBGDXX01`, `LDSINDEA01`, `SDSBGDXX01`, `SDSINDEA01`) |
| `REGION` | `GLOBAL` (single) |

### Scenario workbook — byte-for-byte

| | |
|---|---|
| Source path | `t1_confection/A3_process/SOASIA_OSeMOSYS_Template_v18.xlsx` |
| Source commit | `6e00e8b00144b6859344f54022df34417c075ae9` |
| Destination | `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx` |
| Source SHA-256 | `f9e3ed1fc720c3c2906cd3ffc18e0023ea18df7f977260580c6d7f37a82019a8` |
| Destination SHA-256 | `f9e3ed1fc720c3c2906cd3ffc18e0023ea18df7f977260580c6d7f37a82019a8` |
| Hashes match | yes |
| Size | 257,429 bytes (identical) |

The workbook was **copied as bytes**. It was never opened and resaved, so all 20 sheets —
including `Restrictions`, `Control`, and `Interconnector_Params` — arrive untouched, along
with every formula, format, and defined name a resave would have altered.

### Scenario YAML authorities — copied, then edited only where noted below

| Scenario | Files | Source |
|---|---|---|
| `A_Calibrated_BAU` | `bau_calibration.yaml`, `lid_rule.yaml`, `relax_interconnectors.yaml`, `retirement_schedule.yaml` | `t1_confection/A3_process/rules_scripts/configs/A_Calibrated_BAU/` |
| `B_Optimised_VRE` | `lid_rule.yaml`, `relax_interconnectors.yaml`, `retirement_schedule.yaml`, `set_interconnector_direction.yaml`, `set_interconnector_direction_EXAMPLE_forward.yaml`, `set_interconnector_direction_EXAMPLE_reverse.yaml`, `storage_floors.yaml` | `t1_confection/A3_process/rules_scripts/configs/B_Optimised_VRE/` |
| `C_Target_VRE` | `lid_rule.yaml`, `relax_interconnectors.yaml`, `retirement_schedule.yaml`, `set_vre_targets.yaml`, `storage_floors.yaml` | `t1_confection/A3_process/rules_scripts/configs/C_Target_VRE/` |

Both interconnector-direction **example** files (forward and reverse) were migrated —
they are active teaching material for Exercise 3.5, not backups.

---

## 2. Transformed

### `config/preparation/Config_country_codes.yaml`

Source: `t1_confection/Config_country_codes.yaml`.

| Change | Before | After |
|---|---|---|
| `country_data` | 6 countries (BGD, BTN, IND, NPL, LKA, MDV) | BGD and IND only |
| `countries` | `["BGD", "INDEA"]` | unchanged — already correct |
| `template_generation` | MDV cloned from **LKA** | MMR cloned from **BGD**, isolated (`interconnections: []`), centerpoint 19.8118 / 96.6022 |
| Commented examples | referenced LKA/MDV/BTN/NPL as reference countries | rewritten to reference only BGD/IND, which are the countries this profile actually models |
| `implausible_combinations` comments | referenced Maldives atolls, Sri Lanka, etc. | rewritten for the two modelled countries |
| Generator instruction | `python Z_generate_country_template.py` | `python -m ostram --profile unescap country template` |

The MDV-from-LKA template was **stale and unrunnable**: Sri Lanka was removed from this
model with the other four countries, so the active configuration named a reference country
that no longer exists. MMR-from-BGD is the example the exercises actually teach
(`exercises/add-country.html`), so the active configuration and the exercise now agree.

### `config/preparation/Config_region_consolidation.yaml`

Source: `t1_confection/Config_region_consolidation.yaml`. `enabled: false` preserved. The
commented BRA/MEX examples were replaced with an India example, since the only country in
this profile that *could* have sub-regions is India. Consolidation stays off deliberately:
merging INDEA into a unified India region would erase `TRNBGDXXINDEA`, the object every
interconnector exercise studies.

### `config/compilation/Config_MOMF_T1_A.yaml`

Source: `t1_confection/Config_MOMF_T1_A.yaml` (the repository-root copy, not the
`A3_process/` working copy).

Only the six directory keys changed, from physical relative paths to logical tokens:

| Key | Before | After |
|---|---|---|
| `A1_inputs` | `./A1_Inputs` | `${profile.osemosys_inputs}` |
| `A1_outputs` | `./A1_Outputs` | `${workspace.preparation}/A1_Outputs` |
| `A2_extra_inputs` | `./A2_Extra_Inputs` | `${workspace.preparation}/extra_inputs` |
| `A2_output` | `./A2_Output_Params/` | `${workspace.compilation}/A2_Output_Params/` |
| `A2_output_main_scen` | `./A2_Output_Params/` | `${workspace.compilation}/A2_Output_Params/` |
| `A2_output_NDP` | `./A2_Output_Params/NDC/` | `${workspace.compilation}/A2_Output_Params/NDC/` |

Everything else is unchanged, including the 20-timeslice fabric (`Timeslices`,
`Conversionls`, `Conversionld`, `Conversionlh`) and the four UNESCAP storage technologies
in `xtra_scen.Storage`.

### `config/execution/Config_MOMF_T1_AB.yaml`

Source: `t1_confection/Config_MOMF_T1_AB.yaml`.

| Change | Detail |
|---|---|
| Directory and file keys | rewritten to `${workspace.*}`, `${package.*}`, `${project.*}` tokens |
| `osemosys_model` | `'osemosys_fast_preprocessed.txt'` → `${project.maintained_model}` |
| `otoole_config`, `conv_format` | → `${package.compilation_resources}/conversion_format.yaml` |
| `templates` | → `${package.compilation_templates}` |
| `reserve_margin_xlsx_workbook` | → `${project.execution_inputs}/firm_capacity_fallbacks_by_cr.xlsx` |
| `storage_delay_model_output` | → `${workspace.execution}/…` — the maintained model is never patched in place |
| `concat_csvs: 'concatenate_ostram.py'` | **dropped** — a legacy runner reference; the project-root execution config had already dropped it |
| `solver` | `'cbc'` — **preserved unchanged** |
| `strip_storage_active` | `True` → **`False`** (see below) |
| Stale Tier-1 example | `strip_storage_targets: ["SDSLKAXX01"]` → `["SDSBGDXX01"]` |
| Script-name comments | `patch_storage_delay.py`, `patch_reserve_margin_repair_careful_xlsx.py` → described by what the stage does |

**On `strip_storage_active`.** The source shipped `strip_storage_active: True` with
`strip_storage_mode: "all"` — a no-storage diagnostic. In the source it was inert, because
`storage_delay_active: True` silently turns stripping off, but as written it contradicts a
profile whose declared scope includes four storage technologies and whose B and C scenarios
put floors on them. It is now explicitly `False`, so the intent is legible rather than
depending on a precedence rule two switches away. `storage_delay_active: True` is kept as
the source had it.

### `config/scenarios/C_Target_VRE/set_vre_targets.yaml`

`bau_results_path` changed from `../Executables/A_Calibrated_BAU_0` to
`${workspace.execution}/Executables/A_Calibrated_BAU_0`. Comment references to
`set_vre_targets.py`, `set_min_capacity_floors.py`, and `TECH_TYPES.csv` were reworded to
name the rule or the migrated filename. **No target values were changed** — the INDEA and
BGDXX solar schedules are the source values.

### `config/scenarios/technology_types.csv`

Source: `t1_confection/A3_process/TECH_TYPES.csv` (386 rows) → 92 rows.

Rows were dropped when the technology name carries a region or country this profile does
not model: the 5-character regions `BTNXX`, `INDNE`, `INDNO`, `INDSO`, `INDWE`, `NPLXX`,
`LKAXX`, `MDVXX`, and the 3-letter `MIN*` country suffixes `BTN`, `LKA`, `MDV`, `NPL`.
294 rows dropped, 92 kept. Nothing else was edited — categories, ordering, and the header
are as in the source.

Result by category: `GENERATION` 32, `PRIMARY_NONRENO` 21, `PRIMARY_RENO` 18,
`INTERCONNECTORS` 1 (`TRNBGDXXINDEA`), `STORAGE_LONG` 2, `STORAGE_SHORT` 2, and 2 rows in
each of the seven `TRANSMISSION_*` categories.

Note that `TECH_TYPES` uses post-PWR-cleanup technology names (`PWRCOABGDXX`), while
`inputs/osemosys_global/TECHNOLOGY.csv` carries the pre-cleanup names (`PWRCOABGDXX01`).
That naming difference is inherited from the source, not introduced here. Three kept
entries — `PWRNGSBGDXX`, `PWRNGSINDEA`, `PWRSHPINDEA` — have no counterpart in the current
89-technology set because they are technologies the AO-extension process adds; they appear
in `ao_extension_decisions.csv` for exactly that reason.

### `config/scenarios/ao_extension_decisions.csv`

Source: `t1_confection/A3_process/OSTRAM_AO_Extensions_FILLED.xlsx`, sheet
`1_Extensions_To_Add` (read only; the workbook was not modified).

Five columns extracted verbatim, 16 data rows:

`AO_Code_To_Add`, `Include`, `Override_Template_AO`, `Override_Tech.Name_AO`, `Notes`

Sheets `2_Parameter_Rows_To_Replicate` and `3_Signal_Disagreements` were **not** copied —
both are generated from sheet 1 and the source data, so carrying them would duplicate
derived state.

This CSV is a **historical decision record**, not a baseline authority. Its rows still name
MDV, NPL, BTN, and India's other regions, because that is what was decided at the time it
was filled in. It is kept unfiltered on purpose: trimming a decision record to today's
scope would misrepresent what was actually decided. Every other config authority in this
profile describes only `BGDXX` and `INDEA`.

---

## 3. Created

| File | What it is |
|---|---|
| `profile.yaml` | `ostram-profile-v1`. Explicit `profile:` / `project:` / `package:` authorities, `workspace:` for generated state, and `resolution.implicit_file_fallback: false`. Declares regions BGDXX + INDEA, interconnector TRNBGDXXINDEA, 20 timeslices, `year_range: {start: 2023, end: 2050, count: 28}`, root scenarios BAU/A/B/C, and `runtime.requires_prepare: true`. |
| `config/scenarios/registry.json` | `ostram-scenario-registry-v1`, four roots. `BAU` is the support scenario; A, B, C are decision scenarios. `C_Target_VRE` declares a `result` dependency on a completed `A_Calibrated_BAU`. `derived_scenarios` is empty — none were added. |
| `README.md` | Profile overview, layout, command surface, path-resolution contract. |
| `references/provenance.md` | This file. |

---

## 4. Migrated documents

| Source | Destination |
|---|---|
| `OSTRAM_Training_Exercises.html` | `exercises/training.html` |
| `OSTRAM_Exercise_A_Add_Country.html` | `exercises/add-country.html` |
| `OSTRAM_Exercise_B_Add_Interconnector.html` | `exercises/add-interconnector.html` |
| `OSTRAM_Git_Setup_Guide.html` | `docs/git-setup.html` |
| `OSTRAM_Interconnector_Direction_Results.html` | `references/interconnector-direction-results.html` |

Commands were rewritten to the planned interface:

| Was | Now |
|---|---|
| `python run.py --scenarios "…"` | `python -m ostram --profile unescap run --scenarios "…"` |
| `python t1_confection/A1_Pre_processing_OG_csvs.py` + `A2_AddTx.py` | `python -m ostram example prepare unescap` |
| `python ostram_training_dashboard.py` / `generate_direction_comparison.py` | `python -m ostram example report unescap` |
| `python t1_confection/Z_generate_country_template.py` | `python -m ostram --profile unescap country template` |
| `python t1_confection/templates/MMR/merge_into_inputs.py` | `python -m ostram --profile unescap country merge MMR` |
| `python t1_confection/Z_validate_country_data.py --country MMR` | `python -m ostram --profile unescap country validate MMR` |
| `python t1_confection/A3_process/populate_v18_new_country.py --country MMR` | `python -m ostram --profile unescap country populate-workbook MMR` |

Removed entirely:

- `git clone --branch training-unescap …` — the model ships as a profile; there is no
  training branch to check out and no branch switching in the exercises.
- Manual deletion of generated folders (`rm -rf …/_post_a2_snapshot_*`, `Executables/*`,
  `_run_*`). Generated state lives in the workspace and the stage that owns it replaces it.
- `git checkout -- <file>` and `git checkout -- .` revert recipes. The exercises now quote
  the shipped values inline so a trainee can type them back, and point at file history
  rather than prescribing destructive version-control commands.
- The manual copy-and-label ritual (`cd t1_confection` + `copy …_<DATE>_<LABEL>.csv`),
  replaced by `example report unescap --label <LABEL>`.
- Instructions to edit pipeline source: the `REGION_NAME_MAP` dictionary in
  `3_update_ao_from_extensions.py` and the `TARGET_TECHS` set in `fix_elc_pmode_revert.py`.
  Both are now derived from the country authority the trainee already edits, so the
  exercises explain the behaviour instead of asking for a source patch.
- OS-specific command variants in `training.html`, `add-country.html`, and
  `add-interconnector.html`. Every remaining command is identical on Windows, macOS, and
  Linux, so the OS pickers were removed with the blocks they switched. `docs/git-setup.html`
  keeps its platform-specific installation instructions, which are genuinely different per OS.

Each page carries a "not runnable from this branch yet" notice stating that the
examples-only branch needs the parallel profile-engine branch before the commands work.

`references/interconnector-direction-results.html` keeps its embedded result figures. They
were produced by earlier full-pipeline runs of `B_Optimised_VRE` and are reproduced as a
reference; the page says so and is not regenerated by this branch.

---

## 5. The 2.496 / 2.5 GW discrepancy — documented, not resolved

The `TRNBGDXXINDEA` residual capacity is recorded twice, with two different values:

| Where | Value |
|---|---|
| `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx`, sheet `Interconnector_Params`, `TRNBGDXXINDEA` / `ResidualCapacity` / `User defined` | **2.496 GW** for 2023–2028, stepping 2.996 (2029) → 3.746 (2030) → 4.496 (2033) |
| `config/scenarios/{A,B,C}/relax_interconnectors.yaml`, `overrides.TRNBGDXXINDEA` | **2.50 GW** at 2023 |
| `exercises/training.html`, Exercise 3.1 | "base 2.5 GW freeze" |
| `exercises/add-interconnector.html`, quick-reference table | 2.496 → 4.496 GW (the workbook value) |

2.496 GW is the sum of the two operational projects — Bheramara HVDC (1,000 MW) and Adani
Power Godda HVDC (1,496 MW). 2.50 is that figure rounded.

**The workbook value is preserved exactly.** The workbook was copied as bytes and not
edited; the YAML anchors were copied as written. Neither side was changed to agree with the
other, and nothing was rounded. Reconciling them is a modelling decision — it changes the
2023 anchor of the interconnector relaxation schedule in all three scenarios — and it is
out of scope for an asset migration. It is recorded here so whoever makes that decision
finds both numbers and the reason they differ.

---

## 6. Deliberately omitted

Nothing below was migrated. Each family is listed with why.

### Generated and runtime state

| Family | Source location |
|---|---|
| A1 outputs, per-scenario | `t1_confection/A1_Outputs/A1_Outputs_{A,B,C}/` |
| Post-A2 snapshots | `t1_confection/A1_Outputs/_post_a2_snapshot_*/` |
| A2 compiled parameters | `t1_confection/A2_Output_Params/`, `A2_Outputs_Params_otoole/` |
| Solver working directories and outputs | `t1_confection/Executables/{A,B,C}_0/` |
| Timestamped run folders | `t1_confection/A3_process/_run_20260525_195646/`, `_run_20260526_133935/`, `_run_20260526_160528/`, `_run_20260526_165441/` |
| Concatenated result tables | `concatenate_files/`, `OSTRAM_*Combined_Inputs_Outputs*.csv` |
| Plots, figures, dashboards | `t1_confection/ostram_plots/`, `trn_plots/`, `Figures/`, `figs_A1_A6/` |
| Model copies | `t1_confection/osemosys_fast_preprocessed.txt`, root `OSTRAM_data.txt` |

All of it is derived output. It belongs in the run workspace, and the migrated exercises
now point there.

### Executable code

| Family | Why |
|---|---|
| Legacy runner `run.py` | Replaced by `python -m ostram`; the branch adds no executable Python under `examples/unescap/` |
| Stage scripts `A0`–`A3`, `B1`, `B2`, and the `A3_process/` numbered stages | Engine code; belongs to the package, not to an example profile |
| Rule scripts under `A3_process/rules_scripts/*.py` | Same — the profile ships the rule **configuration**, not the rule implementation |
| `Z_*` and `Z_AUX_*` helpers, `patch_*`, `fix_*`, `check_*`, `concat_*`, `reduce_*`, `slice_*` | Ad-hoc tooling and one-off patches |
| Dashboard and analysis scripts (`ostram_training_dashboard.py`, `ostram_scenario_analysis.py`, `ostram_trn_plotter.py`, `Z_AUX_generate_*`) | Reporting is now `example report unescap` |
| `t1_confection/tests/`, `_test_scenarios_lite.py`, `test_strip_storage.py` | Engine tests; out of the ownership boundary for this branch |

### Deprecated and duplicate configuration

| File | Why |
|---|---|
| `configs/A_Calibrated_BAU/deprecate/retirement_schedule.yaml` | Deprecated folder |
| `configs/B_Optimised_VRE/deprecate/retirement_schedule.yaml` | Deprecated folder |
| `configs/A_Calibrated_BAU/retirement_schedule_-_bau_v2.yaml` | Byte-identical duplicate of `retirement_schedule.yaml` (verified by hash) |
| `configs/B_Optimised_VRE/retirement_schedule_-_opti_v2.yaml` | Byte-identical duplicate of `retirement_schedule.yaml` (verified by hash) |
| `_run_*/Config_MOMF_T1_A.yaml`, `_run_*/rules_scripts/*.yaml` | Snapshots of config inside generated run folders |
| `t1_confection/A3_process/Config_MOMF_T1_A.yaml` | Working copy; the repository-root copy is the authority |
| `t1_confection/Config_tech_equivalences.yaml` | Not referenced by any migrated authority |
| `t1_confection/Miscellaneous/conversion_format.yaml` | otoole config — resolved from `${package.compilation_resources}` |

### Other source material not carried over

| Family | Why |
|---|---|
| `SOASIA_OSeMOSYS_Template_v17.xlsx` | Superseded by v18, which is what was migrated |
| `A-O_Parametrization_*.xlsx`, `OSTRAM_Timeslice_Outputs.xlsx` | Generated A-O workbooks and timeslice output |
| Data workbooks (`CapacityAndDistances.xlsx`, `Shares_*.xlsx`, `RateGrowthDemand_*.xlsx`, `REV_FILTER_ISSUE.xlsx`, `Tech_Country_Matrix.xlsx`, `Interconnectors.xlsx`, `firm_capacity_fallbacks_by_cr.xlsx`) | Preparation-stage source data. `firm_capacity_fallbacks_by_cr.xlsx` already exists at the project root and is referenced through `${project.execution_inputs}` |
| Environment and pipeline definitions (`environment.yaml`, `dvc.yaml`, `dvc.lock`, `.dvcignore`, `.readthedocs.yaml`) | Repository-level; outside the ownership boundary |
| Root documentation (`README.md`, `TECHNICAL_INVENTORY.md`, `docs/`) | Repository-level; outside the ownership boundary |
| `t1_confection/A3_process/{README,USER_GUIDE,LID_RULE}.md` | Documentation for the legacy stage scripts, not for the profile |

---

## 7. Validation performed

| Check | Result |
|---|---|
| Every changed path under `examples/unescap/**` | 95 files, all inside |
| No `.py` under `examples/unescap/` | 0 found |
| CSV count in the profile input directory | 64 |
| CSV filenames and bytes vs source | identical in the working tree (LF-normalised at commit by repository `.gitattributes`) |
| `TECHNOLOGY` / `FUEL` / `STORAGE` / `REGION` | 89 / 43 / 4 / `GLOBAL` |
| Workbook source vs destination SHA-256 | match |
| Every YAML parses | 21/21 |
| Every JSON parses | 1/1 |
| Registry schema, four roots, C→A result dependency, no derived scenarios | as specified |
| `profile.yaml` contract (schema, id, three authorities, no implicit fallback, regions, interconnector, 20 timeslices, `year_range` triple, roots, `requires_prepare`) | as specified |
| No `python run.py`, `run.py`, `t1_confection/`, `training-unescap`, `git checkout --`, or direct legacy script execution in any migrated text | none found |
| No stale region codes in any config authority | none (the AO decision record is excluded by design — see §2) |
| No physical legacy paths in any config authority | none |
| HTML internal links and fragments resolve | 47/47 |
| 2.496 GW preserved verbatim | yes |

Solvers were not run and the full repository test suite was not executed.
