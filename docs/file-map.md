# File map

Each folder and key file of the repository, what it holds and which stage
reads it. Stage names and other terms are in the [glossary](glossary.md).

## Top level

| Path | What it holds | Read by |
|---|---|---|
| `README.md` | The front page. | People. Tests check that it names the commands. |
| `AGENTS.md` | Instructions for AI coding agents. | Agents. |
| `CONTRIBUTING.md` | How to make a change. | People. |
| `LICENSE` | The Apache License, version 2.0. | People. |
| `environment.yaml` | The conda environment definition. | `python -m ostram run` when it checks or creates an environment. |
| `pyproject.toml` | Package metadata and Python dependencies. | `pip install -e .` |
| `dvc.yaml`, `.dvc/`, `.dvcignore` | DVC stage definitions and local DVC state. | `python -m ostram run` without `--skip-pull`. |
| `.gitattributes` | Line-ending rules. It keeps `pwr_min_2023_2026_pin.csv` with CRLF endings. | Git. |
| `.readthedocs.yaml` | Read the Docs build settings. | Read the Docs. |
| `.github/workflows/unescap-macos-acceptance.yml` | The CI workflow that runs the training profile on macOS. | GitHub Actions. |
| `config/` | Configuration. See the table below. | A1 to B2. |
| `docs/` | This documentation. | Sphinx. |
| `examples/unescap/` | The two-region training profile with its own inputs, configuration and registry. | `python -m ostram --profile unescap`. |
| `inputs/` | Model inputs. See the table below. | A1, A3 and B2. |
| `model/osemosys_fast_preprocessed.txt` | The OSeMOSYS model file in GMPL. | B2. The storage delay writes an adjusted copy, and `glpsol` reads that copy. |
| `ostram/` | The Python package. See the table below. | Every command. |
| `tests/` | The solver-free tests in `tests/regression/` and `tests/validation/`. | `python -m unittest discover -s tests -p 'test_*.py'`. |
| `timeslice_module/` | The timeslice workbook generator, its inputs and the reference workbook of the adopted fabric. The pipeline does not read it. | People, by hand. See its README. |
| `tools/generate_scenario_catalogue.py` | Regenerates the scenario catalogue page. | `tests/regression/test_scenario_catalogue.py`. |
| `workspace/` | Generated files. Git ignores it. | Every stage writes here. |

## config

| Path | What it holds | Read by |
|---|---|---|
| `config/profiles/registry.json` | The profile names and their profile files. | The command line option `--profile`. |
| `config/profiles/full.yaml` | The South Asia profile: policies, sources of record, the scenario registry, the shared VRE ceilings and the scenario edits. See the [section map](full-yaml.md). | A3, B2 and the tests. |
| `config/preparation/Config_country_codes.yaml` | Countries, regions, technology mappings, transmission defaults and the internal transmission settings. | A1, A2 and the A3 rules for internal transmission. |
| `config/preparation/Config_region_consolidation.yaml` | An optional merge of sub-regions into one region. Disabled. | A1. |
| `config/compilation/Config_MOMF_T1_A.yaml` | Compiler settings: years, timeslices, the storage set, workbook names and parameter lists. | A1 writes a runtime copy. A3, B1 and B2 read it. |
| `config/execution/Config_MOMF_T1_AB.yaml` | The solver choice and the settings of the model-file adjustments. | B2. |
| `config/scenarios/technology_types.csv` | The technology taxonomy. | A3. |
| `config/scenarios/rules/internal_tx_residuals.csv` | The residual capacity of the internal transmission families per country-region. | The A3 rule `apply_internal_transmission`. |
| `config/scenarios/rules/pwr_min_2023_2026_pin.csv` | The reviewed 2023 to 2026 values written by the base-year pin. | The A3 rule `apply_base_year_pin`. |
| `config/scenarios/A_Calibrated_BAU/` | `retirement_schedule.yaml`, `lid_rule.yaml`, `bau_calibration.yaml` and `relax_interconnectors.yaml`. | The A3 rules of `A_Calibrated_BAU`. |
| `config/scenarios/B_Optimised_VRE/` | `retirement_schedule.yaml`, `lid_rule.yaml`, `relax_interconnectors.yaml` and `storage_floors.yaml`. | The A3 rules of `B_Optimised_VRE`. |
| `config/scenarios/C_Target_VRE/` | The four files above plus `bau_calibration.yaml` and `set_vre_targets.yaml`. | The A3 rules of `C_Target_VRE`. |
| `config/scenarios/B_Opt_DirBidir/set_interconnector_direction.yaml` | The overlay of `B_Opt_DirBidir`. It lists no corridor. | A3, after the scenario edits. |
| `config/scenarios/B_Opt_DirContractual/set_interconnector_direction.yaml` | The overlay of `B_Opt_DirContractual`. | A3, after the scenario edits. |

`BAU` has no folder. Its only rule, the build-cap rule, runs on its built-in
defaults. The test `tests/regression/test_config_files.py` lists every file
under `config/` and fails when a file is added, removed or left unregistered.

## inputs

| Path | What it holds | Read by |
|---|---|---|
| `inputs/osemosys_global/*.csv` | The 64 OSeMOSYS Global set and parameter files that form the base model. | A1, which stages them under `workspace/preparation/og_csvs_inputs/`. |
| `inputs/preparation/centerpoints.csv` | Coordinates of the country centrepoints. | The country template generator. |
| `inputs/preparation/secondary_technologies/Tech_Country_Matrix.xlsx` | Which technology exists in which country, the NGS unification switches and the aggregation rules. | A1. |
| `inputs/preparation/secondary_technologies/` (the four other workbooks) | Installed capacity, generation and technology shares per country. | The secondary-technology editor. A1 reads the installed-capacity workbook only when `add_missing_countries_from_ostram` is true. |
| `inputs/preparation/workbook_templates/` | The templates of the four A-O workbooks and of `A-Xtra_Emissions.xlsx` and `A-Xtra_Storage.xlsx`. | A1. |
| `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx` | The scenario workbook: the Control sheet, one sheet per parameter group, the interconnector parameters and the extension decisions. | A3. |
| `inputs/scenarios/OSTRAM_Timeslice_Inputs.xlsx` | Year splits, demand profiles and capacity factors for the 20 timeslices. | A3. |
| `inputs/execution/firm_capacity_fallbacks_by_cr.xlsx` | The fallback caps of the gas and oil allowance per country-region. | B2. |

## The ostram package

| Path | What it holds |
|---|---|
| `ostram/__main__.py` | The command line: `run`, `transform`, `compile-inputs`, `inspect-resources`, `example`, `country` and `scenario`. |
| `ostram/paths.py` | Path resolution: project root, workspace and the sources of record. |
| `ostram/profiles.py`, `ostram/profile_workspace.py` | Profile loading and profile workspaces. |
| `ostram/examples.py` | The training example commands `example prepare` and `example report`. |
| `ostram/terminal.py` | The run display and the run log. |
| `ostram/validation/profile.py` | Checks of a profile's compiled domain. |
| `ostram/reporting/` | The training dashboard, the direction comparison and the summary workbook. |
| `ostram/pipeline/orchestration.py` | `python -m ostram run`: the environment check, DVC, and the stages A1 to B2 in order. |
| `ostram/pipeline/preparation/` | A1 in `base_inputs.py`, A2 in `transmission.py`, the country template and validation commands and the secondary-technology helpers. |
| `ostram/pipeline/scenarios/` | A3: `materializer.py` builds the roots and then the derived scenarios, `transform.py` runs the root stages, `registry.py` validates the scenario registry, `scenario_edits.py` applies the scenario edits, `transformations/` holds the stage scripts and `rules/` holds the 11 rule modules. |
| `ostram/pipeline/compilation/` | B1: `runner.py`, `compiler.py` and `transforms/`. |
| `ostram/pipeline/execution/` | B2: `runner.py` normalises the CSVs, converts them with otoole and applies the model-file adjustments; `orchestrator.py` writes the matrix, runs the solver and collects results; `preprocess.py`, `model_file_adjustments/`, `gross_import_cap.py`, `annualization.py` and the concatenation modules. |
| `ostram/resources/compilation/conversion_format.yaml` | The otoole conversion schema. It gives -1 to capacity caps that are absent from the datafile. |
| `ostram/resources/compilation/templates/*.csv` | 64 empty CSV templates for the parameters that a compiled scenario lacks. |
| `ostram/resources/preparation/country_templates/MDV/` | The 28 CSVs of the Maldives country template. |

## Generated files

| Path under the workspace | What it holds | Written by |
|---|---|---|
| `preparation/og_csvs_inputs/` | The staged base CSVs. | A1 |
| `preparation/extra_inputs/` | `A-Xtra_Emissions.xlsx` and `A-Xtra_Storage.xlsx`. | A1 |
| `preparation/A1_Outputs/A1_Outputs_<scenario>/` | The A-O workbooks of each root and each derived scenario. | A1, A2, A3 |
| `preparation/A1_Outputs/_post_a2_snapshot_<root>/` | The copy of each root after A2. A3 starts from it. | A2 |
| `preparation/A1_Outputs/_scenario_materialization.json` | The record of the last A3 run. | A3 |
| `scenarios/_run_<stamp>/` | The working folders of one root's A3 run. Removed at the end unless `--keep-workdir` is set. | A3 |
| `compilation/A2_Output_Params/<scenario>/` | The compiled CSVs. | B1 |
| `execution/A2_Outputs_Params_otoole/<scenario>/` | The normalised CSVs. | B2 |
| `execution/Executables/<scenario>_0/` | The datafiles, the adjusted model file, the LP, the solution and `Outputs/`. | B2 |
| `logs/<run-id>/run.log` | The log of one `run`. | `python -m ostram run` |

## Abbreviations

| Abbreviation | Meaning |
|---|---|
| `cr` | country-region |
| `A-O`, `AO` | phase A outputs, the name of the A1 workbooks. It is the prefix of the four scenario workbooks and of the extension decisions. |
| `AR` | activity ratio, as in `A-O_AR_Projections.xlsx` |
| `OG`, `og` | OSeMOSYS Global, as in `inputs/osemosys_global/` and `og_csvs_inputs/` |
| `c2a` | `CapacityToActivityUnit`, as in `add_ao_c2a.py` |
| `ts20` | the 20 timeslices, as in `sync_og_to_ts20.py` |
| `WV` | working version, as in the merged workbook `SOASIA_OSeMOSYS_WV.xlsx` |
| `SOASIA` | South Asia, as in the merged workbook `SOASIA_OSeMOSYS_WV.xlsx` |
| `MOMF` | Multipurpose OSeMOSYS Modelling Framework, as in `Config_MOMF_T1_A.yaml` and `Config_MOMF_T1_AB.yaml` |
| `T1` | tier 1, in the same two file names |
| `A`, `AB` | pipeline phases, in the same two file names |
| `D4` | The prefix of the added sets and constraints in the model file. |
| `v18` | Version 18 of the scenario workbook template. |
| `SDS`, `LDS` | short-duration and long-duration storage |
| `BCK` | backstop |
| `NLI`, `RPO` | new lines and repowered lines, in the transmission families |
| `DSPTRN` | The dispatch node that routes power to and from the interconnectors. |
| `CF` | capacity factor |
| `StorageDelayN5`, `OpenBCK`, `RMCarefulXLSX` | The suffixes that the storage delay, the backstop opening and the gas and oil allowance add to the datafile name. |
| `otoole` | The external package that converts CSVs to the OSeMOSYS datafile. |

Technology and fuel codes such as `PWR`, `TRN`, `ELC` and `MIN` are in the
[data reference](data-reference.md).
