# Latin America and the Caribbean (LAC) area

The `lac` setup models the power systems of 19 Latin American and Caribbean
countries over 2023 to 2050 in twelve time slices, joined by 26 cross-border
interconnectors. Its support scenario `BAU` is the `BAU` model of
[OSTRAM-LAC](https://github.com/Climate-Lead-Group/OSTRAM-LAC), carried into
OSTRAM so that its solver datafile is the same. The decision scenarios
`A_Calibrated_BAU`, `B_Optimised_VRE` and `C_Target_VRE` follow the shape of
the South Asia scenarios of the same letter and the direction of the RELAC
transmission study of the Inter-American Development Bank.

Everything the setup needs is in this folder. Nothing generated lives here: run
outputs go to the workspace given with `--workspace`.

## The model

| | |
|---|---|
| Countries | ARG, BOL, BRA, BRB, CHL, COL, CRI, DOM, ECU, GTM, HND, HTI, MEX, NIC, PAN, PER, PRY, SLV, URY (one node each, `<ISO3>XX`) |
| OSeMOSYS `REGION` | `GLOBAL` (single region) |
| Interconnectors | 26 (`TRN<node><node>`) |
| Years | 2023 to 2050 (28) |
| Time slices | 12: 4 seasons x 1 day type x 3 daily brackets |
| Storage | 40: long duration (`LDS`), short duration (`SDS`) and batteries (`BDS`, Chile and Peru) |
| Compiled sets | 852 technologies, 406 fuels, 172 emissions |
| Solver | CPLEX |

## Layout

```
areas/lac/
  setup.yaml                     authorities, policies, scenario registry
  README.md                      this file
  data_sources.md                where every value comes from
  config/
    compilation/Config_MOMF_T1_A.yaml   time slices, storage set, B1 settings
    execution/Config_MOMF_T1_AB.yaml    solver and B2 adjustments
    preparation/                        country and region configuration
    scenarios/
      technology_types.csv             technology categories
      A_Calibrated_BAU/  B_Optimised_VRE/  C_Target_VRE/   rule settings
  inputs/
    osemosys_global/*.csv          OSeMOSYS Global inputs (64 files)
    scenarios/
      OSTRAM_Scenario_Inputs.xlsx  Control sheet: roots and their rule scripts
      root_workbooks/              the four A-O workbooks every root starts from
    preparation/extra_inputs/      A-Xtra_Emissions.xlsx, A-Xtra_Storage.xlsx
    execution/                     reserve-margin fallbacks (read only when that step is on)
  model/osemosys_fast_preprocessed.txt   the OSeMOSYS model file
```

## How it differs from South Asia

- **Ready-made root workbooks.** The policy `prepared_root_workbooks` makes
  preparation copy `inputs/scenarios/root_workbooks` into every root instead of
  building workbooks from the OSeMOSYS Global CSVs (A1 and A2), and makes A3
  apply only each root's rule scripts. The workbooks are the OSTRAM-LAC `BAU`
  workbooks: the documentation sheet `Fuentes` is removed, and the year cells of
  the eight parameters OSTRAM compiles at full precision are rounded to four
  decimals as OSTRAM-LAC compiles them.
- **Pre-solver values carried as inputs.** OSTRAM-LAC changes some `BAU` solver
  inputs after compilation: a reserve-margin repair, fossil dispatch floors and
  internal transmission costs and limits. Their final values are in the root
  workbook (see the last section of `data_sources.md`).
- **Execution settings.** No base-year build lock; the storage delay rewrites
  only the build cap of storage technologies (`storage_delay_cap_params`); the
  datafile is written with six significant digits (`datafile_float_format`), as
  OSTRAM-LAC writes it; OSTRAM's reserve-margin repair is off.
- **Internal losses.** Internal transmission losses are input activity ratios
  above 1 in the activity-ratio workbooks, per country (for example Argentina
  1.042 for renewable lines and 1.012 for other lines). The South Asia loss rule
  does not run for this setup.

## Scenarios

| Scenario | Rule scripts | Settings |
|---|---|---|
| `BAU` (support) | none | The OSTRAM-LAC `BAU` model. Internal repowering is open from 2028 and unplanned internal lines from 2030. |
| `A_Calibrated_BAU` | `build_cap_rule.py` | `BAU` plus a build lock: generation technologies with residual capacity or a planned build may build only their planned capacity (cap multiplier 0, so each cap is the planned build times 1.01). The study's Tendencial scenario (section 2.3.2). |
| `B_Optimised_VRE` | `build_cap_rule.py` | `BAU` with the South Asia B build-cap relaxation and family ceilings on the LAC families; B also takes the looser of its bounds and A's. Least cost without targets. |
| `C_Target_VRE` | `build_cap_rule.py`, `set_vre_targets.py` | The A settings plus each country's renewable share target from its NDC (`set_vre_targets.yaml`), as generation floors scaled to the solved A. The study's NDC scenario (section 2.3.3). |

### Renewable targets in `C_Target_VRE`

| Country | Target | Document |
|---|---|---|
| Chile | 80% of generation by 2030 | NDC 2025 |
| Bolivia | 75% of generation in the interconnected system by 2035 | NDC 3.0 |
| Honduras | 70% of generation by 2030 | NDC 3.0 |
| Costa Rica | 95% of generation from 2027 | NDC 2025-2035 |
| Barbados | 95% of electricity by 2035 | Second NDC |
| El Salvador | 85.7% of generation by 2030 | updated NDC 2021 |

The URLs are in `config/scenarios/C_Target_VRE/set_vre_targets.yaml`. Each
target is met through solar PV above the country's 2023 renewable share. No
target is set for: Brazil (its 23% non-hydro target for 2030 is already met in
2023), Mexico and Guatemala (targets for "clean" rather than renewable
electricity), Nicaragua and Panama (installed-capacity shares), Argentina and
the Dominican Republic (legal targets for 2025, before the target years of the
model), and Colombia, Ecuador, Haiti, Paraguay, Peru and Uruguay (no
renewable share target found in their NDCs).

The transmission windows of the study (repowering from 2027, new lines from
2030) are carried by `BAU` and kept in all three; repowering opens in 2028 in the
ported model. Each rule file states its values and the ASSUMPTIONs behind them.

## Commands

From the repository root, with the `OSTRAM-env` environment and
`PYTHONUTF8=1`:

```
python -m ostram --workspace workspace/lac --setup lac run --scenarios "BAU" --compile-only --skip-pull
python -m ostram --workspace workspace/lac --setup lac run --scenarios "A_Calibrated_BAU,B_Optimised_VRE" --compile-only --skip-pull
python -m ostram --workspace workspace/lac --setup lac run --scenarios "A_Calibrated_BAU" --skip-pull
python -m ostram --workspace workspace/lac --setup lac run --scenarios "B_Optimised_VRE" --skip-pull
python -m ostram --workspace workspace/lac --setup lac run --scenarios "C_Target_VRE" --skip-pull --a-result-seed <workspace>/setups/lac/execution/Executables/A_Calibrated_BAU_0
```

`C_Target_VRE` needs a solved `A_Calibrated_BAU`. Preparation seeds the roots
only when their snapshots are missing: after changing a root workbook, start
from a new workspace.

## Sources

`data_sources.md` lists the source of every value group, in the format of the
South Asia base-year sources, with the ASSUMPTIONs used across the area. The
rule files of `A_Calibrated_BAU`, `B_Optimised_VRE` and `C_Target_VRE` state the
sources of their own values.
