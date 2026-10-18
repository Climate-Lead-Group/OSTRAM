# Scenario catalogue

The block between the two markers is generated from `config/setups/south_asia.yaml`,
the scenario files `config/scenarios/<name>/scenario.yaml` and
`config/execution/Config_MOMF_T1_AB.yaml` by
`python tools/generate_scenario_catalogue.py`, run from the repository root.
The block documents the delivered scenarios, the rows of the table "Origin of
each scenario" at the end of this page, in the order of the registry. A scenario
registered later is left out until its row is added to the table. The generator
is safe to run at any time. The test
`tests/regression/test_scenario_catalogue.py` fails when the block or the
origin table is out of date. Terms are in the [glossary](../../glossary.md).

## Facts that apply across scenarios

- Every registered derived scenario applies the shared VRE ceilings, including
  the scenarios without "Clipped" in their name. `B_Opt_TxCap150` is therefore
  `B_Opt_Clipped` plus the transmission cap, and the scenario to compare it
  with is `B_Opt_Clipped`.
- The ceilings set flat values. The pipeline never takes the lower of the
  previous value and the ceiling; a ceiling above the previous value raises
  it. The [scenario edits](../../scenario-edits.md) page describes the step.
- `B_Opt_TradeCap15` and `B_Opt_Clipped_GrossImportCap15` are different
  limits. `B_Opt_TradeCap15` sets fixed activity upper limits on four
  Bangladesh import corridors and zero activity on Bangladesh's new and
  repowered internal lines from 2027, through scenario edits. `B_Opt_Clipped_GrossImportCap15` caps the gross received
  imports of each of the six countries at 15 per cent of its annual demand in
  2027 to 2050, through a constraint that B2 builds from the compiled demand
  and output ratios.
- `C_Target_VRE_Clipped` restates the Maldives solar assumption for 2028 that
  `C_Target_VRE` already carries from the scenario workbook.

## Delivered scenarios

<!-- generated: start -->
### BAU

- Parent: none. `BAU` is a support root.
- B2 policy: none.
- Seed needs: none.

### A_Calibrated_BAU

- Parent: none. `A_Calibrated_BAU` is a decision root.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.

### A_Calibrated_BAU_Clipped

- Parent: `A_Calibrated_BAU`.
- Shared VRE ceilings: applied.
- Scenario edits: 0.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to A_Calibrated_BAU. No other changes.

### B_Optimised_VRE

- Parent: none. `B_Optimised_VRE` is a decision root.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.

### B_Opt_Clipped

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 0.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. No other changes.

### B_Opt_DirBidir

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 0.
- Overlay: `config/scenarios/B_Opt_DirBidir/set_interconnector_direction.yaml` from 2027.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. Every interconnector may carry power in both directions; its direction file lists no one-way corridors.

### B_Opt_DirContractual

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 0.
- Overlay: `config/scenarios/B_Opt_DirContractual/set_interconnector_direction.yaml` from 2027.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. Nine interconnectors carry power in one direction only, the direction of their contracts and physical flows. The two Nepal to India interconnectors carry power both ways. Its direction file sets the direction of each interconnector.

### B_Opt_IndiaCosts

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 32.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. The CapitalCost and FixedCost of the existing generation and storage technologies of Bangladesh, Bhutan, Nepal, Sri Lanka and the Maldives, except the backstop, are set to the India values of the same family in each year. Where the India regions differ (coal, hydro and offshore wind), the values of the INDNO region are used. Transmission and fuel prices are unchanged.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `CapitalCost` | copy_from_tech | PWRBIOBGDXX, PWRBIOLKAXX, PWRBIONPLXX, PWRCCSBGDXX, PWRCOABGDXX, PWRCOALKAXX, PWRCOGBGDXX, PWRCOGLKAXX, PWRHYDBGDXX, PWRHYDBTNXX, PWRHYDLKAXX, PWRHYDNPLXX, PWRLDSBGDXX, PWRLDSBTNXX, PWRLDSLKAXX, PWRLDSMDVXX, PWRLDSNPLXX, PWRNGSBGDXX, PWRNGSLKAXX, PWRNGSMDVXX, PWROILBGDXX, PWROILLKAXX, PWROILMDVXX, PWROILNPLXX, PWROTHBGDXX, PWROTHBTNXX, PWROTHLKAXX, PWROTHMDVXX, PWROTHNPLXX, PWRPETBGDXX, PWRPETBTNXX, PWRPETLKAXX, PWRPETMDVXX, PWRPETNPLXX, PWRSDSBGDXX, PWRSDSBTNXX, PWRSDSLKAXX, PWRSDSMDVXX, PWRSDSNPLXX, PWRSPVBGDXX, PWRSPVBTNXX, PWRSPVLKAXX, PWRSPVMDVXX, PWRSPVNPLXX, PWRURNBGDXX, PWRWASBGDXX, PWRWASLKAXX, PWRWASNPLXX, PWRWOFBGDXX, PWRWOFLKAXX, PWRWOFMDVXX, PWRWONBGDXX, PWRWONBTNXX, PWRWONLKAXX, PWRWONMDVXX, PWRWONNPLXX |
| Secondary Techs | `FixedCost` | copy_from_tech | PWRBIOBGDXX, PWRBIOLKAXX, PWRBIONPLXX, PWRCCSBGDXX, PWRCOABGDXX, PWRCOALKAXX, PWRCOGBGDXX, PWRCOGLKAXX, PWRHYDBGDXX, PWRHYDBTNXX, PWRHYDLKAXX, PWRHYDNPLXX, PWRLDSBGDXX, PWRLDSBTNXX, PWRLDSLKAXX, PWRLDSMDVXX, PWRLDSNPLXX, PWRNGSBGDXX, PWRNGSLKAXX, PWRNGSMDVXX, PWROILBGDXX, PWROILLKAXX, PWROILMDVXX, PWROILNPLXX, PWROTHBGDXX, PWROTHBTNXX, PWROTHLKAXX, PWROTHMDVXX, PWROTHNPLXX, PWRPETBGDXX, PWRPETBTNXX, PWRPETLKAXX, PWRPETMDVXX, PWRPETNPLXX, PWRSDSBGDXX, PWRSDSBTNXX, PWRSDSLKAXX, PWRSDSMDVXX, PWRSDSNPLXX, PWRSPVBGDXX, PWRSPVBTNXX, PWRSPVLKAXX, PWRSPVMDVXX, PWRSPVNPLXX, PWRURNBGDXX, PWRWASBGDXX, PWRWASLKAXX, PWRWASNPLXX, PWRWOFBGDXX, PWRWOFLKAXX, PWRWOFMDVXX, PWRWONBGDXX, PWRWONBTNXX, PWRWONLKAXX, PWRWONMDVXX, PWRWONNPLXX |

### B_Opt_IndiaCostsFuel

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 7, after the 32 edits of `B_Opt_IndiaCosts`.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE, then the edits of B_Opt_IndiaCosts: the CapitalCost and FixedCost of the existing generation and storage technologies of Bangladesh, Bhutan, Nepal, Sri Lanka and the Maldives, except the backstop, are set to the India values of the same family in each year. Where the India regions differ (coal, hydro and offshore wind), the values of the INDNO region are used. Then the fuel prices of these five countries (the VariableCost of their MIN fuel-supply technologies) are set to the India prices. Transmission is unchanged.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `CapitalCost` | copy_from_tech | PWRBIOBGDXX, PWRBIOLKAXX, PWRBIONPLXX, PWRCCSBGDXX, PWRCOABGDXX, PWRCOALKAXX, PWRCOGBGDXX, PWRCOGLKAXX, PWRHYDBGDXX, PWRHYDBTNXX, PWRHYDLKAXX, PWRHYDNPLXX, PWRLDSBGDXX, PWRLDSBTNXX, PWRLDSLKAXX, PWRLDSMDVXX, PWRLDSNPLXX, PWRNGSBGDXX, PWRNGSLKAXX, PWRNGSMDVXX, PWROILBGDXX, PWROILLKAXX, PWROILMDVXX, PWROILNPLXX, PWROTHBGDXX, PWROTHBTNXX, PWROTHLKAXX, PWROTHMDVXX, PWROTHNPLXX, PWRPETBGDXX, PWRPETBTNXX, PWRPETLKAXX, PWRPETMDVXX, PWRPETNPLXX, PWRSDSBGDXX, PWRSDSBTNXX, PWRSDSLKAXX, PWRSDSMDVXX, PWRSDSNPLXX, PWRSPVBGDXX, PWRSPVBTNXX, PWRSPVLKAXX, PWRSPVMDVXX, PWRSPVNPLXX, PWRURNBGDXX, PWRWASBGDXX, PWRWASLKAXX, PWRWASNPLXX, PWRWOFBGDXX, PWRWOFLKAXX, PWRWOFMDVXX, PWRWONBGDXX, PWRWONBTNXX, PWRWONLKAXX, PWRWONMDVXX, PWRWONNPLXX |
| Secondary Techs | `FixedCost` | copy_from_tech | PWRBIOBGDXX, PWRBIOLKAXX, PWRBIONPLXX, PWRCCSBGDXX, PWRCOABGDXX, PWRCOALKAXX, PWRCOGBGDXX, PWRCOGLKAXX, PWRHYDBGDXX, PWRHYDBTNXX, PWRHYDLKAXX, PWRHYDNPLXX, PWRLDSBGDXX, PWRLDSBTNXX, PWRLDSLKAXX, PWRLDSMDVXX, PWRLDSNPLXX, PWRNGSBGDXX, PWRNGSLKAXX, PWRNGSMDVXX, PWROILBGDXX, PWROILLKAXX, PWROILMDVXX, PWROILNPLXX, PWROTHBGDXX, PWROTHBTNXX, PWROTHLKAXX, PWROTHMDVXX, PWROTHNPLXX, PWRPETBGDXX, PWRPETBTNXX, PWRPETLKAXX, PWRPETMDVXX, PWRPETNPLXX, PWRSDSBGDXX, PWRSDSBTNXX, PWRSDSLKAXX, PWRSDSMDVXX, PWRSDSNPLXX, PWRSPVBGDXX, PWRSPVBTNXX, PWRSPVLKAXX, PWRSPVMDVXX, PWRSPVNPLXX, PWRURNBGDXX, PWRWASBGDXX, PWRWASLKAXX, PWRWASNPLXX, PWRWOFBGDXX, PWRWOFLKAXX, PWRWOFMDVXX, PWRWONBGDXX, PWRWONBTNXX, PWRWONLKAXX, PWRWONMDVXX, PWRWONNPLXX |
| VariableCost | `VariableCost` | copy_from_tech | MINCOABGD, MINCOABTN, MINCOALKA, MINCOAMDV, MINCOANPL, MINCOGBGD, MINCOGBTN, MINCOGLKA, MINCOGMDV, MINCOGNPL, MINGASBGD, MINGASBTN, MINGASLKA, MINGASMDV, MINGASNPL, MINOILBGD, MINOILBTN, MINOILLKA, MINOILMDV, MINOILNPL, MINOTHBGD, MINOTHBTN, MINOTHLKA, MINOTHMDV, MINOTHNPL, MINPETBGD, MINPETBTN, MINPETLKA, MINPETMDV, MINPETNPL, MINURNBGD, MINURNBTN, MINURNLKA, MINURNMDV, MINURNNPL |

### B_Opt_SolarCapex130

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 1.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. The CapitalCost of every solar technology (PWRSPV, all ten nodes) is multiplied by 1.30 in every year.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `CapitalCost` | multiply | PWRSPV* |

### B_Opt_SolarCapexHi

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 1.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. The CapitalCost of every solar technology (PWRSPV, all ten nodes) is multiplied by 1.10 in every year. To change the increase, edit the factor of its one edit.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `CapitalCost` | multiply | PWRSPV* |

### B_Opt_SolarCapexSpike

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 1.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. The CapitalCost of every solar technology (PWRSPV, all ten nodes) is multiplied by 1.50 in 2028, 2029 and 2030. The other years keep the cost of B_Optimised_VRE.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `CapitalCost` | multiply | PWRSPV* |

### B_Opt_TradeCap15

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 6.
- Overlay: none.
- B2 policy: the historical trade constraints are active; CPLEX numerical emphasis.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. The yearly flow on Bangladesh's four import interconnectors has an upper limit from 2027 to 2050, with no allowance for exports. The activity of Bangladesh's new and repowered internal lines (TRNNLIBGDXX, TRNRPOBGDXX) is set to zero from 2027. ASSUMPTION: the limits were worked out once from an earlier model run, as 15 per cent of Bangladesh's demand split by each interconnector's share of imports in that run. They cannot be reproduced from the repository. This scenario is not the gross import cap of B_Opt_Clipped_GrossImportCap15.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `TotalTechnologyAnnualActivityUpperLimit` | values | TRNBGDXXINDEA, TRNBGDXXINDNE, TRNBTNXXBGDXX, TRNNPLXXBGDXX |
| Demand Techs | `TotalTechnologyAnnualActivityUpperLimit` | values | TRNNLIBGDXX, TRNRPOBGDXX |

### B_Opt_TxCap150

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 40.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. From 2027, the TotalAnnualMaxCapacity and TotalAnnualMaxCapacityInvestment of every cross-border interconnector are the larger of 1.5 times its 2023 residual capacity and its capacity in the base years 2023 to 2026. The base years keep the values of B_Optimised_VRE. The base-year floor is needed because capacity built in the base years lasts into the study period, and the base-year pin needs it. TotalAnnualMinCapacityInvestment is lowered where it would exceed the new limit. Interconnectors inside India are unchanged, and interconnectors with no residual capacity that were never built stay closed. The activity of the new and repowered internal lines outside India (TRNNLI, TRNRPO) is set to zero from 2027. ASSUMPTION: the base-year capacities and the minimum investments were worked out once from an earlier model run. They cannot be reproduced from the repository.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `TotalAnnualMaxCapacity` | set_to_residual_factor_floor | TRNBGDXXINDEA, TRNBGDXXINDNE, TRNBTNXXBGDXX, TRNBTNXXINDEA, TRNBTNXXINDNE, TRNINDEANPLXX, TRNINDNONPLXX, TRNINDSOLKAXX, TRNLKAXXMDVXX, TRNMDVXXINDSO, TRNNPLXXBGDXX |
| Secondary Techs | `TotalAnnualMaxCapacityInvestment` | set_to_residual_factor_floor | TRNBGDXXINDEA, TRNBGDXXINDNE, TRNBTNXXBGDXX, TRNBTNXXINDEA, TRNBTNXXINDNE, TRNINDEANPLXX, TRNINDNONPLXX, TRNINDSOLKAXX, TRNLKAXXMDVXX, TRNMDVXXINDSO, TRNNPLXXBGDXX |
| Secondary Techs | `TotalAnnualMinCapacityInvestment` | values | TRNBTNXXBGDXX, TRNBTNXXINDEA, TRNBTNXXINDNE, TRNINDEANPLXX, TRNINDNONPLXX, TRNINDSOLKAXX, TRNLKAXXMDVXX, TRNNPLXXBGDXX |
| Demand Techs | `TotalTechnologyAnnualActivityUpperLimit` | values | TRNNLIBGDXX, TRNNLIBTNXX, TRNNLILKAXX, TRNNLIMDVXX, TRNNLINPLXX, TRNRPOBGDXX, TRNRPOBTNXX, TRNRPOLKAXX, TRNRPOMDVXX, TRNRPONPLXX |

### C_Target_VRE

- Parent: none. `C_Target_VRE` is a decision root.
- B2 policy: the historical trade constraints are active.
- Seed needs: a solved `A_Calibrated_BAU` result.

### C_Target_VRE_Clipped

- Parent: `C_Target_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 2.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: a solved `A_Calibrated_BAU` result.
- Description: The shared VRE ceilings applied to C_Target_VRE. It also restates the Maldives solar assumption for 2028 that C_Target_VRE already carries: a minimum total of 0.3841 GW, with earlier builds counting towards it. This is a modelling assumption. It is not an official target.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `TotalAnnualMinCapacityInvestment` | values | PWRSPVMDVXX |
| Secondary Techs | `TotalAnnualMinCapacity` | values | PWRSPVMDVXX |

### A_Calibrated_BAU_Clipped_TxFreeze2026

- Parent: `A_Calibrated_BAU`.
- Shared VRE ceilings: applied.
- Scenario edits: 33.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: The shared VRE ceilings applied to A_Calibrated_BAU. From 2027, the capacity of each of the eleven international interconnectors has an upper limit equal to its capacity at the end of 2026, including interconnectors with no capacity. New investment in these interconnectors in 2023 to 2026 is fixed. ASSUMPTION: the end-2026 capacities and the 2023 to 2026 investments were worked out once from an earlier model run of A_Calibrated_BAU_Clipped. They cannot be reproduced from the repository and are not recalculated.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `TotalAnnualMaxCapacity` | values | TRNBGDXXINDEA, TRNBGDXXINDNE, TRNBTNXXBGDXX, TRNBTNXXINDEA, TRNBTNXXINDNE, TRNINDEANPLXX, TRNINDNONPLXX, TRNINDSOLKAXX, TRNLKAXXMDVXX, TRNMDVXXINDSO, TRNNPLXXBGDXX |
| Secondary Techs | `TotalAnnualMinCapacityInvestment` | values | TRNBGDXXINDEA, TRNBGDXXINDNE, TRNBTNXXBGDXX, TRNBTNXXINDEA, TRNBTNXXINDNE, TRNINDEANPLXX, TRNINDNONPLXX, TRNINDSOLKAXX, TRNLKAXXMDVXX, TRNMDVXXINDSO, TRNNPLXXBGDXX |
| Secondary Techs | `TotalAnnualMaxCapacityInvestment` | values | TRNBGDXXINDEA, TRNBGDXXINDNE, TRNBTNXXBGDXX, TRNBTNXXINDEA, TRNBTNXXINDNE, TRNINDEANPLXX, TRNINDNONPLXX, TRNINDSOLKAXX, TRNLKAXXMDVXX, TRNMDVXXINDSO, TRNNPLXXBGDXX |

### B_Opt_Clipped_GrossImportCap15

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 0.
- Overlay: none.
- B2 policy: the historical trade constraints are active; gross imports capped at 0.15 of annual demand in 2027 to 2050 for BGD, BTN, IND, LKA, MDV, NPL.
- Seed needs: none.
- Description: The shared VRE ceilings applied to B_Optimised_VRE. From 2027 to 2050, the electricity each of the six countries receives through its cross-border interconnectors is at most 15 per cent of its yearly demand. India counts as one country.

<!-- generated: end -->

## Origin of each scenario

| Scenario | Origin |
|---|---|
| `BAU` | The support root. It runs only the build-cap rule on its defaults and supplies the restrictions that `A_Calibrated_BAU` inherits. At the end of its A3 run, its interconnector ceilings are set to residual capacity plus committed additions. |
| `A_Calibrated_BAU` | The calibrated business-as-usual root: retirement, the build cap with a flat cap multiplier, the minimum capacity floors, which list nothing on `main`, and interconnector relaxation, with the `BAU` restrictions inherited and the base-year pin applied last. |
| `A_Calibrated_BAU_Clipped` | `A_Calibrated_BAU` with the shared VRE ceilings and nothing else. |
| `B_Optimised_VRE` | The least-cost root: retirement, the build cap with a rising cap multiplier, interconnector relaxation and storage minimum investments. |
| `B_Opt_Clipped` | `B_Optimised_VRE` with the shared VRE ceilings and nothing else. The other B variants are compared with it. |
| `B_Opt_DirBidir` | `B_Opt_Clipped` with an overlay that lists no corridor, so every interconnector flows both ways. It is the reference case for `B_Opt_DirContractual`. |
| `B_Opt_DirContractual` | `B_Opt_Clipped` with an overlay that sets nine corridors to one direction, from their contracts and physical flows, and leaves the two Nepal to India corridors bidirectional. The justification is in the [direction references](interconnector_direction_references.md). |
| `B_Opt_IndiaCosts` | `B_Opt_Clipped` with the capital and fixed costs of the non-India generation and storage technologies set to the India values of the same family and year. The values were first generated from the `B_Optimised_VRE` workbook. |
| `B_Opt_IndiaCostsFuel` | `B_Opt_IndiaCosts` plus the fuel prices of the non-India MIN technologies set to the India prices. |
| `B_Opt_SolarCapex130` | `B_Opt_Clipped` with the solar capital cost multiplied by 1.30 in every year. |
| `B_Opt_SolarCapexHi` | `B_Opt_Clipped` with the solar capital cost multiplied by 1.10 in every year. |
| `B_Opt_SolarCapexSpike` | `B_Opt_Clipped` with the solar capital cost multiplied by 1.50 in 2028, 2029 and 2030 only. |
| `B_Opt_TradeCap15` | `B_Opt_Clipped` with fixed annual activity limits on the four Bangladesh import corridors, computed once as 15 per cent of demand split by the import shares of a historical `B_Optimised_VRE` result, and zero activity on Bangladesh's new and repowered internal lines from 2027. |
| `B_Opt_TxCap150` | `B_Opt_Clipped` with the capacity of each international corridor in 2027 to 2050 capped at the larger of 1.5 times its 2023 residual capacity and its calibrated base-window capacity, plus literal tables of minimum investments and backstop limits. |
| `C_Target_VRE` | The target root: retirement, the minimum capacity floors with the Nepal hydro ceiling, the build cap, the VRE targets sized from the solved A result, interconnector relaxation and storage minimum investments. |
| `C_Target_VRE_Clipped` | `C_Target_VRE` with the shared VRE ceilings, plus the restated Maldives 2028 solar assumption: a minimum total of 0.3841 GW with earlier builds counting towards it. |
| `A_Calibrated_BAU_Clipped_TxFreeze2026` | `A_Calibrated_BAU_Clipped` with upper capacity limits and fixed early investments for the eleven international corridors, derived once from a historical A-clipped reference at the end of 2026. The tables are literal values in its scenario file; the pipeline never recomputes them from a current A result. |
| `B_Opt_Clipped_GrossImportCap15` | `B_Opt_Clipped` with a policy that caps the gross received imports of each country at 15 per cent of its annual demand in 2027 to 2050, for all six countries with India counted nationally. |
