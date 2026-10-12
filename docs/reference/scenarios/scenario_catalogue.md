# Scenario catalogue

The block between the two markers is generated from `config/profiles/full.yaml`
and `config/execution/Config_MOMF_T1_AB.yaml` by
`python tools/generate_scenario_catalogue.py`, run from the repository root.
The test `tests/regression/test_scenario_catalogue.py` fails when the block
or the origin table is out of date. Terms are in the [glossary](../../glossary.md).

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
  Bangladesh import corridors and zero backstop imports from 2027, through
  scenario edits. `B_Opt_Clipped_GrossImportCap15` caps the gross received
  imports of each of the six countries at 15 per cent of its annual demand in
  2027 to 2050, through a constraint that B2 builds from the compiled demand
  and output ratios.
- `C_Target_VRE_Clipped` restates the Maldives solar assumption for 2028 that
  `C_Target_VRE` already carries from the scenario workbook.

## Registered scenarios

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
- Description: Interconnector directions locked to real contractual/physical flows (9 corridors one-way, Nepal<->India pair bidirectional). Shared VRE ceiling layer via scenario edits; direction overlay via set_interconnector_direction.py. Full per-corridor justification + citations in docs/reference/scenarios/interconnector_direction_references.md.

### B_Opt_IndiaCosts

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 112.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: CapitalCost and FixedCost of the existing non-India generation and storage technologies, except the backstop, set to the India values for each family and year. Where the India regions differ (coal, hydro and offshore wind), the INDNO values are used. Transmission and fuel prices are unchanged. Applied after the shared VRE ceilings. The values were first copied from the B_Optimised_VRE workbook and were updated in the September 2026 cost release.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `CapitalCost` | values | PWRBIOBGDXX, PWRBIOLKAXX, PWRBIONPLXX, PWRCCSBGDXX, PWRCOABGDXX, PWRCOALKAXX, PWRCOGBGDXX, PWRCOGLKAXX, PWRHYDBGDXX, PWRHYDBTNXX, PWRHYDLKAXX, PWRHYDNPLXX, PWRLDSBGDXX, PWRLDSBTNXX, PWRLDSLKAXX, PWRLDSMDVXX, PWRLDSNPLXX, PWRNGSBGDXX, PWRNGSLKAXX, PWRNGSMDVXX, PWROILBGDXX, PWROILLKAXX, PWROILMDVXX, PWROILNPLXX, PWROTHBGDXX, PWROTHBTNXX, PWROTHLKAXX, PWROTHMDVXX, PWROTHNPLXX, PWRPETBGDXX, PWRPETBTNXX, PWRPETLKAXX, PWRPETMDVXX, PWRPETNPLXX, PWRSDSBGDXX, PWRSDSBTNXX, PWRSDSLKAXX, PWRSDSMDVXX, PWRSDSNPLXX, PWRSPVBGDXX, PWRSPVBTNXX, PWRSPVLKAXX, PWRSPVMDVXX, PWRSPVNPLXX, PWRURNBGDXX, PWRWASBGDXX, PWRWASLKAXX, PWRWASNPLXX, PWRWOFBGDXX, PWRWOFLKAXX, PWRWOFMDVXX, PWRWONBGDXX, PWRWONBTNXX, PWRWONLKAXX, PWRWONMDVXX, PWRWONNPLXX |
| Secondary Techs | `FixedCost` | values | PWRBIOBGDXX, PWRBIOLKAXX, PWRBIONPLXX, PWRCCSBGDXX, PWRCOABGDXX, PWRCOALKAXX, PWRCOGBGDXX, PWRCOGLKAXX, PWRHYDBGDXX, PWRHYDBTNXX, PWRHYDLKAXX, PWRHYDNPLXX, PWRLDSBGDXX, PWRLDSBTNXX, PWRLDSLKAXX, PWRLDSMDVXX, PWRLDSNPLXX, PWRNGSBGDXX, PWRNGSLKAXX, PWRNGSMDVXX, PWROILBGDXX, PWROILLKAXX, PWROILMDVXX, PWROILNPLXX, PWROTHBGDXX, PWROTHBTNXX, PWROTHLKAXX, PWROTHMDVXX, PWROTHNPLXX, PWRPETBGDXX, PWRPETBTNXX, PWRPETLKAXX, PWRPETMDVXX, PWRPETNPLXX, PWRSDSBGDXX, PWRSDSBTNXX, PWRSDSLKAXX, PWRSDSMDVXX, PWRSDSNPLXX, PWRSPVBGDXX, PWRSPVBTNXX, PWRSPVLKAXX, PWRSPVMDVXX, PWRSPVNPLXX, PWRURNBGDXX, PWRWASBGDXX, PWRWASLKAXX, PWRWASNPLXX, PWRWOFBGDXX, PWRWOFLKAXX, PWRWOFMDVXX, PWRWONBGDXX, PWRWONBTNXX, PWRWONLKAXX, PWRWONMDVXX, PWRWONNPLXX |

### B_Opt_IndiaCostsFuel

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 147.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: CapitalCost and FixedCost of the existing non-India generation and storage technologies, except the backstop, set to the India values for each family and year. Where the India regions differ (coal, hydro and offshore wind), the INDNO values are used. Transmission is unchanged. Non-India fuel prices (VariableCost of the MIN fuel-supply technologies) are also set to the India prices; these are higher for some fuels and years and lower for others. Applied after the shared VRE ceilings. The values were first copied from the B_Optimised_VRE workbook and were updated in the September 2026 cost release.

| Sheet | Parameter | Operation | Technologies |
|---|---|---|---|
| Secondary Techs | `CapitalCost` | values | PWRBIOBGDXX, PWRBIOLKAXX, PWRBIONPLXX, PWRCCSBGDXX, PWRCOABGDXX, PWRCOALKAXX, PWRCOGBGDXX, PWRCOGLKAXX, PWRHYDBGDXX, PWRHYDBTNXX, PWRHYDLKAXX, PWRHYDNPLXX, PWRLDSBGDXX, PWRLDSBTNXX, PWRLDSLKAXX, PWRLDSMDVXX, PWRLDSNPLXX, PWRNGSBGDXX, PWRNGSLKAXX, PWRNGSMDVXX, PWROILBGDXX, PWROILLKAXX, PWROILMDVXX, PWROILNPLXX, PWROTHBGDXX, PWROTHBTNXX, PWROTHLKAXX, PWROTHMDVXX, PWROTHNPLXX, PWRPETBGDXX, PWRPETBTNXX, PWRPETLKAXX, PWRPETMDVXX, PWRPETNPLXX, PWRSDSBGDXX, PWRSDSBTNXX, PWRSDSLKAXX, PWRSDSMDVXX, PWRSDSNPLXX, PWRSPVBGDXX, PWRSPVBTNXX, PWRSPVLKAXX, PWRSPVMDVXX, PWRSPVNPLXX, PWRURNBGDXX, PWRWASBGDXX, PWRWASLKAXX, PWRWASNPLXX, PWRWOFBGDXX, PWRWOFLKAXX, PWRWOFMDVXX, PWRWONBGDXX, PWRWONBTNXX, PWRWONLKAXX, PWRWONMDVXX, PWRWONNPLXX |
| Secondary Techs | `FixedCost` | values | PWRBIOBGDXX, PWRBIOLKAXX, PWRBIONPLXX, PWRCCSBGDXX, PWRCOABGDXX, PWRCOALKAXX, PWRCOGBGDXX, PWRCOGLKAXX, PWRHYDBGDXX, PWRHYDBTNXX, PWRHYDLKAXX, PWRHYDNPLXX, PWRLDSBGDXX, PWRLDSBTNXX, PWRLDSLKAXX, PWRLDSMDVXX, PWRLDSNPLXX, PWRNGSBGDXX, PWRNGSLKAXX, PWRNGSMDVXX, PWROILBGDXX, PWROILLKAXX, PWROILMDVXX, PWROILNPLXX, PWROTHBGDXX, PWROTHBTNXX, PWROTHLKAXX, PWROTHMDVXX, PWROTHNPLXX, PWRPETBGDXX, PWRPETBTNXX, PWRPETLKAXX, PWRPETMDVXX, PWRPETNPLXX, PWRSDSBGDXX, PWRSDSBTNXX, PWRSDSLKAXX, PWRSDSMDVXX, PWRSDSNPLXX, PWRSPVBGDXX, PWRSPVBTNXX, PWRSPVLKAXX, PWRSPVMDVXX, PWRSPVNPLXX, PWRURNBGDXX, PWRWASBGDXX, PWRWASLKAXX, PWRWASNPLXX, PWRWOFBGDXX, PWRWOFLKAXX, PWRWOFMDVXX, PWRWONBGDXX, PWRWONBTNXX, PWRWONLKAXX, PWRWONMDVXX, PWRWONNPLXX |
| VariableCost | `VariableCost` | values | MINCOABGD, MINCOABTN, MINCOALKA, MINCOAMDV, MINCOANPL, MINCOGBGD, MINCOGBTN, MINCOGLKA, MINCOGMDV, MINCOGNPL, MINGASBGD, MINGASBTN, MINGASLKA, MINGASMDV, MINGASNPL, MINOILBGD, MINOILBTN, MINOILLKA, MINOILMDV, MINOILNPL, MINOTHBGD, MINOTHBTN, MINOTHLKA, MINOTHMDV, MINOTHNPL, MINPETBGD, MINPETBTN, MINPETLKA, MINPETMDV, MINPETNPL, MINURNBGD, MINURNBTN, MINURNLKA, MINURNMDV, MINURNNPL |

### B_Opt_SolarCapex130

- Parent: `B_Optimised_VRE`.
- Shared VRE ceilings: applied.
- Scenario edits: 1.
- Overlay: none.
- B2 policy: the historical trade constraints are active.
- Seed needs: none.
- Description: PRIMARY sustained solar CAPEX stress. Solar CapitalCost x1.30 for all PWRSPV<node> (10 nodes), all years, over the shared VRE-ceiling layer. Literature-anchored: NREL ATB Conservative-vs-Moderate spread ~30%, IEA diversification premium, WACC-stress equivalent. Same mechanism as SolarCapexHi, factor 1.30.

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
- Description: Solar CapitalCost multiplied by 'factor' (setting, default 1.10) for all PWRSPV<node>, all 10 nodes, all years. Over the shared VRE-ceiling layer (expected not to bind: dearer solar -> less solar). To change the stress level before CPLEX, edit the single 'factor' value below.

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
- Description: Transient severe solar CAPEX spike x1.50 for 2028-2030 ONLY (other years keep the parent cost). Uses the multiply op's optional per-year 'years' filter (added to scenario_edits.py for this scenario). Over the shared VRE-ceiling layer.

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
- Description: Fixed upper limits on the yearly flow of Bangladesh's four import corridors from 2027 to 2050. They were calculated once as 15 per cent of Bangladesh's demand, split by each corridor's share of imports in a historical B_Optimised_VRE run, with no allowance for exports. Backstop imports to Bangladesh (TRNNLIBGDXX, TRNRPOBGDXX) are set to zero from 2027. Applied after the shared VRE ceilings. This scenario differs from the gross import cap of B_Opt_Clipped_GrossImportCap15.

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
- Description: TotalAnnualMaxCapacity = max(1.5 x ResidualCapacity_2023, calibrated base-window capacity) for every cross-border TRN corridor, applied to the STUDY PERIOD 2027+ only (base years 2023-2026 revert to the source/pinned calibrated mix). The base-window grandfather is required because interconnector capacity is cumulative and long-lived: the WS-4 base-year pin fixes each node's base-year net import, which needs the calibrated base-window corridor capacity, so a bare 1.5x-residual cap on 2027 would retroactively starve the pinned base years (genuine demand-balance infeasibility). India-internal corridors kept at B_Opt. Non-India TRNNLI*/TRNRPO* backstops -> AUL 0 from 2027+ (base years = calibrated). Zero-residual, never-built corridors stay blocked.

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
- Description: The shared VRE ceilings applied to C_Target_VRE. It also restates the Maldives solar assumption for 2028 that C_Target_VRE already carries: a minimum total of 0.3841 GW, with earlier builds counting towards it. This is a legacy modelling assumption. It is not an official target.

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
- Description: Clipped A with all eleven international corridors limited from 2027 to their end-2026 capacity in the historical M0 A_Calibrated_BAU_Clipped run. This is an upper limit. New investment in 2023 to 2026 is fixed to the values of that run. Corridors with zero capacity are included. The reference values are fixed and are not recalculated from a new A run.

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
- Description: Clipped B with annual aggregate gross received-import cap, all six countries per legacy documented region-wide scope; India aggregated nationally.

<!-- generated: end -->

## Origin of each scenario

| Scenario | Origin |
|---|---|
| `BAU` | The support root. It runs only the build-cap rule on its defaults and supplies the restrictions that `A_Calibrated_BAU` inherits. At the end of its A3 run, its interconnector ceilings are set to residual capacity plus committed additions. |
| `A_Calibrated_BAU` | The calibrated business-as-usual root: retirement, the build cap with a flat cap multiplier, the minimum capacity floors, which list nothing on `main`, and interconnector relaxation, with the `BAU` restrictions inherited and the base-year pin applied last. |
| `A_Calibrated_BAU_Clipped` | `A_Calibrated_BAU` with the shared VRE ceilings and nothing else. |
| `B_Optimised_VRE` | The least-cost root: retirement, the build cap with a rising cap multiplier, interconnector relaxation and storage minimum investments. |
| `B_Opt_Clipped` | `B_Optimised_VRE` with the shared VRE ceilings and nothing else. The other B variants are compared with it. |
| `B_Opt_DirBidir` | `B_Opt_Clipped` with an overlay that lists no corridor, so every interconnector flows both ways. It shows that the direction machinery changes nothing on its own. |
| `B_Opt_DirContractual` | `B_Opt_Clipped` with an overlay that sets nine corridors to one direction, from their contracts and physical flows, and leaves the two Nepal to India corridors bidirectional. The justification is in the [direction references](interconnector_direction_references.md). |
| `B_Opt_IndiaCosts` | `B_Opt_Clipped` with the capital and fixed costs of the non-India generation and storage technologies set to the India values of the same family and year. The values were first generated from the `B_Optimised_VRE` workbook and were revised in the September 2026 cost release. |
| `B_Opt_IndiaCostsFuel` | `B_Opt_IndiaCosts` plus the fuel prices of the non-India MIN technologies set to the India prices. |
| `B_Opt_SolarCapex130` | `B_Opt_Clipped` with the solar capital cost multiplied by 1.30 in every year. |
| `B_Opt_SolarCapexHi` | `B_Opt_Clipped` with the solar capital cost multiplied by 1.10 in every year. |
| `B_Opt_SolarCapexSpike` | `B_Opt_Clipped` with the solar capital cost multiplied by 1.50 in 2028, 2029 and 2030 only. |
| `B_Opt_TradeCap15` | `B_Opt_Clipped` with fixed annual activity limits on the four Bangladesh import corridors, computed once as 15 per cent of demand split by the import shares of a historical `B_Optimised_VRE` result, and zero backstop imports from 2027. |
| `B_Opt_TxCap150` | `B_Opt_Clipped` with the capacity of each international corridor in 2027 to 2050 capped at the larger of 1.5 times its 2023 residual capacity and its calibrated base-window capacity, plus literal tables of minimum investments and backstop limits. |
| `C_Target_VRE` | The target root: retirement, the minimum capacity floors with the Nepal hydro ceiling, the build cap, the VRE targets sized from the solved A result, interconnector relaxation and storage minimum investments. |
| `C_Target_VRE_Clipped` | `C_Target_VRE` with the shared VRE ceilings, plus the restated Maldives 2028 solar assumption: a minimum total of 0.3841 GW with earlier builds counting towards it. |
| `A_Calibrated_BAU_Clipped_TxFreeze2026` | `A_Calibrated_BAU_Clipped` with upper capacity limits and fixed early investments for the eleven international corridors, derived once from a historical A-clipped reference at the end of 2026. The tables are literal values in `full.yaml`; the pipeline never recomputes them from a current A result. |
| `B_Opt_Clipped_GrossImportCap15` | `B_Opt_Clipped` with a policy that caps the gross received imports of each country at 15 per cent of its annual demand in 2027 to 2050, for all six countries with India counted nationally. |
