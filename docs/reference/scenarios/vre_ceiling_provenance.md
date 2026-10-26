# OSTRAM — VRE-ceiling & key-assumption provenance / citations

## 1. VRE resource-potential ceilings — reference list

The applied values are the `vre_ceilings` section of `config/setups/south_asia.yaml`.

**India — solar (NISE):**
- NISE/MNRE state-wise solar potential: **748.99 GWp** (748,990 MWp), estimated by NISE from the Wasteland Atlas of India 2010, with 3% of wasteland covered by solar PV. MNRE solar overview (Internet Archive copy, 23 Jan 2026) https://web.archive.org/web/20260123045204/https://mnre.gov.in/en/solar-overview/ · PIB, Ministry of New and Renewable Energy, 7 Feb 2024, state-wise table, total row https://www.pib.gov.in/Pressreleaseshare.aspx?PRID=2003561
- NISE/MNRE ground-mounted solar PV potential (2025): **3,343.37 GWp** on 27,571 km² of feasible wasteland, 6.69% of the 412,458 km² of identified wasteland (Table 5, total row and the paragraph below it, printed p. 32), released 23 Sep 2025. https://nise.res.in/upload/Solar_PV_potential_ground_mounted_2025.pdf
- Model uses the **2025 vintage** (India solar ceilings sum ~1,925 GW, a subset of 3,343).

**India — wind (NIWE):**
- NIWE India Wind Potential Atlas at 150 m agl: **1,163.9 GW** (695.50 GW at 120 m). https://maps.niwe.res.in/media/150m-report.pdf · interactive map, with no national total https://maps.niwe.res.in/resource_map/map/150m/ · https://mnre.gov.in/en/wind-overview/
- The applied India wind ceilings sum to 1,140 GW (INDNO 285, INDSO 445, INDWE 410, INDEA 0 and INDNE 0), close to the NIWE 150 m national total.

**Non-India (per country):**
- **BANGLADESH**: SREDA/Power Division (NREL/USAID), National Solar Energy Roadmap 2021-2041 https://rise.esmap.org/sites/default/files/library/bangladesh/Renewable%20Energy/Bangladesh_National-Solar-Energy-Roadmap.pdf ; land-screened solar ~53 GW (2019). Wind: NREL, Assessing the Wind Energy Potential in Bangladesh (NREL/TP-5000-71077, September 2018), printed p. 62 (PDF p. 72): more than 20,000 km² with a gross wind potential of over 30,000 MW at 5.75 to 7.75 m/s, which the report calls not realistic once unsuitable land is screened out https://docs.nlr.gov/docs/fy18osti/71077.pdf
- **SRI LANKA**: ADB & UNDP, "100% RE by 2050" (2017): 2050 mix ~15,000 MW wind + ~16,000 MW solar https://www.adb.org/sites/default/files/publication/354591/sri-lanka-power-2050v2.pdf ; ADB Energy Sector Assessment, Strategy, and Road Map (2019), PDF p. 12: national solar technical potential 6,000 MW and national exploitable wind potential 5,600 MW https://www.adb.org/sites/default/files/institutional-document/547381/sri-lanka-energy-assessment-strategy-road-map.pdf ; SLSEA RE Resource Development Plan 2021–2026 https://www.energy.gov.lk/images/renewable-energy/renewable-energy-resource-development-plan-en.pdf
- **NEPAL**: SWERA GIS report (UNEP/NREL/Risø) https://policy.asiapacificenergy.org/node/358 ; provincial assessment: Neupane et al., Solar and wind energy potential assessment at provincial level in Nepal, Renewable Energy 181 (2022), 278-291, up to about 47,628 MW of solar and 1,686 MW of wind https://doi.org/10.1016/j.renene.2021.09.027
- **BHUTAN**: Cowlin & Gilman, NREL/TP-6A2-46547 (2009) https://www.osti.gov/biblio/964607 ; RGoB Renewable Energy Master Plan 2017-2032: 12 GW of solar and 760 MW of wind in technical terms, as reported in IRENA and Bhutan's Department of Renewable Energy, Renewables Readiness Assessment: Kingdom of Bhutan (December 2019), printed p. II (PDF p. 4) https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2019/Dec/IRENA_RRA_Bhutan_2019.pdf
- **MALDIVES**: Ministry of Climate Change, Environment and Energy, Energy Roadmap 2024-2033 (large-scale onshore wind ~80 MW niche, Greater Malé) https://media.environment.gov.mv/20241115-pub-energy-roadmap-maldives-2024-2033.pdf ; NREL, Wind Energy Resource Atlas of Sri Lanka and the Maldives (NREL/TP-500-34518, August 2003) https://docs.nlr.gov/docs/fy03osti/34518.pdf

**Uniform multi-country option:**
- Global Solar Atlas (World Bank/ESMAP/Solargis): method and country results in World Bank, Global Photovoltaic Power Potential by Country (June 2020) https://documents1.worldbank.org/curated/en/466331592817725242/pdf/Global-Photovoltaic-Power-Potential-by-Country.pdf · Global Wind Atlas (World Bank/DTU): DTU, Global Wind Atlas v3 dataset and method (250 m grid, ERA5 2008 to 2017, posted 13 Mar 2024) https://data.dtu.dk/articles/dataset/Global_Wind_Atlas_v3/9420803

---

## 2. Per-node provenance table (model value vs best source; ★ = enforced clip)

| Node | Tech | Value (GW) | Best source |
|---|---|--:|---|
| INDNO | Solar | 600 | ASSUMPTION; see the NISE vintage note (NISE-2025, 985.96 GWp) |
| INDWE | Solar | 650 | ASSUMPTION; see the NISE vintage note (NISE-2025, 1,182.60 GWp) |
| INDSO | Solar | 500 | ASSUMPTION; see the NISE vintage note (NISE-2025, 880.41 GWp) |
| INDEA | Solar | 150 | ASSUMPTION; see the NISE vintage note (NISE-2025, 247.87 GWp) |
| INDNE | Solar | 25 | ASSUMPTION; see the NISE vintage note (NISE-2025, 46.51 GWp) |
| BGD | Solar | 40 | ASSUMPTION: 40 GW in every year, a round value below the ~53 GW of land-screened solar in the National Solar Energy Roadmap, so the model cannot use all of the screened land |
| LKA | Solar | 16 ★ | ADB/UNDP 2050 (~16,000 MW) |
| NPL | Solar | 40 | ASSUMPTION: 40 GW in every year, a round value below the 47,628 MW of solar potential in Neupane et al. (2022), so the model cannot use all of the assessed potential |
| BTN | Solar | 1.8102 | ASSUMPTION; see the `B_Optimised_VRE` limits below (RE Master Plan, 12 GW) |
| MDV | Solar | 1 | ASSUMPTION: 1 GW in every year, a round value above the 68.5 MW of solar PV installed and the 90 MW under development in mid-2024 (Maldives Energy Roadmap 2024-2033, PDF p. 9), so solar can grow beyond the 2028 pipeline |
| INDNO | Wind | 285 | NIWE 150 m (N incl. Rajasthan) |
| INDSO | Wind | 445 | NIWE 150 m (TN/KA/AP/TS) |
| INDWE | Wind | 410 | NIWE 150 m (GJ/MH/MP) |
| INDEA | Wind | 0 | ASSUMPTION; see the `B_Optimised_VRE` limits below (NIWE 150 m, 17,449 MW for OD/WB/BR/JH) |
| INDNE | Wind | 0 | ASSUMPTION; see the `B_Optimised_VRE` limits below (NIWE 150 m, 760 MW for the NE states) |
| BGD | Wind | 3 ★ | ASSUMPTION; see the clips below (NREL, >30 GW gross, unscreened) |
| LKA | Wind | 15 | ADB/UNDP 2050 (15,000 MW); ADB 2019: 5,600 MW exploitable wind potential |
| NPL | Wind | 0 | ASSUMPTION; see the `B_Optimised_VRE` limits below (Neupane et al., 1,686 MW) |
| BTN | Wind | 0.5 | ASSUMPTION: 0.5 GW, a round value below the 760 MW of technical wind potential in IRENA's Renewables Readiness Assessment of Bhutan (2019, printed p. II), so the model cannot use all of it |
| MDV | Wind | 0 ★ | ASSUMPTION; see the clips below (Maldives Roadmap, ~80 MW niche) |

### The applied values and the `B_Optimised_VRE` limits

Each ceiling in `south_asia.yaml` is the lower of an atlas-based figure and the
`B_Optimised_VRE` limit. Four values hold the `B_Optimised_VRE` limit:

- ASSUMPTION: Bhutan solar is limited to 1.8102 GW in every year, the `B_Optimised_VRE` limit,
  so the scenarios that apply the ceilings build no more Bhutan solar than `B_Optimised_VRE`.
  Bhutan's Renewable Energy Master Plan gives 12 GW of solar in technical terms (IRENA, 2019, above).
- ASSUMPTION: India East, India North-East and Nepal build no onshore wind (0 GW in every year),
  as in `B_Optimised_VRE`. NIWE's 150 m atlas gives 17,449 MW for Odisha, West Bengal, Bihar and
  Jharkhand and 760 MW for Arunachal Pradesh, Assam and Meghalaya, with 0 for the other
  North-Eastern states (Table 4, printed pp. 24-25), and Neupane et al. (2022) give 1,686 MW for Nepal.

The three clips below are the cases where the atlas figure itself binds.

The pipeline applies every ceiling as a fixed value in all years of a derived scenario. It never
computes the lower of the previous workbook value and the ceiling, and a ceiling above the
previous value raises it. The annual investment bounds of the same technology are lowered to the
ceiling where they exceed it. The [scenario edits](../../scenario-edits.md) page describes the step.

### The three clips
- **LKA solar 16 ★ — well supported**, but it matches a 2050 *scenario build* (ADB/UNDP), not pure technical potential (ADB technical ≈ 6 GW).
- **BGD onshore wind 3 ★.** ASSUMPTION: Bangladesh onshore wind is limited to 3 GW in every year, a round value far below NREL's gross estimate of over 30,000 MW, which NREL calls not realistic once unsuitable land is screened out (NREL/TP-5000-71077, printed p. 62, above).
- **MDV onshore wind 0 ★.** ASSUMPTION: the Maldives builds no onshore wind (0 GW in every year), because the Energy Roadmap 2024-2033 finds large-scale onshore wind possible only on industrial islands of Greater Malé, about 80 MW (PDF p. 23), and Rinaldi et al. (2019) report onshore wind and photovoltaic installations excluded from the archipelago's energy strategy at that time (International Journal of Energy Research, printed p. 2) https://doi.org/10.1002/er.4762

---

## 3. Transmission-loss (3%) provenance

- **Primary source:** Ministry of Power, Government of India, "Several Actions Taken to Reduce Power Transmission Losses", PIB, 4 Dec 2025 (written reply in the Lok Sabha). https://www.pib.gov.in/PressReleasePage.aspx?PRID=2198781&lang=2&reg=48
- **Figures:** transmission losses in India vary between about **3% and 4%**; the region-wise transmission losses for 2023-24 from CEA's General Review (November 2025) are North 3.15%, West 3.12%, South 3.46%, East 4.24% and North-East 3.89%. ASSUMPTION: the internal transmission loss is 3% of the published requirement for every bus-bar demand node (Bangladesh, Bhutan, the five India regions, Nepal and Sri Lanka) in all years, one round value at the lower end of the 3% to 4% that India's Ministry of Power gives; it is below each of the 2023-24 regional values above.
- **Distinguish:** transmission only ≈ **3% to 4%** (above) · **combined T&D in India ≈ 14.5% to 14.9% in 2022 to 2024**, and 16.4% to 20.0% in 2010 to 2020 (World Bank, EG.ELC.LOSS.ZS, https://data.worldbank.org/indicator/EG.ELC.LOSS.ZS?locations=IN) · **AT&C 15.04% in FY2024-25**, down from 21.91% in FY2020-21 (Ministry of Power, Lok Sabha Unstarred Question 4410, answered 19 Mar 2026, https://sansad.in/getFile/loksabhaquestions/annex/187/AU4410_ffcIsh.pdf?source=pqals). The high figures are distribution-dominated; do not conflate them.

### The current demand accounting

The demand rows of Bangladesh, Bhutan, the five India regions, Nepal and Sri Lanka
are measured at the bus bar and hold the published requirement x (1 - loss), with the 3%
above, so the plants and imports that feed them through the internal grid meet the
published requirement. Imports cross the receiving node's internal grid; an
interconnector sent from a bus-bar node draws 1 - loss per unit of activity, because the
published requirement already holds the losses export flows cause. The Maldives row is
consumption and keeps its value.
The `Measured_at` column of the `Demand_Projection` sheet in
`inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx` marks each row, and
`ostram/pipeline/scenarios/rules/apply_internal_tx_losses.py` applies the factor.

## 4. WACC / discount rate (10%) provenance

- **Primary anchor:** IRENA, "Renewable Power Generation Costs in 2018" — "a real cost of capital of 7.5% in OECD countries and China, and **10% in the rest of the world**." https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2019/May/IRENA_Renewable-Power-Generations-Costs-in-2018.pdf
- **Range brackets:** OECD Environment Working Paper No. 245 (18 Jul 2024), p. 20: a WACC of about **5%** for utility-scale solar PV in advanced economies and about **9%** in India, Indonesia or Senegal https://one.oecd.org/document/ENV/WKP(2024)15/REV1/en/pdf ; World Bank "Demystifying the Costs..." used **6%** https://documents1.worldbank.org/curated/en/125521593437517815/pdf/Demystifying-the-Costs-of-Electricity-Generation-Technologies.pdf ; IRENA (2023) trends lower (→5% OECD / 7.5% RoW by 2021) https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2023/May/IRENA_The_cost_of_financing_renewable_power_2023.pdf
- In the model the 10% is a study discount convention, recorded in the row "Sets, defaults and empty inputs" of the README sheet in the scenario workbook. It applies to every technology. The sources above are benchmarks for that convention; they do not establish the financing cost of any one technology or source.

## 5. Storage cost basis

- The storage costs come from NREL, "Utility-Scale Battery Storage", **2024 ATB** v3, rows 20 (energy) and 26 (power), Moderate trajectories 2023 to 2050, in 2022 US dollars. https://atb.nlr.gov/electricity/2024/utility-scale_battery_storage ; projections: Cole & Karmakar, NREL 2023 update https://doi.org/10.2172/1984976
- The 2022 values are converted once to constant 2023 US dollars. ASSUMPTION: every storage cost is multiplied by 131.902/127.224 = 1.0368, a US GDP deflator ratio of 2023 to 2022, to move it from 2022 to 2023 US dollars; the World Bank's US GDP deflator gives 116.14/112.00 = 1.0370 for the same two years (NY.GDP.DEFL.ZS, https://data.worldbank.org/indicator/NY.GDP.DEFL.ZS?locations=US). The power component carries the node factors; the energy component is shared across nodes and converted to the model's energy unit.
- The fixed operating cost is 2.5% per year of the capital cost of a four-hour reference system, including augmentation; the 2024 ATB sets battery fixed O&M at 2.5% of the capital cost in $/kW, including augmentation (page above, Operation and Maintenance Costs). ASSUMPTION: the four-hour system is the reference for every storage technology and node, because the model chooses the storage duration and needs one power-based fixed cost; the ATB itself covers durations of 2, 4, 6, 8 and 10 hours.
- The generic long-duration storage is a study assumption with separate power and energy components. ATB does not establish its costs or lifetimes.
- Sources in the repository: the rows "Monetary basis and interpretation" and "Storage costs and service" of the README sheet in `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx`, and the cell notes of the storage cost rows on its `Secondary_Techs` sheet.

**NISE vintage note.** ASSUMPTION: the India solar ceilings are round values of 54% to 61% of NISE-2025's ground-mounted potential for each region's states, grouped as in CEA's regional grids: INDNO 600 GW of 985.96 GWp, INDWE 650 GW of 1,182.60 GWp, INDSO 500 GW of 880.41 GWp, INDEA 150 GW of 247.87 GWp and INDNE 25 GW of 46.51 GWp (Table 5, printed pp. 31-32). Together they are 1,925 GW, 58% of the national 3,343.37 GWp, so the model cannot use all of the wasteland NISE counts as feasible. They are not an average of the 2014 and 2025 vintages.
