# OSTRAM — VRE-ceiling & key-assumption provenance / citations

## 1. VRE resource-potential ceilings — reference list

The applied values are the `vre_ceilings` section of `config/setups/south_asia.yaml`.

**India — solar (NISE):**
- NISE/MNRE state-wise solar potential (2014): **748.98 GWp** (3% of wasteland + Census 2011). https://mnre.gov.in/en/solar-overview/ · PIB https://www.pib.gov.in/Pressreleaseshare.aspx?PRID=2003561
- NISE/MNRE ground-mounted solar PV potential (2025): **3,343.37 GWp** (6.69% of feasible wasteland ~27,571 km²), released 23 Sep 2025. https://nise.res.in/wp-content/uploads/2025/09/Poster-and-Momento.pdf
- Model uses the **2025 vintage** (India solar ceilings sum ~1,925 GW — a subset of 3,343; impossible under 2014's 749).

**India — wind (NIWE):**
- NIWE India Wind Potential Atlas at 150 m agl: **1,163.9 GW** (695.50 GW at 120 m). https://maps.niwe.res.in/media/150m-report.pdf · map https://maps.niwe.res.in/resource_map/map/150m/ · https://mnre.gov.in/en/wind-overview/
- The applied India wind ceilings sum to 1,140 GW (INDNO 285, INDSO 445, INDWE 410, INDEA 0 and INDNE 0), close to the NIWE 150 m national total.

**Non-India (per country):**
- **BANGLADESH** — SREDA/Power Division (NREL/USAID), National Solar Energy Roadmap 2021–2041 https://rise.esmap.org/sites/default/files/library/bangladesh/Renewable%20Energy/Bangladesh_National-Solar-Energy-Roadmap.pdf ; land-screened solar ~53 GW (2019). Wind: NREL/SREDA gross ">30,000 MW" over 20,000 km² — source-flagged as unfiltered/unrealistic.
- **SRI LANKA** — ADB & UNDP, "100% RE by 2050" (2017): 2050 mix ~15,000 MW wind + ~16,000 MW solar https://www.adb.org/sites/default/files/publication/354591/sri-lanka-power-2050v2.pdf ; ADB Energy Sector Assessment (solar technical ~6,000 MW; wind ~5,600 MW Mannar) https://www.adb.org/sites/default/files/institutional-document/547381/sri-lanka-energy-assessment-strategy-road-map.pdf ; SLSEA RE Resource Development Plan 2021–2026 https://www.energy.gov.lk/images/renewable-energy/renewable-energy-resource-development-plan-en.pdf
- **NEPAL** — SWERA GIS report (UNEP/NREL/Risø) https://policy.asiapacificenergy.org/node/358 ; provincial assessment, Renewable Energy 2021 https://doi.org/10.1016/j.renene.2021.08.109
- **BHUTAN** — Cowlin & Gilman, NREL/TP-6A2-46547 (2009) https://www.osti.gov/biblio/964607 ; RGoB Renewable Energy Master Plan 2017–2032 (~12,000 MW solar, ~760 MW wind).
- **MALDIVES** — Ministry of Environment, Energy Roadmap 2024–2033 (large-scale onshore wind ~80 MW niche, Greater Malé) https://www.environment.gov.mv/v2/wp-content/files/publications/20241107-pub-energy-roadmap-maldives-2024-2033-.pdf ; NREL Wind Atlas of Sri Lanka & Maldives (2003).

**Uniform multi-country option:**
- Global Solar Atlas (World Bank/ESMAP/Solargis) https://globalsolaratlas.info · Global Wind Atlas (World Bank/DTU) https://globalwindatlas.info

---

## 2. Per-node provenance table (model value vs best source; ★ = enforced clip)

| Node | Tech | Value (GW) | Best source |
|---|---|--:|---|
| INDNO | Solar | 600 | NISE-2025 (N states) |
| INDWE | Solar | 650 | NISE-2025 (Rajasthan/Gujarat/MH/MP) |
| INDSO | Solar | 500 | NISE-2025 (AP/KA/TN/TS/KL) |
| INDEA | Solar | 150 | NISE-2025 (Bihar/JH/OD/WB) |
| INDNE | Solar | 25 | NISE-2025 (Assam+7) |
| BGD | Solar | 40 | SREDA / ~53 GW land-screened |
| LKA | Solar | 16 ★ | ADB/UNDP 2050 (~16,000 MW) |
| NPL | Solar | 40 | SWERA / provincial |
| BTN | Solar | 1.8102 | Historical B_Optimised_VRE limit; the atlas figure was 5 (Bhutan RE Master Plan ~12 GW) |
| MDV | Solar | 1 | Maldives Energy Roadmap |
| INDNO | Wind | 285 | NIWE 150 m (N incl. Rajasthan) |
| INDSO | Wind | 445 | NIWE 150 m (TN/KA/AP/TS) |
| INDWE | Wind | 410 | NIWE 150 m (GJ/MH/MP) |
| INDEA | Wind | 0 | Historical B_Optimised_VRE limit; the atlas figure was 8 (NIWE 150 m, OD/WB/BR/JH) |
| INDNE | Wind | 0 | Historical B_Optimised_VRE limit; the atlas figure was 0.5 (NIWE 150 m, NE negligible) |
| BGD | Wind | 3 ★ | NREL/SREDA (>30 GW gross, unrealistic) |
| LKA | Wind | 15 | ADB/UNDP 2050 (15,000 MW); ADB tech ~5,600 MW |
| NPL | Wind | 0 | Historical B_Optimised_VRE limit; the atlas figure was 2 (SWERA / provincial) |
| BTN | Wind | 0.5 | Bhutan RE Master Plan (~760 MW) |
| MDV | Wind | 0 ★ | Maldives Roadmap (~80 MW niche) |

### The applied values and the historical limits

Each ceiling in `south_asia.yaml` was chosen once as the lower of the atlas-based figure and the
historical `B_Optimised_VRE` limit. Four values therefore
hold the historical limit in place of the atlas figure: BTN solar 1.8102 GW (atlas figure 5),
INDEA wind 0 (atlas figure 8), INDNE wind 0 (atlas figure 0.5) and NPL wind 0 (atlas figure 2).
The three clips below are the cases where the atlas figure itself binds.

The pipeline applies every ceiling as a fixed value in all years of a derived scenario. It never
computes the lower of the previous workbook value and the ceiling, and a ceiling above the
previous value raises it. The annual investment bounds of the same technology are lowered to the
ceiling where they exceed it. The [scenario edits](../../scenario-edits.md) page describes the step.

### The three clips
- **LKA solar 16 ★ — well supported**, but it matches a 2050 *scenario build* (ADB/UNDP), not pure technical potential (ADB technical ≈ 6 GW).
- **BGD onshore wind 3 ★ — cite-gap.** Only literature figure is NREL/SREDA ">30 GW gross" (self-flagged unrealistic).
- **MDV onshore wind 0 ★ — defensible but not literally zero.** Roadmap + Rinaldi et al. (2019): onshore wind "excluded from the energy strategy", only ~80 MW niche (Greater Malé).

---

## 3. Transmission-loss (3%) provenance

- **Primary source:** Grid Controller of India (Grid-India/NLDC), "Applicable Transmission Losses" — weekly all-India ISTS loss notices under CERC (Sharing of ISTS Charges & Losses) Regulations 2020. https://posoco.in/en/side-menu-pages/applicable-transmission-losses/
- **Figures:** all-India ISTS loss **3.96%** (27 Apr–3 May 2026), **4.81%** (12–18 Jan 2026); excludes distribution & intra-state. CERC tariff decomposition ~4% attributable to losses.
- **Distinguish:** ISTS transmission-only ≈ **3.5–5%** · national transmission (CEA) ≈ 3.5–4% · **combined T&D ≈ 17–20%** (World Bank EG.ELC.LOSS.ZS) · **AT&C ≈ 15–17%** (PFC/MoP). The high figures are distribution-dominated — do NOT conflate.

### The current demand accounting

The demand of Bangladesh, Bhutan, the five India regions,
Nepal and Sri Lanka for 2023 to 2050 is an upstream requirement. The demand balance of the
model credits the actual internal transfer losses once, while the physical input and output
ratios of the internal transmission families stay as they are. Border losses receive no such
credit, and the Maldives is excluded. The 3% above is the physical loss of the internal
transmission families, and this accounting does not remove it. The sources are the demand
balance in `model/osemosys_fast_preprocessed.txt` and the row "Trade and upstream-demand
accounting" of the README sheet in `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx`.

## 4. WACC / discount rate (10%) provenance

- **Primary anchor:** IRENA, "Renewable Power Generation Costs in 2018" — "a real cost of capital of 7.5% in OECD countries and China, and **10% in the rest of the world**." https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2019/May/IRENA_Renewable-Power-Generations-Costs-in-2018.pdf
- **Range brackets:** OECD/IEA real cost-of-capital ~**4–9%** (up to ~18% stressed) https://one.oecd.org/document/ENV/WKP(2024)15/REV1/en/pdf ; World Bank "Demystifying the Costs..." used **6%** https://documents1.worldbank.org/curated/en/125521593437517815/pdf/Demystifying-the-Costs-of-Electricity-Generation-Technologies.pdf ; IRENA (2023) trends lower (→5% OECD / 7.5% RoW by 2021) https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2023/May/IRENA_The_cost_of_financing_renewable_power_2023.pdf
- In the model the 10% is a study discount convention, recorded in the row "Sets, defaults and empty inputs" of the README sheet in the scenario workbook. It applies to every technology. The sources above are benchmarks for that convention; they do not establish the financing cost of any one technology or source.

## 5. Storage cost basis

- The storage costs come from NREL, "Utility-Scale Battery Storage", **2024 ATB** v3, rows 20 (energy) and 26 (power), Moderate trajectories 2023 to 2050, in 2022 US dollars. https://atb.nrel.gov/electricity/2024/utility-scale_battery_storage ; projections: Cole & Karmakar, NREL 2023 update https://www.nrel.gov/docs/fy23osti/85332.pdf
- The 2022 values are converted once to constant 2023 US dollars with the DESNZ 2024 US GDP deflator ratio 131.902/127.224. The power component carries the node factors; the energy component is shared across nodes and converted to the model's energy unit.
- The fixed operating cost is 2.5% of a four-hour reference system per year, including augmentation. This is a power-based approximation, because the model chooses the storage duration.
- The generic long-duration storage is a study assumption with separate power and energy components. ATB does not establish its costs or lifetimes.
- Sources in the repository: the rows "Monetary basis and interpretation" and "Storage costs and service" of the README sheet in `inputs/scenarios/OSTRAM_Scenario_Inputs.xlsx`, and the cell notes of the storage cost rows on its `Secondary_Techs` sheet.

**NISE vintage note.** The India solar ceilings (INDEA 150, INDNE 25, INDNO 600, INDSO 500 and INDWE 650 GW) are round values benchmarked to the order of magnitude of NISE-2025. They are neither a precise aggregation of NISE state values nor an average of the 2014 and 2025 vintages.
