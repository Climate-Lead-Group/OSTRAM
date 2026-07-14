# OSTRAM — VRE-ceiling & key-assumption provenance / citations
### Deep-research verification pass, 2026-07-12 (companion to OSTRAM_METHODOLOGY.md)

**Scope & status.** Provenance for the study's load-bearing, previously-uncited assumptions — the per-node
VRE resource ceilings, the internal-transmission loss, the discount rate/WACC, storage cost — plus a
reference check of the D4 report's added citations. **No model value was changed by this pass**: it documents
and, where needed, corrects *provenance/wording*. All ceiling numbers remain exactly as solved. Cite-gaps are
flagged honestly. Primary/official sources preferred; URLs checkable.

---

## 1. VRE resource-potential ceilings — reference list

**Label correction (important):** the model's `vre_ceilings.csv` `atlas_source` column tags all solar
"NISE-2025" and all wind "NIWE 150 m", but **NISE and NIWE are MNRE-India bodies (India only)** — the five
non-India nodes (BGD, LKA, NPL, BTN, MDV) must be cited to their own national studies or a uniform atlas.

**India — solar (NISE):**
- NISE/MNRE state-wise solar potential (2014): **748.98 GWp** (3% of wasteland + Census 2011). https://mnre.gov.in/en/solar-overview/ · PIB https://www.pib.gov.in/Pressreleaseshare.aspx?PRID=2003561
- NISE/MNRE ground-mounted solar PV potential (2025): **3,343.37 GWp** (6.69% of feasible wasteland ~27,571 km²), released 23 Sep 2025. https://nise.res.in/wp-content/uploads/2025/09/Poster-and-Momento.pdf
- Model uses the **2025 vintage** (India solar ceilings sum ~1,925 GW — a subset of 3,343; impossible under 2014's 749).

**India — wind (NIWE):**
- NIWE India Wind Potential Atlas at 150 m agl: **1,163.9 GW** (695.50 GW at 120 m). https://maps.niwe.res.in/media/150m-report.pdf · map https://maps.niwe.res.in/resource_map/map/150m/ · https://mnre.gov.in/en/wind-overview/
- Model India-wind ceilings sum ~1,149 GW ≈ the NIWE 150 m national total (strong concordance).

**Non-India (per country):**
- **BANGLADESH** — SREDA/Power Division (NREL/USAID), National Solar Energy Roadmap 2021–2041 https://rise.esmap.org/sites/default/files/library/bangladesh/Renewable%20Energy/Bangladesh_National-Solar-Energy-Roadmap.pdf ; land-screened solar ~53 GW (2019). Wind: NREL/SREDA gross ">30,000 MW" over 20,000 km² — source-flagged as unfiltered/unrealistic.
- **SRI LANKA** — ADB & UNDP, "100% RE by 2050" (2017): 2050 mix ~15,000 MW wind + ~16,000 MW solar https://www.adb.org/sites/default/files/publication/354591/sri-lanka-power-2050v2.pdf ; ADB Energy Sector Assessment (solar technical ~6,000 MW; wind ~5,600 MW Mannar) https://www.adb.org/sites/default/files/institutional-document/547381/sri-lanka-energy-assessment-strategy-road-map.pdf ; SLSEA RE Resource Development Plan 2021–2026 https://www.energy.gov.lk/images/renewable-energy/renewable-energy-resource-development-plan-en.pdf
- **NEPAL** — SWERA GIS report (UNEP/NREL/Risø) https://policy.asiapacificenergy.org/node/358 ; provincial assessment, Renewable Energy 2021 https://doi.org/10.1016/j.renene.2021.08.109
- **BHUTAN** — Cowlin & Gilman, NREL/TP-6A2-46547 (2009) https://www.osti.gov/biblio/964607 ; RGoB Renewable Energy Master Plan 2017–2032 (~12,000 MW solar, ~760 MW wind).
- **MALDIVES** — Ministry of Environment, Energy Roadmap 2024–2033 (large-scale onshore wind ~80 MW niche, Greater Malé) https://www.environment.gov.mv/v2/wp-content/files/publications/20241107-pub-energy-roadmap-maldives-2024-2033-.pdf ; NREL Wind Atlas of Sri Lanka & Maldives (2003).

**Uniform multi-country option (recommended baseline for the 5 non-India nodes):**
- Global Solar Atlas (World Bank/ESMAP/Solargis) https://globalsolaratlas.info · Global Wind Atlas (World Bank/DTU) https://globalwindatlas.info

---

## 2. Per-node provenance table (model value vs best source; ★ = enforced clip)

| Node | Tech | Value (GW) | Best source | Support | Confidence |
|---|---|--:|---|---|---|
| INDNO | Solar | 600 | NISE-2025 (N states) | ✓ order-of-mag | partial |
| INDWE | Solar | 650 | NISE-2025 (Rajasthan/Gujarat/MH/MP) | ✓ direction | partial |
| INDSO | Solar | 500 | NISE-2025 (AP/KA/TN/TS/KL) | ✓ direction | partial |
| INDEA | Solar | 150 | NISE-2025 (Bihar/JH/OD/WB) | DIFFER on 2014; ✓ on 2025 | partial |
| INDNE | Solar | 25 | NISE-2025 (Assam+7) | ✓ | partial |
| BGD | Solar | 40 | SREDA / ~53 GW land-screened | ✓ (conservative) | verified (dir) |
| LKA | Solar | 16 ★ | ADB/UNDP 2050 (~16,000 MW) | ✓ near-exact | verified |
| NPL | Solar | 40 | SWERA / provincial | ✓ (very conservative) | verified (dir) |
| BTN | Solar | 5 | Bhutan RE Master Plan (~12 GW) | ✓ (conservative) | verified (dir) |
| MDV | Solar | 1 | Maldives Energy Roadmap | ✓ (island-constrained) | partial |
| INDNO | Wind | 285 | NIWE 150 m (N incl. Rajasthan) | ✓ direction | partial |
| INDSO | Wind | 445 | NIWE 150 m (TN/KA/AP/TS) | ✓ direction | partial |
| INDWE | Wind | 410 | NIWE 150 m (GJ/MH/MP) | ✓ direction | partial |
| INDEA | Wind | 8 | NIWE 150 m (OD/WB/BR/JH) | ✓ (small) | partial |
| INDNE | Wind | 0.5 | NIWE 150 m (NE negligible) | ✓ | partial |
| BGD | Wind | 3 ★ | NREL/SREDA (>30 GW gross, unrealistic) | **cite-gap** — modeler screen | flag |
| LKA | Wind | 15 | ADB/UNDP 2050 (15,000 MW); ADB tech ~5,600 MW | ✓ vs 2050 scenario | partial |
| NPL | Wind | 2 | SWERA / provincial | ✓ direction | partial |
| BTN | Wind | 0.5 | Bhutan RE Master Plan (~760 MW) | ✓ (conservative) | verified (dir) |
| MDV | Wind | 0 ★ | Maldives Roadmap (~80 MW niche) | **flag** — ≈0 defensible, not literally 0 | flag |

### The three clips
- **LKA solar 16 ★ — well supported**, but it matches a 2050 *scenario build* (ADB/UNDP), not pure technical potential (ADB technical ≈ 6 GW). State which concept the ceiling represents.
- **BGD onshore wind 3 ★ — cite-gap.** Only literature figure is NREL/SREDA ">30 GW gross" (self-flagged unrealistic). Document 3 GW as a modeler's conservative land-screen; cite NREL/SREDA for the gross-vs-realistic caveat.
- **MDV onshore wind 0 ★ — defensible but not literally zero.** Roadmap + Rinaldi et al. (2019): onshore wind "excluded from the energy strategy", only ~80 MW niche (Greater Malé). Cite as "≈0 at grid scale".

### Citation strategy (recommendation only — values unchanged)
Lead with the **Global Solar/Wind Atlas** as the uniform baseline for the 5 non-India nodes; keep **NISE-2025 (solar) + NIWE 150 m (wind)** for India; add the specific national source on each of the three clips (ADB/UNDP → LKA; NREL/SREDA → BGD; Maldives Roadmap → MDV). "Global-atlas baseline + national-study clips" is more audit-friendly than ad-hoc national studies.

---

## 3. Transmission-loss (3%) provenance — recast required

- **Primary source:** Grid Controller of India (Grid-India/NLDC), "Applicable Transmission Losses" — weekly all-India ISTS loss notices under CERC (Sharing of ISTS Charges & Losses) Regulations 2020. https://posoco.in/en/side-menu-pages/applicable-transmission-losses/
- **Figures:** all-India ISTS loss **3.96%** (27 Apr–3 May 2026), **4.81%** (12–18 Jan 2026); excludes distribution & intra-state. CERC tariff decomposition ~4% attributable to losses.
- **Distinguish:** ISTS transmission-only ≈ **3.5–5%** · national transmission (CEA) ≈ 3.5–4% · **combined T&D ≈ 17–20%** (World Bank EG.ELC.LOSS.ZS) · **AT&C ≈ 15–17%** (PFC/MoP). The high figures are distribution-dominated — do NOT conflate.
- **Action:** our flat **3% is at the low edge**; cite as "~3–4%" (Grid-India). **Drop / heavily qualify "single largest driver of system cost"** — transmission losses are small vs distribution and vs generation capital+fuel; that phrase is an over-claim (it appears in the D4 report; the methodology doc's "+4.1% on A, dwarfs the interconnector/D5 effects" is a correctly-scoped model result and stays). Note: our 3% is applied to *intra-node* internal-tx; the ISTS figure is *inter-state* — same order, slightly different layer.

## 4. WACC / discount rate (10%) provenance

- **Primary anchor:** IRENA, "Renewable Power Generation Costs in 2018" — "a real cost of capital of 7.5% in OECD countries and China, and **10% in the rest of the world**." https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2019/May/IRENA_Renewable-Power-Generations-Costs-in-2018.pdf
- **Range brackets:** OECD/IEA real cost-of-capital ~**4–9%** (up to ~18% stressed) https://one.oecd.org/document/ENV/WKP(2024)15/REV1/en/pdf ; World Bank "Demystifying the Costs..." used **6%** https://documents1.worldbank.org/curated/en/125521593437517815/pdf/Demystifying-the-Costs-of-Electricity-Generation-Technologies.pdf ; IRENA (2023) trends lower (→5% OECD / 7.5% RoW by 2021) https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2023/May/IRENA_The_cost_of_financing_renewable_power_2023.pdf
- **Verdict:** 10% real is defensible as IRENA's rest-of-world default, at the **conservative/higher** end of current estimates. No SA-specific single "10%" source exists. Country-level dataset for sensitivity: "Historical and future projected costs of capital..." https://www.ncbi.nlm.nih.gov/pmc/articles/PMC12678392/

## 5. Storage cost basis

- NREL, "Utility-Scale Battery Storage," **2024 ATB** (2022 USD; Li-ion; 2–10 h; 85% RTE; FOM 2.5% capex). https://atb.nrel.gov/electricity/2024/utility-scale_battery_storage ; projections: Cole & Karmakar, NREL 2023 update https://www.nrel.gov/docs/fy23osti/85332.pdf . Corroborate with BNEF/IRENA.

---

## 6. D4 report reference verification & corrections

- **[25]** Timilsina & Toman, "Potential gains from expanding regional electricity trade in South Asia," *Energy Economics* 60:6–14, 2016. DOI 10.1016/j.eneco.2016.08.023. ✓ peer-reviewed.
- **[26]** Timilsina, Toman, Karacsonyi & de Tena Diego, "How much could South Asia benefit from regional electricity cooperation and trade?" World Bank PRWP 7341, 2015. SSRN 2623783 / RePEc wbk:wbrwps:7341. ✓ **Add middle initials: Toman, Michael A.; Karacsonyi, Jorge G.** Note [25] is the journal version of [26] (same study).
- **[27]** Koirala & Rahut, "Reimagining South Asia's electricity system amid growing energy market volatility," *Asia Pathways* (ADBI blog), Jul 2022. https://www.asiapathways-adbi.org/2022/07/reimagining-south-asias-electricity-system-amid-growing-energy-market-volatility/ — **NOT peer-reviewed; re-tag as grey-literature/commentary**; lean on Timilsina & Toman (2016) for any quantitative claim.

**Additional verified CBET studies (directional corroboration only — never compare magnitudes):**
- Wijayatunga, Chattopadhyay & Fernando, "Cross-Border Power Trading in South Asia: A Techno-Economic Rationale," ADB SA WP-38, 2015. https://www.adb.org/sites/default/files/publication/173198/south-asia-wp-038.pdf
- Singh, Jamasb, Nepal & Toman, "Electricity cooperation in South Asia: Barriers to cross-border trade," *Energy Policy* 120:741–748, 2018. DOI 10.1016/j.enpol.2017.12.048
- Timilsina & Toman, "Carbon Pricing and CBET for Climate Change Mitigation in South Asia," *EEEP* 7(2), 2018. RePEc aen:eeepjl:eeep7-2-timilsina
- Timilsina, "Regional electricity trade for hydropower development in South Asia," *IJWRD* 37(3):392–410, 2021. DOI 10.1080/07900627.2018.1515065
- Aryal & Dhakal, "Medium-term assessment of cross-border trading potential of Nepal's renewable energy... TIMES," *Energy Policy* 168:113109, 2022. DOI 10.1016/j.enpol.2022.113109
- SARI/EI (IRADe/USAID), "BIMSTEC Energy Outlook 2035" (2021) https://irade.org/Bimstec%20Final%20Report%20Energy%20Outlook-2035.pdf
- (EU analogue) Newbery, Strbac & Viehoff, "The benefits of integrating European electricity markets," *Energy Policy* 94:253–263, 2016. DOI 10.1016/j.enpol.2016.03.047
- **Could not verify:** a standalone peer-reviewed "Rahman/Fernando" SA CBET journal article — cite the ADB report instead.

## 7. Cite-gaps & caveats (state, don't paper over)
1. **BGD onshore wind = 3 GW** — no direct source (only gross-unrealistic >30 GW); modeler's conservative screen. ★clip.
2. **MDV onshore wind ≈ 0 GW** — defensible qualitatively; ~80 MW niche exists; not literally zero. ★clip.
3. **LKA solar = 16 GW** — well-matched but a *scenario-build* figure, not technical potential (~6 GW). ★clip.
4. **India solar ceilings** reconcile only with the **NISE-2025** vintage; the exact state→region aggregation is order-of-magnitude, not line-item verified.
5. **"3% loss = single largest driver of system cost"** — unsupported over-claim; recast to ~3–4% (Grid-India) and remove the superlative.
6. **WACC 10%** — no SA-specific single source; defended via IRENA rest-of-world (conservative).
7. All cross-study comparisons are **directional only** (different horizons, discount rates, resolutions). CBET → lower cost / higher RE share is robustly corroborated in *direction*.

*Provenance pass 2026-07-12. No ceiling or model value changed. Companion to `OSTRAM_METHODOLOGY.md` §3 and the gaps in §8.*

---

**NISE vintage note (added 2026-07-12).** The India solar ceilings (150 / 25 / 600 / 500 / 650 GW) are deliberately-round **non-binding headroom guards** — set well above B_Opt's actual buildout (e.g. INDNO 600 vs buildout 319; INDWE 650 vs 185; INDEA 150 vs 19) and benchmarked to the **NISE-2025 order of magnitude**, *not* a precise NISE state-aggregation and *not* an average of the 2014/2025 vintages (the round hundreds are the tell; a real aggregation gives messy decimals). Because India-solar buildout sits far below these caps, their exact value **does not affect any result** — only the three *binding* non-India clips (LKA solar 16, BGD wind 3, MDV wind 0) require precise sourcing. So the "which NISE vintage?" question is immaterial for the India nodes.
