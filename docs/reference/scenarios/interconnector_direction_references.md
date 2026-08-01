# Interconnector Direction — Contractual Reference & Justification

**Scenario:** `B_Opt_DirContractual`
**Mechanism:** `set_interconnector_direction.py` (zeros the disabled mode's Input/Output
ActivityRatio in `A-O_AR_Projections.xlsx` + base-year values; `Projection.Mode = "User defined"`).
**Convention:** corridor code `TRN<SRC><DST>`; `forward` = SRC→DST, `reverse` = DST→SRC,
`bidirectional` = both modes active (omit from config).

Each corridor's direction is set to its **real-world governing flow**, established from primary
sources (PPAs, commissioned assets, official plans) via three research passes (2026-07). Genuinely
seasonal corridors are left bidirectional; conceptual corridors are set to the only physically
plausible direction (blocking impossible export legs) and are inactive in B_Opt regardless.

| # | Corridor | Setting | Direction kept | Basis | Confidence |
|---|----------|---------|----------------|-------|------------|
| 1 | TRNBGDXXINDEA | reverse | India (ER) → Bangladesh | Adani Godda 1,600 MW dedicated radial line, 25-yr PPA with BPDB (signed 2017), 100% exported to BD; Bheramara–Baharampur HVDC 2×500 MW operating since 2013/2018. BD has never exported this way. | High |
| 2 | TRNBGDXXINDNE | reverse | India (NER) → Bangladesh | Surjyamaninagar (Tripura)–Comilla 400 kV delivering 160 MW India→BD since 2016 (Palatana/ONGC). | High |
| 3 | TRNBTNXXINDEA | forward | Bhutan → India (ER) | Tala (1,020), Chukha (336), Mangdechhu (720 MW) evacuate to West Bengal; ~70% of Bhutan generation exported to India. Minor dry-winter reversal only. | High |
| 4 | TRNBTNXXINDNE | forward | Bhutan → India (NER) | Kurichhu (60 MW)–Gelephu–Salakati (Assam) 132 kV; same wet-export pattern. | High |
| 5 | TRNBTNXXBGDXX | forward | Bhutan → Bangladesh | Dorjilung 1,125 MW trilateral (Bhutan exporter, BD buyer per IEPMP 2023). Planned/pre-construction; direction of intent unambiguous. | Med |
| 6 | TRNNPLXXBGDXX | forward | Nepal → Bangladesh | Tripartite deal (NEA/NVVN/BPDB) signed 3 Oct 2024: 40 MW Nepal→BD wheeled via India, first flow 15 Nov 2024. | High |
| 7 | TRNINDSOLKAXX | forward | India → Sri Lanka | Madurai–Mannar ±320 kV 1,000 MW VSC HVDC; designed bidirectional but CEB LTGEP 2025–2044 models it as an SL **import** link (in-service ~2034). SL→India export is aspirational and **uncontracted** — flagship enabler (Adani Mannar/Pooneryn wind) cancelled 13 Feb 2025. Forward encodes the real import-dominant flow and blocks the uncontracted reverse. | High |
| 8 | TRNMDVXXINDSO | reverse | India → Maldives | **Purely conceptual**: only a *proposed* (never signed) 2022 OSOWOG MoU, downgraded to a "feasibility study to explore participation" (Oct 2024); absent from India's Apr 2026 cross-border roster. Maldives is ~100% import-dependent, diesel-dominated, deficit — **zero export surplus**, so only India→MDV is physically plausible. Reverse blocks the impossible MDV→India export leg. Inactive in B_Opt. | High |
| 9 | TRNLKAXXMDVXX | forward | Sri Lanka → Maldives | **No electricity project exists** (the only SL–MDV submarine cable is telecom/fibre, MSC 2021). Direction **inferred** from Maldives' deficit-only status (no export surplus) — if ever built, SL→MDV. Inactive in B_Opt. | High (that no project exists); direction inferred |
| 10 | TRNINDEANPLXX | bidirectional | Nepal ↔ India (ER) | Genuinely seasonal: Nepal exports run-of-river hydro in monsoon (~Jun–Nov), imports in dry winter (~Dec–May). Net importer FY2022/23 → net exporter FY2024/25. Forcing one direction would misrepresent it. | High |
| 11 | TRNINDNONPLXX | bidirectional | Nepal ↔ India (NR) | Same seasonal reversal; Butwal–Gorakhpur 400 kV under construction (target 2026–30). | High |

## Citations (URLs)

1. **TRNBGDXXINDEA** — ADB, 2013 (Bheramara HVDC): https://www.adb.org/news/india-electricity-flows-bangladesh-first-south-asian-hvdc-cross-border-link · Adani press release, 2023: https://www.adani.com/newsroom/media-releases · Global Energy Monitor: https://www.gem.wiki/Adani_Godda_power_station
2. **TRNBGDXXINDNE** — SASEC, 2015/16 (Tripura→BD): https://sasec.asia/index.php?nid=138 · The Business Standard, 2024: https://www.tbsnews.net/bangladesh/energy/tripura-trims-power-export-bangladesh-over-dues-860296
3. **TRNBTNXXINDEA** — Energy in Bhutan (Wikipedia): https://en.wikipedia.org/wiki/Energy_in_Bhutan · MFA Bhutan (hydropower relations): https://www.mfa.gov.bt/rbedelhi/bhutan-india-relations/bhutan-india-hydropower-relations/ · iRADe working paper: https://www.irade.org/Bhutan%20Working%20Paper.pdf
4. **TRNBTNXXINDNE** — USEA / Bhutan Power Corp: https://usea.org/sites/default/files/event-/Bhutan%20Power%20Corporation.pdf · BPC Power Data Book 2023: https://www.bpc.bt/wp-content/themes/bpc/assets/downloads/Power%20Data%20Book%202023.pdf
5. **TRNBTNXXBGDXX** — The Bhutanese (Dorjilung trilateral): https://thebhutanese.bt/bangladesh-bhutan-and-india-agree-to-1024-mw-dorjilung-project/ · Tata Power–DGPC, 2025: https://www.tatapower.com/news-and-media/media-releases · World Bank, 2026: https://www.worldbank.org/en/news/press-release/2026/05/05/
6. **TRNNPLXXBGDXX** — Kathmandu Post, 2024 (trilateral deal): https://kathmandupost.com/national/2024/10/03/nepal-india-and-bangladesh-sign-trilateral-electricity-trade-deal · The Diplomat, 2024: https://thediplomat.com/2024/11/with-start-of-trilateral-hydropower-trade-south-asia-begins-historic-cooperation/
7. **TRNINDSOLKAXX** — SolarQuarter (Apr 2025 MoU): https://solarquarter.com/2025/04/11/india-sri-lanka-hvdc-interconnection-mou-a-strategic-power-link-in-south-asia/ · India MEA transcript, 5 Apr 2025: https://www.hcicolombo.gov.in/section/speeches-and-interviews/ · India–SL joint vision (bidirectional language), 21 Jul 2023 (SL MFA): https://mfa.gov.lk/en/india-sri-lanka-economic-partnership-vision/ · Daily FT (import-first), 16 Jun 2025: https://www.ft.lk/front-page/Sri-Lanka-and-India-confirm-technical-parameters-for-power-grid-interconnection/44-778103 · CEB LTGEP 2025–2044: https://www.ceb.lk/front_img/img_reports/1748839124LTGEP-2025-2044-FINAL_c.pdf · Adani wind withdrawal, 13 Feb 2025 (Business Standard): https://www.business-standard.com/companies/news/adani-green-withdraws-from-wind-energy-transmission-projects-in-sri-lanka-125021300885_1.html · India Ministry of Power (VSC HVDC design): https://powermin.gov.in/en/content/interconnection-neighbouring-countries
8. **TRNMDVXXINDSO** — PIB India, 2022 (proposed MoU): https://www.pib.gov.in/PressReleaseIframePage.aspx?PRID=1820208 · Maldives President's Office, 7 Oct 2024 ("feasibility study"): https://presidency.gov.mv/Press/Article/31823 · SolarQuarter, Apr 2026 (India roster excludes Maldives): https://solarquarter.com/2026/04/03/india-strengthens-regional-power-ties-with-cross-border-projects-and-strategic-mous/ · World Bank, 2023 (Maldives energy): https://blogs.worldbank.org/en/endpovertyinsouthasia/why-maldives-5-mw-solar-project-game-changer
9. **TRNLKAXXMDVXX** — Submarine Networks (MSC = telecom, not power): https://www.submarinenetworks.com/en/systems/intra-asia/msc · (direction inferred from Maldives deficit-only status; see #8 sources)
10. **TRNINDEANPLXX** — Kathmandu Post, 2023 (net importer): https://kathmandupost.com/national/2023/08/29/nepal-continues-to-be-net-importer-of-electricity · Kathmandu Post, 2024 (rising exports): https://kathmandupost.com/money/2024/08/20/india-raises-imports-of-nepal-s-energy-to-nearly-1-000-mw
11. **TRNINDNONPLXX** — POWERGRID (Butwal–Gorakhpur): https://www.powergrid.in/en/node/3146 · Kathmandu Post, 2025 (400 kV lines): https://kathmandupost.com/money/2025/10/29/nepal-india-sign-deal-to-build-two-400kv-transmission-lines

## Caveats

- **Conceptual corridors (#8, #9):** No committed project exists in any direction; the setting encodes
  the only physically plausible flow (deficit island imports) and blocks the impossible export leg.
  Both are inactive (0 build) in B_Opt, so the setting is faithful but non-binding.
- **India↔Sri Lanka (#7):** Physically/contractually designed *bidirectional*; `forward` encodes the
  documented near/mid-term import-dominant reality and blocks the uncontracted reverse (consistent with
  the Bangladesh treatment). Also pre-construction (in-service ~2034) — a timing fact, out of scope here.
- **Node mapping (#10 vs #11):** The flagship Dhalkebar–Muzaffarpur line physically lands in Bihar =
  India's *Eastern* region, so it arguably belongs to `TRNINDEANPLXX`; `TRNINDNONPLXX` (Butwal–Gorakhpur)
  is the true Northern link. Both are bidirectional, so this does not affect the direction map — noted
  for a possible separate node-mapping review.

_Generated 2026-07 from three primary-source research passes; every non-inferred direction is cited above._
