# Maldives Demand Timeslice Profile — Status & Data Assessment

**Country:** Maldives  
**Model:** OSTRAM (South Asia Power System Connectivity Model)  
**Date:** 2026-03-11  
**Status:** v1 profile available (yearbook-based); higher-resolution STELCO dashboard data exists but assessed as low priority

---

## Current deliverable

The file `Maldives_new_data.xlsx` contains compiled data for Maldives covering capacity factors, demand forecasts, historical production, and seasonality. The data was assembled from the following sources:

- **Capacity factors:** IRENA Energy Profile — Maldives (updated Sep 2025), reference year 2023. Covers diesel (26%), solar PV (18%), wind (20%), bioenergy (1%).
- **Demand forecasts:** MCCEE Energy Road Map 2024–2033 (ADB, Nov 2024). Includes national consumption projections, segment-level peak demand (Greater Malé, other inhabited islands, resort islands), and RE targets (33% by end-2028).
- **Historical production:** IRENA national generation series (2018–2023) and STELCO Statistical Yearbook Table 12.1 — sectoral electricity utilisation in Malé (2010–2024), broken out by residential, manufacturing & commercial, government buildings, and public places & schools.
- **Seasonality:** STELCO Statistical Yearbook Table 12.2 (monthly electricity production by locality, 2024) and Table 12.3 (monthly fuel consumption by locality, 2024), plus long-term climate data (GHI, temperature, rainfall, wind, humidity) from Solargis/World Bank and CRU/Met Office, aligned to monsoon seasons.

**Primary source:** [Statistical Yearbook of Maldives 2023](https://statisticsmaldives.gov.mv/yearbook/2023/) and STELCO Statistical Yearbook tables (2024 edition).

---

## What the yearbook data provides for OSTRAM

### Seasonal demand fractions
The Seasonality sheet contains monthly production for **all STELCO-served localities** (approximately 60 islands), tagged by region. Monthly totals can be aggregated into OSTRAM seasons to derive demand fractions. The data reveals a clear seasonal pattern: production peaks during the SW Monsoon (May–Oct) when higher humidity drives cooling demand, and is lower during the NE Monsoon / dry season (Nov–Apr). This is well-resolved at the monthly level.

### Technology capacity factors by season
The climate seasonality data (Section C of the Seasonality sheet) provides the physical basis for technology-level CF variation:

- **Solar PV:** GHI ranges from ~4.7 kWh/m²/day (Nov) to ~6.2 kWh/m²/day (Mar). NE Monsoon months produce ~20–25% more solar output than SW Monsoon months due to reduced cloud cover. Monthly GHI values can be mapped directly to seasonal PV capacity factors.
- **Diesel:** Fuel consumption data by locality and month (Table 12.3) can be used to derive effective diesel CFs per season, though diesel operates as baseload/peaking across all seasons.
- **Wind:** Mean wind speeds range from ~3.1 m/s (Apr) to ~5.8 m/s (Jun). SW Monsoon brings best wind resource, but annual average (~4.2 m/s at 10m) is below typical utility-scale thresholds.

### Demand growth trajectory
Historical sectoral data for Malé (2010–2024) shows steady growth at ~5%/yr. The MCCEE Road Map provides segment-level forecasts to 2028. However, there is some tension between the national 5% base case and the Greater Malé segment forecast (~7.9%/yr), which may need reconciliation for the model's demand projection.

---

## What the yearbook data does NOT provide

### Intra-day (hourly) load shape
The yearbook provides monthly totals only — there is no breakdown by hour of day. This means **daypart splits (e.g., Day/Night/Peak) must be estimated**, not computed from data. This is the main gap.

### Generation by source at sub-annual resolution
National generation by fuel type is available only as annual totals (IRENA series). The monthly production data from STELCO Table 12.2 is aggregate (all sources combined per locality), not broken out by diesel vs. solar vs. other. Monthly fuel consumption (Table 12.3) provides a diesel proxy but does not isolate RE generation by month.

---

## Additional data source: STELCO real-time dashboard

STELCO operates a generation dashboard at `https://stelco.com.mv/generation` that displays sub-hourly demand data (approximately 15-minute resolution) for Malé & Hulhumalé, overlaid with weather data from the Open-Meteo API.

### What the dashboard offers
- Demand timeseries at sub-hourly resolution (tested range: Jan 2025 – Mar 2026)
- Weather overlay (temperature, humidity, wind, precipitation)
- Peak/minimum demand statistics (e.g., peak ~111 MW, minimum ~66 MW for the tested period)

### Why it is not being pursued

1. **Limited geographic scope.** The dashboard covers only Malé & Hulhumalé. The yearbook data covers all ~60 STELCO-served localities nationally, which is what OSTRAM needs for a country-level node.

2. **Marginal incremental value.** The main gap the dashboard would fill is intra-day load shape. However, Maldives has near-constant equatorial temperatures (seasonal range ~1.3°C) and no significant industrial load variation, so the diurnal demand pattern is relatively stable year-round. A generic tropical island diurnal profile from literature is a reasonable proxy for daypart splits.

3. **Data extraction difficulty.** The site returns HTTP 403 to programmatic requests (web scrapers, API fetches). Extracting the data would require either:
   - Reverse-engineering the frontend API calls via browser developer tools and replaying them with session authentication
   - Browser automation (Selenium/Playwright) to interact with the date-range selector and capture responses
   
   This is feasible but labour-intensive for data that covers only one locality.

4. **No generation-by-source breakdown.** The dashboard appears to show aggregate demand only, not generation dispatched by fuel type. It would not resolve the technology-level CF gap.

### Recommendation
**Park the STELCO dashboard scraping as a "nice to have."** If higher-resolution intra-day profiles become critical later (e.g., for battery storage sizing or detailed VRE integration analysis), revisit the dashboard extraction. For OSTRAM's timeslice parameterization, the yearbook data is sufficient.

---

## Using the yearbook data for OSTRAM timeslice construction

The following approach is recommended for building Maldives timeslice parameters from the available data:

| Parameter | Method | Confidence |
|-----------|--------|------------|
| Seasonal demand fractions | Aggregate monthly production (Table 12.2) into OSTRAM seasons | **High** — based on metered production data for all localities |
| Daypart demand splits | Estimate from literature (tropical island diurnal profiles) or adopt a standard 40/35/25 Day/Evening/Night split, adjusted for Maldives context | **Medium** — estimated, not measured |
| Solar PV CF by season | Derive from monthly GHI values (Solargis long-term averages) | **High** — well-established irradiation dataset |
| Diesel CF by season | Derive from monthly fuel consumption (Table 12.3) relative to installed capacity | **Medium–High** — fuel data is metered; capacity assumptions introduce some uncertainty |
| Wind CF by season | Derive from monthly mean wind speeds and published power curves | **Low–Medium** — wind deployment is minimal; CF is indicative only |
| YearSplit | Calculate from number of hours per OSTRAM season × daypart | **High** — arithmetic |

---

## Data scarcity rating

| Dimension | Rating | Notes |
|-----------|--------|-------|
| Annual demand & generation totals | ★★★★★ | IRENA + STELCO yearbook, verified and consistent |
| Sectoral demand breakdown | ★★★★☆ | Malé only (Table 12.1); no sectoral split for outer islands |
| Monthly production by locality | ★★★★★ | 60+ localities, 12 months, metered data |
| Monthly fuel consumption | ★★★★☆ | Available by locality; enables diesel efficiency estimates |
| Seasonal climate parameters | ★★★★★ | GHI, wind, temperature, rainfall — well-sourced, long-term averages |
| Intra-day (hourly) load shape | ★★☆☆☆ | Not available from yearbook; STELCO dashboard exists but limited scope and hard to extract |
| Generation by source (sub-annual) | ★★☆☆☆ | Annual totals only by source; monthly data is aggregate |
| Demand forecasts | ★★★★☆ | MCCEE Road Map provides segment-level forecasts to 2028; some internal tension between segments |
| RE pipeline / planned capacity | ★★★☆☆ | Government 33% RE target announced; project-level pipeline not yet compiled |

---

## Action item summary

| # | Action | Owner | Priority |
|---|--------|-------|----------|
| 1 | Build OSTRAM timeslice demand fractions from monthly production data (Table 12.2) | CLG modeling team | High |
| 2 | Select and document daypart split assumptions (literature-based tropical island profile) | CLG modeling team | High |
| 3 | Compute seasonal solar PV CFs from monthly GHI data | CLG modeling team | High |
| 4 | Derive seasonal diesel CFs from fuel consumption data (Table 12.3) and installed capacity | CLG modeling team | Medium |
| 5 | Reconcile demand growth trajectory (national 5% vs. Greater Malé 7.9% segment forecast) | CLG modeling team | Medium |
| 6 | Compile RE project pipeline for NDC scenario (solar tenders, WTE projects, battery storage) | Maldives country lead | Medium |
| 7 | If intra-day resolution becomes critical, revisit STELCO dashboard extraction | CLG modeling team | Low |
