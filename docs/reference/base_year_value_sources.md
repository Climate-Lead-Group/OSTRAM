# Where the base-year values come from

The table `config/scenarios/rules/pwr_min_2023_2026_pin.csv` sets 2,644 values: generation limits, capacity factors, stocks and investment limits for 2023 to 2026, and residual capacity, corridor capacity and three required builds up to 2050. The base-year rule writes these values for `A_Calibrated_BAU`, `B_Optimised_VRE` and `C_Target_VRE`, and again after a derived scenario's edits; [Scenario rules](../rules.md) describes it. To find a value, open its country, then the block for its technology, parameter and years. The block's label is the row's `authority_lineage_class`, and the block names the documents, the derivation and the assumptions behind the value.

## Assumptions used across the table

Capacity factors use four seasons, S1 January, February and December (90 days), S2 March to May (92), S3 June to September (122) and S4 October and November (61), and five day parts, 00:00 to 06:00, 06:00 to 17:00, 17:00 to 20:00, 20:00 to 22:00 and 22:00 to 24:00. Generation in GWh converts to PJ at 0.0036 PJ/GWh.

**A1.** In 2023 to 2025, each technology's annual generation must lie between 98% and 102% of its published generation. The 2% band allows for rounding and reporting differences in the published figures; it is a calibration tolerance, not a measured uncertainty.

**A2.** 2026 annual generation is not yet published, so each technology must generate at least 95% of its 2025 figure, with no upper limit.

**A3.** Where a country publishes by fiscal year, each model year takes the fiscal year ending in it. Bangladesh 2023 uses the mean of the two fiscal years it overlaps.

**A4.** A plant commissioned during a year counts for the share of the year it was in service.

**A5.** Each hydro capacity factor is observed seasonal generation divided by installed capacity, season hours and an availability: 0.95 for large hydro, 0.92 for India South and India West large hydro, and 0.9 for small hydro and for Bangladesh. The model multiplies capacity and availability back in its annual limit, so existing plants can generate the published amount, except in Bangladesh, where the model's availability is 1 and plants can generate up to 11% more than observed from 2025. Output is the same at all times of day within a season, except in Bangladesh, which follows PGCB's half-hourly hydro curve, and Sri Lanka, which follows a fixed daily shape.

**A6.** 2026 hydrology and weather are not yet published, so 2026 capacity factors repeat 2025. The model year has 365 days.

**A7.** India North and South solar follow the average daily shape from Renewables.ninja at one point per region (Delhi for North); the point can be changed.

**A8.** Each region's thermal generation (CEA) is split between coal and oil in proportion to installed capacity times a utilisation of 0.65 for coal and 0.25 for oil. Oil receives 0% in East, North and West, 1.8% in North-East and 0.35% in South, so the split moves less than 2% of any region's thermal total.

**A9.** NEA publishes national generation, not generation by technology. Each year's total is split across hydro, solar, bagasse and thermal by day-weighted in-service capacity times a typical capacity factor (hydro 0.45, solar 0.1925, bagasse 0.17, thermal 0.25), each share capped at its physical maximum, which gives hydro about 97.5%. Because the split is an allocation, hydro, solar and bagasse have only a lower limit of 98% of their share, and thermal has no limit.

**A10.** Waste-to-energy in India East and North-East has a minimum generation of zero in 2023 to 2026, because CEA's generation categories cannot be matched to the plants in these two regions. Zero means no required minimum, not zero output.

**A11.** Existing plants with no published retirement date stay in service to 2050. Solar plants retire after 25 years.

**A12.** Plants commissioned in 2023 to 2026 are already in existing capacity at their published dates, so the model may not build new capacity in Bangladesh and Bhutan in 2023 to 2026, or in Nepal in 2023 to 2025. Where the model must build named capacity in those years, that capacity is subtracted from the published stock.

**A13.** States are grouped into the model's five Indian regions as in CEA's regional grids. MNRE's off-grid waste-to-energy is counted with grid-connected waste-to-energy.

**A14.** Corridor capacity is CEA's sum of inter-regional line ratings, with HVDC links at their reverse-direction rating; CEA notes this sum is not the actual transfer capability. Capacity stays constant after the last listed addition, to 2050.

**A15.** Technologies with no published annual generation for their model region in 2023 to 2026 have no annual generation limit; the model sets their output within their capacity.

**A16.** In 2025, Bangladesh generation by fuel has a lower limit of 98% of BPDB's FY2024-25 figure and no upper limit, as in 2026. Calendar-2025 demand, 105,155 GWh, is 3.9% above BPDB's FY2024-25 net generation including imports, 101,187 GWh, because the fiscal year runs from July 2024 to June 2025 while demand grows 6.8% from 2024 to 2025; upper limits at 102% of the fiscal figures could not meet calendar demand.

## Bangladesh

### `PWRCOABGDXX`, `PWRURNBGDXX` TotalAnnualMaxCapacityInvestment, 2026

- Values: 1.2 to 1.32 (2 distinct values); GW; 2 rows.
- Label: `BASE_YEAR_CAPACITY_FREEZE`.
- Sources: [R1] About RNPL; [R2] section Rooppur.
- Derivation: Each 2026 build cap is the capacity of the plant that reaches commercial operation in 2026: Patuakhali coal, 2 x 660 = 1,320 MW = 1.32 GW, and Rooppur unit 1 nuclear, 1,200 MW = 1.2 GW.
- Assumptions: A12. Bangladesh may build 1.32 GW of coal (Patuakhali) and 1.2 GW of nuclear (Rooppur unit 1) in 2026, the calendar year of their commercial operation dates. Patuakhali generated on test from January and April 2025 and reached initial operation in September 2025; it enters in 2026 because Bangladesh's 2025 generation follows BPDB's FY2024-25 figures, which include at most its test-run output.

### `PWRCOABGDXX`, `PWRHYDBGDXX`, `PWRNGSBGDXX`, `PWROILBGDXX` TotalTechnologyAnnualActivityLowerLimit, 2023

- Values: 2.53134 to 163.332288 (4 distinct values); PJ per year; 4 rows.
- Label: `BPDB_FY2022_23_FY2023_24_MEAN`.
- Sources: [R3] and [R4] Energy Generation by Fuel Type, printed p.14 (PDF p.15).
- Derivation: The 2023 reference is the mean of BPDB's FY2022-23 and FY2023-24 generation, and the lower limit is 0.98 x GWh x 0.0036 PJ/GWh. Oil is furnace oil plus HSD in each year: (20,650 + 11,797) / 2 = 16,223.5 GWh. Example: gas, (46,013 + 46,579) / 2 = 46,296 GWh x 0.0036 x 0.98 = 163.332288 PJ per year.
- Assumptions: A1, A3.

### `PWRCOABGDXX`, `PWRHYDBGDXX`, `PWRNGSBGDXX`, `PWROILBGDXX`, `PWRSPVBGDXX`, `PWRURNBGDXX`, `PWRWONBGDXX` TotalTechnologyAnnualActivityLowerLimit, 2024 to 2026

- Values: 0 to 164.330712 (19 distinct values); PJ per year; 20 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R4] and [R5] Energy Generation by Fuel Type, printed p.14 (PDF p.15).
- Derivation: BPDB's generation by fuel for FY2023-24 (2024) and FY2024-25 (2025), in GWh, x 0.0036 PJ/GWh; the lower limit is 0.98 x PJ in 2024 and 2025 and 0.95 x the 2025 PJ in 2026. Oil is furnace oil plus HSD (2024: 11,357 + 440 = 11,797 GWh; 2025: furnace oil 10,860 GWh, the HSD line of -10 GWh being auxiliary use), and nuclear is 0 because BPDB prints no nuclear generation. Example: coal 2024, 19,138 GWh x 0.0036 x 0.98 = 67.518864 PJ per year.
- Assumptions: A1, A2, A3.

### `PWRCOABGDXX`, `PWRHYDBGDXX`, `PWRNGSBGDXX`, `PWROILBGDXX` TotalTechnologyAnnualActivityUpperLimit, 2023

- Values: 2.63466 to 169.998912 (4 distinct values); PJ per year; 4 rows.
- Label: `BPDB_FY2022_23_FY2023_24_MEAN`.
- Sources: [R3] and [R4] Energy Generation by Fuel Type, printed p.14 (PDF p.15).
- Derivation: The 2023 reference is the mean of BPDB's FY2022-23 and FY2023-24 generation, and the upper limit is 1.02 x GWh x 0.0036 PJ/GWh. Example: gas, (46,013 + 46,579) / 2 = 46,296 GWh x 0.0036 x 1.02 = 169.998912 PJ per year.
- Assumptions: A1, A3.

### `PWRCOABGDXX`, `PWRHYDBGDXX`, `PWRNGSBGDXX`, `PWROILBGDXX`, `PWRSPVBGDXX`, `PWRURNBGDXX`, `PWRWONBGDXX` TotalTechnologyAnnualActivityUpperLimit, 2024 and 2025

- Values: 0 to 171.038088 (7 distinct values); PJ per year; 8 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R4] and [R5] Energy Generation by Fuel Type, printed p.14 (PDF p.15).
- Derivation: The upper limit is 1.02 x BPDB's FY2023-24 generation x 0.0036 PJ/GWh for all seven technologies in 2024; in 2025 only nuclear has an upper limit, 0. Example: gas 2024, 46,579 GWh x 0.0036 x 1.02 = 171.038088 PJ per year.
- Assumptions: A1, A3.

### `PWRHYDBGDXX` CapacityFactor, 2024 to 2026

- Values: 0.1421422687 to 0.9366800671 (40 distinct values); capacity factor (fraction); 60 rows.
- Label: `PGCB_FISCAL_PROFILE_WITH_SAME_PERIOD_CAPACITY`.
- Sources: [R6] and [R7] tab En-Curve, columns HYDRO and SOLAR, 48 half-hourly MW values per day; the series is one file per day on the PGCB ERP portal from the curve for 1 July 2023 to the curve for 30 June 2025, and these are its first and last files; [R5] Installed Capacity by Plant & Fuel Type, printed p.13 (PDF p.14).
- Derivation: Each half-hour of PGCB's daily energy curve falls in a season and a day part (00:00 to 06:00, 06:00 to 17:00, 17:00 to 20:00, 20:00 to 22:00 and 22:00 to 24:00); the load factor is the mean hydro MW in the slice / 230 MW of hydro capacity, and the capacity factor is that / 0.9. Example: 2024 S1, 00:00 to 06:00, 1,092 half-hours from December 2023 to February 2024 average 66.54 MW: 66.54 / 230 / 0.9 = 0.321457. Half-hours reported twice with different values are left out; identical repeats are counted once.
- Assumptions: A3, A5, A6.

### `PWRSPVBGDXX` CapacityFactor, 2024 to 2026

- Values: 0 to 0.5431140672 (26 distinct values); capacity factor (fraction); 60 rows.
- Label: `PGCB_FISCAL_PROFILE_WITH_SAME_PERIOD_CAPACITY`.
- Sources: [R6] and [R7] tab En-Curve, columns HYDRO and SOLAR, 48 half-hourly MW values per day; the series is one file per day on the PGCB ERP portal from the curve for 1 July 2023 to the curve for 30 June 2025, and these are its first and last files; [R4] Projects completed (Solar), printed p.31 (PDF p.32); [R5] Grid-tied renewable energy projects, printed p.34 (PDF p.35).
- Derivation: The capacity factor is the mean over the slice of half-hourly solar MW / the grid-connected AC solar stock in service that day, built from BPDB and EGCB commissioning dates: 458 MW, plus 75 MW from 1 April 2024, 1 MW from 25 May 2024, 68 MW from 14 July 2024 and 100 MW from 23 October 2024. Example: 2024 S1, 06:00 to 17:00, 2,002 half-hours from December 2023 to February 2024 with a stock of 458 MW give 0.375450.
- Assumptions: A3, A6.

### `PWRSPVBGDXX` ResidualCapacity, 2023, 2025 and 2026

- Values: 0.458; 0.6683452054794521; 0.7570273972602739; GW; 3 rows.
- Label: `BPDB_COHORTS_AND_BCRECL_ACTUAL_COD`.
- Sources: [R5] Implementation, Planning & Development of Grid-tied Renewable Energy Projects, printed p.34 (PDF p.35); Projects completed (Solar) and Projects under Construction (Solar); [R4] Projects completed (Solar), printed p.31 (PDF p.32); [R8] Project Overview.
- Derivation: Solar in service before July 2023 is 702 MW at 30 June 2025 less the 76 MW commissioned in FY2023-24 and the 168 MW commissioned in FY2024-25: 458 MW, the 2023 value. Plants count for their days in service in the fiscal year: 2025 = 458 + 75 + 1 + 68 x 352/365 + 100 x 251/365 = 668.345 MW and 2026 = 702 + 65 x 309/365 = 757.027 MW.
- Assumptions: A3, A4.

### `PWRSPVBGDXX` ResidualCapacity, 2024

- Values: 0.4767486338797814; GW; 1 rows.
- Label: `BPDB_COHORTS_AND_BCRECL_ACTUAL_COD`.
- Sources: [R4] Projects completed (Solar), printed p.31 (PDF p.32); [R9] General Information table, item 11, printed p.37 (PDF p.38); [R5] Implementation, Planning & Development of Grid-tied Renewable Energy Projects, printed p.34 (PDF p.35); Projects completed (Solar) and Projects under Construction (Solar).
- Derivation: 2024 is FY2023-24 (366 days): 458 MW + Sonagazi 75 MW x 91/366 (from 1 April 2024, EGCB's commercial operation date) + Barisal 1 MW x 37/366 (from 25 May 2024) = 476.749 MW = 0.476749 GW.
- Assumptions: A3, A4.

### `PWRSPVBGDXX` ResidualCapacity, 2027 to 2034

- Values: 0.767; GW; 8 rows.
- Label: `CAPACITY_CHRONOLOGY`.
- Sources: [R5] printed pp.47-49 (PDF pp.48-50), table 'Generation Capacity (MW) As of 30 Jun 2025', solar rows (Drive file linked from BPDB's annual-reports page); [R5] printed p.33 (PDF p.34), list of power plants under construction, Private Sector row 8 'Pabna 64.55 MW Solar Park (BCRECL)'; [R8] 'Project Details' table.
- Derivation: Fourteen plants in service at 30 June 2025 sum to 6 + 75 + 1 + 6 + 3 + 20 + 8 + 50 + 100 + 35 + 30 + 100 + 200 + 68 = 702 MW; with BCRECL Pabna, 65 MW (BPDB's 64.55 MW rounded; commercial operation 26 August 2025), the stock is 767 MW = 0.767 GW. 2027 to 2034 are the years in which all fifteen plants are in service for whole years.
- Assumptions: A11.

### `PWRSPVBGDXX` ResidualCapacity, 2035 to 2050

- Values: 0.309 to 0.539 (3 distinct values); GW; 16 rows.
- Label: `BCRECL_ACTUAL_COD_AND_RETAINED_30_YEAR_LIFE`.
- Sources: [R5] Implementation, Planning & Development of Grid-tied Renewable Energy Projects, printed p.34 (PDF p.35); Projects completed (Solar) and Projects under Construction (Solar); [R4] Projects completed (Solar), printed p.31 (PDF p.32); [R8] Project Overview; [R3] printed p.44, grid-tied solar projects completed.
- Derivation: Plants commissioned from March 2024 (Sonagazi 75, Barisal 1, Sirajganj 68, Pabna 100 and BCRECL Pabna 65 MW, 309 MW in all) stay to 2050; Teesta (200 MW) leaves in 2048 and Intraco (30 MW) in 2047. 2035 to 2046 = 309 + 200 + 30 = 539 MW, 2047 = 509 MW and 2048 to 2050 = 309 MW.
- Assumptions: A11. 228 MW of Bangladesh solar plants in service before July 2023 have no published commissioning year; they are given 2010 and leave the stock in 2035. Teesta (200 MW) and Intraco (30 MW) leave 25 years after their commercial operation dates, in 2048 and 2047. Plants commissioned in 2024 and 2025 have a 30-year life, so they stay in service to 2050.

### `PWRWONBGDXX` ResidualCapacity, 2023

- Values: 0.063; GW; 1 rows.
- Label: `USER_DEFINED_RESIDUAL_LOST_IN_SCENARIO_BUILD`.
- Sources: [R10] PDF p.29, table 'Onshore wind energy, CAP (MW)', row Bangladesh, column 2023; [R4] printed p.32 (PDF p.33), 'Projects completed (Wind)'.
- Derivation: IRENA gives Bangladesh 3 MW of onshore wind in 2023 and 63 MW in 2024; BPDB gives the 60 MW Cox's Bazar plant a commercial operation date of 3 March 2024. The 2023 value is 63 MW = 0.063 GW.
- Assumptions: Bangladesh onshore wind is 63 MW from 2023, counting the 60 MW Cox's Bazar plant from the base year, although BPDB gives its commercial operation date as 3 March 2024; IRENA reports 3 MW at end-2023, so the 2023 stock is 60 MW above the published figure.

### `PWRWONBGDXX` ResidualCapacity, 2024 to 2050

- Values: 0.063; GW; 27 rows.
- Label: `USER_DEFINED_RESIDUAL_LOST_IN_SCENARIO_BUILD`.
- Sources: [R10] PDF p.29, table 'Onshore wind energy, CAP (MW)', row Bangladesh, columns 2024 and 2025; [R4] printed p.32 (PDF p.33), 'Projects completed (Wind)'; [R5] printed pp.48-49 (PDF pp.49-50), table 'Generation Capacity (MW) As of 30 Jun 2025', rows 'Sirajgonj 2 MW Wind Plant' and 'US-DK Wind PP 60 MW, Cox's Bazar' (Drive file linked from BPDB's annual-reports page).
- Derivation: 63 MW / 1,000 = 0.063 GW, IRENA's figure for 2024 and 2025, held to 2050. It is BPDB's 60 MW Cox's Bazar plant plus about 3 MW of small units in IRENA's series.
- Assumptions: A11. Bangladesh's 63 MW national onshore wind stock reported for 2024 and 2025 stays constant to 2050; the 2 MW Sirajganj unit (commissioned 26 April 2025) is within that national total.

## Bhutan

### `PWRHYDBTNXX` CapacityFactor, 2023 to 2026

- Values: 0.1898224066 to 0.9753438849 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `RMA_VERIFIED_PRODUCTION_COD_AND_AF_COUNTED_ONCE`.
- Sources: [R11] tab Operations-Hydropower Projects, monthly production by plant (2023); [R12] tab Operations-Hydropower Projects, monthly production by plant (2024); [R13] tab Operations-Hydropower Projects, monthly production by plant (2025, and PHP-II December 2024).
- Derivation: Seasonal production is the sum of RMA's monthly production by plant; load factor = production / (year-end stock GW x 8,760 h x season share 0.246576, 0.252056, 0.334247 or 0.167123), and the capacity factor is that / 0.95 in all five day parts. Example: 2023 S1, 1,005.48 GWh / (2.33 GW x 8,760 h x 0.246576) / 0.95 = 0.210300.
- Assumptions: A5, A6. Dagachhu and Kurichhu production for December 2025 is printed as 0 in the August 2026 bulletin and is taken as 0, so the 2025 and 2026 winter capacity factor may be slightly low; Suchhu counts from its April 2025 commissioning.

### `PWRHYDBTNXX` ResidualCapacity, 2023 to 2026

- Values: 2.33; 2.788; 3.486; 3.486; GW; 4 rows.
- Label: `RMA_PRODUCTION_YEAR_END_STOCK_COD_EXPOSURE_IN_PROFILE_ONCE`.
- Sources: [R14] Note 1 General Corporate Information, printed p.59 (PDF p.64); DHPC profile, printed p.15 (PDF p.20); Nikachhu commercial operation, printed pp.25 and 33 (PDF pp.30, 38); first two PHP-II units December 2024, printed p.24 (PDF p.29); [R15] unit commissioning table; [R16] Introduction, printed p.3 (PDF p.5).
- Derivation: Plants in service before 2023 are Tala 1,020, Chukha 336, Kurichhu 60, Basochhu 64, Mangdechhu 720 and Dagachhu 126 MW, 2,326 MW, plus 4 MW of mini and micro hydro: 2,330 MW in 2023. 2024 adds Nikachhu 118 MW and Punatsangchhu-II units I and II, 340 MW (27 December 2024): 2,788 MW. 2025 adds units III to VI, 680 MW, and Suchhu 18 MW: 3,486 MW, also in 2026.
- Assumptions: Bhutan's mini and micro hydro is represented by 4 MW added to the named plants; DGPC reports 9 MW of embedded generation including wind and solar, and using 9 MW would add 0.2% to Bhutan's hydro stock.

### `PWRHYDBTNXX` ResidualCapacity, 2027 to 2050

- Values: 3.486 (1 distinct value); GW; 24 rows.
- Label: `DGPC_DHYE_AND_APPROVED_UNIT_COD_CALENDAR_EXPOSURE_ONCE`.
- Sources: [R14] Note 1 General Corporate Information, printed p.59 (PDF p.64); DHPC profile, printed p.15 (PDF p.20); Nikachhu commercial operation, printed pp.25 and 33 (PDF pp.30, 38); first two PHP-II units December 2024, printed p.24 (PDF p.29); [R15] unit commissioning table; [R16] Introduction, printed p.3 (PDF p.5).
- Derivation: From 2027 the stock is 3,486 MW = 3.486 GW: 2,326 MW of plants in service before 2023, 4 MW of mini and micro hydro, Nikachhu 118 MW, Punatsangchhu-II 6 x 170 = 1,020 MW and Suchhu 18 MW, with no retirement before 2050. BPC's Power Data Book labels Suchhu 20 MW and Dagachhu 2 x 126 MW; DGPC's figures, 18 MW and 2 x 63 MW, are used.
- Assumptions: A11. Bhutan's mini and micro hydro is represented by 4 MW added to the named plants; DGPC reports 9 MW of embedded generation including wind and solar, and using 9 MW would add 0.2% to Bhutan's hydro stock. Plants under construction in 2025 are left to the model's investment choices.

### `PWRHYDBTNXX`, `PWRSPVBTNXX` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 0.059679 to 55.89914904 (6 distinct values); PJ per year; 6 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R17] Table 2.1, printed p.24 (PDF p.28).
- Derivation: Hydro is the sum of the plant generation rows of BPC's Table 2.1 and solar is Sephu plus DSP, in MU (GWh), x 0.0036 PJ/GWh; the lower limit is 0.98 x PJ in 2023 to 2025 and 0.95 x the 2025 PJ in 2026. Example: hydro 2023, 10,516.02 GWh x 0.0036 x 0.98 = 37.100519 PJ per year. The 2025 hydro sum, 15,844.43 MU, includes Suchhu (41.09 MU), which the printed total of 15,803.33 MU leaves out; embedded mini and micro hydro and wind (16.95 MU) are not counted as hydro.
- Assumptions: A1, A2.

### `PWRHYDBTNXX`, `PWRSPVBTNXX` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 0.0640764 to 58.18074696 (4 distinct values); PJ per year; 4 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R17] Table 2.1, printed p.24 (PDF p.28).
- Derivation: Hydro is the sum of the plant generation rows of BPC's Table 2.1 and solar is Sephu plus DSP, in MU (GWh); the upper limit is 1.02 x GWh x 0.0036 PJ/GWh. Example: hydro 2023, 10,516.02 GWh x 0.0036 x 1.02 = 38.614825 PJ per year. The 2025 hydro sum, 15,844.43 MU, includes Suchhu (41.09 MU), which the printed total of 15,803.33 MU leaves out; embedded mini and micro hydro and wind (16.95 MU) are not counted as hydro.
- Assumptions: A1, A2.

### `PWRSPVBTNXX` ResidualCapacity, 2023 to 2029

- Values: 0.002 (2023 to 2024); 0.010328986301 (2025); 0.02438 (2026 to 2029); GW; 7 rows.
- Label: `MOENR_2025_07_19_AND_MOF_2026_27_P157`.
- Sources: [R18] News post, 19 July 2025; [R19] Annexure, VI. Solar and Wind Projects under Operation and Construction, printed p.157 (PDF p.173); [R20] Directors' Report, 1. Introduction, printed p.16 (PDF p.17).
- Derivation: DSP = 3.23 GWh / (8,760 h x 0.18) = 2.05 MW, taken as 2 MW: 0.002 GW in 2023 and 2024. 2025 = 2 + 17.38 x 166/365 (Phase I from 19 July) + 5 x 31/365 (Phase II from 1 December) = 10.329 MW, and 2026 to 2029 = 2 + 22.38 = 24.38 MW.
- Assumptions: A4, A11. Desuung (DSP) solar is represented by 2 MW, from its 3.23 MU sold to BPC in 2024 at an 18% capacity factor, with a commissioning year of 2005, leaving the stock in 2030; the 2026-27 budget report lists 4.11 MWp of DSP-SI I and II plus 1.5 MWp at Paro airport in operation. Sephu Phase II (5 MWp) counts from 1 December 2025, the latest date consistent with the budget report's November 2025 statement, and Phase I from 19 July 2025.

### `PWRSPVBTNXX` ResidualCapacity, 2030 to 2050

- Values: 0.02238 (2030 to 2049); 0 (2050); GW; 21 rows.
- Label: `MOENR_2025_07_19_AND_MOF_2026_27_P157`.
- Sources: [R18] News post, 19 July 2025; [R19] Annexure, VI. Solar and Wind Projects under Operation and Construction, printed p.157 (PDF p.173); [R21] 1. Background, printed p.2 (PDF p.4).
- Derivation: Sephu solar is 17.38 + 5 = 22.38 MWp = 0.02238 GW from 2030 to 2049, and 0 in 2050, 25 years after its 2025 commissioning. The 2 MW DSP plant leaves in 2030.
- Assumptions: A11. Sephu solar is counted at its peak DC rating (MWp).

## India

### `PWRBIOINDEA` ResidualCapacity, 2023 to 2050

- Values: 0.52684; GW; 28 rows.
- Label: `USER_DEFINED_RESIDUAL_LOST_IN_SCENARIO_BUILD`.
- Sources: [R22] PDF page 1 (single table); columns 'BM Power/Bagasse Cogen. (MW)' and 'BM Cogen. (Non-Bagasse) (MW)'; rows Bihar, Jharkhand, Odisha, West Bengal.
- Derivation: Bagasse and non-bagasse cogeneration over Bihar, Jharkhand, Odisha and West Bengal in MNRE's table at 31 March 2023: (112.50 + 12.20) + 4.30 + (50.40 + 8.82) + (300.00 + 38.62) = 526.84 MW = 0.52684 GW in every year.
- Assumptions: A11, A13.

### `PWRBIOINDEA`, `PWRBIOINDNE`, `PWRBIOINDNO`, `PWRBIOINDSO`, `PWRBIOINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 0 to 19.23691392 (17 distinct values); PJ per year; 20 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Bagasse regional Sub Total, current cumulative Apr-Dec2025, Biomass regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year biomass and bagasse generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The lower limit is GWh x 0.98 x 0.0036 PJ/GWh in 2023 to 2025 and 0.95 x the 2025 GWh x 0.0036 in 2026. Example: `PWRBIOINDWE` 2023, (994.26 + 835.16 + 684.22 + 2,939.00) = 5,452.64 GWh x 0.98 x 0.0036 = 19.236914 PJ per year.
- Assumptions: A1, A2.

### `PWRBIOINDEA`, `PWRBIOINDNE`, `PWRBIOINDNO`, `PWRBIOINDSO`, `PWRBIOINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 0 to 20.02209408 (13 distinct values); PJ per year; 15 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Bagasse regional Sub Total, current cumulative Apr-Dec2025, Biomass regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year biomass and bagasse generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The upper limit is GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWRBIOINDWE` 2023, 5,452.64 GWh x 1.02 x 0.0036 = 20.022094 PJ per year.
- Assumptions: A1, A2.

### `PWRBIOINDNE` ResidualCapacity, 2023 to 2050

- Values: 0.0158; GW; 28 rows.
- Label: `USER_DEFINED_RESIDUAL_LOST_IN_SCENARIO_BUILD`.
- Sources: [R22] PDF page 1 (single table); columns 'BM Power/Bagasse Cogen. (MW)' and 'BM Cogen. (Non-Bagasse) (MW)'; rows Assam, Meghalaya.
- Derivation: Bagasse and non-bagasse cogeneration in MNRE's table at 31 March 2023: Assam 2.00 + Meghalaya 13.80 = 15.80 MW = 0.0158 GW in every year.
- Assumptions: A11, A13.

### `PWRCOAINDEA`, `PWRCOAINDNE`, `PWRCOAINDNO`, `PWRCOAINDSO`, `PWRCOAINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 14.8468698 to 1861.15575744 (20 distinct values); PJ per year; 20 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R34] PDF p.1; regional or state rows; current month; [R35] PDF pp.1 to 44, state station tables; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R36] PDF p.1; regional or state rows; April-December cumulative; [R37] PDF pp.1 to 45, state station tables; regional or state rows; April-December cumulative.
- Derivation: Regional thermal generation is the calendar-year sum of CEA's January, February and March figures and its April to December cumulative figure; oil receives thermal x oil weight / (coal weight + oil weight), each weight being installed GW times the utilisation in A8, and coal the rest; lower limit = GWh x 0.98 x 0.0036 PJ/GWh (0.95 x the 2025 GWh in 2026). Example: `PWRCOAINDWE` 2023, thermal 40,406.22 + 37,819.02 + 41,548.74 + 383,808.84 = 503,582.82 GWh, oil weight 0, so coal = 503,582.82 GWh x 0.98 x 0.0036 = 1,776.640189 PJ per year.
- Assumptions: A1, A2, A8.

### `PWRCOAINDEA`, `PWRCOAINDNE`, `PWRCOAINDNO`, `PWRCOAINDSO`, `PWRCOAINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 15.94084968 to 1937.12129856 (15 distinct values); PJ per year; 15 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R34] PDF p.1; regional or state rows; current month; [R35] PDF pp.1 to 44, state station tables; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R36] PDF p.1; regional or state rows; April-December cumulative; [R37] PDF pp.1 to 45, state station tables; regional or state rows; April-December cumulative.
- Derivation: Regional thermal generation is the calendar-year sum of CEA's January, February and March figures and its April to December cumulative figure; oil receives thermal x oil weight / (coal weight + oil weight), each weight being installed GW times the utilisation in A8, and coal the rest; upper limit = GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWRCOAINDWE` 2023, 503,582.82 GWh of thermal, all coal, x 1.02 x 0.0036 = 1,849.156115 PJ per year.
- Assumptions: A1, A2, A8.

### `PWRHYDINDEA` CapacityFactor, 2023 to 2026

- Values: 0.14927271697391 to 0.6348967227908 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R33], [R38] to [R54] tab Large Hydro, Eastern Region Sub Total, monthly current/prior and cumulative current/prior; [R55] to [R90] PDF p.1, Sub-report 2, Summary - Region Wise, EASTERN REGION HYDRO, Monitored Capacity (MW).
- Derivation: Seasonal generation is read from the Large Hydro tab of CEA's monthly renewable generation reports for the Eastern Region, and each capacity factor is seasonal GWh / (stock GW x season hours x availability 0.95). Example: 2023 S1, 2,148.54 GWh / (5.988 GW x 2,160 h x 0.95) = 0.174858. The 5.988 GW stock is CEA's monitored large hydro capacity for the Eastern Region, 5,987.75 MW, in each of the 36 monthly 18-column reports for January 2023 to December 2025.
- Assumptions: A5, A6.

### `PWRHYDINDEA`, `PWRHYDINDNE`, `PWRHYDINDNO`, `PWRHYDINDSO`, `PWRHYDINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 21.97386576 to 287.61681312 (20 distinct values); PJ per year; 20 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R34] PDF p.1; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R36] PDF p.1; regional or state rows; April-December cumulative.
- Derivation: Calendar-year large hydro generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The lower limit is GWh x 0.98 x 0.0036 PJ/GWh in 2023 to 2025 and 0.95 x the 2025 GWh x 0.0036 in 2026. Example: `PWRHYDINDNO` 2023, (3,210.57 + 3,480.26 + 3,692.09 + 63,579.49) = 73,962.41 GWh x 0.98 x 0.0036 = 260.939382 PJ per year.
- Assumptions: A1, A2.

### `PWRHYDINDEA`, `PWRHYDINDNE`, `PWRHYDINDNO`, `PWRHYDINDSO`, `PWRHYDINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 22.87075824 to 299.35627488 (15 distinct values); PJ per year; 15 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R34] PDF p.1; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R36] PDF p.1; regional or state rows; April-December cumulative.
- Derivation: Calendar-year large hydro generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The upper limit is GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWRHYDINDNO` 2023, 73,962.41 GWh x 1.02 x 0.0036 = 271.589970 PJ per year.
- Assumptions: A1, A2.

### `PWRHYDINDNE` CapacityFactor, 2023 to 2026

- Values: 0.1927680023388 to 0.74519759774536 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R55] to [R90] Sub-report 2, Summary - Region Wise, North Eastern Region HYDRO, PDF pp.1-2.
- Derivation: Each month's hydro generation in the North Eastern Region is divided by that month's monitored capacity in GW, the results are summed over the season, and the sum is divided by the season hours and an availability of 0.95. Example: 2023 S1, (269.42 / 2.027 + 218.68 / 2.027 + 313.70 / 2.027) / 2,160 h / 0.95 = 0.192768.
- Assumptions: A5, A6.

### `PWRHYDINDNO` CapacityFactor, 2023 to 2026

- Values: 0.2147883997357 to 0.75163947615789 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R55] to [R90] Sub-report 2, Summary - Region Wise, Northern Region HYDRO, PDF pp.1-2.
- Derivation: Each month's hydro generation in the Northern Region is divided by that month's monitored capacity in GW, the results are summed over the season, and the sum is divided by the season hours and an availability of 0.95. Example: 2023 S1, (3,279.99 / 19.6963 + 3,472.23 / 19.6963 + 3,434.94 / 19.7263) / 2,160 h / 0.95 = 0.251924.
- Assumptions: A5, A6.

### `PWRHYDINDSO` CapacityFactor, 2023 to 2026

- Values: 0.17162104913187 to 0.56417906574904 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R55] to [R90] Sub-report 2, Summary - Region Wise, Southern Region HYDRO, PDF pp.1-2.
- Derivation: Each month's hydro generation in the Southern Region is divided by that month's monitored capacity in GW, the results are summed over the season, and the sum is divided by the season hours and an availability of 0.92. Example: 2023 S1, (2,597.36 / 11.7471 + 2,258.85 / 11.7471 + 1,331.46 / 11.7471) / 2,160 h / 0.92 = 0.265065.
- Assumptions: A5, A6.

### `PWRHYDINDWE` CapacityFactor, 2023 to 2026

- Values: 0.2126291779529 to 0.49141599996632 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R55] to [R90] Sub-report 2, Summary - Region Wise, Western Region HYDRO, PDF pp.1-2.
- Derivation: Each month's hydro generation in the Western Region is divided by that month's monitored capacity in GW, the results are summed over the season, and the sum is divided by the season hours and an availability of 0.92. Example: 2023 S1, (1,203.12 / 7.392 + 1,654.03 / 7.392 + 844.43 / 7.392) / 2,160 h / 0.92 = 0.251990.
- Assumptions: A5, A6.

### `PWRNGSINDEA`, `PWRNGSINDNE`, `PWRNGSINDNO`, `PWRNGSINDSO`, `PWRNGSINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 0 to 55.0558512 (17 distinct values); PJ per year; 20 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R35] PDF pp.1 to 44, state station tables; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R37] PDF pp.1 to 45, state station tables; regional or state rows; April-December cumulative.
- Derivation: Calendar-year gas generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The lower limit is GWh x 0.98 x 0.0036 PJ/GWh in 2023 to 2025 and 0.95 x the 2025 GWh x 0.0036 in 2026. Example: `PWRNGSINDNE` 2023, (926.52 + 868.36 + 924.03 + 7,487.30) = 10,206.21 GWh x 0.98 x 0.0036 = 36.007509 PJ per year.
- Assumptions: A1, A2.

### `PWRNGSINDEA`, `PWRNGSINDNE`, `PWRNGSINDNO`, `PWRNGSINDSO`, `PWRNGSINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 0 to 57.3030288 (13 distinct values); PJ per year; 15 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R35] PDF pp.1 to 44, state station tables; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R37] PDF pp.1 to 45, state station tables; regional or state rows; April-December cumulative.
- Derivation: Calendar-year gas generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The upper limit is GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWRNGSINDNE` 2023, 10,206.21 GWh x 1.02 x 0.0036 = 37.477203 PJ per year.
- Assumptions: A1, A2.

### `PWROILINDEA`, `PWROILINDNE`, `PWROILINDNO`, `PWROILINDSO`, `PWROILINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 0 to 3.09677256 (9 distinct values); PJ per year; 20 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R34] PDF p.1; regional or state rows; current month; [R35] PDF pp.1 to 44, state station tables; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R36] PDF p.1; regional or state rows; April-December cumulative; [R37] PDF pp.1 to 45, state station tables; regional or state rows; April-December cumulative.
- Derivation: Regional thermal generation is the calendar-year sum of CEA's January, February and March figures and its April to December cumulative figure; oil receives thermal x oil weight / (coal weight + oil weight), each weight being installed GW times the utilisation in A8, and coal the rest; lower limit = GWh x 0.98 x 0.0036 PJ/GWh (0.95 x the 2025 GWh in 2026). Example: `PWROILINDSO` 2023, thermal 18,540.88 + 18,809.79 + 21,403.00 + 178,502.22 = 237,255.89 GWh, oil = 237,255.89 x 0.108415 / (30.2234 + 0.108415) = 848.02 GWh x 0.98 x 0.0036 = 2.991815 PJ per year.
- Assumptions: A1, A2, A8.

### `PWROILINDEA`, `PWROILINDNE`, `PWROILINDNO`, `PWROILINDSO`, `PWROILINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 0 to 3.22317144 (7 distinct values); PJ per year; 15 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R34] PDF p.1; regional or state rows; current month; [R35] PDF pp.1 to 44, state station tables; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R36] PDF p.1; regional or state rows; April-December cumulative; [R37] PDF pp.1 to 45, state station tables; regional or state rows; April-December cumulative.
- Derivation: Regional thermal generation is the calendar-year sum of CEA's January, February and March figures and its April to December cumulative figure; oil receives thermal x oil weight / (coal weight + oil weight), each weight being installed GW times the utilisation in A8, and coal the rest; upper limit = GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWROILINDSO` 2023, 848.02 GWh of oil x 1.02 x 0.0036 = 3.113929 PJ per year.
- Assumptions: A1, A2, A8.

### `PWROILINDNE` ResidualCapacity, 2023 to 2050

- Values: 0.036; GW; 28 rows.
- Label: `USER_DEFINED_RESIDUAL_LOST_IN_SCENARIO_BUILD`.
- Sources: [R91] PDF page 1, region table; row 'North Eastern Region - Sub Total', column 'Thermal - Diesel'.
- Derivation: CEA prints 36.00 MW of diesel for the North Eastern Region at 31 March 2023: 0.036 GW in every year.
- Assumptions: A11.

### `PWRSHPINDEA` CapacityFactor, 2023 to 2026

- Values: 0.14892589599877 to 0.45101053559251 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R33], [R38] to [R54], [R92] and [R93] tab Small Hydro, Eastern Region Sub Total, monthly current/prior and cumulative current/prior; [R94] PDF p.1, Small Hydro Power column; Andaman & Nicobar, Bihar, Jharkhand, Odisha, Sikkim, West Bengal.
- Derivation: Seasonal generation is read from the Small Hydro tab of CEA's monthly renewable generation reports for the Eastern Region, and each capacity factor is seasonal GWh / (stock GW x season hours x availability 0.9). Example: 2023 S1, 108.01 GWh / (0.3442 GW x 2,160 h x 0.9) = 0.161420.
- Assumptions: A5, A6. India East small hydro stock is 0.3442 GW (344.20 MW) in 2023 to 2026; MNRE's state-wise installed capacity at 31 December 2023 gives 349.24 MW for Andaman and Nicobar, Bihar, Jharkhand, Odisha, Sikkim and West Bengal, so the model stock is 1.44% lower. Without Andaman and Nicobar (5.25 MW), which is not connected to the mainland grid, MNRE's figure is 343.99 MW, 0.06% below the model stock. The stock only scales the capacity factor, which the model multiplies back by the same stock in its annual limit (A5), so generation still equals the published amount.

### `PWRSHPINDEA`, `PWRSHPINDNE`, `PWRSHPINDNO`, `PWRSHPINDSO`, `PWRSHPINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 1.09883088 to 17.38520784 (19 distinct values); PJ per year; 19 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Small Hydro regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year small hydro generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The lower limit is GWh x 0.98 x 0.0036 PJ/GWh in 2023 to 2025 and 0.95 x the 2025 GWh x 0.0036 in 2026. Example: `PWRSHPINDSO` 2023, (214.62 + 163.44 + 160.76 + 2,198.06) = 2,736.88 GWh x 0.98 x 0.0036 = 9.655713 PJ per year.
- Assumptions: A1, A2.

### `PWRSHPINDEA`, `PWRSHPINDNE`, `PWRSHPINDNO`, `PWRSHPINDSO`, `PWRSHPINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 1.14368112 to 18.09480816 (14 distinct values); PJ per year; 14 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Small Hydro regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year small hydro generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The upper limit is GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWRSHPINDSO` 2023, 2,736.88 GWh x 1.02 x 0.0036 = 10.049823 PJ per year.
- Assumptions: A1, A2.

### `PWRSHPINDNE` CapacityFactor, 2023 to 2026

- Values: 0.093513984742035 to 0.29461222093325 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R33], [R38] to [R54], [R92] and [R93] tab Small Hydro, North Eastern Region Sub Total, monthly current/prior and cumulative current/prior; [R94] PDF p.1, Small Hydro Power column; Arunachal Pradesh, Assam, Manipur, Meghalaya, Mizoram, Nagaland, Tripura.
- Derivation: Seasonal generation is read from the Small Hydro tab of CEA's monthly renewable generation reports for the North Eastern Region, and each capacity factor is seasonal GWh / (stock GW x season hours x availability 0.9). Example: 2023 S1, 62.65 GWh / (0.29935 GW x 2,160 h x 0.9) = 0.107666.
- Assumptions: A5, A6. India North-East small hydro stock is 0.29935 GW (299.35 MW) in 2023 to 2026; MNRE's state-wise installed capacity at 31 December 2023 gives 321.85 MW for Arunachal Pradesh, Assam, Manipur, Meghalaya, Mizoram, Nagaland and Tripura, so the model stock is 6.99% lower. The stock only scales the capacity factor, which the model multiplies back by the same stock in its annual limit (A5), so generation still equals the published amount.

### `PWRSHPINDNO` CapacityFactor, 2023 to 2026

- Values: 0.18503094440003 to 0.51663796463992 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R33], [R38] to [R54], [R92] and [R93] tab Small Hydro, Northern Region Sub Total, monthly current/prior and cumulative current/prior; [R94] PDF p.1, Small Hydro Power column; Chandigarh, Delhi, Haryana, Himachal Pradesh, Jammu & Kashmir, Ladakh, Punjab, Rajasthan, Uttar Pradesh, Uttarakhand.
- Derivation: Seasonal generation is read from the Small Hydro tab of CEA's monthly renewable generation reports for the Northern Region, and each capacity factor is seasonal GWh / (stock GW x season hours x availability 0.9). Example: 2023 S1, 699.21 GWh / (1.7189 GW x 2,160 h x 0.9) = 0.209249. The 1.7189 GW stock is MNRE's small hydro capacity at 31 December 2023 summed over the region's states, 1,713.50 MW, 0.32% below it.
- Assumptions: A5, A6.

### `PWRSHPINDNO` TotalTechnologyAnnualActivityLowerLimit, 2023

- Values: 15.79534992; PJ per year; 1 rows.
- Label: `VERIFIED_ASSEMBLY_ERRATUM_NORTH_SHP_2023`.
- Sources: [R23] PDF p.30; Northern Region; Small Hydro; current 2023 column; [R24] PDF p.41; Northern Region; Small Hydro; prior 2023 column; [R25] PDF p.29; Northern Region; Small Hydro; current 2023 column; [R26] PDF p.42; Northern Region; Small Hydro; current 2023 column.
- Derivation: Northern Region small hydro in calendar 2023 is 225.70 + 228.94 + 242.52 + 3,779.98 = 4,477.14 GWh (January, February and March, then April to December cumulative). The lower limit is 4,477.14 x 0.98 x 0.0036 = 15.795350 PJ per year.
- Assumptions: A1.

### `PWRSHPINDNO` TotalTechnologyAnnualActivityUpperLimit, 2023

- Values: 16.44005808; PJ per year; 1 rows.
- Label: `VERIFIED_ASSEMBLY_ERRATUM_NORTH_SHP_2023`.
- Sources: [R23] PDF p.30; Northern Region; Small Hydro; current 2023 column; [R24] PDF p.41; Northern Region; Small Hydro; prior 2023 column; [R25] PDF p.29; Northern Region; Small Hydro; current 2023 column; [R26] PDF p.42; Northern Region; Small Hydro; current 2023 column.
- Derivation: Northern Region small hydro in calendar 2023 is 225.70 + 228.94 + 242.52 + 3,779.98 = 4,477.14 GWh (January, February and March, then April to December cumulative). The upper limit is 4,477.14 x 1.02 x 0.0036 = 16.440058 PJ per year.
- Assumptions: A1.

### `PWRSHPINDSO` CapacityFactor, 2023 to 2026

- Values: 0.055665835651173 to 0.48579570070472 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R33], [R38] to [R54], [R92] and [R93] tab Small Hydro, Southern Region Sub Total, monthly current/prior and cumulative current/prior; [R94] PDF p.1, Small Hydro Power column; Andhra Pradesh, Karnataka, Kerala, Lakshadweep, Pondicherry, Tamil Nadu, Telangana.
- Derivation: Seasonal generation is read from the Small Hydro tab of CEA's monthly renewable generation reports for the Southern Region, and each capacity factor is seasonal GWh / (stock GW x season hours x availability 0.9). Example: 2023 S1, 530.78 GWh / (1.9333 GW x 2,160 h x 0.9) = 0.141226. The 1.9333 GW stock is MNRE's small hydro capacity at 31 December 2023 summed over the region's states, 1,928.48 MW, 0.25% below it.
- Assumptions: A5, A6.

### `PWRSHPINDWE` CapacityFactor, 2023 to 2026

- Values: 0.25264712262053 to 0.43925890821833 (12 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_MONTHLY_PLF_WITH_RETAINED_AF_DECOMPOSITION`.
- Sources: [R33], [R38] to [R54], [R92] and [R93] tab Small Hydro, Western Region Sub Total, monthly current/prior and cumulative current/prior; [R94] PDF p.1, Small Hydro Power column; Chhatisgarh, Goa, Gujarat, Madhya Pradesh, Maharashtra, Dadar & Nagar Haveli/Daman & Diu.
- Derivation: Seasonal generation is read from the Small Hydro tab of CEA's monthly renewable generation reports for the Western Region, and each capacity factor is seasonal GWh / (stock GW x season hours x availability 0.9). Example: 2023 S1, 390.70 GWh / (0.6739 GW x 2,160 h x 0.9) = 0.298230. The 0.6739 GW stock is MNRE's small hydro capacity at 31 December 2023 summed over the region's states, 673.68 MW, 0.03% below it.
- Assumptions: A5, A6.

### `PWRSPVINDEA`, `PWRSPVINDNE`, `PWRSPVINDNO`, `PWRSPVINDSO`, `PWRSPVINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 1.13188824 to 232.3177416 (20 distinct values); PJ per year; 20 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Solar regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year solar generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The lower limit is GWh x 0.98 x 0.0036 PJ/GWh in 2023 to 2025 and 0.95 x the 2025 GWh x 0.0036 in 2026. Example: `PWRSPVINDNO` 2023, (3,539.94 + 3,733.67 + 4,172.76 + 34,805.56) = 46,251.93 GWh x 0.98 x 0.0036 = 163.176809 PJ per year.
- Assumptions: A1, A2.

### `PWRSPVINDEA`, `PWRSPVINDNE`, `PWRSPVINDNO`, `PWRSPVINDSO`, `PWRSPVINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 1.17808776 to 241.8000984 (15 distinct values); PJ per year; 15 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Solar regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year solar generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The upper limit is GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWRSPVINDNO` 2023, 46,251.93 GWh x 1.02 x 0.0036 = 169.837087 PJ per year.
- Assumptions: A1, A2.

### `PWRSPVINDNO` CapacityFactor, 2023 to 2026

- Values: 0 to 0.52945116045079 (31 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_GENERATION_MNRE_GRID_STOCK_NINJA_DAYLIGHT_SHAPE`.
- Sources: [R94] to [R96] PDF p.1; Ground Mounted Solar, Rooftop/RTS, Hybrid Solar and Off-grid columns; Haryana, Himachal Pradesh, Jammu & Kashmir, Ladakh, Punjab, Rajasthan, Uttar Pradesh, Uttarakhand, Chandigarh, Delhi; [R33], [R38] to [R54], [R92] and [R93] tab Solar, Northern Region Sub Total; monthly current/prior and cumulative current/prior; [R97] Section 2; Section 4; printed p.1263, Figure 18; [R98] Renewables.ninja PV model; [R99] PV simulation at Delhi 28.67 N 77.23 E (North), MERRA-2, 2014 to 2022, 10% system losses, tilt equal to latitude, azimuth 180.
- Derivation: Seasonal solar generation for the Northern Region (CEA, Solar tab) is divided by the grid-connected stock at 31 December of the same year (MNRE ground-mounted + rooftop + hybrid solar, off-grid left out) and the season hours, then spread over the five day parts with the Renewables.ninja daily shape for Delhi, scaled so that its seasonal mean is the seasonal factor. Example: 2023 S1, day part 06:00 to 17:00, 10,650.44 GWh / (23.61227 GW x 2,160 h) x 0.378969 / 0.173696 = 0.455607.
- Assumptions: A6, A7.

### `PWRSPVINDNO` ResidualCapacity, 2023 to 2025

- Values: 23.41941 to 32.94885 (3 distinct values); GW; 3 rows.
- Label: `MNRE_YEAR_END_GRID_STOCK_LESS_RETAINED_MANDATORY_INVESTMENT`.
- Sources: [R94] to [R96] PDF p.1; Ground Mounted Solar, Rooftop/RTS, Hybrid Solar and Off-grid columns; Haryana, Himachal Pradesh, Jammu & Kashmir, Ladakh, Punjab, Rajasthan, Uttar Pradesh, Uttarakhand, Chandigarh, Delhi.
- Derivation: Grid-connected stock is ground-mounted + rooftop + hybrid solar over the Northern Region states in MNRE's year-end tables, off-grid left out: 23.61227, 32.1884 and 43.77581 GW at 31 December 2023, 2024 and 2025. Each year's value is that stock less the new capacity the model must build from 2023 to that year (0.19286 GW in 2023, 1.295 GW in 2024, 9.3391 GW in 2025), so 2023 = 23.61227 - 0.19286 = 23.41941 GW and 2025 = 43.77581 - 10.82696 = 32.94885 GW.
- Assumptions: A11, A12.

### `PWRSPVINDNO` ResidualCapacity, 2026 to 2050

- Values: 22.76455 to 32.94885 (4 distinct values); GW; 25 rows.
- Label: `MNRE_OBSERVED_NET_COHORTS_RETAINED_25_YEAR_MODEL_LIFE`.
- Sources: [R94] to [R96] PDF p.1; Ground Mounted Solar, Rooftop/RTS, Hybrid Solar and Off-grid columns; Haryana, Himachal Pradesh, Jammu & Kashmir, Ladakh, Punjab, Rajasthan, Uttar Pradesh, Uttarakhand, Chandigarh, Delhi.
- Derivation: The 2025 value, 32.94885 GW, holds from 2026 to 2047. The capacity added in 2023, 2024 and 2025 above the 22.76455 GW that stays to 2050 (0.65486, 7.28113 and 2.24831 GW) retires 25 years later: 2048 = 32.94885 - 0.65486 = 32.29399 GW, 2049 = 25.01286 GW and 2050 = 22.76455 GW.
- Assumptions: A11, A12. India North solar stock from 2026 to 2047 is the published stock at 31 December 2025 (43.78 GW, MNRE) less the new capacity the model must build in 2023 to 2025 (10.83 GW), so the model's 2025 capacity equals MNRE's figure. Capacity added in 2023, 2024 and 2025 retires in 2048, 2049 and 2050; the remaining 22.76 GW stays in service to 2050.

### `PWRSPVINDSO` CapacityFactor, 2023 to 2026

- Values: 0 to 0.43376060151633 (25 distinct values); capacity factor (fraction); 80 rows.
- Label: `CEA_GENERATION_MNRE_GRID_STOCK_NINJA_DAYLIGHT_SHAPE`.
- Sources: [R94] to [R96] PDF p.1; Ground Mounted Solar, Rooftop/RTS, Hybrid Solar and Off-grid columns; Andhra Pradesh, Karnataka, Kerala, Tamil Nadu, Telangana, Lakshadweep, Puducherry; [R33], [R38] to [R54], [R92] and [R93] tab Solar, Southern Region Sub Total; monthly current/prior and cumulative current/prior; [R97] Section 2; Section 4; printed p.1263, Figure 18; [R98] Renewables.ninja PV model; [R99] PV simulation at Bangalore 12.9724 N 77.5806 E (South), MERRA-2, 2014 to 2022, 10% system losses, tilt equal to latitude, azimuth 180.
- Derivation: Seasonal solar generation for the Southern Region (CEA, Solar tab) is divided by the grid-connected stock at 31 December of the same year (MNRE ground-mounted + rooftop + hybrid solar, off-grid left out) and the season hours, then spread over the five day parts with the Renewables.ninja daily shape for Bangalore, scaled so that its seasonal mean is the seasonal factor. Example: 2023 S1, day part 06:00 to 17:00, 10,683.67 GWh / (26.74001 GW x 2,160 h) x 0.469493 / 0.215343 = 0.403277.
- Assumptions: A6, A7.

### `PWRSPVINDSO` ResidualCapacity, 2023 to 2025

- Values: 26.64001 to 34.18073 (3 distinct values); GW; 3 rows.
- Label: `MNRE_YEAR_END_GRID_STOCK_LESS_RETAINED_MANDATORY_INVESTMENT`.
- Sources: [R94] to [R96] PDF p.1; Ground Mounted Solar, Rooftop/RTS, Hybrid Solar and Off-grid columns; Andhra Pradesh, Karnataka, Kerala, Tamil Nadu, Telangana, Lakshadweep, Puducherry.
- Derivation: Grid-connected stock is ground-mounted + rooftop + hybrid solar over the Southern Region states in MNRE's year-end tables, off-grid left out: 26.74001, 29.21533 and 35.65423 GW at 31 December 2023, 2024 and 2025. Each year's value is that stock less the new capacity the model must build from 2023 to that year (0.1 GW in 2023, 0.17 GW in 2024, 1.2035 GW in 2025), so 2023 = 26.74001 - 0.1 = 26.64001 GW and 2025 = 35.65423 - 1.4735 = 34.18073 GW.
- Assumptions: A11, A12.

### `PWRSPVINDSO` ResidualCapacity, 2026 to 2050

- Values: 25.035 to 34.18073 (4 distinct values); GW; 25 rows.
- Label: `MNRE_OBSERVED_NET_COHORTS_RETAINED_25_YEAR_MODEL_LIFE`.
- Sources: [R94] to [R96] PDF p.1; Ground Mounted Solar, Rooftop/RTS, Hybrid Solar and Off-grid columns; Andhra Pradesh, Karnataka, Kerala, Tamil Nadu, Telangana, Lakshadweep, Puducherry.
- Derivation: The 2025 value, 34.18073 GW, holds from 2026 to 2047. The capacity added in 2023, 2024 and 2025 above the 25.035 GW that stays to 2050 (1.60501, 2.30532 and 5.2354 GW) retires 25 years later: 2048 = 34.18073 - 1.60501 = 32.57572 GW, 2049 = 30.2704 GW and 2050 = 25.035 GW.
- Assumptions: A11, A12. India South solar stock from 2026 to 2047 is the published stock at 31 December 2025 (35.65 GW, MNRE) less the new capacity the model must build in 2023 to 2025 (1.47 GW), so the model's 2025 capacity equals MNRE's figure. Capacity added in 2023, 2024 and 2025 retires in 2048, 2049 and 2050; the remaining 25.04 GW stays in service to 2050.

### `PWRURNINDEA`, `PWRURNINDNE`, `PWRURNINDNO`, `PWRURNINDSO`, `PWRURNINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 0 to 84.1528548 (13 distinct values); PJ per year; 20 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R34] PDF p.1; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R36] PDF p.1; regional or state rows; April-December cumulative.
- Derivation: Calendar-year nuclear generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The lower limit is GWh x 0.98 x 0.0036 PJ/GWh in 2023 to 2025 and 0.95 x the 2025 GWh x 0.0036 in 2026. Example: `PWRURNINDSO` 2023, (2,206.49 + 2,062.88 + 2,354.98 + 17,013.23) = 23,637.58 GWh x 0.98 x 0.0036 = 83.393382 PJ per year.
- Assumptions: A1, A2.

### `PWRURNINDEA`, `PWRURNINDNE`, `PWRURNINDNO`, `PWRURNINDSO`, `PWRURNINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 0 to 87.5876652 (10 distinct values); PJ per year; 15 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.26 and 27; regional or state rows; current month; [R34] PDF p.1; regional or state rows; current month; [R25] PDF pp.25 and 26; regional or state rows; current month; [R26] PDF pp.33 and 34; regional or state rows; April-December cumulative; [R27] PDF pp.32 and 33; regional or state rows; current month; [R24] PDF pp.33 and 34; regional or state rows; current month; [R28] PDF pp.32 and 33; regional or state rows; current month; [R29] PDF pp.38 and 39; regional or state rows; April-December cumulative; [R30] PDF pp.38 and 39; regional or state rows; current month; [R31] PDF pp.37 and 38; regional or state rows; current month; [R32] PDF pp.39 and 40; regional or state rows; current month; [R36] PDF p.1; regional or state rows; April-December cumulative.
- Derivation: Calendar-year nuclear generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The upper limit is GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWRURNINDSO` 2023, 23,637.58 GWh x 1.02 x 0.0036 = 86.797194 PJ per year.
- Assumptions: A1, A2.

### `PWRWASINDEA` ResidualCapacity, 2023 to 2050

- Values: 0.0058; GW; 28 rows.
- Label: `USER_DEFINED_RESIDUAL_LOST_IN_SCENARIO_BUILD`.
- Sources: [R22] PDF page 1 (single table); columns 'Waste to Energy (MW)' and 'Waste to Energy (Off-grid) (MW)'; rows Bihar, West Bengal.
- Derivation: Grid and off-grid waste-to-energy over the Eastern Region states in MNRE's table at 31 March 2023: 1.32 + 4.48 = 5.80 MW = 0.0058 GW in every year.
- Assumptions: A11, A13.

### `PWRWASINDEA`, `PWRWASINDNE`, `PWRWASINDNO`, `PWRWASINDSO`, `PWRWASINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 0.00143208 to 5.64779304 (15 distinct values); PJ per year; 15 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Others regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year waste-to-energy (CEA "Others") generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The upper limit is GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWRWASINDEA` 2023, (120.99 + 121.94 + 137.96 + 1,157.18) = 1,538.07 GWh x 1.02 x 0.0036 = 5.647793 PJ per year.
- Assumptions: A1, A2.

### `PWRWASINDNO` ResidualCapacity, 2023 to 2050

- Values: 0.24141; GW; 28 rows.
- Label: `USER_DEFINED_RESIDUAL_LOST_IN_SCENARIO_BUILD`.
- Sources: [R22] PDF page 1 (single table); columns 'Waste to Energy (MW)' and 'Waste to Energy (Off-grid) (MW)'; rows Haryana, Himachal Pradesh, Punjab, Rajasthan, Uttar Pradesh, Uttarakhand, Delhi.
- Derivation: Waste-to-energy over the Northern Region states in MNRE's table at 31 March 2023: grid 11.20 + 10.75 + 84.00 = 105.95 MW and off-grid 7.57 + 1.00 + 15.37 + 3.83 + 98.47 + 9.22 = 135.46 MW, total 241.41 MW = 0.24141 GW in every year.
- Assumptions: A11, A13.

### `PWRWASINDNO`, `PWRWASINDSO`, `PWRWASINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 0.181692 to 2.58277824 (12 distinct values); PJ per year; 12 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Others regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year waste-to-energy (CEA "Others") generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The lower limit is GWh x 0.98 x 0.0036 PJ/GWh in 2023 to 2025 and 0.95 x the 2025 GWh x 0.0036 in 2026. Example: `PWRWASINDNO` 2023, (35.98 + 43.19 + 47.33 + 503.65) = 630.15 GWh x 0.98 x 0.0036 = 2.223169 PJ per year.
- Assumptions: A1, A2.

### `PWRWASINDSO` ResidualCapacity, 2023 to 2050

- Values: 0.18876; GW; 28 rows.
- Label: `USER_DEFINED_RESIDUAL_LOST_IN_SCENARIO_BUILD`.
- Sources: [R22] PDF page 1 (single table); columns 'Waste to Energy (MW)' and 'Waste to Energy (Off-grid) (MW)'; rows Andhra Pradesh, Karnataka, Kerala, Tamil Nadu, Telangana.
- Derivation: Waste-to-energy over the Southern Region states in MNRE's table at 31 March 2023: grid 53.16 + 1.00 + 6.40 + 45.80 = 106.36 MW and off-grid 29.20 + 13.85 + 0.23 + 24.65 + 14.47 = 82.40 MW, total 188.76 MW = 0.18876 GW in every year.
- Assumptions: A11, A13.

### `PWRWASINDWE` ResidualCapacity, 2023 to 2050

- Values: 0.11806; GW; 28 rows.
- Label: `USER_DEFINED_RESIDUAL_LOST_IN_SCENARIO_BUILD`.
- Sources: [R22] PDF page 1 (single table); columns 'Waste to Energy (MW)' and 'Waste to Energy (Off-grid) (MW)'; rows Chhatisgarh, Goa, Gujarat, Madhya Pradesh, Maharashtra.
- Derivation: Waste-to-energy over the Western Region states in MNRE's table at 31 March 2023: grid 0.34 + 7.50 + 15.40 + 12.59 = 35.83 MW and off-grid 0.41 + 25.93 + 12.19 + 43.70 = 82.23 MW, total 118.06 MW = 0.11806 GW in every year.
- Assumptions: A11, A13.

### `PWRWONINDEA`, `PWRWONINDNE`, `PWRWONINDNO`, `PWRWONINDSO`, `PWRWONINDWE` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2026

- Values: 0 to 180.23973408 (13 distinct values); PJ per year; 20 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Wind regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year wind generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The lower limit is GWh x 0.98 x 0.0036 PJ/GWh in 2023 to 2025 and 0.95 x the 2025 GWh x 0.0036 in 2026. Example: `PWRWONINDSO` 2023, (1,808.23 + 1,417.57 + 1,887.13 + 32,090.85) = 37,203.78 GWh x 0.98 x 0.0036 = 131.254936 PJ per year.
- Assumptions: A1, A2.

### `PWRWONINDEA`, `PWRWONINDNE`, `PWRWONINDNO`, `PWRWONINDSO`, `PWRWONINDWE` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2025

- Values: 0 to 187.59645792 (10 distinct values); PJ per year; 15 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R23] PDF pp.30 to 32; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; prior 2023 column; [R25] PDF pp.29 to 31; regional or state rows; current month; [R26] PDF pp.42 to 44; regional or state rows; April-December cumulative; [R27] PDF pp.40 to 42; regional or state rows; current month; [R24] PDF pp.41 to 43; regional or state rows; current month; [R28] PDF pp.40 to 42; regional or state rows; current month; [R29] PDF pp.47 to 49; regional or state rows; April-December cumulative; [R30] PDF pp.46 to 48; regional or state rows; current month; [R31] PDF pp.45 to 47; regional or state rows; current month; [R32] PDF pp.47 to 49; regional or state rows; current month; [R33] Wind regional Sub Total, current cumulative Apr-Dec2025; regional or state rows; April-December cumulative.
- Derivation: Calendar-year wind generation for each region is January + February + March (monthly figures) + April to December (cumulative figure) from CEA's reports, in GWh; where a printed regional sub-total differs from its state rows, the state rows are summed. The upper limit is GWh x 1.02 x 0.0036 PJ/GWh. Example: `PWRWONINDSO` 2023, 37,203.78 GWh x 1.02 x 0.0036 = 136.612280 PJ per year.
- Assumptions: A1, A2.

### `TRNINDEAINDWE` TotalAnnualMaxCapacity, 2023 to 2024

- Values: 21.19; GW; 2 rows.
- Label: `INTERCONNECTOR_PARAMS_AND_CAP_TRN_TO_RESIDUAL`.
- Sources: [R100] printed p.92 (PDF p.134), section 7.5 Table 7.8 'Inter-Regional Transmission Capacity considering capacity of HVDC links in reverse direction (MW)', row West-East; [R101] printed p.181 (PDF p.185), EAST-WEST Sub-total; [R102] PDF p.4, table 'Growth of Inter- Regional Transmission Capacity (220 kV and above)', column ER-WR, rows 'Up to March-2023' and 'Up to March-2024'.
- Derivation: CEA's East-West inter-regional capacity is 21,190 MW = 21.19 GW in 2023 and 2024. The corridor has no HVDC link, so the one-direction sum and the reverse-direction sum are the same.
- Assumptions: A14.

### `TRNINDEAINDWE` TotalAnnualMaxCapacity, 2025 to 2050

- Values: 22.79; GW; 26 rows.
- Label: `INTERCONNECTOR_PARAMS_AND_CAP_TRN_TO_RESIDUAL`.
- Sources: [R100] printed p.92 (PDF p.134), section 7.5 Table 7.8 'Inter-Regional Transmission Capacity considering capacity of HVDC links in reverse direction (MW)', row West-East; [R100] printed p.398 onwards (PDF pp.440-441), Annex 7.3 'Inter-regional Transmission Links and Capacity (MW) till 2026-27', EAST-WEST row 'Jeypore-Jagdalpur 400 kV D/c line (Under Construction)'; [R102] PDF p.4, table 'Growth of Inter- Regional Transmission Capacity (220 kV and above)', column ER-WR, rows 'Up to March-2025' and 'Up to 31st March - 2026'; [R114] PDF p.4, table 'Growth of Inter- Regional Transmission Capacity (220 kV and above)', column ER-WR, row 'Up to 30th June - 2025'.
- Derivation: 21,190 MW + the Jeypore-Jagdalpur 400 kV double-circuit line, 1,600 MW = 22,790 MW = 22.79 GW from 2025 to 2050. CEA's growth statement shows 22,790 MW at 31 March 2026. The Jeypore-Jagdalpur line entered service in June 2025 (CEA's monthly growth statement for June 2025, ER-WR 22,790 MW), so it counts from 2025.
- Assumptions: A3, A11, A14.

### `TRNINDNOINDWE` TotalAnnualMaxCapacity, 2023

- Values: 29.22; GW; 1 rows.
- Label: `INTERCONNECTOR_PARAMS_AND_CAP_TRN_TO_RESIDUAL`.
- Sources: [R100] printed p.92 (PDF p.134), section 7.5 Table 7.8 'Inter-Regional Transmission Capacity considering capacity of HVDC links in reverse direction (MW)', row North-West, column 'At the end of 2021-22 (31.03.2022)'; [R102] PDF p.4, table 'Growth of Inter- Regional Transmission Capacity (220 kV and above)', column WR-NR, rows '2017-2022' and 'Up to March-2023'; [R101] printed p.181 (PDF p.185), WEST- NORTH Sub-total.
- Derivation: 29,220 MW = 29.22 GW in 2023: Table 7.8 counts HVDC links at their reverse capability, and the West-to-North sum, 36,720 MW, is 7,500 MW higher, mainly because the 6,000 MW Champa-Kurukshetra HVDC link cannot run in reverse. The West-North sum is unchanged at 36,720 MW to March 2023.
- Assumptions: A14.

### `TRNINDNOINDWE` TotalAnnualMaxCapacity, 2024 to 2025

- Values: 30.82; GW; 2 rows.
- Label: `INTERCONNECTOR_PARAMS_AND_CAP_TRN_TO_RESIDUAL`.
- Sources: [R100] printed p.92 (PDF p.134), section 7.5 Table 7.8 'Inter-Regional Transmission Capacity considering capacity of HVDC links in reverse direction (MW)', row North-West, column 'At the end of 2023-24 (31.03.2024)'; [R100] printed p.398 onwards (PDF pp.440-441), Annex 7.3 'Inter-regional Transmission Links and Capacity (MW) till 2026-27', WEST- NORTH row 'Neemuch PS - Chhittorgarh 400 kV D/c line (Commissioned)'; [R102] PDF p.4, table 'Growth of Inter- Regional Transmission Capacity (220 kV and above)', column WR-NR, rows 'Up to March-2024' and 'Up to March-2025'.
- Derivation: 29,220 MW + the Neemuch PS-Chittorgarh 400 kV double-circuit line, 1,600 MW = 30,820 MW = 30.82 GW, printed in Table 7.8 for 31 March 2024; it holds in 2024 and 2025.
- Assumptions: A14.

### `TRNINDNOINDWE` TotalAnnualMaxCapacity, 2026 to 2027

- Values: 30.82; GW; 2 rows.
- Label: `INTERCONNECTOR_PARAMS_AND_CAP_TRN_TO_RESIDUAL`.
- Sources: [R100] printed p.92 (PDF p.134), section 7.5 Table 7.8 'Inter-Regional Transmission Capacity considering capacity of HVDC links in reverse direction (MW)', row North-West, column 'At the end of 2023-24 (31.03.2024)'; [R102] PDF p.4, table 'Growth of Inter- Regional Transmission Capacity (220 kV and above)', column WR-NR, rows 'Up to March-2024' and 'Up to 31st March - 2026'.
- Derivation: 30,820 MW = 30.82 GW in 2026 and 2027. CEA's one-direction West-North sum is 38,320 MW in March 2024 and still 38,320 MW on 31 March 2026.
- Assumptions: A14. The North-West corridor stays at its in-service capacity of 30.82 GW in 2026 and 2027: CEA shows no addition between 31 March 2024 and 31 March 2026, including the two 765 kV Mandsaur lines (8.4 GW), and the plan lines count from 2028.

### `TRNINDNOINDWE` TotalAnnualMaxCapacity, 2028 to 2050

- Values: 47.62; GW; 23 rows.
- Label: `INTERCONNECTOR_PARAMS_AND_CAP_TRN_TO_RESIDUAL`.
- Sources: [R100] printed p.92 (PDF p.134), section 7.5 Table 7.8, row North-West, column 'At the end of 2026-27 (31.03.2027)'; [R100] printed p.398 onwards (PDF pp.440-441), Annex 7.3 'Inter-regional Transmission Links and Capacity (MW) till 2026-27', WEST- NORTH rows 'Beawar - Mandasaur 765 kV D/c line (Under Bidding)', 'Rishabhdeo - Mandasaur 765 kV D/c line (Under Bidding)', 'Sirohi - Mandasaur 765 kV D/c line (Planned)' and 'Sasan - Prayagraj 765 kV D/c line (Planned)'.
- Derivation: 47,620 MW = 47.62 GW from 2028 to 2050, Table 7.8's value for 31 March 2027: 30,820 MW in service plus the four 765 kV lines of Annex 7.3, 4 x 4,200 = 16,800 MW.
- Assumptions: A11, A14. The four 765 kV North-West lines in CEA's 2024 transmission plan (16.8 GW: two under bidding and two planned, due by 31 March 2027) are in service from 2028, the first full calendar year after the plan date, and the corridor holds 47.62 GW to 2050.

### `TRNINDNOINDWE` TotalAnnualMinCapacityInvestment, 2026 and 2028

- Values: 0; 12.6; GW; 2 rows.
- Label: `CEA_IN_SERVICE_STATEMENT_AND_NEP_PLAN_DATE`.
- Sources: [R102] PDF p.4, table 'Growth of Inter- Regional Transmission Capacity (220 kV and above)', column WR-NR, row 'Up to 31st March - 2026'; [R100] printed p.398 onwards (PDF pp.440-441), Annex 7.3, WEST- NORTH rows 'Beawar - Mandasaur 765 kV D/c line (Under Bidding)', 'Rishabhdeo - Mandasaur 765 kV D/c line (Under Bidding)', 'Sirohi - Mandasaur 765 kV D/c line (Planned)' and 'Sasan - Prayagraj 765 kV D/c line (Planned)'.
- Derivation: 2026 = 0 GW, because CEA shows no North-West addition by 31 March 2026. 2028 = 4,200 + 4,200 MW (the two Mandsaur lines under bidding) + 4,200 MW (one of the two planned lines) = 12,600 MW = 12.6 GW.
- Assumptions: A14. The model must build 12.6 GW on the North-West corridor in 2028: the two 765 kV Mandsaur lines under bidding (8.4 GW) and one of the two planned lines (Sirohi-Mandsaur or Sasan-Prayagraj, 4.2 GW). The other planned line is allowed but not required.

## Nepal

### `PWRBIONPLXX` ResidualCapacity, 2023 to 2050

- Values: 0.006; GW; 28 rows.
- Label: `NEA_OFFICIAL_ROSTER_COD_WEIGHTED_ONCE`.
- Sources: [R103] PDF p.161 BAGASSE (IPPs) and p.154 (printed 144) Total Bagasse (IPPs); [R104] PDF p.206 (printed 196) BAGASSE (IPPs); [R105] PDF p.168 BAGASSE (IPPs), image page.
- Derivation: Two bagasse plants, Indushankar Chini Udhyog and Everest Sugar and Chemical Industries, 3,000 kW each, were in service before FY2022-23: 6,000 kW = 0.006 GW in every year.
- Assumptions: A11.

### `PWRHYDNPLXX` ResidualCapacity, 2023 to 2050

- Values: 2.27309920274 to 3.388251 (4 distinct values); GW; 28 rows.
- Label: `NEA_OFFICIAL_ROSTER_COD_WEIGHTED_ONCE`.
- Sources: [R105] PDF p.162 (capacity summary, image page) and pp.165-168 (IPP and subsidiary operating list, image pages); [R104] PDF p.199 (printed 189, capacity summary) and pp.202-206 (printed 192-196, operating list with COD); [R103] PDF p.154 (printed 144, capacity summary) and pp.157-161 (operating list with COD).
- Derivation: Nepal's hydro stock is NEA's own plants (583.160 MW less 1.661 MW not operating) plus each subsidiary and IPP plant weighted by its days in service in the fiscal year, from NEA's operating lists with commercial operation dates. Example: 2023 = 0.583160 - 0.001661 + 0.478100 (subsidiaries) + 1.213500 (IPPs, day-weighted) = 2.273099 GW. From 2026 the stock is the FY2024-25 list, 3,389,912 kW less 1,661 kW = 3.388251 GW, with no retirement.
- Assumptions: A3, A4, A11.

### `PWRHYDNPLXX`, `PWRSPVNPLXX`, `PWRBIONPLXX` TotalAnnualMaxCapacityInvestment, 2023 to 2025

- Values: 0; GW; 9 rows.
- Label: `NEA_OFFICIAL_ROSTER_REPLACES_MODEL_PLACEHOLDER_INVESTMENTS`.
- Sources: [R105] PDF pp.162, 165-168; [R104] PDF pp.199, 202-206; [R103] PDF pp.154, 157-161.
- Derivation: Every plant on NEA's three operating lists, including those commissioned during the year, is already in ResidualCapacity, so the most new investment allowed in 2023 to 2025 is 0 GW.
- Assumptions: A3, A12. Plants commissioned from mid-July to December 2025 fall outside FY2024/25 and are not counted in 2025.

### `PWRHYDNPLXX`, `PWRSPVNPLXX`, `PWRBIONPLXX` TotalAnnualMinCapacityInvestment, 2023 to 2025

- Values: 0; GW; 9 rows.
- Label: `NEA_OFFICIAL_ROSTER_REPLACES_MODEL_PLACEHOLDER_INVESTMENTS`.
- Sources: [R105] PDF pp.162, 165-168; [R104] PDF pp.199, 202-206; [R103] PDF pp.154, 157-161.
- Derivation: Every plant on NEA's three operating lists, including those commissioned during the year, is already in ResidualCapacity, so the required new investment in 2023 to 2025 is 0 GW.
- Assumptions: A3, A12. Plants commissioned from mid-July to December 2025 fall outside FY2024/25 and are not counted in 2025.

### `PWRHYDNPLXX`, `PWRSPVNPLXX`, `PWRBIONPLXX` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2025

- Values: 0.034042864464 to 48.055541258592 (9 distinct values); PJ per year; 9 rows.
- Label: `NEPAL_NATIONAL_TOTAL_AND_OFFICIAL_OPERATING_ROSTER`.
- Sources: [R105] PDF p.13 (printed 11), Managing Director's review; same figures on PDF p.127; [R104] PDF p.12 (printed 2), Managing Director's review; same figures on PDF p.162; [R103] PDF p.119, Operational Performance.
- Derivation: NEA's national generation is NEA plants + subsidiaries + IPP purchases: 2,930 + 2,488 + 5,118 = 10,536 GWh in FY2022-23, 12,072 GWh in FY2023-24 and 13,959 GWh in FY2024-25. Each total is split by day-weighted capacity times a typical capacity factor (hydro 0.45, solar 0.1925, bagasse 0.17, thermal 0.25), each share capped at its physical maximum, and the lower limit is 0.98 x share x 0.0036 PJ/GWh. Example: 2023 hydro, 10,268.02 GWh x 0.98 x 0.0036 = 36.225563 PJ per year.
- Assumptions: A3, A4, A9.

### `PWRSPVNPLXX` ResidualCapacity, 2023 to 2050

- Values: 0 to 0.1419404 (11 distinct values); GW; 28 rows.
- Label: `NEA_OFFICIAL_ROSTER_COD_WEIGHTED_ONCE`.
- Sources: [R103] PDF p.154 (printed 144) Total Solar (NEA) and Total Solar (IPPs), p.161 solar IPP list with COD; FY2022/23 PDF p.168 and FY2023/24 PDF p.206 solar IPP lists; [R106] PDF p.14, Component 1 paragraph; [R107] PDF p.3, indicator IN01163233 table; p.1 component narrative.
- Derivation: NEA solar is 19.4 MW all year plus 5.6 MW for 319 of 365 days in 2023, and IPP solar plants are weighted by their days in service: 2023 = 19.400 + 4.894 + 45.568 = 69.862 MW, and 2026 = 25,000 + 116,940.4 kW = 0.1419404 GW. Each IPP plant retires 25 years after its commissioning year and the NEA 25 MW in 2048, so 2048 = 0.055 GW, 2049 = 0.035 GW and 2050 = 0.
- Assumptions: A3, A4, A11. The last 5.6 MW of the NEA solar farms counts from 1 September 2022, since the World Bank reports the final site in operation in August 2022 without a day.

## Sri Lanka

### `PWRCOALKAXX`, `PWRHYDLKAXX`, `PWROILLKAXX`, `PWRSPVLKAXX`, `PWRWONLKAXX` TotalTechnologyAnnualActivityLowerLimit, 2023 to 2024

- Values: 2.794176 to 24.339672 (10 distinct values); PJ per year; 10 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R108] PDF p.6 (printed p.2) text: total hydro 5,951 GWh in 2023 and 6,899 GWh in 2024, CEB hydro 5,426 GWh, mini hydro 1,473 GWh in 2024 and 1,378 GWh in 2023, wind 818 and 792 GWh; PDF p.7 (printed p.3) chart NCRE Generation 2023 and 2024; PDF p.73 (printed p.69) Summary Statistics of Generation 2024, gross generation and energy used at power stations; [R109] PDF p.3 (printed p.1) GENERATION STATISTICS, table Net Generation (GWh) by ownership and source, columns 2024 and 2025; PDF p.2 Summary of Statistics 2025.
- Derivation: Coal is CEB coal, hydro CEB major hydro + IPP mini hydro, oil CEB oil + IPP oil + small islands, solar IPP grid-connected + rooftop, and wind CEB + IPP wind, in GWh; the lower limit is 0.98 x GWh x 0.0036 PJ/GWh. Example: coal 2023, 5,150.767 GWh gross less 505.242 GWh station use = 4,646 GWh x 0.98 x 0.0036 = 16.391088 PJ per year.
- Assumptions: A1.

### `PWRCOALKAXX`, `PWRHYDLKAXX`, `PWROILLKAXX`, `PWRSPVLKAXX`, `PWRWONLKAXX` TotalTechnologyAnnualActivityLowerLimit, 2025 to 2026

- Values: 2.64708 to 27.274968 (10 distinct values); PJ per year; 10 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R109] PDF p.3 (printed p.1) GENERATION STATISTICS, table Net Generation (GWh) by ownership and source, columns 2024 and 2025; PDF p.2 Summary of Statistics 2025.
- Derivation: Sources map to technologies as in the 2023 and 2024 limits; the lower limit is 0.98 x the 2025 GWh x 0.0036 PJ/GWh in 2025 and 0.95 x the 2025 GWh x 0.0036 in 2026. Example: hydro 2025, (6,306 + 1,425) GWh x 0.98 x 0.0036 = 27.274968 PJ per year.
- Assumptions: A1, A2.

### `PWRCOALKAXX`, `PWRHYDLKAXX`, `PWROILLKAXX`, `PWRSPVLKAXX`, `PWRWONLKAXX` TotalTechnologyAnnualActivityUpperLimit, 2023 to 2024

- Values: 2.908224 to 25.333128 (10 distinct values); PJ per year; 10 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R108] PDF p.6 (printed p.2) text: total hydro 5,951 GWh in 2023 and 6,899 GWh in 2024, CEB hydro 5,426 GWh, mini hydro 1,473 GWh in 2024 and 1,378 GWh in 2023, wind 818 and 792 GWh; PDF p.7 (printed p.3) chart NCRE Generation 2023 and 2024; PDF p.73 (printed p.69) Summary Statistics of Generation 2024, gross generation and energy used at power stations; [R109] PDF p.3 (printed p.1) GENERATION STATISTICS, table Net Generation (GWh) by ownership and source, columns 2024 and 2025; PDF p.2 Summary of Statistics 2025.
- Derivation: Coal is CEB coal, hydro CEB major hydro + IPP mini hydro, oil CEB oil + IPP oil + small islands, solar IPP grid-connected + rooftop, and wind CEB + IPP wind, in GWh; the upper limit is 1.02 x GWh x 0.0036 PJ/GWh. Example: oil 2023, 3,134.7 GWh x 1.02 x 0.0036 = 11.510618 PJ per year.
- Assumptions: A1.

### `PWRCOALKAXX`, `PWRHYDLKAXX`, `PWROILLKAXX`, `PWRSPVLKAXX`, `PWRWONLKAXX` TotalTechnologyAnnualActivityUpperLimit, 2025

- Values: 2.842128 to 28.388232 (5 distinct values); PJ per year; 5 rows.
- Label: `GENERATION_CALIBRATION_EVIDENCE`.
- Sources: [R109] PDF p.3 (printed p.1) GENERATION STATISTICS, table Net Generation (GWh) by ownership and source, columns 2024 and 2025; PDF p.2 Summary of Statistics 2025.
- Derivation: Sources map to technologies as in the 2023 and 2024 limits; the 2025 upper limit is 1.02 x GWh x 0.0036 PJ/GWh, and 2026 has none. Example: coal 2025, 4,870 GWh x 1.02 x 0.0036 = 17.882640 PJ per year.
- Assumptions: A1.

### `PWRHYDLKAXX` CapacityFactor, 2023 to 2026

- Values: 0.1788500127191547 to 0.8373299102841576 (60 distinct values); 20 time slices per year; capacity factor (fraction); 80 rows.
- Label: `PUCSL_SOURCE_YEAR_HYDRO_WITH_RETAINED_AF`.
- Sources: [R110] PDF p.12 (printed 11) chart Major Hydro monthly GWh 2023 vs 2024; PDF p.13 (printed 12) chart Mini Hydro 2023 vs 2024; PDF p.4 (printed 3) Renewable Energy Capacity as at 31.12.2024; PDF p.5 (printed 4) Annual Cumulative Renewable Generation 2024; [R111] PDF p.12 chart Major Hydro 2024 vs 2025; PDF p.13 chart Mini Hydro 2024 vs 2025; PDF pp.3-4 installed capacity and Annual Cumulative Renewable Generation 2025; [R112] PDF p.4 (printed 3) Renewable Energy Capacity as at 31.12.2023; PDF p.5 (printed 4) Annual Cumulative Renewable Generation 2023.
- Derivation: Monthly major and mini hydro generation is read from PUCSL's line charts, and each year's twelve months sum to the printed annual totals, 5,963, 6,899 and 7,708 GWh. The seasonal factor is season GWh / (year-end stock GW x 24 h x season days), spread over the five day parts in proportion to a fixed daily hydro shape and divided by an availability of 0.95. Example: 2023 S1, 1,669.1 GWh / (1.835 GW x 24 h x 90 days) = 0.421104 before availability.
- Assumptions: A5, A6.

### `PWRHYDLKAXX` ResidualCapacity, 2023 to 2050

- Values: 1.8345 to 1.963 (3 distinct values); GW; 28 rows.
- Label: `PUCSL_YEAR_END_HYDRO_STOCK_EXPOSURE_IN_MONTHLY_PROFILE_ONCE`.
- Sources: [R112] PDF p.4 (printed 3) chart Renewable Energy Capacity (MW) as at 31.12.2023; [R110] PDF p.4 (printed 3) Renewable Energy Capacity (MW) as at 31.12.2024; [R111] PDF p.3 (printed 2) Installed Capacity (MW) by Source as at 31.12.2025, and PDF p.4 Renewable Energy Capacity as at 31.12.2025.
- Derivation: The stock is major hydro + IPP mini hydro at 31 December, less the new hydro the model must build (0.01 GW in 2025): 2023 = 1,412.8 + 421.7 = 1,834.5 MW, 2024 = 1,531 + 432 = 1,963 MW and 2025 = 1,531 + 437 - 10 = 1,958 MW, held to 2050. PUCSL prints 1,531 MW on PDF p.3 and 1,533 MW on PDF p.4 of its end-2025 report; 1,531 MW is used.
- Assumptions: A4, A11, A12.

### `PWRHYDLKAXX`, `PWRLDSLKAXX` TotalAnnualMinCapacityInvestment, 2034

- Values: `PWRHYDLKAXX` 0; `PWRLDSLKAXX` 0.6; GW; 2 rows.
- Label: `PLANNED_GENERATION_AND_EXISTING_LONG_STORAGE_MAPPING`.
- Sources: [R113] Table 10.1 Base Case Scenario 2025-2044, year 2034, PDF p.171 (printed 10-3); Table 10.2 Annual Capacity Additions, year 2034, pumped storage column, PDF p.174 (printed 10-6); Chapter 10 summary item c), PDF p.169 (printed 10-1); project data 3 x 200 MW Maha Oya Pumped Storage, PDF p.124.
- Derivation: CEB's Base Case plan adds the Maha Oya pumped storage plant, 3 x 200 MW = 600 MW, in 2034 (Tables 10.1 and 10.2): 0.6 GW is the required 2034 build of long-duration storage `PWRLDSLKAXX`, and hydro `PWRHYDLKAXX` has 0 GW.
- Assumptions: The 600 MW Maha Oya pumped storage plant is built in 2034, as in CEB's Base Case plan, and is modelled as long-duration storage, not hydro generation.

### `PWRSPVLKAXX` ResidualCapacity, 2025 to 2050

- Values: 0.947 to 2.394 (4 distinct values); GW; 26 rows.
- Label: `CEB_SD2025_P0_P1_CEB_ROOFTOP_GRID_PLUS_GOVERNED_LECO_LESS_MANDATORY_INVESTMENT`.
- Sources: [R109] PDF p.2 Summary of Statistics 2025, CEB Rooftop Solar PV Total Capacity; PDF p.3 (printed p.1) Generation Statistics, IPP NCRE (Grid Connected Solar) installed capacity; [R112] PDF p.4 (printed 3) Renewable Energy Capacity (MW) as at 31.12.2023; [R110] and [R111] PDF p.4 (printed 3) Renewable Energy Capacity (MW) as at 31.12.2024 and as at 31.12.2025.
- Derivation: Solar stock for year Y is CEB rooftop + grid-connected solar + LECO rooftop at 31 December of Y-1, less the 0.2 GW the model must build in 2025: 2025 = 1,141 + 173 + 259 = 1,573 MW, and 2026 to 2048 = 1,935 + 307 + 352 - 200 = 2,394 MW. The additions counted in 2025 (0.626 GW) and 2026 (0.821 GW) retire after 25 years, so 2049 = 1.768 GW and 2050 = 0.947 GW.
- Assumptions: A4, A11. Sri Lanka solar stock in each year from 2025 is the published stock at the end of the year before (CEB rooftop and grid-connected solar, and LECO rooftop from PUCSL), less the 0.2 GW of new solar the model must build in 2025, so solar added during a year counts from the next year. CEB marks its end-2025 rooftop figure of 1,935 MW as provisional; PUCSL reports 2,120 MW for the same date, which would raise the 2026 to 2048 stock by 0.19 GW. Additions retire after 25 years; the end-2023 stock of 0.947 GW stays in service to 2050.

## Rows set by modelling choices

These rows follow one assumption each and cite no document of their own.

### Technologies with no published generation for their region

- Label: `ORIGINAL_PRODUCTION_ROW_ID_AUDIT`.
- Covers: `TotalTechnologyAnnualActivityUpperLimit` and `TotalTechnologyAnnualActivityLowerLimit` of India (242 rows) and Sri Lanka (54 rows) power technologies with no published annual generation in their model region, 2023 to 2026. These rows carry no value, so the technology has no annual generation limit.
- Rows: 296.
- Assumption: A15.

### Investment in Bangladesh and Bhutan in 2023 to 2026

- Label: `BASE_YEAR_CAPACITY_FREEZE`.
- Covers: `TotalAnnualMaxCapacityInvestment` of Bangladesh (2023 to 2026, 28 rows) and Bhutan (2023 to 2026, 12 rows) and `TotalAnnualMinCapacityInvestment` of Bangladesh (2024 to 2026, 17 rows) and Bhutan (2024 to 2026, 6 rows), all 0 GW. The 2026 coal and nuclear build caps of Bangladesh have their own block above.
- Rows: 63.
- Assumption: A12.

### Nepal generation upper limits in 2023 to 2025

- Label: `REVIEWED_NEPAL_LOWER_ONLY`.
- Covers: `TotalTechnologyAnnualActivityUpperLimit` of `PWRBIONPLXX`, `PWRHYDNPLXX` and `PWRSPVNPLXX`, 2023 to 2025. These rows carry no value, so Nepal's generation has a lower limit only.
- Rows: 9.
- Assumption: A9.

### Waste-to-energy in India East and North-East

- Label: `REVIEWED_WASTE_MAPPING`.
- Covers: `TotalTechnologyAnnualActivityLowerLimit` of `PWRWASINDEA` and `PWRWASINDNE`, 2023 to 2026, all 0 PJ per year.
- Rows: 8.
- Assumption: A10.

### Bangladesh and Bhutan generation upper limits in 2026

- Label: `LOWER_ONLY_2026`.
- Covers: `TotalTechnologyAnnualActivityUpperLimit` of six Bangladesh technologies (coal, hydro, gas, oil, solar and wind) and two Bhutan technologies (hydro and solar), 2026. These rows carry no value, so 2026 generation has a lower limit only.
- Rows: 8.
- Assumption: A2.

### Bangladesh generation upper limits in 2025

- Label: `LOWER_ONLY_2025`.
- Covers: `TotalTechnologyAnnualActivityUpperLimit` of six Bangladesh technologies (coal, hydro, gas, oil, solar and wind), 2025. These rows carry no value, so 2025 generation has a lower limit only.
- Rows: 6.
- Assumption: A16.

## Documents

Each document carries the date it was last retrieved. A document whose address no longer opens is listed with the date it was retrieved.

- **R1** Rural Power Company Limited - Norinco International Power Limited (RNPL). RNPL Company Overview (Patuakhali 1,320 MW coal power plant). <https://rnpl.com.bd/company-overview/>. Retrieved 23 October 2026.
- **R2** World Nuclear Association. Nuclear Power in Bangladesh, country profile (updated 13 October 2026). <https://world-nuclear.org/information-library/country-profiles/countries-a-f/bangladesh>. Retrieved 23 October 2026.
- **R3** Bangladesh Power Development Board (BPDB). BPDB Annual Report 2022-23 (revised). <https://objectstorage.ap-dcc-gazipur-1.oraclecloud15.com/n/axvjbnqprylg/b/V2Ministry/o/office-bpdb/2024/12/651fabb73895432c99a7f6760e1b381d.pdf>. Retrieved 23 October 2026.
- **R4** Bangladesh Power Development Board (BPDB). BPDB Annual Report 2023-24 (official copy linked from BPDB's annual-report page, hosted on Google Drive). <https://drive.google.com/file/d/1HIDtQmLmHL3EziZF8Rmezv3neEsLOUDn/view>. Retrieved 23 October 2026.
- **R5** Bangladesh Power Development Board (BPDB). BPDB Annual Report 2024-25 (official copy linked from BPDB's annual-report page, hosted on Google Drive). <https://drive.google.com/file/d/1iydhcFHm7YLBYJAZVYhFbBGd7LHB_zaE/view>. Retrieved 23 October 2026.
- **R6** Power Grid Bangladesh PLC (PGCB), National Load Dispatch Center. PGCB NLDC Daily Report (Excel), tab En-Curve 'DAILY ENERGY CURVE', curve for 1 July 2023. <https://erp.powergrid.gov.bd/web/files/download?location=erp%2Fweb%2Freport_docs%2F3825.xlsm>. Retrieved 23 October 2026.
- **R7** Power Grid Bangladesh PLC (PGCB), National Load Dispatch Center. PGCB NLDC Daily Report (Excel), tab En-Curve 'DAILY ENERGY CURVE', curve for 30 June 2025. <https://erp.powergrid.gov.bd/web/files/download?location=erp%2Fweb%2Freport_docs%2F4635.xlsx>. Retrieved 23 October 2026.
- **R8** Bangladesh-China Renewable Energy Company (Pvt.) Ltd (BCRECL). Pabna 64 MW Solar Park, project page. <https://www.bcrecl.com.bd/project/pabna-64-mw-solar-park>. Retrieved 23 October 2026.
- **R9** Electricity Generation Company of Bangladesh Ltd (EGCB). EGCB Annual Report 2024 (FY2023-24), Sonagazi 75 MW Solar Power Plant general information. <https://objectstorage.ap-dcc-gazipur-1.oraclecloud15.com/n/axvjbnqprylg/b/V2Ministry/o/office-egcb/2024/12/2abd0c09d1b7441587a1e7f8bc1b94e2.pdf>. Retrieved 23 October 2026.
- **R10** International Renewable Energy Agency (IRENA). Renewable Capacity Statistics 2026. <https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2026/Mar/IRENA_DAT_RE_Capacity_Statistics_2026.pdf>. Retrieved 23 October 2026.
- **R11** Royal Monetary Authority of Bhutan (RMA). RMA Monthly Statistical Bulletin January 2025. <https://www.rma.org.bt/media/Publication/Macro-economic%20Data/Monthly%20Statistical%20Bulletin%20January%202025.xlsx>. Retrieved 23 October 2026.
- **R12** Royal Monetary Authority of Bhutan (RMA). RMA Monthly Statistical Bulletin February 2026. <https://www.rma.org.bt/media/Publication/Macro-economic%20Data/Monthly%20Statistical%20Bulletin%20February%202026.xlsx>. Retrieved 23 October 2026.
- **R13** Royal Monetary Authority of Bhutan (RMA). RMA Monthly Statistical Bulletin August 2026. <https://www.rma.org.bt/media/Publication/Macro-economic%20Data/Monthly%20Statistical%20Bulletin%20August%202026.xlsx>. Retrieved 23 October 2026.
- **R14** Druk Green Power Corporation Limited (DGPC). DGPC Annual Report 2024. <https://www.drukgreen.bt/wp-content/uploads/2026/05/DGPC-annual-report-2024-1-1.pdf>. Retrieved 23 October 2026.
- **R15** Punatsangchhu-II Hydroelectric Project Authority (PHPA-II). Punatsangchhu-II Hydroelectric Project, Electro-Mechanical page. <https://phpa2.gov.bt/electro-mechanical/>. Retrieved 17 October 2026.
- **R16** Druk Green Power Corporation Limited (DGPC), filed with the Electricity Regulatory Authority (ERA). Proposal of Domestic Generation Tariff of DGPC Power Plants (July 2025-June 2028), December 2025. <https://era.gov.bt/wp-content/uploads/2025/12/DGPC-Tariff-Proposal-2025.pdf>. Retrieved 23 October 2026.
- **R17** Bhutan Power Corporation Limited (BPC). Power Data Book 2025, Table 2.1 Major Hydropower Generation, Embedded Generations and Solar Power (MU). <https://www.bpc.bt/wp-content/themes/bpc/assets/downloads/PDB_2025.pdf>. Retrieved 23 October 2026.
- **R18** Ministry of Energy and Natural Resources (MoENR), Department of Energy. Sephu Solar Project (Department of Energy announcement of Phase I commissioning), 19 July 2025. <https://www.moenr.gov.bt/?p=15007>. Retrieved 23 October 2026.
- **R19** Ministry of Finance, Royal Government of Bhutan. Budget Report for FY 2026-27 (National Budget Financial Year 2026-27). <https://mof.gov.bt/wp-content/uploads/2026/05/Budget-Report-FY-2026-27-English-1.pdf>. Retrieved 23 October 2026.
- **R20** Bhutan Power Corporation Limited (BPC). BPC Annual Report 2024. <https://www.bpc.bt/wp-content/themes/bpc/assets/downloads/BPC_Annual_Report_2024.pdf>. Retrieved 23 October 2026.
- **R21** Electricity Regulatory Authority (ERA), Bhutan. Tariff Review Report of 22.38 MWp Sephu Solar Project, July 2025. <https://era.gov.bt/wp-content/uploads/2025/12/Sephu_Review_Report-1.pdf>. Retrieved 23 October 2026.
- **R22** Ministry of New and Renewable Energy (MNRE), P&C Division, Government of India. State-wise installed capacity of Renewable Power as on 31.03.2023. <https://cdnbbsr.s3waas.gov.in/s3716e1b8c6cd17b771da77391355749f3/uploads/2023/04/2023041165.pdf>. Retrieved 23 October 2026.
- **R23** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, January 2023 (Internet Archive copy). <https://web.archive.org/web/20261026034215/https://cea.nic.in/api/dashboardReports/download-pdf/en/295192f6-1d47-4013-92f0-4b885444a495>. Retrieved 26 October 2026.
- **R24** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, February 2024 (Internet Archive copy). <https://web.archive.org/web/20240813003439/https://cea.nic.in/wp-content/uploads/resd/2024/03/Broad_Overview_of_RE_Generation_February_2024.pdf>. Retrieved 26 October 2026.
- **R25** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, March 2023 (Internet Archive copy). <https://web.archive.org/web/20240229184450/https://cea.nic.in/wp-content/uploads/resd/2023/03/Broad_overview_of_RE_Generation_March_2023-1.pdf>. Retrieved 26 October 2026.
- **R26** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, December 2023 (Internet Archive copy). <https://web.archive.org/web/20240714201354/https://cea.nic.in/wp-content/uploads/resd/2024/01/Broad_Overview_of_RE_Generation_December_2023.pdf>. Retrieved 26 October 2026.
- **R27** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, January 2024 (Internet Archive copy). <https://web.archive.org/web/20261026034253/https://cea.nic.in/api/dashboardReports/download-pdf/en/21f7ee42-68e4-4198-93bc-5a0255e780a4>. Retrieved 26 October 2026.
- **R28** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, March 2024 (Internet Archive copy). <https://web.archive.org/web/20240704041727/https://cea.nic.in/wp-content/uploads/resd/2024/04/Broad_Overview_of_RE_Generation_March_2024.pdf>. Retrieved 26 October 2026.
- **R29** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, December 2024 (Internet Archive copy). <https://web.archive.org/web/20250423074614/https://cea.nic.in/wp-content/uploads/resd/2024/12/Broad_Overview_of_RE_Generation_Dec_2024.pdf>. Retrieved 26 October 2026.
- **R30** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, January 2025 (Internet Archive copy). <https://web.archive.org/web/20261026034331/https://cea.nic.in/api/dashboardReports/download-pdf/en/59a96a87-b482-448c-8145-e8eeb6e985e9>. Retrieved 26 October 2026.
- **R31** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, February 2025 (Internet Archive copy). <https://web.archive.org/web/20250426045818id_/https://cea.nic.in/wp-content/uploads/resd/2025/03/Broad_Overview_of_RE_Generation_February_2025.pdf>. Retrieved 23 October 2026.
- **R32** Central Electricity Authority. Broad Overview of Monthly Renewable Energy Generation, March 2025 (Internet Archive copy). <https://web.archive.org/web/20260131172922/https://cea.nic.in/wp-content/uploads/resd/2025/04/Broad_Overview_of_RE_Generation_March_2025.pdf>. Retrieved 26 October 2026.
- **R33** Central Electricity Authority. Monthly Report of Renewable Generation, December 2025 (Internet Archive copy). <https://web.archive.org/web/20261026034408/https://cea.nic.in/api/dashboardReports/download-excel/en/0268bded-105f-4641-b00e-a17ecb398364>. Retrieved 26 October 2026.
- **R34** Central Electricity Authority / National Power Portal. 18-column actual generation report, February 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_act/2023/FEB//18_col_act-2_2023-FEB.pdf>. Retrieved 23 October 2026.
- **R35** Central Electricity Authority / National Power Portal. 18-column actual generation report, February 2023, sub-report 16 (state-wise details). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_act/2023/FEB//18_col_act-16_2023-FEB.pdf>. Retrieved 23 October 2026.
- **R36** Central Electricity Authority / National Power Portal. 18-column actual generation report, December 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_act/2025/DEC//18_col_act-2_2025-DEC.pdf>. Retrieved 23 October 2026.
- **R37** Central Electricity Authority / National Power Portal. 18-column actual generation report, December 2025, sub-report 16 (state-wise details). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_act/2025/DEC//18_col_act-16_2025-DEC.pdf>. Retrieved 23 October 2026.
- **R38** Central Electricity Authority. Monthly Report of Renewable Generation, May 2023 (Internet Archive copy). <https://web.archive.org/web/20261026034452/https://cea.nic.in/api/dashboardReports/download-excel/en/0b4b687d-0e26-4a58-9500-dd36fd0a3ed0>. Retrieved 26 October 2026.
- **R39** Central Electricity Authority. Monthly Report of Renewable Generation, September 2023 (Internet Archive copy). <https://web.archive.org/web/20261026034535/https://cea.nic.in/api/dashboardReports/download-excel/en/a779d54c-472b-4b09-a3a0-fb10c7a7a7df>. Retrieved 26 October 2026.
- **R40** Central Electricity Authority. Monthly Report of Renewable Generation, November 2023 (Internet Archive copy). <https://web.archive.org/web/20261026034619/https://cea.nic.in/api/dashboardReports/download-excel/en/062e52e9-a6cb-4824-ba91-55590a713c53>. Retrieved 26 October 2026.
- **R41** Central Electricity Authority. Monthly Report of Renewable Generation, December 2023 (Internet Archive copy). <https://web.archive.org/web/20261026034702/https://cea.nic.in/api/dashboardReports/download-excel/en/867dabcb-be51-4179-a917-d9d4d07a654e>. Retrieved 26 October 2026.
- **R42** Central Electricity Authority. Monthly Report of Renewable Generation, January 2024 (Internet Archive copy). <https://web.archive.org/web/20261026034745/https://cea.nic.in/api/dashboardReports/download-excel/en/069f445d-e5e3-4e59-9819-6a90948f72e4>. Retrieved 26 October 2026.
- **R43** Central Electricity Authority. Monthly Report of Renewable Generation, February 2024 (Internet Archive copy). <https://web.archive.org/web/20261026034829/https://cea.nic.in/api/dashboardReports/download-excel/en/a0c37d75-8512-4623-b107-144a04e2d7e9>. Retrieved 26 October 2026.
- **R44** Central Electricity Authority. Monthly Report of Renewable Generation, March 2024 (Internet Archive copy). <https://web.archive.org/web/20261026034912/https://cea.nic.in/api/dashboardReports/download-excel/en/ae4a4022-6056-4d34-99cb-85bd39c2c2ec>. Retrieved 26 October 2026.
- **R45** Central Electricity Authority. Monthly Report of Renewable Generation, May 2024 (Internet Archive copy). <https://web.archive.org/web/20240705031112/https://cea.nic.in/wp-content/uploads/resd/2024/05/RE_Monthly_Generation_report_May_2024.xlsx>. Retrieved 26 October 2026.
- **R46** Central Electricity Authority. Monthly Report of Renewable Generation, September 2024 (Internet Archive copy). <https://web.archive.org/web/20261026035030/https://cea.nic.in/api/dashboardReports/download-excel/en/9e39d078-39dd-409f-9b43-9cb6d3a67bb0>. Retrieved 26 October 2026.
- **R47** Central Electricity Authority. Monthly Report of Renewable Generation, November 2024 (Internet Archive copy). <https://web.archive.org/web/20261026035110/https://cea.nic.in/api/dashboardReports/download-excel/en/25282775-abaa-42fe-ac22-4d5b68677fef>. Retrieved 26 October 2026.
- **R48** Central Electricity Authority. Monthly Report of Renewable Generation, December 2024 (Internet Archive copy). <https://web.archive.org/web/20250130033735/https://cea.nic.in/wp-content/uploads/resd/2024/12/RE_Monthly_Generation_report_December_2024-1.xlsx>. Retrieved 26 October 2026.
- **R49** Central Electricity Authority. Monthly Report of Renewable Generation, January 2025 (Internet Archive copy). <https://web.archive.org/web/20260626004213/https://cea.nic.in/wp-content/uploads/resd/2025/01/RE_Monthly_Generation_report_January_2025.xlsx>. Retrieved 26 October 2026.
- **R50** Central Electricity Authority. Monthly Report of Renewable Generation, February 2025 (Internet Archive copy). <https://web.archive.org/web/20250426045658/https://cea.nic.in/wp-content/uploads/resd/2025/02/RE_Monthly_Generation_report_February_2025.xlsx>. Retrieved 26 October 2026.
- **R51** Central Electricity Authority. Monthly Report of Renewable Generation, March 2025 (Internet Archive copy). <https://web.archive.org/web/20261026035228/https://cea.nic.in/api/dashboardReports/download-excel/en/f7d9c05b-e86d-4ec8-b72c-c2cc1f06c182>. Retrieved 26 October 2026.
- **R52** Central Electricity Authority. Monthly Report of Renewable Generation, May 2025 (Internet Archive copy). <https://web.archive.org/web/20261026035308/https://cea.nic.in/api/dashboardReports/download-excel/en/102c54b4-f331-43c3-8aa1-c423fbe2d7dc>. Retrieved 26 October 2026.
- **R53** Central Electricity Authority. Monthly Report of Renewable Generation, September 2025 (Internet Archive copy). <https://web.archive.org/web/20261026035345/https://cea.nic.in/api/dashboardReports/download-excel/en/0461548c-9e7d-45da-af62-cb89734fd193>. Retrieved 26 October 2026.
- **R54** Central Electricity Authority. Monthly Report of Renewable Generation, November 2025 (Internet Archive copy). <https://web.archive.org/web/20261026035512/https://cea.nic.in/api/dashboardReports/download-excel/en/bad576fa-6969-44ab-b74d-9be95d807a66>. Retrieved 26 October 2026.
- **R55** Central Electricity Authority. 18-column tentative monthly generation report, January 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/JAN//18_col_ten-2_2023-JAN.pdf>. Retrieved 23 October 2026.
- **R56** Central Electricity Authority. 18-column tentative monthly generation report, February 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/FEB//18_col_ten-2_2023-FEB.pdf>. Retrieved 23 October 2026.
- **R57** Central Electricity Authority. 18-column tentative monthly generation report, March 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/MAR//18_col_ten-2_2023-MAR.pdf>. Retrieved 23 October 2026.
- **R58** Central Electricity Authority. 18-column tentative monthly generation report, April 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/APR//18_col_ten-2_2023-APR.pdf>. Retrieved 23 October 2026.
- **R59** Central Electricity Authority. 18-column tentative monthly generation report, May 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/MAY//18_col_ten-2_2023-MAY.pdf>. Retrieved 23 October 2026.
- **R60** Central Electricity Authority. 18-column tentative monthly generation report, June 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/JUN//18_col_ten-2_2023-JUN.pdf>. Retrieved 23 October 2026.
- **R61** Central Electricity Authority. 18-column tentative monthly generation report, July 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/JUL//18_col_ten-2_2023-JUL.pdf>. Retrieved 23 October 2026.
- **R62** Central Electricity Authority. 18-column tentative monthly generation report, August 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/AUG//18_col_ten-2_2023-AUG.pdf>. Retrieved 23 October 2026.
- **R63** Central Electricity Authority. 18-column tentative monthly generation report, September 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/SEP//18_col_ten-2_2023-SEP.pdf>. Retrieved 23 October 2026.
- **R64** Central Electricity Authority. 18-column tentative monthly generation report, October 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/OCT//18_col_ten-2_2023-OCT.pdf>. Retrieved 23 October 2026.
- **R65** Central Electricity Authority. 18-column tentative monthly generation report, November 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/NOV//18_col_ten-2_2023-NOV.pdf>. Retrieved 23 October 2026.
- **R66** Central Electricity Authority. 18-column tentative monthly generation report, December 2023, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2023/DEC//18_col_ten-2_2023-DEC.pdf>. Retrieved 23 October 2026.
- **R67** Central Electricity Authority. 18-column tentative monthly generation report, January 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/JAN//18_col_ten-2_2024-JAN.pdf>. Retrieved 23 October 2026.
- **R68** Central Electricity Authority. 18-column tentative monthly generation report, February 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/FEB//18_col_ten-2_2024-FEB.pdf>. Retrieved 23 October 2026.
- **R69** Central Electricity Authority. 18-column tentative monthly generation report, March 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/MAR//18_col_ten-2_2024-MAR.pdf>. Retrieved 23 October 2026.
- **R70** Central Electricity Authority. 18-column tentative monthly generation report, April 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/APR//18_col_ten-2_2024-APR.pdf>. Retrieved 23 October 2026.
- **R71** Central Electricity Authority. 18-column tentative monthly generation report, May 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/MAY//18_col_ten-2_2024-MAY.pdf>. Retrieved 23 October 2026.
- **R72** Central Electricity Authority. 18-column tentative monthly generation report, June 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/JUN//18_col_ten-2_2024-JUN.pdf>. Retrieved 23 October 2026.
- **R73** Central Electricity Authority. 18-column tentative monthly generation report, July 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/JUL//18_col_ten-2_2024-JUL.pdf>. Retrieved 23 October 2026.
- **R74** Central Electricity Authority. 18-column tentative monthly generation report, August 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/AUG//18_col_ten-2_2024-AUG.pdf>. Retrieved 23 October 2026.
- **R75** Central Electricity Authority. 18-column tentative monthly generation report, September 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/SEP//18_col_ten-2_2024-SEP.pdf>. Retrieved 23 October 2026.
- **R76** Central Electricity Authority. 18-column tentative monthly generation report, October 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/OCT//18_col_ten-2_2024-OCT.pdf>. Retrieved 23 October 2026.
- **R77** Central Electricity Authority. 18-column tentative monthly generation report, November 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/NOV//18_col_ten-2_2024-NOV.pdf>. Retrieved 23 October 2026.
- **R78** Central Electricity Authority. 18-column tentative monthly generation report, December 2024, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2024/DEC//18_col_ten-2_2024-DEC.pdf>. Retrieved 23 October 2026.
- **R79** Central Electricity Authority. 18-column tentative monthly generation report, January 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/JAN//18_col_ten-2_2025-JAN.pdf>. Retrieved 23 October 2026.
- **R80** Central Electricity Authority. 18-column tentative monthly generation report, February 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/FEB//18_col_ten-2_2025-FEB.pdf>. Retrieved 23 October 2026.
- **R81** Central Electricity Authority. 18-column tentative monthly generation report, March 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/MAR//18_col_ten-2_2025-MAR.pdf>. Retrieved 23 October 2026.
- **R82** Central Electricity Authority. 18-column tentative monthly generation report, April 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/APR//18_col_ten-2_2025-APR.pdf>. Retrieved 23 October 2026.
- **R83** Central Electricity Authority. 18-column tentative monthly generation report, May 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/MAY//18_col_ten-2_2025-MAY.pdf>. Retrieved 23 October 2026.
- **R84** Central Electricity Authority. 18-column tentative monthly generation report, June 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/JUN//18_col_ten-2_2025-JUN.pdf>. Retrieved 23 October 2026.
- **R85** Central Electricity Authority. 18-column tentative monthly generation report, July 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/JUL//18_col_ten-2_2025-JUL.pdf>. Retrieved 23 October 2026.
- **R86** Central Electricity Authority. 18-column tentative monthly generation report, August 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/AUG//18_col_ten-2_2025-AUG.pdf>. Retrieved 23 October 2026.
- **R87** Central Electricity Authority. 18-column tentative monthly generation report, September 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/SEP//18_col_ten-2_2025-SEP.pdf>. Retrieved 23 October 2026.
- **R88** Central Electricity Authority. 18-column tentative monthly generation report, October 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/OCT//18_col_ten-2_2025-OCT.pdf>. Retrieved 23 October 2026.
- **R89** Central Electricity Authority. 18-column tentative monthly generation report, November 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/NOV//18_col_ten-2_2025-NOV.pdf>. Retrieved 23 October 2026.
- **R90** Central Electricity Authority. 18-column tentative monthly generation report, December 2025, sub-report 2 (region-wise summary). <https://npp.gov.in/public-reports/cea/monthly/generation/18_col_ten/2025/DEC//18_col_ten-2_2025-DEC.pdf>. Retrieved 23 October 2026.
- **R91** Central Electricity Authority (CEA), Ministry of Power, Government of India. All India Installed Capacity (in MW) of Power Stations, as on 31.03.2023 (Utilities) (Internet Archive copy). <http://web.archive.org/web/20240229215056/https://cea.nic.in/wp-content/uploads/installed/2023/03/IC_March_2023.pdf>. Retrieved 23 October 2026.
- **R92** Central Electricity Authority. Monthly Report of Renewable Generation, January 2023 (Internet Archive copy). <https://web.archive.org/web/20261026035554/https://cea.nic.in/api/dashboardReports/download-excel/en/01c9c450-3b5f-458d-b489-7ead26418701>. Retrieved 26 October 2026.
- **R93** Central Electricity Authority. Monthly Report of Renewable Generation, March 2023 (Internet Archive copy). <https://web.archive.org/web/20261026035709/https://cea.nic.in/api/dashboardReports/download-excel/en/c368a1c1-befa-4870-ab72-57cacf8d67b2>. Retrieved 26 October 2026.
- **R94** Ministry of New and Renewable Energy. State-wise installed capacity of Renewable Power as on 31.12.2023. <https://cdnbbsr.s3waas.gov.in/s3716e1b8c6cd17b771da77391355749f3/uploads/2024/01/202401091363179073.pdf>. Retrieved 23 October 2026.
- **R95** Ministry of New and Renewable Energy. State-wise installed capacity of Renewable Power as on 31.12.2024. <https://cdnbbsr.s3waas.gov.in/s3716e1b8c6cd17b771da77391355749f3/uploads/2025/01/202501081447570936.pdf>. Retrieved 23 October 2026.
- **R96** Ministry of New and Renewable Energy. State-wise installed capacity of Renewable Power as on 31.12.2025. <https://cdnbbsr.s3waas.gov.in/s3716e1b8c6cd17b771da77391355749f3/uploads/2026/01/202601091741212092.pdf>. Retrieved 23 October 2026.
- **R97** Pfenninger and Staffell, Energy 114 (2016). Long-term patterns of European PV output using 30 years of validated hourly reanalysis and satellite data, authors' copy. <https://www.pfenninger.org/publications-pdf/Pfenninger%20and%20Staffell_2016_Long-term%20patterns%20of%20European%20PV%20output.pdf>. Retrieved 23 October 2026.
- **R98** Pfenninger and Staffell, Energy 114 (2016). Long-term patterns of European PV output using 30 years of validated hourly reanalysis and satellite data, Energy 114, 1251-1265 (the Renewables.ninja PV model). <https://doi.org/10.1016/j.energy.2016.08.060>. Retrieved 23 October 2026.
- **R99** Renewables.ninja. Renewables.ninja PV simulation tool. <https://www.renewables.ninja/>. Retrieved 23 October 2026.
- **R100** Central Electricity Authority (CEA), Ministry of Power, Government of India. National Electricity Plan (Transmission), Volume II, 2023-2032 (Internet Archive copy). <http://web.archive.org/web/20241106202726/http://cea.nic.in/wp-content/uploads/notification/2024/10/National_Electricity_Plan_Volume_II_Transmission.pdf>. Retrieved 23 October 2026.
- **R101** Central Electricity Authority (CEA), Ministry of Power, Government of India. CEA Annual Report 2022-23, Annexure-3A Details of Inter-regional Transmission lines as on 31.03.2023 (Internet Archive copy). <http://web.archive.org/web/20240201164518/https://cea.nic.in/wp-content/uploads/annual_reports/2023/Approved_CEA_Annual_Report_2022_23.pdf>. Retrieved 23 October 2026.
- **R102** Central Electricity Authority (CEA), Ministry of Power, Government of India. Growth of Inter-Regional Transmission Capacity (220 kV and above), monthly transmission growth statement, March 2026 (Internet Archive copy). <http://web.archive.org/web/20260715234209/https://cea.nic.in/wp-content/uploads/transmission/2026/03/TL_GS_March___2026.pdf>. Retrieved 23 October 2026.
- **R103** Nepal Electricity Authority. A Year in Review Fiscal Year 2024/2025. <https://nea.org.np/uploads/shares/publication/77022320.pdf>. Retrieved 17 October 2026.
- **R104** Nepal Electricity Authority. A Year in Review Fiscal Year 2023/2024. <https://www.nea.org.np/uploads/shares/annual_report/60896609.pdf>. Retrieved 17 October 2026.
- **R105** Nepal Electricity Authority. A Year in Review Fiscal Year 2022/2023. <https://www.nea.org.np/uploads/shares/annual_report/NEA_Annual_Report_2023.pdf>. Retrieved 17 October 2026.
- **R106** World Bank. Implementation Completion and Results Report ICR00006241, Nepal: Grid Solar and Energy Efficiency Project. <https://documents1.worldbank.org/curated/en/099101223111521685/pdf/BOSIB0fd3a39cd04b0a92105f2d25ccdebf.pdf>. Retrieved 23 October 2026.
- **R107** World Bank. Disclosable Version of the ISR, Nepal: Grid Solar and Energy Efficiency (P146344), Sequence 14, ISR52198. <https://documents1.worldbank.org/curated/en/099311006302220076/pdf/P1463440279d430f70a03904bd0140290a6.pdf>. Retrieved 23 October 2026.
- **R108** Ceylon Electricity Board. Sales and Generation Data Book 2024. <https://www.ceb.lk/front_img/img_reports/1763693550Sales_and_Generation_Data_Book_2024.pdf>. Retrieved 23 October 2026.
- **R109** Ceylon Electricity Board. Statistical Digest 2025 (Web Edition). <https://www.ceb.lk/front_img/img_reports/1783568926SD-2025-Finalized-%28Web-Edition%29.pdf>. Retrieved 23 October 2026.
- **R110** Public Utilities Commission of Sri Lanka. Renewable Generation Report Quarter 04 (Oct 2024 - Dec 2024). <https://www.pucsl.gov.lk/wp-content/uploads/2025/04/Renewable-Generation-Report-Q4-2024.pdf>. Retrieved 23 October 2026.
- **R111** Public Utilities Commission of Sri Lanka. Renewable Generation Report Quarter 04 2025 (Oct 2025 - Dec 2025). <https://www.pucsl.gov.lk/wp-content/uploads/2026/04/Renewable-Generation-Report-Q4-2025.pdf>. Retrieved 23 October 2026.
- **R112** Public Utilities Commission of Sri Lanka. Renewable Generation Report Q4 (Oct 2023 - Dec 2023). <https://www.pucsl.gov.lk/wp-content/uploads/2024/04/Renewable-Generation-Report-Q4-2023-Final.pdf>. Retrieved 23 October 2026.
- **R113** Ceylon Electricity Board. Long Term Generation Expansion Plan 2025-2044 (May 2025). <https://www.ceb.lk/front_img/img_reports/1748839124LTGEP-2025-2044-FINAL_c.pdf>. Retrieved 23 October 2026.
- **R114** Central Electricity Authority (CEA), Ministry of Power, Government of India. Growth Summary of Transmission Lines (220 kV and above), monthly transmission growth statement, June 2025 (Internet Archive copy). <https://web.archive.org/web/20261026034143/https://cea.nic.in/api/dashboardReports/download-pdf/en/fe1eb66f-9bc6-46ea-a708-754f4ede6824>. Retrieved 26 October 2026.
