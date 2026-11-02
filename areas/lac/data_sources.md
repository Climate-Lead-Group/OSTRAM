# Where the LAC values come from

This page covers the Latin America and Caribbean (LAC) area: 19 countries (Argentina ARG, Barbados BRB, Bolivia BOL, Brazil BRA, Chile CHL, Colombia COL, Costa Rica CRI, Dominican Republic DOM, Ecuador ECU, El Salvador SLV, Guatemala GTM, Haiti HTI, Honduras HND, Mexico MEX, Nicaragua NIC, Panama PAN, Paraguay PRY, Peru PER and Uruguay URY) and the scenarios of the OSTRAM-LAC model it comes from: Reference (BAU), Vegetativo A (INV), Optimal (OPT) and Vegetativo B (VGB). The `lac` area carries the Reference (BAU) values as its support scenario `BAU`; the blocks about INV, OPT and VGB describe the source model and are kept for reference. The values sit in three workbooks, `A-O_Parametrization`, `A-O_Demand` and `A-Xtra_Storage`. To find a value, open the section for its workbook and sheet, then the block for its technologies and parameter. Each block names the technologies, parameters and years, the external source, how the value is built, and the assumptions behind it. A block marked "needs source" names the part of the value that has no external source.

## Assumptions used across the table

Capacity converts to energy at 31.536 PJ per GW-year (8 760 hours x 3.6 TJ per GWh). Generation in GWh converts to PJ at 0.0036 PJ/GWh. Costs are in US dollars; million USD per GW equals USD per kW.

**LA1.** The domestic fuel supply technologies (`MIN<fuel><country>`, 133 technologies) have an annual activity upper limit (`TotalTechnologyAnnualActivityUpperLimit`) of 0 PJ in every year from 2023 to 2050 in the 19 countries, so they supply no fuel; fossil plants take their fuel through mode 2 from the international supply technologies (`MIN<fuel>INT`), priced on the VariableCost sheet. Their other Primary Techs parameters set no constraint. Reason: OSeMOSYS Global v2.0.0a limits fuel supply only through an annual activity upper limit read from a per-country limits table, and that table has no entries in v2.0.0a, so every limit is 0. Check: OSeMOSYS Global repository, https://github.com/OSeMOSYS/osemosys_global.

**LA2.** Each country takes the generation costs of the closest IEA World Energy Outlook 2020 region: South American countries take the values for Brazil, and Mexico, Central America and the Caribbean (including Barbados) take the values for the United States. Costs (USD of 2019 per kW) are interpolated linearly between the WEO years 2019, 2030 and 2040 and held at the 2040 value from 2041 to 2050. This covers `CapitalCost` and `FixedCost` of `PWR*` and `ELC*` in the 19 countries. Reason: the WEO 2020 cost table gives values for a few world regions, not for each country, and its last year is 2040.

**LA3.** The backstop technology `PWRBCK*` stands for energy not supplied. It is a penalty generator, not a real plant, in the 19 countries in 2023 to 2050. Its `CapitalCost` and `FixedCost` are very high sentinel values with no economic meaning. Its `VariableCost` is 1 750 million USD per PJ (about 6 300 USD/MWh): the ENTSO-E scarcity price cap of 5 500 EUR/MWh for target year 2028 in the European Resource Adequacy Assessment 2025 (ERAA), converted to 1 743 million USD per PJ and rounded to 1 750. ENTSO-E uses this cap as a proxy for the cost of energy not served in its economic viability module, a least-cost expansion model like this one. Reason: the penalty keeps the optimisation feasible when supply cannot meet demand; it is a conservative penalty, not an estimate of the social cost of outages. Check: ACER (2026), Decision No 06/2026, Annex III, states that these caps are well below the value of lost load (VoLL) that member states use for their reliability standards.

**LA4.** Interconnector costs (`TRN{A}XX{B}XX`, 26 links, 2023 to 2050, constant over time): `CapitalCost` = converter cost (95 400 USD/MW for alternating current; 297 509 USD/MW for direct current) + line cost (779 USD per MW and km for alternating current; 238 for direct current; 295 for subsea direct current) x the distance between the centroids of the two countries. `FixedCost` per year = 0.35% of `CapitalCost`. Example: Argentina-Brazil, 616.7 million USD per GW. These are the default values of OSeMOSYS Global v2.0.0a. Reason: one set of unit costs applied to the distance between two countries gives each link a cost that follows its length and its technology.

**LA5.** Annual cap on new generation capacity (`TotalAnnualMaxCapacityInvestment`, `PWR*`, 19 countries, 2026 to 2050, GW per year): cap = base percentage x demand growth multiplier x the country's existing capacity. The base percentage is 3.75% in 2026 to 2030, 2.25% in 2031 to 2040 and 3.75% in 2041 to 2050. Lower percentages apply to wind, solar and biomass (1.0 to 0.5%), geothermal and hydro (0.5 to 0.3%) and wave (0.3 to 0.2%). The cap is never below the required minimum investment x 1.01; where a required minimum is above the cap, the cap becomes minimum x 1.01. Reason: without an annual cap the least-cost solution adds capacity at disproportionate and unrealistic rates. The factor 1.01 keeps the cap just above any required build, so both limits can hold.

**LA6.** In the Reference scenario (BAU) only, new solar PV (`PWRSPV*`) and onshore wind (`PWRWON*`) capacity has an annual ceiling in GW per year, one value per country for each of three periods (2023 to 2030, 2031 to 2040 and 2041 to 2050), applied in 2025 to 2050 in the 19 countries. The values are the `TotalAnnualMaxCapacityInvestment` entries for these technologies on Secondary Techs. Reason: as for LA5, without an annual ceiling the least-cost solution adds solar and wind capacity at disproportionate and unrealistic rates.

**LA7.** The required minimum investment (`TotalAnnualMinCapacityInvestment`) in 2023 to 2039, for generation (`PWR*`) and for internal transmission (`PWRTRN*`, `RNWTRN*`) in the Reference (BAU) and Vegetativo A (INV) scenarios, counts only projects that are under construction, awarded in a tender, covered by a signed contract or power purchase agreement, or already in service. Projects that are only indicative have a minimum of zero. Reason: indicative projects are not commitments; with a zero minimum the model chooses whether to build them.

**LA8.** In the Vegetativo A (INV), Optimal (OPT) and Vegetativo B (VGB) scenarios, every parameter on Primary Techs, Secondary Techs, Demand Techs and Capacities takes the Reference (BAU) value in 2023 to 2026 (OPT) or 2023 to 2030 (INV and VGB). The exceptions are the interconnector investment cap from 2027 and, in VGB, minimum generation and minimum investment. Reason: the scenarios share one history and do not diverge before 2027.

**LA9.** Demand in 2025 to 2050 for 18 countries (all except Haiti) grows from the OLADE 2024 base at the annual rates of each national plan's high-demand scenario. Bolivia, Ecuador and Nicaragua publish no high scenario, so their central scenario is used; Paraguay's plan has a single demand path, used for both. Beyond the last year of a plan, demand grows at the compound annual growth rate of the nearest segment of that plan. The same extra growth of the high scenario over the reference path scales the interconnector exchange limits. Reason: a plan without a high scenario offers only its central path, and most plans end before 2050.

**LA10.** International prices of coal, crude oil and natural gas are multiplied by 1.3 in the 19 countries in 2023 to 2050. Reason: OSeMOSYS Global applies this surcharge to imported fuels, to reflect the extra cost of importing fuel over the international benchmark price.

## A-O_Parametrization, Primary Techs

### `MIN*` fuel supply

- Technologies: `MIN*`, fuel supply (coal, gas, oil, diesel, uranium, biomass and others).
- Parameters: `CapitalCost`, `FixedCost`, `ResidualCapacity`, `TotalAnnualMaxCapacity`, `TotalAnnualMaxCapacityInvestment`, `TotalAnnualMinCapacityInvestment`, `TotalTechnologyAnnualActivityLowerLimit`, `TotalTechnologyAnnualActivityUpperLimit`, `AvailabilityFactor`, `ReserveMarginTagFuel`, `ReserveMarginTagTechnology`.
- Years: 2023-2050.
- Source: OSeMOSYS Global v2.0.0a (April 2025), https://github.com/OSeMOSYS/osemosys_global.
- Detail: OSeMOSYS Global gives fuel supply technologies only an annual activity upper limit, taken from a per-country limits table. That table is empty in v2.0.0a, so every row holds 0, which allows these domestic technologies no activity; fuel comes from the international supply technologies (`MIN<fuel>INT`). The other fuel parameters set no constraint. Fuel costs are on the VariableCost sheet.
- Assumptions: ASSUMPTION LA1.

## A-O_Parametrization, Secondary Techs

### `PWR*` generation and `ELC*` distribution, capital and fixed costs

- Technologies: `PWR*` generation (biomass, coal, carbon capture, cogeneration, solar thermal, geothermal, hydro, natural gas, fuel oil, other, diesel, solar PV, nuclear, waste, wave, offshore wind, onshore wind and backstop) and `ELC*` (electricity distribution).
- Parameters: `CapitalCost`, `FixedCost`.
- Years: 2023-2050.
- Source: IEA, World Energy Outlook 2020, cost table for power generation technologies (USD of 2019 per kW), https://www.iea.org/reports/world-energy-outlook-2020, via OSeMOSYS Global v2.0.0a (April 2025), https://github.com/OSeMOSYS/osemosys_global.
- Detail: Capital cost and annual fixed operating cost by technology, taken by OSeMOSYS Global v2.0.0a from WEO 2020. Each country takes the closest WEO region: South American countries take the values for Brazil; Mexico, Central America and the Caribbean (including Barbados) take the values for the United States. The WEO values for 2019, 2030 and 2040 are interpolated linearly and held constant after 2040 (for example, solar PV in South America: 510 USD/kW in 2023, falling to 370 in 2050; onshore wind 1 090 falling to 1 040). Natural gas and coal average several WEO plant types. The backstop technology (`PWRBCK`) carries very high sentinel costs, not an external value.
- Assumptions: ASSUMPTION LA2; ASSUMPTION LA3 for `PWRBCK`.

### `PWR*` generation and `ELC*` distribution, availability and reserve margin tags

- Technologies: `PWR*` generation (same list as above) and `ELC*` (electricity distribution).
- Parameters: `AvailabilityFactor`, `ReserveMarginTagFuel`, `ReserveMarginTagTechnology`.
- Years: 2023-2050.
- Source: OSeMOSYS Global v2.0.0a default value table, https://github.com/OSeMOSYS/osemosys_global. Needs source: the default values themselves have no cited origin.
- Detail: Availability factor by technology type from the OSeMOSYS Global default table: 0.9 for nuclear, coal, gas, geothermal, hydro, solar and wind; 0.8 for fuel oil, diesel, cogeneration and other; 0.5 for biomass; constant over time. The reserve margin tags mark which technologies and fuels count towards the system reserve, as set in OSeMOSYS Global.

### `PWRLDS*`, `PWRSDS*`, `PWRBDS*` storage, power component of cost

- Technologies: `PWRLDS*`, `PWRSDS*`, `PWRBDS*` (long-duration storage, short-duration storage and batteries).
- Parameters: `CapitalCost`, `FixedCost`.
- Years: 2023-2050.
- Source: NREL Annual Technology Baseline (ATB) 2024, National Renewable Energy Laboratory, United States (URL needed).
- Detail: Power component of storage cost (USD per kW). Long duration falls from 4 228 to 1 843 USD/kW; short duration is 944 USD/kW. These differ from the OSeMOSYS Global defaults (1 897 USD/kW constant for long duration; 711.5 USD/kW for short duration), which the model does not use. Batteries (`BDS`, Chile and Peru only) take the short-duration values.

### `PWR*` generation, residual capacity

- Technologies: `PWR*` generation.
- Parameters: `ResidualCapacity`.
- Years: 2023-2050 (constant value).
- Source: OLADE, installed capacity by source (annual statistics of the OLADE energy information system) (URL needed).
- Detail: Installed capacity by country and technology in 2023, in MW, converted to GW (divided by 1 000) and held constant in all years. Capacity fuelled by oil products is split between diesel (`PET`) and fuel oil plus bunker fuel (`OIL`) using each country's historical shares.
- Needs source: holding 2023 capacity constant to 2050, and the historical shares used for the `PET`/`OIL` split.

### `PWR*` generation, annual activity limits

- Technologies: `PWR*` generation.
- Parameters: `TotalTechnologyAnnualActivityLowerLimit`, `TotalTechnologyAnnualActivityUpperLimit`.
- Years: 2023-2050.
- Source: OLADE, electricity generation by source (share of each technology in 2023) (URL needed), and the projected demand of this model (see A-O_Demand, Demand_Projection).
- Detail: Minimum annual generation by technology = the country's total required generation (demand + exports - imports) x the technology's share of 2023 generation according to OLADE, adjusted to renewable share targets. The upper limit is 5% above the minimum. From 2026 the minimum stays constant at its 2024 value.
- Needs source: the renewable share targets, the 5% band, and holding the minimum at its 2024 value from 2026.

### `TRN{A}XX{B}XX` interconnectors, annual activity limits

- Technologies: `TRN{A}XX{B}XX` (26 electricity interconnectors between countries).
- Parameters: `TotalTechnologyAnnualActivityLowerLimit`, `TotalTechnologyAnnualActivityUpperLimit`.
- Years: 2023-2050.
- Source: OLADE, energy balance matrix (electricity imports and exports by country) (URL needed), and the high-demand scenario of each country's official planning document (see A-O_Parametrization and A-O_Demand, demand by country).
- Detail: Bilateral exchanges are estimated by splitting each country's OLADE imports and exports among its interconnected neighbours (GWh converted to PJ). Lower limit = exchange - 5%; upper limit = exchange + 5%; constant from the last year with data. Each limit is then scaled by the extra growth of the high-demand scenario of the two countries over the reference path. Exceptions: Colombia-Panama (0.4 GW and 12.614 PJ from 2029) and Bolivia-Brazil (0.42 GW and 13.245 PJ from 2032).
- Assumptions: ASSUMPTION LA9 for the demand scaling.
- Needs source: the rule that splits national imports and exports among neighbours, the 5% band, and the exchange values for the Colombia-Panama and Bolivia-Brazil exceptions.

### `TRN{A}XX{B}XX` interconnectors, residual capacity

- Technologies: `TRN{A}XX{B}XX` (26 electricity interconnectors between countries).
- Parameters: `ResidualCapacity`.
- Years: 2023-2050 (with steps in the entry year of each expansion).
- Source: Global Transmission Database v1.1 (Brinkerink et al., 2024; current capacity of each interconnector), via OSeMOSYS Global v2.0.0a, https://github.com/OSeMOSYS/osemosys_global (database URL needed); revised with each country's official transmission planning documents (capacity and entry year of expansions and new lines; see A-O_Parametrization, Demand Techs, MinCapInv Tx by country).
- Detail: Existing transmission capacity of each interconnector, in GW. 14 corridors keep the Global Transmission Database value (for example Argentina-Brazil 2.276 GW, Brazil-Paraguay 8.751 GW, Argentina-Uruguay 3.89 GW); 5 of them have no existing or planned line: Bolivia-Paraguay, Brazil-Colombia, Brazil-Peru, Colombia-Peru and Dominican Republic-Haiti. 12 corridors take the capacity and entry year from the planning documents: Colombia-Panama 0.4 GW from 2029; Bolivia-Brazil 0.42 GW from 2032; Guatemala-Mexico 0.24 GW rising to 0.64 GW in 2032; Colombia-Ecuador 0.54 GW; Ecuador-Peru 0.6 GW from 2027; Chile-Peru 0.2 GW in 2032 and 1.2 GW in 2040; stepwise expansions of the Central American system: Costa Rica-Nicaragua 0.6 GW in 2029; Costa Rica-Panama and Honduras-Nicaragua 0.45 GW in 2029 and 0.6 GW in 2031; Guatemala-El Salvador 0.45/0.6 GW in 2031/2033; Honduras-El Salvador 0.45/0.6 GW in 2032/2034; Guatemala-Honduras 0.45/0.6 GW in 2033/2035.
- Needs source: values that some countries' plans do not give, which were supplied by Inter-American Development Bank (IDB) contributors.

### `TRN{A}XX{B}XX` interconnectors, capital and fixed costs

- Technologies: `TRN{A}XX{B}XX` (26 electricity interconnectors between countries).
- Parameters: `CapitalCost`, `FixedCost`.
- Years: 2023-2050.
- Source: OSeMOSYS Global v2.0.0a default unit costs of transmission lines x distance between countries, https://github.com/OSeMOSYS/osemosys_global.
- Detail: Capital cost of each interconnector = converter cost (95 400 USD/MW for alternating current; 297 509 USD/MW for direct current) + line cost (779 USD per MW and km for alternating current; 238 for direct current; 295 for subsea direct current) x the distance between the centroids of the two countries. Annual fixed cost = 0.35% of capital cost. Constant over time (for example Argentina-Brazil 616.7 million USD per GW).
- Assumptions: ASSUMPTION LA4.

### `PWR*` generation, annual cap on new capacity

- Technologies: `PWR*` generation.
- Parameters: `TotalAnnualMaxCapacityInvestment`.
- Years: 2026-2050.
- Source: ASSUMPTION LA5.
- Detail: Annual cap on new capacity by technology and country = base percentage (3.75% in 2026-2030, 2.25% in 2031-2040 and 3.75% in 2041-2050) x demand growth multiplier x the country's existing capacity. Lower percentages apply to wind, solar and biomass (1.0 to 0.5%), geothermal and hydro (0.5 to 0.3%) and wave (0.3 to 0.2%). The cap is never below the investment floor (minimum x 1.01). In scenarios other than Reference, the historical years take the Reference values (ASSUMPTION LA8).
- Assumptions: ASSUMPTION LA5, ASSUMPTION LA8.

### `PWRSPV*`, `PWRWON*` solar and wind ceilings (Reference scenario BAU only)

- Technologies: `PWRSPV*`, `PWRWON*` (solar PV and onshore wind, 19 countries).
- Parameters: `TotalAnnualMaxCapacityInvestment`.
- Years: 2025-2050.
- Source: ASSUMPTION LA6.
- Detail: Annual ceilings on new capacity (GW per year) for three periods: 2023-2030, 2031-2040 and 2041-2050.
- Assumptions: ASSUMPTION LA6.

### Generation projects dated 2040 to 2050 (Optimal scenario OPT only)

- Technologies: 24 generation technologies with projects dated between 2040 and 2050.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2040-2050.
- Source: official electricity planning documents of each country with projects dated 2040-2050: Guatemala, Plan de Expansión de la Generación 2020-2050 (MEM / CNEE); Paraguay, Plan Maestro de Generación 2024-2043 (ANDE); Chile, Planificación Energética de Largo Plazo (PELP) 2023-2027 (Ministerio de Energía); Uruguay, Plan de Expansión de la Generación 2024-2043 (UTE); Costa Rica, Plan de Expansión de la Generación (PEG) 2024-2040 (ICE); Argentina, Lineamientos y Escenarios para la Transición Energética al 2050 (Secretaría de Energía); Brazil, Plano Nacional de Energia 2050 (EPE). All URLs needed.
- Detail: Floors on new capacity per year, taken from the projects and indicative expansions that the national documents date between 2040 and 2050. Where a floor was above the investment cap, the cap becomes floor x 1.01.
- Assumptions: ASSUMPTION LA5 for the factor 1.01.

### `PWR*` generation, required minimum investment (Reference BAU and Vegetativo A INV)

- Technologies: `PWR*` generation (Reference BAU and Vegetativo A INV scenarios).
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: official generation planning documents of each country (see the MinCapInv PWR blocks by country below).
- Detail: Required minimum investment in new capacity per year. Only projects under construction, awarded, with a signed power purchase agreement, or already in service count; projects that are only indicative are set to zero.
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Argentina

- Technologies: `PWR*`, Argentina.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Nuclear Argentino 2025 (URL needed); Secretaría de Energía, MATER and RenovAr programmes (URL needed).
- Detail: Argentina has no unified ten-year generation plan in force in 2025.
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Bolivia

- Technologies: `PWR*`, Bolivia.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Eléctrico Referencial del Estado Plurinacional de Bolivia 2035 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Brazil

- Technologies: `PWR*`, Brazil.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: PDE 2034, Plano Decenal de Expansão de Energia 2034 (EPE) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Barbados

- Technologies: `PWR*`, Barbados.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Barbados National Energy Policy (BNEP) 2019-2030 (URL needed); Barbados Energy Transition and Investment Plan (SEforALL 2025) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Chile

- Technologies: `PWR*`, Chile.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: PELP 2023-2027, Planificación Energética de Largo Plazo (Ministerio de Energía de Chile) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Colombia

- Technologies: `PWR*`, Colombia.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan de Expansión de Referencia Generación-Transmisión 2024-2038 / Plan Indicativo 2023-2037 (UPME) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Costa Rica

- Technologies: `PWR*`, Costa Rica.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan de Expansión de la Generación Eléctrica (PEG) 2024-2040 (ICE) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Dominican Republic

- Technologies: `PWR*`, Dominican Republic.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Energético Nacional 2022-2036 (CNE / MEM) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Ecuador

- Technologies: `PWR*`, Ecuador.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Maestro de Electricidad (PME) 2023-2032 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Guatemala

- Technologies: `PWR*`, Guatemala.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan de Expansión Indicativo del Sistema de Generación 2024-2054 (MEM / CNEE) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Honduras

- Technologies: `PWR*`, Honduras.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Indicativo de Expansión de la Generación (PIEG) 2024-2033 (ENEE / CND) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Mexico

- Technologies: `PWR*`, Mexico.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: PRODESEN 2024-2038 (URL needed); CFE, Plan de Fortalecimiento y Expansión 2025-2030 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Nicaragua

- Technologies: `PWR*`, Nicaragua.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan de Expansión de la Generación Eléctrica de Nicaragua 2019-2033 (MEM) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Panama

- Technologies: `PWR*`, Panama.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Indicativo de Generación 2025-2039, Tomo II (ASEP / ETESA) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Peru

- Technologies: `PWR*`, Peru.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Energético Nacional 2014-2025 (in force) (URL needed); MINEM portfolio of 26 projects to 2027 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Paraguay

- Technologies: `PWR*`, Paraguay.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Maestro de Generación ANDE 2021-2040 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, El Salvador

- Technologies: `PWR*`, El Salvador.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Indicativo de la Expansión de la Generación Eléctrica 2021-2031 (DGEHM / CNE) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv PWR, Uruguay

- Technologies: `PWR*`, Uruguay.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan de Expansión de UTE (URL needed); Hoja de Ruta del Hidrógeno Verde (MIEM), Segunda Transición Energética (URL needed).
- Assumptions: ASSUMPTION LA7.

## A-O_Parametrization, VariableCost

### `PWRBCK*` backstop

- Technologies: `PWRBCK*` (backstop technology standing for energy not supplied).
- Parameters: `VariableCost`.
- Years: 2023-2050.
- Source: ACER (2026), Decision No 06/2026, European Resource Adequacy Assessment 2025 (ERAA), Annex III (ENTSO-E scarcity price cap) (URL needed).
- Detail: Variable cost of 1 750 million USD per PJ (about 6 300 USD/MWh), used as the penalty for energy not served. OSeMOSYS Global's default for this cost is a sentinel value (999 999) with no economic meaning; the model uses the value below instead. The value is the scarcity price cap of 5 500 EUR/MWh that ENTSO-E uses for target year 2028 in ERAA 2025 as a proxy for the cost of energy not served in its economic viability module, a least-cost expansion model like this one. Converted to dollars it equals 1 743 million USD per PJ, rounded to 1 750. According to ACER (2026), these caps are well below the value of lost load (VoLL) that member states use for their reliability standards, so the value is a conservative penalty that keeps the optimisation feasible, not an estimate of the social cost of outages.
- Assumptions: ASSUMPTION LA3.

### `MIN*` fuels, `TRN*` interconnectors, `PWRLDS*`, `PWRSDS*`, `PWRBDS*` storage

- Technologies: `MIN*` (fuels), `TRN*` (interconnectors), `PWRLDS*`, `PWRSDS*`, `PWRBDS*` (storage).
- Parameters: `VariableCost`.
- Years: 2023-2050.
- Source: World Bank, Commodity Markets Outlook, October 2024 (international prices of coal, crude oil and natural gas) (URL needed), via OSeMOSYS Global v2.0.0a, https://github.com/OSeMOSYS/osemosys_global. Uranium, biomass and waste: OSeMOSYS Global default values. Interconnectors and storage: OSeMOSYS Global default values.
- Detail: Commodity Markets Outlook prices converted to million USD per PJ (coal divided by 29.31 GJ per tonne; oil divided by 6.12 GJ per barrel; gas divided by 1.055 GJ per MMBtu) and multiplied by 1.3, the surcharge OSeMOSYS Global applies to imported fuels; for example coal 5.9 and oil/diesel 13.5 million USD per PJ, constant over time. Uranium 2, biomass 10 and waste 10 million USD per PJ (fixed OSeMOSYS Global values). Interconnectors: 4 USD/MWh (1.111 million USD per PJ), an OSeMOSYS Global default parameter.
- Assumptions: ASSUMPTION LA10.
- Needs source: the uranium, biomass and waste prices, the interconnector cost of 4 USD/MWh, the storage variable cost, and the heat contents 29.31 GJ per tonne of coal and 6.12 GJ per barrel of oil.

## A-O_Parametrization, Fixed Horizon Parameters

### `PWR*` generation, `MIN*` fuels, `ELC*` distribution

- Technologies: `PWR*` (generation), `MIN*` (fuels), `ELC*` (distribution).
- Parameters: `CapacityToActivityUnit`, `OperationalLife`.
- Years: none (not year-indexed).
- Source: operational life: NREL Annual Technology Baseline, as used in the OSeMOSYS Global v2.0.0a default value table, https://github.com/OSeMOSYS/osemosys_global (ATB URL needed); conversion factor: physical constant.
- Detail: Capacity to energy conversion factor: 31.536 PJ per GW-year (8 760 hours x 3.6 TJ per GWh). Operational life by technology from the OSeMOSYS Global default table, based on the NREL ATB: 30 years for most technologies, hydro 100, nuclear 60, wave 20; technologies with no value take 1.
- Needs source: the operational life of 1 year for technologies with no value.

### Internal transmission network

- Technologies: `RNWTRN*`, `RNWRPO*`, `RNWNLI*`, `PWRTRN*`, `TRNRPO*`, `TRNNLI*` (each country's internal transmission network: existing lines, repowering and new lines).
- Parameters: `CapacityToActivityUnit`, `OperationalLife`.
- Years: none (not year-indexed).
- Source: needs source (operational life).
- Detail: Capacity to energy conversion factor 31.536 PJ per GW-year (physical constant) and operational life of 50 years, the same in the 19 countries.
- Needs source: the operational life of 50 years.

## A-O_Parametrization, Demand Techs

### Internal transmission network, capital and fixed costs

- Technologies: `RNWTRN*`, `PWRTRN*` (existing lines), `RNWRPO*`, `TRNRPO*` (repowering), `RNWNLI*`, `TRNNLI*` (new lines), internal transmission network of each of the 19 countries.
- Parameters: `CapitalCost`, `FixedCost`.
- Years: 2023-2050 (constant value).
- Source: needs source.
- Detail: Values differ by country. Capital cost of new lines between 20 and 463 million USD per GW depending on the country (lowest in Barbados, the Dominican Republic and Haiti; highest in Chile); existing lines = 1.2 x new lines; repowering = 0.45 x new lines; annual fixed cost proportional.
- Needs source: the country capital costs of new lines, the factors 1.2 and 0.45, and the fixed cost ratio.

### Internal transmission network, residual and maximum capacity

- Technologies: `RNWTRN*`, `PWRTRN*` (existing lines of the internal transmission network, renewable and non-renewable part, 19 countries); `RNWRPO*`, `TRNRPO*`, `RNWNLI*`, `TRNNLI*` (repowering and new lines).
- Parameters: `ResidualCapacity`, `TotalAnnualMaxCapacity`.
- Years: 2023-2050.
- Source: official transmission planning documents of each country (see the MinCapInv Tx blocks by country below), and OSeMOSYS Global v2.0.0a, https://github.com/OSeMOSYS/osemosys_global.
- Detail: Existing transmission capacity of each country, in GW, different by country (from 0.2 GW in Barbados to 163 GW in Brazil), split between the part that carries renewable generation (`RNWTRN*`) and the non-renewable part (`PWRTRN*`). Repowering and new lines start from zero. The renewable network of Brazil (`RNWTRNBRAXX`) is 85.40 GW. The maximum total capacity (`TotalAnnualMaxCapacity`) grows year by year along a path set for each country on the same basis.
- Needs source: the renewable/non-renewable split, the `RNWTRNBRAXX` value of 85.40 GW, and the country paths for `TotalAnnualMaxCapacity`.

### Internal transmission network, investment cap (Vegetativo A INV only)

- Technologies: `PWRTRN*`, `TRNNLI*`, `TRNRPO*`, `RNWTRN*`, `RNWRPO*`, `RNWNLI*` (internal transmission network).
- Parameters: `TotalAnnualMaxCapacityInvestment`.
- Years: 2027-2050.
- Source: needs source (derived from the model's own Reference scenario (BAU) result with a reduction factor).
- Detail: Annual cap on transmission investment = new capacity built by the model in the Reference scenario. 2027-2030: the same value; 2031-2050: multiplied by a factor that falls linearly from 0.99 to 0.80. Peru, Costa Rica, Panama and Colombia keep a factor of 1.0 (Costa Rica + 5%). Before 2027 there is no cap.
- Needs source: the reason for the factors 0.99 to 0.80, the factor 1.0 for four countries and the extra 5% for Costa Rica.

### Internal transmission network, required minimum investment (Reference BAU and Vegetativo A INV)

- Technologies: `PWRTRN*`, `RNWTRN*` (internal transmission network; Reference BAU and Vegetativo A INV scenarios).
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: official electricity transmission planning documents of each country (see the MinCapInv Tx blocks by country below).
- Detail: Required minimum investment in transmission per year. Only works under construction, with an awarded tender, with a signed contract, or already in service count; indicative plans are set to zero.
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Argentina

- Technologies: `PWRTRN*`, `RNWTRN*`, Argentina.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Nacional de Expansión del Transporte Eléctrico 2035 (Secretaría de Energía, December 2022) (URL needed); Plan Nacional de Ampliación del Transporte Eléctrico 2024-2025 (Resolución 715/2025) (URL needed).
- Detail: No works meet the rule in ASSUMPTION LA7, so the minimum is zero.
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Bolivia

- Technologies: `PWRTRN*`, `RNWTRN*`, Bolivia.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Eléctrico Referencial 2035 del Estado Plurinacional de Bolivia (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Brazil

- Technologies: `PWRTRN*`, `RNWTRN*`, Brazil.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: EPE, Programa de Expansão da Transmissão (PET) / Plano de Expansão de Longo Prazo (PELP), Ciclo 2024, 1º Semestre (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Barbados

- Technologies: `PWRTRN*`, `RNWTRN*`, Barbados.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Barbados National Energy Policy (BNEP) 2019-2030 (URL needed).
- Detail: No works meet the rule in ASSUMPTION LA7, so the minimum is zero.
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Chile

- Technologies: `PWRTRN*`, `RNWTRN*`, Chile.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: CNE, Informe Técnico Final, Plan de Expansión Anual de Transmisión 2024 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Colombia

- Technologies: `PWRTRN*`, `RNWTRN*`, Colombia.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: UPME, Plan Maestro de Modernización y Expansión de la Infraestructura de Transmisión Eléctrica, Tomo II, Plan de Expansión de Transmisión 2024-2038 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Costa Rica

- Technologies: `PWRTRN*`, `RNWTRN*`, Costa Rica.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: ICE, Plan de Expansión de la Transmisión (PET) 2024-2034 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Dominican Republic

- Technologies: `PWRTRN*`, `RNWTRN*`, Dominican Republic.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: ETED, Plan de Expansión de la Transmisión 2024-2028 (URL needed); Plan Energético Nacional 2025-2038 (MEM / CNE) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Ecuador

- Technologies: `PWRTRN*`, `RNWTRN*`, Ecuador.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: CELEC EP Transelectric, Plan Maestro de Electricidad 2023-2032 (chapter 5, transmission expansion) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Guatemala

- Technologies: `PWRTRN*`, `RNWTRN*`, Guatemala.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: CNEE / MEM, Plan de Expansión del Sistema de Transporte (PET) 2024-2054 (URL needed).
- Detail: No works meet the rule in ASSUMPTION LA7, so the minimum is zero.
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Honduras

- Technologies: `PWRTRN*`, `RNWTRN*`, Honduras.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: ENEE / CND, Plan de Expansión de la Red de Transmisión (PERT) 2024-2033 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Mexico

- Technologies: `PWRTRN*`, `RNWTRN*`, Mexico.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: PRODESEN 2024-2038, Programa de Desarrollo del Sistema Eléctrico Nacional (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Nicaragua

- Technologies: `PWRTRN*`, `RNWTRN*`, Nicaragua.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Indicativo de la Expansión del Sistema Nacional de Transmisión (SNT) 2022-2037 (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Panama

- Technologies: `PWRTRN*`, `RNWTRN*`, Panama.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan de Expansión del Sistema Interconectado Nacional (PESIN) 2025-2039, ETESA (ASEP public consultation CP-009-2025) (URL needed).
- Detail: No works meet the rule in ASSUMPTION LA7, so the minimum is zero.
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Peru

- Technologies: `PWRTRN*`, `RNWTRN*`, Peru.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: COES, Actualización del Plan de Transmisión 2023-2032 (binding works) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Paraguay

- Technologies: `PWRTRN*`, `RNWTRN*`, Paraguay.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Maestro de Transmisión ANDE 2021-2030 (part of the Plan Maestro de Generación, Transmisión y Distribución 2021-2040) (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, El Salvador

- Technologies: `PWRTRN*`, `RNWTRN*`, El Salvador.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan Estratégico de Expansión de ETESAL (URL needed).
- Assumptions: ASSUMPTION LA7.

### MinCapInv Tx, Uruguay

- Technologies: `PWRTRN*`, `RNWTRN*`, Uruguay.
- Parameters: `TotalAnnualMinCapacityInvestment`.
- Years: 2023-2039.
- Source: Plan de Inversiones UTE 2024 (URL needed); Cierre del Anillo Norte Tacuarembó-Salto Grande 500 kV (URL needed).
- Detail: No works meet the rule in ASSUMPTION LA7, so the minimum is zero.
- Assumptions: ASSUMPTION LA7.

## A-O_Parametrization, Capacities

### `PWRCSP*`, `PWRHYD*`, `PWRSPV*`, `PWRWOF*`, `PWRWON*` capacity factors

- Technologies: `PWRCSP*`, `PWRHYD*`, `PWRSPV*`, `PWRWOF*`, `PWRWON*` (solar thermal, hydro, solar PV, offshore wind and onshore wind) x 12 time slices.
- Parameters: `CapacityFactor`.
- Years: 2023-2050.
- Source: Renewables.ninja (hourly profiles of solar PV, solar thermal and wind), https://www.renewables.ninja, and monthly hydro profiles, from the PLEXOS-World 2015 dataset (Brinkerink et al., 2021) (URL needed), via OSeMOSYS Global v2.0.0a, https://github.com/OSeMOSYS/osemosys_global. Plant locations: WRI Global Power Plant Database (URL needed). Calibration: IRENA 2015 generation statistics (URL needed).
- Detail: Hourly series for 2015 generated with Renewables.ninja at the plant locations of the Global Power Plant Database (WRI) and adjusted to actual 2015 generation according to IRENA. Hydro: a 15-year average monthly profile calibrated to each country's historical generation. OSeMOSYS Global averages the hours of each time slice (season x part of day) to give the capacity factor of the slice. A capacity factor can be adjusted where the minimum generation required by the historical share cannot be reached with installed capacity.
- Needs source: the rule and the values of any such adjustment.

## A-O_Parametrization, Yearsplit

### `YearSplit`

- Technologies: none (12 time slices: 4 seasons x 3 parts of day).
- Parameters: `YearSplit`.
- Years: 2023-2050.
- Source: time structure defined in OSeMOSYS Global v2.0.0a (4 seasons x 3 parts of day on the 2015 calendar), https://github.com/OSeMOSYS/osemosys_global.
- Detail: Share of the hours of the year that falls in each time slice, computed by OSeMOSYS Global from the definition of seasons and parts of day; it is a structural value, not external data. Normalised so that the shares sum to 1 in each year.

## A-O_Parametrization, DaySplit

### `DaySplit`

- Technologies: none (3 parts of day).
- Parameters: `DaySplit`.
- Years: 2023-2050.
- Source: time structure defined in OSeMOSYS Global v2.0.0a (3 parts of day), https://github.com/OSeMOSYS/osemosys_global.
- Detail: Share of the day in each part of day (hours in the part divided by 24), from the OSeMOSYS Global definition; it is a structural value, not external data. Normalised.

## A-O_Parametrization, Primary Techs, Secondary Techs, Demand Techs and Capacities (scenarios Vegetativo A INV, Optimal OPT and Vegetativo B VGB)

### All technologies and parameters, historical years

- Technologies: all.
- Parameters: all.
- Years: 2023-2026 (OPT) / 2023-2030 (INV and VGB).
- Source: ASSUMPTION LA8 (values equal to the Reference scenario BAU of this model).
- Detail: The historical years take the Reference scenario values so that the scenarios do not diverge before 2027. Exceptions: the interconnector investment cap from 2027 and, in Vegetativo B, minimum generation and minimum investment.
- Assumptions: ASSUMPTION LA8.

## A-O_Demand, Demand_Projection

### `ELC{ISO}XX03` demand, 2023 and 2024

- Technologies: `ELC{ISO}XX03` (electricity demand of each of the 19 countries).
- Parameters: `SpecifiedAnnualDemand`.
- Years: 2023-2024.
- Source: OLADE, electricity generation by source, and OLADE, energy balance matrix (URL needed).
- Detail: Base demand (PJ) = the country's total electricity generation in 2023 and 2024 (OLADE annual statistics, GWh x 0.0036) + electricity imports - electricity exports in the same year (OLADE energy balance matrix). Barbados has no interconnection, so its exchange is zero.

### `ELC{ISO}XX03` demand, 2025 to 2050 (18 countries)

- Technologies: `ELC{ISO}XX03` (18 countries; Haiti has its own block).
- Parameters: `SpecifiedAnnualDemand`.
- Years: 2025-2050.
- Source: high-demand scenario of each country's official electricity planning document (see A-O_Parametrization and A-O_Demand, demand by country).
- Detail: Demand(year) = OLADE 2024 base demand x cumulative growth, applying year by year the growth rate of the high-demand scenario of each national document. Bolivia, Ecuador, Nicaragua and Paraguay publish no high scenario, so their central scenario is used. For years beyond the document's horizon, the trend continues at the compound growth rate of the nearest segment.
- Assumptions: ASSUMPTION LA9.
- Needs source: the reason for using the high-demand scenario rather than the central one.

### `ELCHTIXX03` demand, Haiti, 2025 to 2050

- Technologies: `ELCHTIXX03` (Haiti).
- Parameters: `SpecifiedAnnualDemand`.
- Years: 2025-2050.
- Source: IEA, World Energy Outlook 2025, Net Zero Emissions (NZE) scenario for the 19 countries of the region, https://www.iea.org/reports/world-energy-outlook-2025.
- Detail: Haiti's demand follows the IEA projection. Demand jumps from 3.35 PJ in 2024 (OLADE value) to 6.67 PJ in 2025 (IEA value).

## A-O_Parametrization and A-O_Demand, demand by country (Demand_Projection sheet and interconnector scaling)

All blocks in this section set the high-scenario demand growth used for `SpecifiedAnnualDemand` in 2025 to 2050 and for scaling the interconnector exchange limits (ASSUMPTION LA9).

### Argentina

- Technologies: `ELCARGXX03` and Argentina's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Lineamientos y Escenarios para la Transición Energética al 2050 (Secretaría de Energía de Argentina) (URL needed).
- Detail: Ambitious scenario (Escenario Ambicioso, E3), Table 2, p. 38; the central scenario is the Base scenario (Escenario Base, E1). The 2022-2030 segment is common to both, section 7.2.1, pp. 29-30.
- Assumptions: ASSUMPTION LA9.

### Bolivia

- Technologies: `ELCBOLXX03` and Bolivia's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan Eléctrico Referencial 2035 (Bolivia) (URL needed).
- Detail: Referencial scenario, Table 8, pp. 34-36 and Annex 4, p. 93. The document publishes no high scenario, so the Referencial scenario is used.
- Assumptions: ASSUMPTION LA9.

### Brazil

- Technologies: `ELCBRAXX03` and Brazil's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plano Decenal de Expansão de Energia 2034 (PDE 2034, EPE) (URL needed); Plano Nacional de Energia 2050 (PNE 2050, EPE) (URL needed).
- Detail: High scenario: Cenário Superior of PDE 2034, Figure 2-23 / Box 2.2, p. 62 (4.0% per year). Central scenario: Cenário de Referência, Table 2-3, p. 61, extended to 2050 with Figure 9 of PNE 2050.
- Assumptions: ASSUMPTION LA9.

### Barbados

- Technologies: `ELCBRBXX03`.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Barbados Action Plan / Roadmap 2030 (URL needed).
- Detail: High scenario, section 6.2, p. 35 (the document is narrative, with no numeric table). Central scenario: Base.
- Assumptions: ASSUMPTION LA9.

### Chile

- Technologies: `ELCCHLXX03` and Chile's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Planificación Energética de Largo Plazo (PELP) 2023-2027 (Ministerio de Energía de Chile) (URL needed).
- Detail: Transición Acelerada scenario, Figure 28 (values read from the chart). Central scenario: Carbono Neutralidad.
- Assumptions: ASSUMPTION LA9.

### Colombia

- Technologies: `ELCCOLXX03` and Colombia's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan de Expansión de Referencia Generación 2025-2039 (UPME) (URL needed).
- Detail: IC Superior 68% scenario, Tables 32 and 33, pp. 93 and 95. Central scenario: Demanda Media, Tables 13, 30 and 31.
- Assumptions: ASSUMPTION LA9.

### Costa Rica

- Technologies: `ELCCRIXX03` and Costa Rica's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan de Expansión de la Generación 2024-2040 (ICE) (URL needed).
- Detail: Alta scenario, Generación SEN, Table 6.2, pp. 60-61. Central scenario: Media.
- Assumptions: ASSUMPTION LA9.

### Dominican Republic

- Technologies: `ELCDOMXX03` and the interconnector with Haiti.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan Energético Nacional 2022-2037 (CNE) (URL needed).
- Detail: Alternativo scenario, chapter 6.1.6, Table 182. Central scenario: Tendencial, Table 181.
- Assumptions: ASSUMPTION LA9.

### Ecuador

- Technologies: `ELCECUXX03` and Ecuador's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan Maestro de Electricidad 2023-2032 (Ecuador) (URL needed).
- Detail: Hipótesis 2, Caso Base (demand at generator terminals), chapter 3, Table 3-26, p. 106. The document publishes no high scenario, so the Caso Base is used.
- Assumptions: ASSUMPTION LA9.

### Guatemala

- Technologies: `ELCGTMXX03` and Guatemala's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan de Expansión de la Generación 2020-2050 (MEM / CNEE) (URL needed).
- Detail: ALTO scenario, Table 1, p. 35 (annual series 2020-2050 in GWh). Central scenario: MEDIO.
- Assumptions: ASSUMPTION LA9.

### Honduras

- Technologies: `ELCHNDXX03` and Honduras's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan Indicativo de Expansión de la Generación 2026-2035, and its annex (URL needed).
- Detail: Alta IC 95% scenario, Table 26, p. 54 and annex Table C-3, p. 38. Central scenario: Media, Table 25 and Table C-2.
- Assumptions: ASSUMPTION LA9.

### Haiti

- Technologies: `ELCHTIXX03` (OLADE data for 2023-2024 only).
- Parameters: demand growth, high scenario.
- Years: none.
- Source: IEA, World Energy Outlook 2025 (NZE) for 2025-2050, https://www.iea.org/reports/world-energy-outlook-2025.
- Detail: See the Haiti block under A-O_Demand, Demand_Projection.

### Mexico

- Technologies: `ELCMEXXX03` and the interconnector with Guatemala.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Programa de Desarrollo del Sistema Eléctrico Nacional (PRODESEN) 2024-2038 (SENER) (URL needed).
- Detail: Alto scenario, Table (Cuadro) 3.4, p. 63 (mean annual rate 2.9%). Central scenario: Planeación, Tables (Cuadros) 3.1, 3.4 and 3.5, pp. 38, 63 and 64 (net consumption of the national electricity system, SEN).
- Assumptions: ASSUMPTION LA9.

### Nicaragua

- Technologies: `ELCNICXX03` and Nicaragua's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan Indicativo de Expansión de la Generación 2021-2035 (MEM) (URL needed).
- Detail: Demanda Media of the national interconnected system (SIN), Table 5, p. 10. The document publishes no high scenario, so the Media scenario is used.
- Assumptions: ASSUMPTION LA9.

### Panama

- Technologies: `ELCPANXX03` and Panama's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan de Expansión del Sistema Interconectado Nacional 2025-2039, Tomo II (ETESA) (URL needed).
- Detail: Optimista-Alto scenario, Table 2.1, p. 22. Central scenario: Moderado-Medio.
- Assumptions: ASSUMPTION LA9.

### Peru

- Technologies: `ELCPERXX03` and Peru's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Annex D.1 (Proyecciones de Demanda, Futuros Extremos) of report COES/DP-02-2022 (Plan de Transmisión del SEIN 2023-2032) (URL needed).
- Detail: Muy Optimista scenario, Table D.12, p. 12. Central scenario: Base, Table D.8, p. 10. The document ends in 2032; later years continue at the same rate.
- Assumptions: ASSUMPTION LA9.

### Paraguay

- Technologies: `ELCPRYXX03` and Paraguay's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan Maestro de Generación 2024-2043 (ANDE) (URL needed).
- Detail: Section 4.1.2, Figure 4.7, DTM curve (values read from the chart; only the year 2043, 57 561 GWh, is labelled). The document has a single demand path, used as both the high and the central scenario.
- Assumptions: ASSUMPTION LA9.

### El Salvador

- Technologies: `ELCSLVXX03` and El Salvador's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan Indicativo de la Expansión de la Generación 2024-2038 (CNE) (URL needed).
- Detail: Demanda Alta, Table 14, pp. 46-47. Central scenario: Demanda Base, Table 8, pp. 34-35.
- Assumptions: ASSUMPTION LA9.

### Uruguay

- Technologies: `ELCURYXX03` and Uruguay's interconnectors.
- Parameters: demand growth, high scenario.
- Years: 2025-2050.
- Source: Plan de Expansión de la Generación 2024-2043 (UTE) (URL needed).
- Detail: PES nov/23 + Demanda 50 MW scenario, Table I, section 3.1, pp. 16-17. Central scenario: Incondicional + Demanda 50 MW.
- Assumptions: ASSUMPTION LA9.

## A-O_Demand, Profiles

### `ELC{ISO}{RR}01/02/03` demand profiles

- Technologies: `ELC{ISO}{RR}01/02/03` (hourly demand profile of each country) x 12 time slices.
- Parameters: `SpecifiedDemandProfile`.
- Years: 2023-2050.
- Source: 2015 hourly electricity demand series from the PLEXOS-World 2015 dataset (Brinkerink et al., 2021) (URL needed), via OSeMOSYS Global v2.0.0a, https://github.com/OSeMOSYS/osemosys_global.
- Detail: Hourly demand profile of each country in 2015. About half of the profiles come from historical system operator data; the rest were estimated from countries with similar characteristics. OSeMOSYS Global sums the hours of each time slice (season x part of day) and normalises so that the profile sums to 1 in each year; the same profile applies to every year.

## A-Xtra_Storage, Fixed Horizon Parameters

### `LDS*`, `SDS*`, `BDS*` storage

- Technologies: `LDS*`, `SDS*`, `BDS*` (long-duration storage, short-duration storage and batteries).
- Parameters: `StorageLevelStart`, `OperationalLifeStorage`.
- Years: none (not year-indexed).
- Source: operational life: NREL Annual Technology Baseline, as used in the OSeMOSYS Global v2.0.0a default value table, https://github.com/OSeMOSYS/osemosys_global (ATB URL needed); starting level: OSeMOSYS Global default value.
- Detail: Operational life of 60 years for long-duration storage and 30 years for short-duration storage and batteries (OSeMOSYS Global default table, based on the NREL ATB). Starting storage level 0 (default value). Batteries (`BDS`, Chile and Peru only) take the short-duration values.
- Needs source: the starting storage level of 0.

## A-Xtra_Storage, CapitalCostStorage

### `LDS*` long-duration storage, energy component of cost

- Technologies: `LDS*` (long-duration storage, 19 countries).
- Parameters: `CapitalCostStorage`.
- Years: 2023-2050.
- Source: NREL Annual Technology Baseline (ATB) 2024, National Renewable Energy Laboratory, United States (URL needed).
- Detail: Energy component of storage cost: 20 833.33 million USD per PJ, equal to 75 USD per kWh, constant. The OSeMOSYS Global default, 52 694 million USD per PJ (about 190 USD/kWh), is not used. The power component is on the Secondary Techs sheet of A-O_Parametrization (`PWRLDS*`).

### `SDS*` short-duration storage and `BDS*` batteries, energy component of cost

- Technologies: `SDS*` (short-duration storage, 19 countries) and `BDS*` (batteries, Chile and Peru).
- Parameters: `CapitalCostStorage`.
- Years: 2023-2050.
- Source: NREL Annual Technology Baseline (ATB) 2024, National Renewable Energy Laboratory, United States (URL needed).
- Detail: Energy component of cost: 32 777.77 million USD per PJ, equal to 118 USD per kWh, constant. The OSeMOSYS Global default, 42 051 million USD per PJ (about 151 USD/kWh), is not used. The power component (944 USD/kW) is on the Secondary Techs sheet of A-O_Parametrization (`PWRSDS*`, `PWRBDS*`).

### `LDS*`, `SDS*` existing storage

- Technologies: `LDS*`, `SDS*` (existing storage).
- Parameters: `ResidualStorageCapacity`.
- Years: 2023-2050.
- Source: US DOE / Sandia National Laboratories, Global Energy Storage Database (GESDB), https://gesdb.sandia.gov, via OSeMOSYS Global v2.0.0a, https://github.com/OSeMOSYS/osemosys_global.
- Detail: Existing storage capacity by country from the commissioned projects in GESDB, classified by OSeMOSYS Global as long duration (pumped hydro, compressed air, hydrogen) or short duration (batteries, flywheels, thermal); MWh converted to PJ. Only Argentina, Bolivia, Brazil, Haiti and Mexico have non-zero values.

## A-Xtra_Storage, TechnologyStorage

### Links between storage plants and their reservoirs

- Technologies: `PWRLDS*` / `PWRSDS*` / `PWRBDS*` linked to `LDS*` / `SDS*` / `BDS*` (link between each storage plant and its reservoir).
- Parameters: `TechnologyToStorage`, `TechnologyFromStorage`.
- Years: none (not year-indexed).
- Source: model structure defined by OSeMOSYS Global v2.0.0a, https://github.com/OSeMOSYS/osemosys_global.
- Detail: Says which technology charges (mode 1) and discharges (mode 2) each storage reservoir; it is part of the model structure, not data.

## Values written by the OSTRAM-LAC pre-solver steps

The OSTRAM-LAC model changes some `BAU` solver inputs after compilation, in steps OSTRAM has no exact equivalent for. The `lac` root workbook `A-O_Parametrization.xlsx` holds their final values, so OSTRAM compiles them directly. Each block names the step, the values and their basis.

### Reserve margin and reserve credits

- Technologies and parameters: `ReserveMargin` (sheet System Parameters); `ReserveMarginTagTechnology` of `PWRBCK*` and `PWRCCS*`; `TotalAnnualMaxCapacityInvestment` of `PWRPET*`, `PWROIL*` and `PWRNGS*` where the compiled value was 0 or 9999.
- Years: 2023-2050 (`ReserveMargin`, tags); 2026-2050 (build caps).
- Source: needs source.
- Detail: `ReserveMargin` is 0.15 in every year. The backstop counts fully towards the reserve (tag 1.0) and carbon capture plants at 0.9. Oil, diesel and gas plants whose build cap was 0 or 9999 from 2026 get a build cap of 0.5 GW per year, raised to the planned build where that is larger, except technologies with a planned build in any year.
- Assumptions: ASSUMPTION LA11.

### Fossil dispatch floors

- Technologies: 51 coal, gas, oil and diesel technologies (`PWRCOA*`, `PWRNGS*`, `PWROIL*`, `PWRPET*`).
- Parameter: `TotalTechnologyAnnualActivityLowerLimit`.
- Years: 2027-2050.
- Source: needs source.
- Detail: floor (PJ) = (residual capacity + planned builds, GW) x capacity factor x 31.536, never below the floor the workbook already held. The capacity factor is held from 2027 to 2030 and then falls linearly to 0.05 (or to its starting value when that is lower) in 2050. The starting capacity factor is plan-based where a national plan gives one, else implied by historical generation, else a fuel default.
- Assumptions: ASSUMPTION LA12.

### Internal transmission costs and limits

- Technologies: internal transmission (`PWRTRN*`, `RNWTRN*`, `TRNNLI*`, `RNWNLI*`, `TRNRPO*`, `RNWRPO*`) and the 26 interconnectors (`TRN{A}XX{B}XX`).
- Parameters: `CapitalCost`, `TotalAnnualMaxCapacityInvestment`, `TotalAnnualMaxCapacity` (repowering), `TotalAnnualMinCapacityInvestment` (`RNWTRNBRAXX`).
- Years: 2023-2050.
- Source: IEA, Latin America Energy Outlook 2023, https://www.iea.org/reports/latin-america-energy-outlook-2023 (regional grid investment, the calibration anchor); IEA, Building the Future Transmission Grid (2025), https://www.iea.org/reports/building-the-future-transmission-grid (line replacement rates).
- Detail: `CapitalCost` of the six internal families = OSeMOSYS Global base cost x m x 1.003^(year - 2025) from 2025, where m = 1.7954 scales the planned internal transmission builds of 2024-2029 to an average investment of 3 080 million USD per year. Repowering (`TRNRPO*`, `RNWRPO*`) may add up to 1% of the 2023 internal stock of its family per year from 2028, nothing before, with no total cap. Existing internal lines (`PWRTRN*`, `RNWTRN*`) may grow only by their planned builds. Interconnectors may grow only by their planned builds from 2027. Unplanned new lines (`TRNNLI*`, `RNWNLI*`) keep a total capacity of zero until 2029. The Brazil bipole planned build is 5.23 GW in 2029 and 5.0 GW in 2030.
- Assumptions: ASSUMPTION LA13.

**LA11.** Regional reserve margin 0.15 (fraction, every year, the single model region), backstop reserve credit 1.0, carbon capture reserve credit 0.9, and a build cap of 0.5 GW per year for oil, diesel and gas plants that had no finite cap, from 2026, in the 19 countries. Reason: the backstop and capture plants are dispatchable firm capacity, and the 0.5 GW cap gives dispatchable plants a bounded build allowance so the reserve constraint can be met.

**LA12.** Fossil plants keep a minimum generation from 2027 to 2050: their capacity factor is held at its starting value to 2030 and falls linearly to 0.05 by 2050. Reason: from 2027 the workbook floors stop for most fossil fleets, and without a floor plants with planned builds could stand idle; the decline lets the least-cost mix replace them over time while keeping a small operating level.

**LA13.** Calibration factor m = 1.7954 on internal transmission capital costs, anchor 3 080 million USD per year (2024-2029 average of planned internal transmission investment), real cost escalation 0.3% per year from 2025, repowering of 1% of the existing stock per year from 2028, and unplanned new lines from 2030, in the 19 countries. Reason: the anchor ties the cost of planned internal lines to regional grid investment in the IEA Latin America Energy Outlook 2023; 1% a year is below the line replacement rate the IEA reports for emerging economies; the escalation reflects the real cost trend of conductor metals; the dates allow for project lead times. Check: IEA Latin America Energy Outlook 2023 (URL above).
