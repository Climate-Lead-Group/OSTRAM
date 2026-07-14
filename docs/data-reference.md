# Data Reference

This page documents the data file formats, naming conventions, and OSeMOSYS parameter structure used by OSTRAM, based on the current South/Southeast Asia model (Bangladesh, Bhutan, India [5 regions], Nepal, Sri Lanka, Maldives).

## Naming Conventions

### Technology Codes

Technology names in OSTRAM follow a structured format that encodes the technology type, energy source, country, and region:

```
{PREFIX}{SOURCE}{COUNTRY}{REGION}
```

| Component | Length | Description | Examples |
|-----------|--------|-------------|----------|
| Prefix | 3 chars | Technology category | PWR, MIN, RNW, ELC, TRN |
| Source | 3 chars | Energy source | BIO, COA, SPV, WON, HYD |
| Country | 3 chars | ISO-3 country code | BGD, IND, LKA |
| Region | 2 chars | Sub-region | XX (default, single-region countries); EA, NE, NO, SO, WE (India's 5 regions) |

**Examples (as they appear in the compiled model, after A1's PWR-suffix cleanup):**

| Code | Meaning |
|------|---------|
| `PWRBIOBGDXX` | Power + Biomass + Bangladesh + Default region |
| `PWRHYDINDEA` | Power + Hydroelectric + India-East |
| `RNWSPVLKAXX` | Renewable resource + Solar PV + Sri Lanka + Default region |
| `MINCOAIND` | Mining + Coal + India (mining/emission codes are country-level only -- they are **not** split by India's 5 sub-regions) |

:::{note}
Raw `OG_csvs_inputs/` CSVs still carry a trailing `01`/`00` PWR-duplicate suffix (e.g. `PWRBIOBGDXX01`) before Stage A1's `pwr_cleanup_mode` merges/drops the duplicate and strips the suffix. Storage codes (`LDS`/`SDS`) keep a `01` suffix even in the final compiled model (e.g. `LDSBGDXX01`) -- PWR and storage codes are cleaned up differently.
:::

### Technology Prefixes

| Prefix | Description |
|--------|-------------|
| `PWR` | Power generation |
| `MIN` | Mining (tradable commodity extraction) -- country-level, not region-split |
| `RNW` | Renewable resource supply (region-split) |
| `ELC` | Electricity distribution/fuel bus |
| `TRN` | Cross-border interconnector (see below) |
| `RNWTRN`/`RNWRPO`/`RNWNLI`/`PWRTRN`/`TRNRPO`/`TRNNLI` | Per-country grid-tier conversion technologies added by Stage A2 (existing/repowered/new-line, renewable/non-renewable) -- **not** interconnectors themselves, despite the `TRN`-containing names; see Fuel Codes below |
| `DSPTRN` | Dispatch technology (routes electricity to/from cross-border interconnections) |

### Cross-Border Interconnector Codes

The literal cross-border transmission line between two countries/regions uses a different, longer pattern than the per-country technologies above:

```
TRN{ORIGIN 5-char}{DEST 5-char}
```

Each 5-char block is a 3-letter country code + 2-letter region code (13 characters total after `TRN`). Examples from the live data: `TRNBGDXXINDEA` (Bangladesh ↔ India-East), `TRNINDEAINDNE` (India-East ↔ India-North-East). These pre-exist in the raw input data and have their fuel codes rewritten by Stage A2; `TRNNLI*`-prefixed per-country technologies (not to be confused with this pair-level `TRN` code) represent *unplanned candidate* lines rather than existing/committed ones.

### Fuel Codes

Fuel names follow a similar pattern, with a 2-digit tier suffix instead of a region suffix on electricity buses:

```
{SOURCE}{COUNTRY}{REGION}
```

**Examples:**

| Code | Meaning |
|------|---------|
| `BIOBGDXX` | Biomass fuel, Bangladesh |
| `COAINDEA` | Coal fuel, India-East |
| `ELCBGDXX00` | Renewable power plant output, Bangladesh |
| `ELCBGDXX01` | Non-renewable power plant output, Bangladesh |
| `ELCBGDXX02` | Transmission line output (feeds domestic demand and export dispatch), Bangladesh |
| `ELCBGDXX03` | Dispatch-ready for interconnection (export side), Bangladesh |
| `ELCBGDXX04` | Imported electricity (import side, from a cross-border interconnector), Bangladesh |

The four-tier `ELC*00`-`ELC*04` electricity scheme (five codes: 00, 01, 02, 03, 04) is central to how cross-border trade is modeled -- see **Model Architecture** below for the full flow.

### Emission Codes

```
CO2{COUNTRY}
```

Example: `CO2BGD` = CO2 emissions for Bangladesh. Note `OG_csvs_inputs/EMISSION.csv` still carries a full global list of `CO2<ISO3>` codes inherited from the broader source dataset (e.g. `CO2ARG`, `CO2ARE`) -- only the six codes for the active countries (`CO2BGD`, `CO2BTN`, `CO2IND`, `CO2NPL`, `CO2LKA`, `CO2MDV`) are populated/relevant in the current model; India's emission code is country-level, not split by region.

### Storage Codes

```
{TYPE}{COUNTRY}{REGION}{SUFFIX}
```

| Code | Meaning |
|------|---------|
| `LDSBGDXX01` | Long Duration Storage, Bangladesh |
| `SDSINDEA01` | Short Duration Storage, India-East |

---

## CSV File Structure

### SET Files

SET files define the elements of each OSeMOSYS set. They have a single column:

```csv
VALUE
MINCOABGD
MINCOABTN
MINCOAIND
```

SET files in `OG_csvs_inputs/`:

| File | Content |
|------|---------|
| `TECHNOLOGY.csv` | All technology codes |
| `FUEL.csv` | All fuel codes |
| `EMISSION.csv` | All emission codes |
| `STORAGE.csv` | All storage codes |
| `REGION.csv` | Region codes (typically just `GLOBAL`) |
| `YEAR.csv` | Model years (2023--2050) |
| `TIMESLICE.csv` | Timeslice codes (S1D1..S4D5, 20 total) |
| `SEASON.csv` | Season codes (1, 2, 3, 4) |
| `DAYTYPE.csv` | Day type codes (1) |
| `DAILYTIMEBRACKET.csv` | Daily time bracket codes (1, 2, 3, 4, 5) |
| `MODE_OF_OPERATION.csv` | Mode codes (1, 2) |

### Parameter Files

Parameter files contain data values indexed by OSeMOSYS dimensions. The column structure varies by parameter:

**4-column format** (Region, Technology, Year, Value):

```csv
REGION,TECHNOLOGY,YEAR,VALUE
GLOBAL,PWRBCKBGDXX,2023,999999.0
GLOBAL,PWRBIOBGDXX,2023,1500.0
```

Used by: `CapitalCost`, `FixedCost`, `VariableCost`, `ResidualCapacity`, `TotalAnnualMaxCapacity`, `TotalAnnualMaxCapacityInvestment`, `AvailabilityFactor`, and others.

**6-column format** (Region, Technology, Fuel, Mode, Year, Value):

```csv
REGION,TECHNOLOGY,FUEL,MODE_OF_OPERATION,YEAR,VALUE
GLOBAL,PWRBIOBGDXX,BIOBGDXX,1,2023,3.67
```

Used by: `InputActivityRatio`, `OutputActivityRatio`, `EmissionActivityRatio`.

**Other formats:**

- `CapacityFactor`: REGION, TECHNOLOGY, TIMESLICE, YEAR, VALUE
- `SpecifiedDemandProfile`: REGION, FUEL, TIMESLICE, YEAR, VALUE
- `TradeRoute`: REGION, FUEL, REGION, YEAR, VALUE

---

## Temporal Structure

OSTRAM uses a hierarchical temporal structure:

```
Year (2023-2050)
  └── Season (4 seasons)
      └── Day Type (1 type)
          └── Daily Time Bracket (5 brackets)
```

This produces **20 timeslices** (4 seasons x 5 brackets), expanded from an earlier 12-timeslice (4x3) structure:

| Timeslice | Season | Bracket |
|-----------|--------|---------|
| S1D1 .. S1D5 | 1 | 1 .. 5 |
| S2D1 .. S2D5 | 2 | 1 .. 5 |
| S3D1 .. S3D5 | 3 | 1 .. 5 |
| S4D1 .. S4D5 | 4 | 1 .. 5 |

### Conversion Matrices

Three matrices map timeslices to temporal dimensions, defined directly in `Config_MOMF_T1_A.yaml`:

- **Conversionls** (20x4): One-hot mapping of each timeslice to its season -- `S1D1..S1D5 → [1,0,0,0]`, `S2D1..S2D5 → [0,1,0,0]`, `S3D1..S3D5 → [0,0,1,0]`, `S4D1..S4D5 → [0,0,0,1]`.
- **Conversionld** (20x1): All values are `1` -- there is only one `DayType`, so the mapping is degenerate.
- **Conversionlh** (20x5): One-hot mapping of each timeslice to its daily time bracket, cycling every 5 timeslices within each season -- `D1 → [1,0,0,0,0]`, `D2 → [0,1,0,0,0]`, ..., `D5 → [0,0,0,0,1]`, repeating identically across all 4 seasons.

---

## Excel Model Files

### A-O_Parametrization.xlsx

The main parameter file per scenario. Sheets (confirmed from the current workbook structure):

| Sheet | Content |
|-------|---------|
| Fixed Horizon Parameters | CapacityToActivityUnit, OperationalLife |
| Primary Techs | Costs, capacities, limits for primary (MIN/RNW) supply technologies |
| Secondary Techs | Cost and capacity parameters for power (PWR) technologies |
| Capacities | CapacityFactor by timeslice |
| Yearsplit | YearSplit |
| DaySplit | DaySplit |
| VariableCost | Variable cost data by mode |
| Other_Techs | Miscellaneous technology parameters outside Primary/Secondary/Demand |
| Demand Techs | Parameters for the A2-generated transmission/dispatch technologies |
| System Parameters | ReserveMargin (optional; written by A3's `A0_insert_reserve_margin.py`) |
| Vehicle Techs / Vehicle Groups / Transport Fuel Distribution | Transport parameters, only read when `Use_Transport: true` (currently `false`) |

### A-O_Demand.xlsx

Demand data per scenario:

| Sheet | Content |
|-------|---------|
| Demand_Projection | SpecifiedAnnualDemand values |
| Profiles | SpecifiedDemandProfile timeslice distribution |

### A-O_AR_Model_Base_Year.xlsx / A-O_AR_Projections.xlsx

Base year and projection activity ratios:

| Sheet | Content |
|-------|---------|
| Primary | InputActivityRatio / OutputActivityRatio for primary supply |
| Secondary | Activity ratios for power generation, incl. TRN interconnector fuel-tier rewrites |
| Demand Techs | Activity ratios for the A2-generated transmission/dispatch technologies |
| Distribution Transport / Transport / Transport Groups | Transport activity ratios (only read if `Use_Transport: true`) |

---

## OSeMOSYS Parameters Reference

### Cost Parameters

| Parameter | Unit | Description |
|-----------|------|-------------|
| `CapitalCost` | M$/GW | Overnight capital cost |
| `FixedCost` | M$/GW/yr | Annual fixed O&M cost |
| `VariableCost` | M$/PJ | Variable O&M cost |
| `CapitalCostStorage` | M$/GW | Storage capital cost |

### Capacity Parameters

| Parameter | Unit | Description |
|-----------|------|-------------|
| `ResidualCapacity` | GW | Existing installed capacity |
| `TotalAnnualMaxCapacity` | GW | Maximum total capacity allowed |
| `TotalAnnualMaxCapacityInvestment` | GW | Maximum new capacity per year (the "investment lid" A3 rule scripts write to) |
| `TotalAnnualMinCapacity` | GW | Minimum required capacity |
| `TotalAnnualMinCapacityInvestment` | GW | Minimum new capacity per year |

### Performance Parameters

| Parameter | Unit | Description |
|-----------|------|-------------|
| `AvailabilityFactor` | fraction | Maximum available fraction of capacity |
| `CapacityFactor` | fraction | Capacity factor by timeslice |
| `CapacityToActivityUnit` | PJ/GW/yr | Conversion factor (typically 31.536) |
| `OperationalLife` | years | Technology lifetime |
| `InputActivityRatio` | - | Fuel input per unit activity |
| `OutputActivityRatio` | - | Fuel output per unit activity |

### Demand Parameters

| Parameter | Unit | Description |
|-----------|------|-------------|
| `SpecifiedAnnualDemand` | PJ | Annual energy demand |
| `SpecifiedDemandProfile` | fraction | Timeslice distribution (must sum to 1.0) |

### Emission Parameters

| Parameter | Unit | Description |
|-----------|------|-------------|
| `EmissionActivityRatio` | Mt/PJ | Emissions per unit activity |
| `AnnualEmissionLimit` | Mt | Maximum annual emissions |
| `ModelPeriodEmissionLimit` | Mt | Maximum total emissions |

### Activity Limits

| Parameter | Unit | Description |
|-----------|------|-------------|
| `TotalTechnologyAnnualActivityLowerLimit` | PJ | Minimum annual generation (used by `set_vre_targets.py` to enforce VRE floors) |
| `TotalTechnologyAnnualActivityUpperLimit` | PJ | Maximum annual generation |

### Storage Parameters

| Parameter | Unit | Description |
|-----------|------|-------------|
| `StorageLevelStart` | PJ | Initial storage level |
| `StorageMaxChargeRate` | GW | Maximum charging rate |
| `StorageMaxDischargeRate` | GW | Maximum discharging rate |
| `OperationalLifeStorage` | years | Storage technology lifetime |
| `ResidualStorageCapacity` | GW | Existing storage capacity |

---

## Model Architecture

OSTRAM uses a single-region (`GLOBAL`) architecture where geographic granularity is embedded in the technology and fuel naming conventions. Each country/region's technologies operate independently, connected only through explicit cross-border interconnector (`TRN`) technologies.

```
Mining (MIN)  →  Power (PWR)  →  ELC*00/01 (plant output)
                                        ↓
                        Grid-tier conversion (RNWTRN/PWRTRN/RNWRPO/TRNRPO/RNWNLI/TRNNLI)
                                        ↓
                                   ELC*02 (line output)  →  Demand
                                        ↓
                              DSPTRN Mode 1 (export)
                                        ↓
                                   ELC*03 (dispatch-ready)
                                        ↓
                      Cross-border interconnector (TRN{origin}{dest})
                                        ↓
                                   ELC*04 (imported, at destination)
                                        ↓
                              DSPTRN Mode 2 (import)
                                        ↓
                                   ELC*03 (dispatch-ready again, available locally)
```

### Energy Flow

1. **Mining technologies** (`MIN`) extract primary commodities (coal, gas, oil, etc.), country-level (not region-split).
2. **Power technologies** (`PWR`) convert fuels into electricity: `ELC*00` for renewable output, `ELC*01` for non-renewable output.
3. **Renewable resource technologies** (`RNW`) supply renewable primary energy, region-split.
4. **Grid-tier conversion technologies** (`RNWTRN`/`PWRTRN`/`RNWRPO`/`TRNRPO`/`RNWNLI`/`TRNNLI`, one set per country-region, added by Stage A2) convert `ELC*00`/`ELC*01` into `ELC*02`, the tier available to meet domestic demand.
5. **Dispatch technology** (`DSPTRN`) routes electricity to/from cross-border interconnections: Mode 1 converts `ELC*02` → `ELC*03` (dispatch-ready for export); Mode 2 converts `ELC*04` (imported) → `ELC*03` (dispatch-ready again, for local use).
6. **Cross-border interconnector technologies** (`TRN{origin}{dest}`, a pre-existing pair-level code, distinct from the per-country A2-generated technologies above) transport electricity between countries/regions, consuming `ELC*03` at the origin and producing `ELC*04` at the destination.
7. **Storage technologies** (`LDS`/`SDS`) balance supply and demand across timeslices, exchanging with the `ELC*00` renewable-output tier.
