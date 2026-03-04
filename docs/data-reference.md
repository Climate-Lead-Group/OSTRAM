# Data Reference

This page documents the data file formats, naming conventions, and OSeMOSYS parameter structure used by OSTRAM.

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
| Region | 2 chars | Sub-region | XX (default), EA, NE, NO |

**Examples:**

| Code | Meaning |
|------|---------|
| `PWRBIOBGDXX` | Power + Biomass + Bangladesh + Default region |
| `MINCOAINDEA` | Mining + Coal + India-East |
| `RNWSPVLKAXX` | Renewable + Solar PV + Sri Lanka + Default region |
| `TRNRPOARGBOLXX` | Transmission Repowered + Argentina-Bolivia + Default region |

### Technology Prefixes

| Prefix | Description |
|--------|-------------|
| `PWR` | Power generation |
| `MIN` | Mining (commodity extraction) |
| `RNW` | Renewable energy supply |
| `ELC` | Electricity distribution |
| `TRN` | Transmission (cross-border) |

### Fuel Codes

Fuel names follow a similar pattern:

```
{SOURCE}{COUNTRY}{REGION}
```

**Examples:**

| Code | Meaning |
|------|---------|
| `BIOBGDXX` | Biomass fuel, Bangladesh |
| `COAINDEA` | Coal fuel, India-East |
| `ELC00BGDXX` | Renewable electricity, Bangladesh |
| `ELC01BGDXX` | Non-renewable electricity, Bangladesh |
| `ELC02BGDXX` | Transmission output electricity, Bangladesh |

### Emission Codes

```
CO2{COUNTRY}
```

Example: `CO2BGD` = CO2 emissions for Bangladesh.

### Storage Codes

```
{TYPE}{COUNTRY}{REGION}
```

| Code | Meaning |
|------|---------|
| `LDSBGDXX` | Long Duration Storage, Bangladesh |
| `SDSBGDXX` | Short Duration Storage, Bangladesh |

---

## CSV File Structure

### SET Files

SET files define the elements of each OSeMOSYS set. They have a single column:

```csv
VALUE
MINCOABGD
MINCOABTN
MINCOAINDEA
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
| `TIMESLICE.csv` | Timeslice codes (S1D1, S1D2, ..., S4D3) |
| `SEASON.csv` | Season codes (1, 2, 3, 4) |
| `DAYTYPE.csv` | Day type codes (1) |
| `DAILYTIMEBRACKET.csv` | Daily time bracket codes (1, 2, 3) |
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
          └── Daily Time Bracket (3 brackets)
```

This produces **12 timeslices** (4 seasons x 3 brackets):

| Timeslice | Season | Bracket | Description |
|-----------|--------|---------|-------------|
| S1D1 | 1 | 1 | Season 1, Bracket 1 |
| S1D2 | 1 | 2 | Season 1, Bracket 2 |
| S1D3 | 1 | 3 | Season 1, Bracket 3 |
| S2D1 | 2 | 1 | Season 2, Bracket 1 |
| ... | ... | ... | ... |
| S4D3 | 4 | 3 | Season 4, Bracket 3 |

### Conversion Matrices

Three matrices map timeslices to temporal dimensions:

- **Conversionls** (12x4): Maps timeslices to seasons (identity-like).
- **Conversionld** (12x1): Maps timeslices to day types (all 1).
- **Conversionlh** (12x3): Maps timeslices to daily time brackets.

---

## Excel Model Files

### A-O_Parametrization.xlsx

The main parameter file per scenario. Contains multiple sheets:

| Sheet | Content |
|-------|---------|
| Fixed Horizon Parameters | CapacityToActivityUnit, OperationalLife |
| Secondary Techs | Cost and capacity parameters for power technologies |
| Demand Techs | Parameters for demand technologies |
| Capacities | ResidualCapacity, TotalAnnualMaxCapacity, etc. |
| VariableCost | Variable cost data |
| TotalAnnualMaxCapacityInvestment | Investment limits |

### A-O_Demand.xlsx

Demand data per scenario:

| Sheet | Content |
|-------|---------|
| Demand_Projection | SpecifiedAnnualDemand values |
| Profiles | SpecifiedDemandProfile timeslice distribution |

### A-O_AR_Model_Base_Year.xlsx

Base year activity ratios:

| Sheet | Content |
|-------|---------|
| Primary | InputActivityRatio / OutputActivityRatio for primary supply |
| Secondary | Activity ratios for power generation |
| Demands | Activity ratios for demand technologies |

### A-O_AR_Projections.xlsx

Projection activity ratios (same structure as base year, with projection modes).

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
| `TotalAnnualMaxCapacityInvestment` | GW | Maximum new capacity per year |
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
| `TotalTechnologyAnnualActivityLowerLimit` | PJ | Minimum annual generation |
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

OSTRAM uses a single-region (`GLOBAL`) architecture where geographic granularity is embedded in the technology and fuel naming conventions. Each country's technologies operate independently, connected only through explicit transmission (TRN) technologies.

```
Mining (MIN)  →  Power (PWR)  →  Electricity (ELC)  →  Demand
                      ↕
              Transmission (TRN)  ←→  Other Countries
                      ↕
               Renewables (RNW)
```

### Energy Flow

1. **Mining technologies** extract primary commodities (coal, gas, oil, etc.).
2. **Power technologies** convert fuels into electricity.
3. **Renewable technologies** supply renewable electricity.
4. **Transmission technologies** route electricity between countries.
5. **Electricity distribution** delivers power to meet demand.
6. **Storage technologies** balance supply and demand across timeslices.
