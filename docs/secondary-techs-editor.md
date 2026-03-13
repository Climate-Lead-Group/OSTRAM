# Secondary Technologies Editor

The Secondary Technologies Editor provides a user-friendly Excel interface for modifying technology parameters across scenarios, with support for automatic OSTRAM source data integration.

## Overview

The editor workflow has two steps:

1. **Generate the template** (`D1`) -- Creates `Secondary_Techs_Editor.xlsx`.
2. **Apply changes** (`D2`) -- Reads the filled template and updates the model files.

---

## Step 1: Generate the Editor Template

**Script:** `t1_confection/D1_generate_editor_template.py`

From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python t1_confection/D1_generate_editor_template.py
```

### What It Creates

The script generates `Secondary_Techs_Editor.xlsx` with these sheets:

| Sheet | Purpose |
|-------|---------|
| **Instructions** | User guide with editing instructions |
| **Documentation** | Full technical documentation on calculations |
| **OSTRAM_Config** | Toggle switches for automatic OSTRAM data integration |
| **Demand_Growth** | Demand growth rate configuration per country |
| **Scenarios_Demand_Growth** | Scenario-specific demand growth rates |
| **Renewability_Targets** | Renewable percentage targets per year/country |
| **Technology_Weights** | Custom distribution of renewable/non-renewable technologies |
| **Editor** | Main editing area with dropdown lists |
| **Interconnections** | Transmission interconnection technologies (only if TRN technologies exist) |
| *(Hidden sheets)* | Validation data for dropdown lists |

### How It Works

The script:
1. Reads all `A-O_Parametrization.xlsx` files across scenarios.
2. Collects existing parameter data and transmission interconnections.
3. Builds the Excel template with:
   - **Dropdown lists** for Scenario, Country, Technology, and Parameter selection.
   - **Auto-fill formulas** (VLOOKUP) for the Tech column.
   - **Year columns** (2021--2050) for entering values.

---

## Step 2: Edit the Template

### Manual Editing (Editor Sheet)

In the **Editor** sheet (and similarly in the **Interconnections** sheet for transmission technologies):

1. **Select a Scenario**: Choose from the auto-discovered scenarios (based on existing `A1_Outputs_*` folders), or ALL (applies to every scenario).
2. **Select a Country**: Pick a country from the dropdown.
3. **Select a Technology**: Choose by Tech.Name (descriptive name). The Tech code auto-populates.
4. **Select a Parameter**: Choose which parameter to modify (e.g., CapitalCost, ResidualCapacity).
5. **Enter values**: Fill in the year columns (2021--2050) with your desired values.

### OSTRAM Configuration (OSTRAM_Config Sheet)

The **OSTRAM_Config** sheet provides toggle switches for automatic data population:

| Parameter | Values | Description |
|-----------|--------|-------------|
| `ResidualCapacitiesFromOSTRAM` | YES / NO | Auto-populate ResidualCapacity from installed capacity data |
| `PetroleumSplitMode` | `OIL_only` / `Split_PET_OIL` | How to handle petroleum capacity allocation |
| `DemandFromOSTRAM` | YES / NO | Auto-populate electricity demand from generation data |
| `ActivityLowerLimitFromOSTRAM` | YES / NO | Auto-populate TotalTechnologyAnnualActivityLowerLimit |
| `ActivityUpperLimitFromOSTRAM` | YES / NO | Auto-populate TotalTechnologyAnnualActivityUpperLimit |
| `TradeBalanceDemandAdjustment` | YES / NO | Adjust demand based on trade balance data |
| `InterconnectionsControl` | ON / OFF | Enable/disable interconnection technologies |

### Petroleum Split Modes

| Mode | Behavior |
|------|----------|
| `OIL_only` | All petroleum capacity is assigned to OIL (Fuel oil) |
| `Split_PET_OIL` | Capacity is split between PET (Diesel) and OIL (Fuel oil + Bunker) using proportions from `Shares_PET_OIL_Split.xlsx` |

### Demand Configuration (Demand_Growth Sheet)

When `DemandFromOSTRAM` is YES, configure growth rates per country:

- Uses OSTRAM generation data as the base.
- Applies linear growth: `Demand(year) = Demand(2023) * (1 + rate * (year - 2023))`.
- Growth rates are specified per country in the **Demand_Growth** sheet.
- Scenario-specific overrides are available in **Scenarios_Demand_Growth**.

### Renewability Targets (Renewability_Targets Sheet)

When Activity Limits from OSTRAM are enabled:

- Define target renewable percentages per year and country.
- The system interpolates between specified target years.
- Targets affect the distribution of activity limits across technologies.

### Technology Weights (Technology_Weights Sheet)

Customize how activity limits are distributed among technologies:

- Define weights for individual renewable and non-renewable technologies.
- Weights determine each technology's share of the total activity limit.

---

## Step 3: Apply Changes

**Script:** `t1_confection/D2_update_secondary_techs.py`

From an **Anaconda Prompt** (with the `OG-MOMF-env` environment activated):

```bash
python t1_confection/D2_update_secondary_techs.py
```

### What It Does

1. Reads the filled `Secondary_Techs_Editor.xlsx`.
2. Reads OSTRAM configuration toggles.
3. For each scenario:
   - Creates a **backup** of the parametrization file.
   - Applies manual edits from the Editor sheet.
   - If OSTRAM integration is enabled:
     - Reads capacity data (MW to GW conversion).
     - Reads generation data (GWh to PJ conversion).
     - Applies petroleum split logic.
     - Calculates activity limits with renewability targets.
     - Validates limits against available capacities.
   - Updates `A-O_Parametrization.xlsx`.
   - Updates `A-O_Demand.xlsx` (if demand integration is enabled).
4. Sets **Projection.Mode** to "User defined" for modified parameters.

### Unit Conversions

| Source | Target | Conversion |
|--------|--------|------------|
| MW (source capacity) | GW (model) | / 1000 |
| GWh (source generation) | PJ (model) | * 0.0036 |

### Safety Features

- **Automatic backups**: One backup per scenario before applying changes.
- **Activity limit validation**: Verifies that limits do not exceed available capacity.
- **Detailed logging**: Full log output with country identification for each operation.

---

## Related Files

| File | Description |
|------|-------------|
| `D1_generate_editor_template.py` | Generates the Excel template |
| `D2_update_secondary_techs.py` | Applies changes to scenario files |
| `Secondary_Techs_Editor.xlsx` | The editor template (generated) |
| `OSTRAM - Installed Capacity by Source - Annual.xlsx` | Installed capacity source data |
| `OSTRAM - Electric Generation by Source - Annual.xlsx` | Electricity generation source data |
| `Shares_PET_OIL_Split.xlsx` | Petroleum/oil split proportions per scenario |
| `Shares_Power_Generation_Technologies.xlsx` | Power generation technology shares |
