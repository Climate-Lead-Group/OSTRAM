# Secondary Technologies Editor

The Secondary Technologies Editor provides a user-friendly Excel interface for modifying technology parameters across scenarios, with support for automatic OSTRAM source data integration.

:::{note}
This is a **manual, optional** layer, independent of the automated Stage A3 scenario engine (retirement schedules, investment lids, VRE targets, etc. -- see {doc}`pipeline`). D1 reads whatever currently exists in each scenario's `A1_Outputs_<scenario>/` folder, regardless of whether it was produced by plain A1/A2 or by the full A3 rule-script chain, and D2 writes edits back into those same files. A3 never invokes D1/D2, and D1/D2 never touch `A3_process/` internals. Use this workflow for one-off overrides, interconnection ON/OFF toggling, or OSTRAM-source auto-population not covered by an A3 rule script.
:::

## Overview

The editor workflow has two steps, neither of which takes command-line arguments -- both scripts operate on whatever scenario folders currently exist under `A1_Outputs/`:

1. **Generate the template** (`D1`) -- Creates `Secondary_Techs_Editor.xlsx`.
2. **Apply changes** (`D2`) -- Reads the filled template and updates the model files.

---

## Step 1: Generate the Editor Template

**Script:** `t1_confection/D1_generate_editor_template.py`

From an **Anaconda Prompt** (with the `OSTRAM-env` environment activated):

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

The **OSTRAM_Config** sheet provides toggle switches for automatic data population. All are `YES`/`NO` switches (default `NO` except `PetroleumSplitMode`):

| Parameter | Values | Default | Description |
|-----------|--------|---------|-------------|
| `ResidualCapacitiesFromOSTRAM` | YES / NO | NO | Auto-populate ResidualCapacity from installed capacity data |
| `PetroleumSplitMode` | `OIL_only` / `Split_PET_OIL` | Split_PET_OIL | How to handle petroleum capacity allocation |
| `DemandFromOSTRAM` | YES / NO | NO | Auto-populate electricity demand from generation data |
| `ActivityLowerLimitFromOSTRAM` | YES / NO | NO | Auto-populate TotalTechnologyAnnualActivityLowerLimit |
| `ActivityUpperLimitFromOSTRAM` | YES / NO | NO | Auto-populate TotalTechnologyAnnualActivityUpperLimit |
| `TradeBalanceDemandAdjustment` | YES / NO | NO | Adjust demand based on trade balance data |
| `InterconnectionsControl` | YES / NO | NO | Master switch for the interconnection controls below |

:::{note}
`InterconnectionsControl` itself is `YES`/`NO`, like every other master toggle in this sheet. `ON`/`OFF` is used only for the per-row **Status** column inside the separate **Interconnections** sheet (one row per TRN technology), not for this master switch.
:::

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

From an **Anaconda Prompt** (with the `OSTRAM-env` environment activated):

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

## Optional Pre-Fill Step: Interconnection Limits From Flows

**Script:** `t1_confection/Z_AUX_D1b_set_trn_limits_from_flows.py`

A standalone, manual tool meant to run **between D1 and D2** (not invoked automatically by either). It reads bilateral electricity flow data from `Matriz Balance energético/flujos_energia_estimados_optimizacion.xlsx` (sheet "Flujos por Interconexión") and writes `TotalTechnologyAnnualActivityLowerLimit`/`UpperLimit` (±5% of the historical flow, converted GWh → PJ) directly into the already-generated `Secondary_Techs_Editor.xlsx`'s **Editor** sheet, as a pre-fill for TRN interconnection rows before you run D2. It takes no command-line arguments.

```bash
python t1_confection/D1_generate_editor_template.py
python t1_confection/Z_AUX_D1b_set_trn_limits_from_flows.py   # optional pre-fill
# ... review/adjust the Editor sheet ...
python t1_confection/D2_update_secondary_techs.py
```

## Related Files

| File | Description |
|------|-------------|
| `D1_generate_editor_template.py` | Generates the Excel template |
| `D2_update_secondary_techs.py` | Applies changes to scenario files |
| `Z_AUX_D1b_set_trn_limits_from_flows.py` | Optional pre-fill of TRN interconnection activity limits from flow data |
| `Secondary_Techs_Editor.xlsx` | The editor template (generated) |
| `OSTRAM - Installed Capacity by Source - Annual.xlsx` | Installed capacity source data (MW → GW) |
| `OSTRAM - Electric Generation by Source - Annual.xlsx` | Electricity generation source data (GWh → PJ) |
| `Shares_PET_OIL_Split.xlsx` | Petroleum/oil split proportions per scenario |
| `Shares_Power_Generation_Technologies.xlsx` | Power generation technology shares |
| `flujos_energia_estimados_optimizacion.xlsx` (`Matriz Balance energético/`) | Bilateral trade-balance and interconnection flow data (GWh → PJ), used by `TradeBalanceDemandAdjustment` and by `Z_AUX_D1b_set_trn_limits_from_flows.py` |

:::{note}
`CapacityAndDistances.xlsx` and `RateGrowthDemand_RenovabilityGoals.xlsx` also exist at `t1_confection/` root but are not currently read by D1, D2, or any other script -- they appear to be staged data for future use, not active inputs.
:::
