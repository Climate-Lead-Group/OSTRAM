# Auxiliary Tools

OSTRAM includes several utility scripts (prefixed with `Z_`) for data maintenance, visualization, and support tasks. None of the scripts on this page take command-line arguments unless noted otherwise -- most operate on hardcoded paths relative to `t1_confection/`.

## Configuration Loader

**Script:** `t1_confection/Z_AUX_config_loader.py`

A centralized module (not run directly) that provides cached access to `Config_country_codes.yaml`. All other scripts import functions from this module instead of reading the YAML directly.

### Available Functions

| Function | Returns | Description |
|----------|---------|-------------|
| `get_country_data()` | `dict` | Raw `country_data` block (iso3 → english_name/ostram_name) |
| `get_countries()` | `list[str]` | Sorted list of active country/region codes |
| `get_country_names()` | `dict[str, str]` | `{iso3: english_name}` |
| `get_iso_country_map()` | `dict[str, str]` | `{iso3: english_name}` including special entries |
| `get_ostram_country_mapping()` | `dict[str, str]` | `{ostram_name: iso3}` |
| `get_ostram_country_mapping_normalized()` | `dict[str, str]` | Accent-stripped version of the above |
| `get_shares_country_mapping()` | `dict[str, str]` | Country-name mapping for the Shares workbook |
| `get_first_year()` | `int` | Model start year (default: 2023) |
| `get_add_missing_countries_from_ostram()` | `bool` | Whether to fill missing countries from OSTRAM source data |
| `get_pwr_cleanup_mode()` | `str \| bool` | `"drop"`, `"merge"`, or `False` |
| `get_force_empty_max_capacity_investment_pwr()` | `bool` | Whether to force `Projection.Mode = EMPTY` for PWR `TotalAnnualMaxCapacityInvestment` |
| `get_ostram_tech_mapping()` | `dict[str, str]` | OSTRAM tech name to model code mapping |
| `get_code_to_energy()` | `dict[str, str]` | `{tech_code: description}` |
| `get_renewable_fuels()` | `set[str]` | Set of renewable fuel codes |
| `get_shares_tech_mapping()` | `dict[str, str]` | Shares file name to model code mapping |
| `get_enable_dsptrn()` | `bool` | Whether DSPTRN dispatch injection is enabled |
| `get_model_countries_list()` | `list[str]` | Raw `countries` list (may mix 3-char and 5-char region codes) |
| `get_multi_region_map()` | `dict[str, list[str]]` | `{iso3: [region_codes]}`, e.g. `{'IND': ['EA','NE','NO','SO','WE']}` |
| `get_raw_config()` | `dict` | The full raw YAML dictionary |

### Usage in Scripts

```python
from Z_AUX_config_loader import get_countries, get_first_year

countries = get_countries()  # ['BGD', 'BTN', 'INDEA', ...]
year = get_first_year()      # 2023
```

---

## Demand Profile Normalizer

**Script:** `t1_confection/Z_AUX_fix_excel_profiles.py`

Fixes rounding drift in SpecifiedDemandProfile sheets that can cause OSeMOSYS model errors. Profiles must sum to exactly 1.0 per fuel/technology per year.

### Usage

```bash
python t1_confection/Z_AUX_fix_excel_profiles.py
```

### What It Does

1. Iterates over a **hardcoded** scenario list (`main()`, currently `["BAU", "NDC", "NDC+ELC", "NDC_NoRPO"]`).
2. Opens each `A1_Outputs_<scenario>/A-O_Demand.xlsx` file, if it exists.
3. For each profile sheet, normalizes values so that each fuel/technology column sums to exactly 1.0 per year.
4. Creates a **timestamped backup** before modifying any file.

:::{warning}
The hardcoded scenario list is stale -- it dates from the model's earlier LATAM phase. None of `NDC`, `NDC+ELC`, `NDC_NoRPO` exist in the current model; only `BAU` (of the four hardcoded names) is real. Real current scenarios are `BAU`, `A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE`. Update the list in the script before relying on it for anything beyond `BAU`.
:::

### Tolerance

Values within `0.0001` of 1.0 are considered acceptable. Values outside this range are corrected by proportional scaling.

---

## Interactive Dashboard Generator

**Script:** `t1_confection/Z_AUX_generate_interactive_dashboards_aggregated.py`

Generates standalone HTML dashboards with embedded Plotly.js charts for analyzing power (PWR) technology results. Prompts interactively (`input()`) for which CSV files to load from the current directory -- it is not scriptable/headless.

### Usage

```bash
python t1_confection/Z_AUX_generate_interactive_dashboards_aggregated.py
```

### What It Produces

Standalone HTML files containing:

- **Renewability share charts**: Percentage of renewable vs. non-renewable power generation.
- **Total sum charts**: Aggregated capacity or generation by technology type.
- **Temporal evolution charts**: How the technology mix changes over the model horizon.

All charts are interactive (zoom, hover, filter) and require no external dependencies -- they embed Plotly.js directly in the HTML.

### PWR Technology Validation

The dashboard uses a regex pattern to identify valid power technologies:

```
^PWR(BIO|WAS|CSP|SPV|GEO|HYD|WAV|WON|WOF|URN|NGS|COA|COG|OIL|PET|CCS|OTH)[A-Z]{3}XX$
```

:::{warning}
This pattern hardcodes a literal `XX` region suffix. It correctly matches the five single-region countries (`BGDXX`, `BTNXX`, `LKAXX`, `MDVXX`, `NPLXX`) but **silently excludes all of India's 5 regions** (`INDEA`, `INDNE`, `INDNO`, `INDSO`, `INDWE`), since their region suffix isn't literally `XX`. Anyone using this dashboard on the current model should be aware that India's power technologies won't appear in the output. Widening the pattern to also accept `EA|NE|NO|SO|WE` would fix this.
:::

---

## CSV Sorter

**Script:** `t1_confection/Z_AUX_sort_csv.py`

Sorts all CSV files in a directory by all columns. Used to ensure deterministic file ordering for reproducibility and version control.

### Usage

```bash
python t1_confection/Z_AUX_sort_csv.py
```

When run interactively, it prompts for a folder path. The script can also be imported and used programmatically:

```python
from Z_AUX_sort_csv import sort_csv_files_in_folder

sort_csv_files_in_folder("path/to/csv/folder")
```

---

## Region Consolidation (Brazil) -- Legacy, Unused

**Script:** `t1_confection/Z_AUX_united_regions.py`

A specialized, manually-configured script for consolidating Brazilian sub-regions (CN, NW, NE, CW, SO, SE, WE) into a unified XX region, dating from the model's earlier LATAM phase.

:::{note}
This script is genuinely dead code in the current South/Southeast Asia model -- it is not imported or invoked anywhere in the active pipeline, and it is entirely hardcoded to Brazil (`"BRA"`, `brazil_regions = [...]`, `"BRACN"` → `"BRAXX"` string substitutions). For general region consolidation (e.g. merging India's 5 regions), use the configurable mechanism in `Config_region_consolidation.yaml` instead -- see {doc}`country-management`.
:::

### Usage

The script uses boolean flags at the top of the file to control which files to process:

```python
parametrization = False  # Process A-O_Parametrization.xlsx
demand = True            # Process A-O_Demand.xlsx
storage = False          # Process A-Xtra_Storage.xlsx
```

### Processing Rules

- **Cost parameters** (CapitalCost, FixedCost, VariableCost): Averaged across regions.
- **Capacity parameters** (ResidualCapacity, TotalAnnualMaxCapacity): Summed across regions.
- **Interconnections**: TRN codes are normalized to alphabetical country-pair ordering.

---

## Capital Annualization

**Script:** `t1_confection/Z_AUX_capital_annualization_script.py`

Post-processing script that converts lump-sum capital investment into an annualized cost stream in the final combined results. Runs automatically as part of the B2 execution stage when `annualize_capital: True` in `Config_MOMF_T1_AB.yaml`.

### What It Does

Reads the `CapitalInvestment` column from `OSTRAM_Combined_Inputs_Outputs.csv`, computes a Capital Recovery Factor `CRF = r(1+r)^n / ((1+r)^n - 1)` (default discount rate `0.0639`, asset lifetime `15` years), then for each positive investment year distributes `investment × CRF` as an annual payment across `year .. year + lifetime - 1`, accumulating overlapping investment cohorts into a new `CapitalInvestmentAnnualized` column. Grouping columns (Future/Scenario/REGION/TECHNOLOGY/...) are auto-narrowed to whichever actually vary. The result is written back over the same input file -- there is no backup for this step.

Can also be called programmatically:

```python
from Z_AUX_capital_annualization_script import annualize_capital_investment

annualize_capital_investment(
    input_file_path="OSTRAM_Combined_Inputs_Outputs.csv",
    discount_rate=0.0639,
    asset_lifetime=15,
)
```

---

## RES (Reference Energy System) Diagram

**Script:** `t1_confection/Z_AUX_generate_RES_diagram.py`

Reads `Config_country_codes.yaml` for country names and the `BAU` scenario's `A-O_AR_Model_Base_Year.xlsx` (sheets `Primary`, `Secondary`, `Demand Techs`, `Transport Groups`) to build a fuel → technology → fuel flow map. Regions are discovered dynamically from `ELC[A-Z]{5}\d{2}` fuel codes -- there is no hardcoded country list, so it works automatically as countries/regions are added or removed.

### Usage

```bash
python t1_confection/Z_AUX_generate_RES_diagram.py
```

### What It Produces

A standalone, self-contained Sankey diagram at `Figures/RES_Diagram.html`, with a multi-region checkbox filter, pan/zoom, PNG export, and green/brown/blue color-coding for renewable/fossil/nuclear technologies. Node labels show technology/fuel codes with full names on hover.

---

## Transmission Maps and Dispatch Chart

**Script:** `t1_confection/Z_AUX_generate_transmission_maps.py`

Reads `OSTRAM_Combined_Inputs_Outputs.csv` and `Miscellaneous/centerpoints.csv` (the geographic centerpoints generated for each country/region, including any added via `Z_generate_country_template.py`). Identifies cross-border interconnectors with the pattern `^TRN[A-Z]{5}[A-Z]{5}$`.

### Usage

```bash
python t1_confection/Z_AUX_generate_transmission_maps.py
```

### What It Produces

Two standalone HTML files in `Figures/`:

- **`TransmissionMaps.html`**: three tabs (Capacity in GW/TW, directional Flow in PJ/TWh/GWh with arrowheads, Load-Capacity Ratio with a green→red color scale), year/scenario selectors, a Scattergeo world map, and PNG export.
- **`DispatchChart.html`**: a stacked-area generation-mix-by-timeslice chart per scenario/year/region, with a PJ/GWh/MWh unit toggle. Curtailment is inferred as generation minus demand on the `ELC*02` bus.

---

## Interconnections Dashboard

**Script:** `t1_confection/Z_AUX_interconnections_dashboard.py`

Reads `OSTRAM_Combined_Inputs_Outputs.csv`. Hardcodes the current scenario order (`BAU`, `A_Calibrated_BAU`, `B_Optimised_VRE`, `C_Target_VRE`) and the current 10-node country/region list -- update both if scenarios or countries change. Excludes `TRNNLI*` (unplanned candidate lines) from the flow analysis, showing only existing/committed interconnections.

### Usage

```bash
python t1_confection/Z_AUX_interconnections_dashboard.py
```

### What It Produces

A standalone `Figures/interconnections_dashboard.html` with KPI cards, an annual trend chart, a per-line heatmap, a year-scoped Sankey diagram with a scenario dropdown, delta-vs-BAU bars, net-flow-per-node bars, a capacity trend chart, and a seasonal heatmap. Includes a PJ/TWh unit toggle (1 PJ = 1/3.6 TWh).

---

## Interconnection Limits From Flows

**Script:** `t1_confection/Z_AUX_D1b_set_trn_limits_from_flows.py`

Pre-fills TRN interconnection activity limits in `Secondary_Techs_Editor.xlsx` from historical bilateral flow data. This is part of the D1/D2 manual editing workflow rather than a standalone analysis tool -- see {doc}`secondary-techs-editor` for details.
