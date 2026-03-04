# Auxiliary Tools

OSTRAM includes several utility scripts (prefixed with `Z_`) for data maintenance, visualization, and support tasks.

## Configuration Loader

**Script:** `t1_confection/Z_AUX_config_loader.py`

A centralized module (not run directly) that provides cached access to `Config_country_codes.yaml`. All other scripts import functions from this module instead of reading the YAML directly.

### Available Functions

| Function | Returns | Description |
|----------|---------|-------------|
| `get_countries()` | `list[str]` | Sorted list of active country codes |
| `get_country_names()` | `dict[str, str]` | `{iso3: english_name}` |
| `get_iso_country_map()` | `dict[str, str]` | `{iso3: english_name}` including special entries |
| `get_olade_country_mapping()` | `dict[str, str]` | `{olade_name: iso3}` |
| `get_olade_country_mapping_normalized()` | `dict[str, str]` | Accent-stripped version of the above |
| `get_first_year()` | `int` | Model start year (default: 2023) |
| `get_pwr_cleanup_mode()` | `str \| bool` | `"drop"`, `"merge"`, or `False` |
| `get_code_to_energy()` | `dict[str, str]` | `{tech_code: description}` |
| `get_renewable_fuels()` | `set[str]` | Set of renewable fuel codes |
| `get_add_missing_countries_from_olade()` | `bool` | Whether to fill from OLADE |
| `get_olade_tech_mapping()` | `dict[str, str]` | OLADE name to model code mapping |
| `get_shares_tech_mapping()` | `dict[str, str]` | Shares file name to model code mapping |
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

1. Iterates over all scenario directories (`A1_Outputs_BAU`, `A1_Outputs_NDC`, etc.).
2. Opens each `A-O_Demand.xlsx` file.
3. For each profile sheet, normalizes values so that each fuel/technology column sums to exactly 1.0 per year.
4. Creates a **timestamped backup** before modifying any file.

### Tolerance

Values within `0.0001` of 1.0 are considered acceptable. Values outside this range are corrected by proportional scaling.

---

## Interactive Dashboard Generator

**Script:** `t1_confection/Z_AUX_generate_interactive_dashboards_aggregated.py`

Generates standalone HTML dashboards with embedded Plotly.js charts for analyzing power (PWR) technology results.

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
^PWR(BIO|WAS|CSP|GEO|HYD|SPV|WON|WOF|COA|GAS|OIL|PET|URN|...)([A-Z]{3})XX$
```

Only technologies matching this pattern are included in the visualizations.

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

## Region Consolidation (Brazil)

**Script:** `t1_confection/Z_AUX_united_regions.py`

A specialized, manually-configured script for consolidating Brazilian sub-regions (CN, NW, NE, CW, SO, SE, WE) into a unified XX region.

:::{note}
This is a legacy script with hardcoded Brazil-specific logic. For general region consolidation, use the configurable system in `Config_region_consolidation.yaml` (see {doc}`country-management`).
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

Post-processing script that annualizes capital costs in the model results. Runs automatically as part of the B2 execution stage when `annualize_capital: True` in `Config_MOMF_T1_AB.yaml`.

### What It Does

Calculates `AccumulatedTotalAnnualMinCapacityInvestment` via cumulative sum of annual capacity investments, converting lump-sum capital costs into annualized values.
