# Country Management

OSTRAM includes tools for validating country data, generating templates for new countries, and consolidating sub-regional data.

## Technology-Country Matrix

The Technology-Country Matrix controls which technology-country combinations are active in the model. See {doc}`pipeline` (Stage A0) for generation instructions.

### Editing the Matrix

After generating `Tech_Country_Matrix.xlsx`:

1. Open the **Matrix** sheet.
2. Change any cell from `YES` to `NO` to disable a technology for a country.
3. Change from `NO` to `YES` to enable it.
4. Implausible combinations (highlighted in red, from `implausible_combinations` in `Config_country_codes.yaml`) can be overridden if desired.

### NGS Unification

In the **NGS_Unification** sheet:

- Set a country to `YES` to merge CCG (Combined Cycle Gas) and OCG (Open Cycle Gas) into a single NGS (Natural Gas) technology.
- Set to `NO` to keep them separate.
- Aggregation follows the rules defined in the **Aggregation_Rules** sheet.

---

## Country Data Validator

**Script:** `t1_confection/Z_validate_country_data.py`

Verifies that a country has complete and consistent data in the OSeMOSYS input CSV files. The country list is driven entirely by `Config_country_codes.yaml` (via `Z_AUX_config_loader.get_countries()`) -- it is not hardcoded to any particular region.

### Usage

```bash
# Validate all countries
python t1_confection/Z_validate_country_data.py

# Validate a specific country or region
python t1_confection/Z_validate_country_data.py --country LKA
python t1_confection/Z_validate_country_data.py --country INDEA

# Quiet mode (summary only)
python t1_confection/Z_validate_country_data.py --quiet
```

### Command-Line Options

| Flag | Description |
|------|-------------|
| `--country`, `-c` | ISO-3 (or 5-char region) code, or `all` (default: all) |
| `--report`, `-r` | Accepted by the CLI but currently a no-op -- not wired to any report-generation code |
| `--quiet`, `-q` | Suppress verbose per-check output |

### Validations Performed

The validator checks:

1. **SET membership**: Country appears in TECHNOLOGY, FUEL, EMISSION, and STORAGE sets.
2. **Technology type counts**: Minimum expected counts per prefix (`PWR` ≥ 15, `MIN` ≥ 5, `RNW` ≥ 7), plus warnings for missing backstop/`PWRTRN` entries.
3. **Required parameters**: Data exists for all required OSeMOSYS parameters (CapitalCost, FixedCost, VariableCost, ResidualCapacity, OperationalLife, CapacityToActivityUnit, etc.) -- 22 parameters checked.
4. **Value ranges**: Parameter values fall within physically reasonable ranges (12 parameters checked).
5. **Demand profiles**: SpecifiedDemandProfile sums to approximately 1.0 per fuel/tech per year.
6. **Storage**: Storage technologies (LDS + SDS) have matching parameters and link presence.
7. **Referential integrity**: Technologies referenced in parameter files exist in the TECHNOLOGY set, and fuels exist in the FUEL set (flags orphan codes).

:::{note}
The fuel-completeness check hardcodes the `XX` region suffix (e.g. `ELC{country}XX01`/`02`) when building its expected-fuel patterns. For single-region countries (BGD, BTN, NPL, LKA, MDV) this is correct; for India's 5-char region codes (`INDEA`, `INDNE`, ...) this under-checks unless you invoke the validator per full region code rather than the bare `IND` code.
:::

### Output Format

Results are displayed as:

- **PASS**: Check succeeded.
- **FAIL**: Critical issue that will cause model errors.
- **WARN**: Potential issue that should be reviewed.

---

## New Country Template Generator

**Script:** `t1_confection/Z_generate_country_template.py`

Creates a complete set of CSV files with the minimum structure needed to add a new country (or a new region of an existing multi-region country) to the model, using an existing country as a reference.

### Usage

```bash
# Read all entries from Config_country_codes.yaml's template_generation list
python t1_confection/Z_generate_country_template.py

# Override via command line (single country)
python t1_confection/Z_generate_country_template.py --new MDV --ref LKA --region XX -i LKA

# A country with no interconnections at all
python t1_confection/Z_generate_country_template.py --new AUS --ref LKA -i
```

### Command-Line Options

| Flag | Description |
|------|-------------|
| `--new`, `-n` | 3-letter ISO code for the new country |
| `--ref`, `-r` | 3-letter ISO code of the reference country to clone (defaults to `ARG`, a legacy fallback from the model's earlier LATAM phase -- always pass this explicitly) |
| `--output`, `-o` | Output directory path (default: `templates/<new_code>`) |
| `--interconnections`, `-i` | List of neighbor country codes for TRN links (space-separated; pass with no values for zero interconnections) |
| `--region` | Region suffix (default: `XX`) |
| `--lat` | Centerpoint latitude |
| `--lon` | Centerpoint longitude |

### YAML Configuration

Alternatively (and preferably, since it supports multiple countries per run), configure in `Config_country_codes.yaml`'s `template_generation` list -- see {doc}`configuration` for the full anatomy and examples, including the multi-region pattern used for countries like India.

```yaml
template_generation:
  - new_country: MDV
    reference_country: LKA
    region: XX
    centerpoint_lat: 1.924992
    centerpoint_lon: 73.399658
    interconnections:
      - LKA
```

This is the currently active entry in the live configuration (Maldives, cloned from Sri Lanka, one interconnection).

### What It Generates

The script creates a `templates/{CODE}/` directory containing:

1. **SET CSVs**: TECHNOLOGY, FUEL, EMISSION, STORAGE entries for the new country.
2. **Parameter CSVs**: All OSeMOSYS parameter files with the new country's data (cloned and adapted from the reference).
3. **`merge_into_inputs.py`**: A generated helper script (with hardcoded absolute paths for the machine it was generated on) that merges the new country's files into `OG_csvs_inputs/` -- by concatenation, not upsert, so re-running it after a first successful merge will duplicate rows. It also merges the new centerpoint into `Miscellaneous/centerpoints.csv` and prints the suggested validation command.

### Interconnection Handling

| YAML/CLI value | Behavior |
|----------------|----------|
| Key omitted entirely | **Legacy mode**: blindly copies the reference country's TRN rows verbatim, with no remapping. Only correct if the new country's topology should be identical to the reference's. |
| `interconnections: []` (empty list) | No interconnections for the new country. |
| `interconnections: [LKA, IND]` | Creates TRN links to the specified neighbors, remapping country/region codes. |
| Fewer/more neighbors than the reference | Dynamically adjusts TRN entries: matches by reference neighbor where possible, reuses/discards the remainder. |

### TRN Code Structure

Transmission technology codes follow this pattern:

```
TRN{ORIGIN 5-char}{DEST 5-char}
```

Each 5-char block is a 3-letter country code + 2-letter region code. For example: `TRNBGDXXINDEA` = transmission between Bangladesh (region XX) and India-East (region EA).

The generator correctly handles:
- Position-aware country/region code replacement.
- Alphabetical ordering of country pairs in codes (and a `MODE_OF_OPERATION` swap when a country's position flips between origin/destination).
- Fuel and mode-of-operation code transformations.

### Integration Workflow

After generating the template:

```bash
# 1. Generate the template
python t1_confection/Z_generate_country_template.py

# 2. Review and customize the generated CSVs
# Edit files in templates/MDV/ as needed

# 3. Merge into the main dataset
cd templates/MDV/
python merge_into_inputs.py

# 4. Validate the new country's data
python t1_confection/Z_validate_country_data.py --country MDV --report
```

---

## Region Consolidation

Region consolidation merges multiple sub-regional datasets into a single unified region. This is useful when a country is modeled with geographic granularity (like India's current 5 regions) but you want aggregated results.

### Configuration

Edit `Config_region_consolidation.yaml`. Currently `enabled: false` with an empty `countries:` block -- India's 5 regions are kept separate in the active configuration. To consolidate them:

```yaml
enabled: true

countries:
  IND:
    regions: ["EA", "NE", "NO", "SO", "WE"]
    unified_region: "XX"
```

### Aggregation Rules

When consolidating regions, parameters are combined according to these rules:

**Averaged** (rate-like parameters): `AvailabilityFactor`, `CapacityFactor`, `CapitalCost`, `FixedCost`, `VariableCost`, `InputActivityRatio`, `OutputActivityRatio`, `CapacityToActivityUnit`, `OperationalLife`, `SpecifiedDemandProfile`, `EmissionActivityRatio`, `ReserveMarginTagFuel`/`Technology`, storage rate parameters, and others.

**Summed** (quantity-like parameters): `ResidualCapacity`, `SpecifiedAnnualDemand`, `TotalAnnualMaxCapacity`, `TotalAnnualMaxCapacityInvestment`, `TotalAnnualMinCapacityInvestment`, `StorageLevelStart`, `ResidualStorageCapacity`, activity limits.

**Disabled** (skipped): `CapacityOfOneTechnologyUnit`, `RETagTechnology`, `TotalAnnualMinCapacity`, model-period activity limits.

See {doc}`configuration` for the complete current parameter lists.

### Processing

When enabled, region consolidation runs as part of Stage A1 preprocessing. It:

1. Groups data by country (across sub-regions).
2. Applies averaging or summing per the rules.
3. Replaces region codes with the unified region code.
4. Removes internal interconnections that become self-loops after merging.

:::{note}
The earlier hardcoded Brazil-specific consolidation helper is retired. It is
unrelated to the generic maintained mechanism described above.
:::
