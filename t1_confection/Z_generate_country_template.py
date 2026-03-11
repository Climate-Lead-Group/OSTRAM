"""
Template generator for adding a new country to OG_csvs_inputs.

Creates a set of CSV files with the minimum required data structure
for a new country, using an existing country as a reference template.

Usage:
    python Z_generate_country_template.py                  # reads from Config_country_codes.yaml
    python Z_generate_country_template.py --new NCC --ref ARG -i BOL PRY   # CLI overrides

Configure the 'template_generation' section in Config_country_codes.yaml,
then just run: python Z_generate_country_template.py

This script does NOT modify the original files in OG_csvs_inputs.
It creates new CSV files in the output directory that can be manually
reviewed and then merged into the input files.

Author: Climate Lead Group, Andrey Salazar-Vargas
"""

import pandas as pd
import yaml
import os
import sys
import argparse
from collections import defaultdict

# ============================================================================
# Configuration
# ============================================================================

SCRIPT_DIR = os.path.dirname(__file__)
INPUT_DIR = os.path.join(SCRIPT_DIR, "OG_csvs_inputs")
CONFIG_PATH = os.path.join(SCRIPT_DIR, "Config_country_codes.yaml")

# Sets files to modify
SETS_FILES = ["TECHNOLOGY", "FUEL", "EMISSION", "STORAGE"]

# Parameter files that need country-specific data
PARAM_FILES_WITH_TECH = [
    "CapitalCost", "FixedCost", "VariableCost",
    "ResidualCapacity", "CapacityFactor", "AvailabilityFactor",
    "InputActivityRatio", "OutputActivityRatio", "EmissionActivityRatio",
    "OperationalLife", "CapacityToActivityUnit",
    "TotalAnnualMaxCapacity", "TotalAnnualMaxCapacityInvestment",
    "TotalTechnologyAnnualActivityUpperLimit",
    "ReserveMarginTagTechnology",
]

PARAM_FILES_WITH_FUEL = [
    "SpecifiedAnnualDemand", "SpecifiedDemandProfile",
    "ReserveMarginTagFuel",
]

PARAM_FILES_WITH_STORAGE = [
    "CapitalCostStorage", "OperationalLifeStorage",
    "StorageLevelStart", "ResidualStorageCapacity",
    "TechnologyToStorage", "TechnologyFromStorage",
]


# ============================================================================
# Helper functions
# ============================================================================

def load_csv(filename):
    """Load a CSV file from the input directory."""
    filepath = os.path.join(INPUT_DIR, filename + ".csv")
    if not os.path.exists(filepath):
        return None
    return pd.read_csv(filepath)


def replace_country_in_string(value, old_cc, new_cc):
    """Replace country code in a string value."""
    return str(value).replace(old_cc, new_cc)


def get_country_rows(df, country_code, search_cols=None):
    """Get all rows containing the country code."""
    if search_cols is None:
        search_cols = [c for c in df.columns
                       if c in ["TECHNOLOGY", "FUEL", "STORAGE", "EMISSION"]]
    if not search_cols:
        return pd.DataFrame(columns=df.columns)

    mask = pd.Series(False, index=df.index)
    for col in search_cols:
        mask |= df[col].astype(str).str.contains(country_code, na=False)
    return df[mask].copy()


# ============================================================================
# TRN interconnection helpers
# ============================================================================

def parse_trn_tech(tech_name):
    """Parse a TRN interconnection technology name into components.

    Args:
        tech_name: e.g. 'TRNARGXXBOLXX'

    Returns:
        tuple (cc1, rr1, cc2, rr2) or None if not a valid TRN interconnection.
        PWRTRN technologies (intra-country) return None.
    """
    s = str(tech_name)
    if not s.startswith("TRN") or s.startswith("PWRTRN") or len(s) != 13:
        return None
    body = s[3:]  # e.g. 'ARGXXBOLXX'
    return (body[0:3], body[3:5], body[5:8], body[8:10])


def build_trn_tech(cc1, rr1, cc2, rr2):
    """Build a TRN technology name ensuring alphabetical ordering.

    Returns:
        tuple (tech_name, is_swapped) where is_swapped is True if
        cc1+rr1 was moved to the second position.
    """
    key1 = cc1 + rr1
    key2 = cc2 + rr2
    if key1 <= key2:
        return f"TRN{key1}{key2}", False
    else:
        return f"TRN{key2}{key1}", True


def extract_ref_interconnections(ref_cc):
    """Extract all TRN interconnections involving the reference country.

    Returns:
        list of (tech_name, neighbor_cc, neighbor_rr, ref_position)
        where ref_position is 1 if ref_cc is Country1, 2 if Country2.
    """
    tech_df = load_csv("TECHNOLOGY")
    if tech_df is None:
        return []

    interconnections = []
    for tech in tech_df["VALUE"].astype(str):
        parsed = parse_trn_tech(tech)
        if parsed is None:
            continue
        cc1, rr1, cc2, rr2 = parsed
        if cc1 == ref_cc:
            interconnections.append((tech, cc2, rr2, 1))
        elif cc2 == ref_cc:
            interconnections.append((tech, cc1, rr1, 2))

    return interconnections


def parse_interconnection_args(args_list):
    """Parse interconnection CLI arguments into (country, region) tuples.

    Accepts 3-letter codes (assumes region XX) or 5-letter codes.
    Examples: 'BOL' -> ('BOL','XX'), 'BRASO' -> ('BRA','SO')

    Returns:
        list of (neighbor_cc, neighbor_rr) tuples.
    """
    result = []
    for item in args_list:
        item = item.strip().upper()
        if len(item) == 5:
            result.append((item[:3], item[3:]))
        elif len(item) == 3:
            result.append((item, "XX"))
        else:
            print(f"  WARNING: Cannot parse interconnection '{item}', "
                  f"expected 3 or 5 chars. Skipping.")
    return result


def build_interconnection_mapping(ref_cc, ref_rr, new_cc, new_rr,
                                  ref_interconnections, new_neighbors):
    """Build a mapping from reference interconnections to new ones.

    For each new neighbor:
      1. If the neighbor also exists in the reference, use that interconnection.
      2. Otherwise use the first unused reference interconnection.
      3. If all refs are used, reuse the first one.

    Reference interconnections without a match are excluded (fewer case).

    Returns:
        list of dicts with keys: ref_tech, new_tech, ref_neighbor,
        new_neighbor, ref_position, new_position.
    """
    mapping = []
    used_refs = set()

    for nbr_cc, nbr_rr in new_neighbors:
        new_tech, swapped = build_trn_tech(new_cc, new_rr, nbr_cc, nbr_rr)
        new_position = 2 if swapped else 1

        # Try exact neighbor match first
        template = None
        for ref_tech, ref_nbr_cc, ref_nbr_rr, ref_pos in ref_interconnections:
            if ref_nbr_cc == nbr_cc and ref_nbr_rr == nbr_rr:
                template = (ref_tech, ref_nbr_cc, ref_nbr_rr, ref_pos)
                used_refs.add(ref_tech)
                break

        # Fallback: first unused reference interconnection
        if template is None:
            for ref_tech, ref_nbr_cc, ref_nbr_rr, ref_pos in ref_interconnections:
                if ref_tech not in used_refs:
                    template = (ref_tech, ref_nbr_cc, ref_nbr_rr, ref_pos)
                    used_refs.add(ref_tech)
                    break

        # Last resort: reuse the first reference interconnection
        if template is None and ref_interconnections:
            first = ref_interconnections[0]
            template = first

        if template is None:
            print(f"  WARNING: No reference interconnection available as "
                  f"template for {new_cc}{new_rr}<->{nbr_cc}{nbr_rr}")
            continue

        mapping.append({
            "ref_tech": template[0],
            "new_tech": new_tech,
            "ref_neighbor": (template[1], template[2]),
            "new_neighbor": (nbr_cc, nbr_rr),
            "ref_position": template[3],
            "new_position": new_position,
        })

    return mapping


def transform_trn_row(row, mapping_entry, ref_cc, ref_rr, new_cc, new_rr):
    """Transform a row from a reference TRN interconnection to a new one.

    Handles TECHNOLOGY replacement, FUEL code replacement for both sides,
    and MODE_OF_OPERATION swap when the new country changes position.
    """
    new_row = row.copy()
    ref_nbr_cc, ref_nbr_rr = mapping_entry["ref_neighbor"]
    new_nbr_cc, new_nbr_rr = mapping_entry["new_neighbor"]
    need_mode_swap = mapping_entry["ref_position"] != mapping_entry["new_position"]

    # Replace TECHNOLOGY column
    if "TECHNOLOGY" in new_row.index:
        new_row["TECHNOLOGY"] = mapping_entry["new_tech"]

    # Swap MODE_OF_OPERATION if position changed
    if need_mode_swap and "MODE_OF_OPERATION" in new_row.index:
        mode = new_row["MODE_OF_OPERATION"]
        new_row["MODE_OF_OPERATION"] = 2 if mode == 1 else 1

    # Replace FUEL codes for both sides of the interconnection
    if "FUEL" in new_row.index:
        fuel = str(new_row["FUEL"])
        fuel = fuel.replace(ref_cc + ref_rr, new_cc + new_rr)
        fuel = fuel.replace(ref_nbr_cc + ref_nbr_rr, new_nbr_cc + new_nbr_rr)
        new_row["FUEL"] = fuel

    return new_row


# ============================================================================
# Template generation
# ============================================================================

def generate_sets_template(ref_cc, new_cc, interconnection_mapping=None):
    """Generate new set entries by cloning from reference country.

    When interconnection_mapping is provided, TRN interconnection entries
    are replaced by the mapped set instead of blindly cloned.
    """
    templates = {}

    for set_name in SETS_FILES:
        df = load_csv(set_name)
        if df is None:
            continue

        existing = set(df["VALUE"].astype(str))
        ref_entries = [v for v in existing if ref_cc in v]

        new_entries = []
        for entry in ref_entries:
            # If mapping is active, skip TRN interconnection entries
            if interconnection_mapping is not None and parse_trn_tech(entry):
                continue

            new_entry = replace_country_in_string(entry, ref_cc, new_cc)
            if new_entry not in existing:
                new_entries.append(new_entry)

        # Add mapped TRN interconnection entries for TECHNOLOGY set
        if set_name == "TECHNOLOGY" and interconnection_mapping is not None:
            for m in interconnection_mapping:
                if m["new_tech"] not in existing:
                    new_entries.append(m["new_tech"])

        if new_entries:
            templates[set_name] = pd.DataFrame({"VALUE": new_entries})

    return templates


def generate_param_templates(ref_cc, new_cc, interconnection_mapping=None,
                             ref_rr="XX", new_rr="XX"):
    """Generate parameter templates by cloning from reference country.

    When interconnection_mapping is provided, TRN interconnection rows are
    transformed using the mapping (correct fuel codes + mode swaps) instead
    of simple string replacement.
    """
    templates = {}

    # Parameters with TECHNOLOGY column
    for param in PARAM_FILES_WITH_TECH:
        df = load_csv(param)
        if df is None or len(df) == 0:
            continue

        ref_rows = get_country_rows(df, ref_cc, ["TECHNOLOGY"])
        if len(ref_rows) == 0:
            continue

        if interconnection_mapping is not None:
            # Separate TRN interconnection rows from normal rows
            trn_mask = ref_rows["TECHNOLOGY"].apply(
                lambda t: parse_trn_tech(str(t)) is not None
            )
            non_trn = ref_rows[~trn_mask].copy()
            trn_rows = ref_rows[trn_mask]

            # Normal rows: simple replacement
            for col in non_trn.columns:
                if non_trn[col].dtype == object:
                    non_trn[col] = non_trn[col].apply(
                        lambda x: replace_country_in_string(x, ref_cc, new_cc)
                    )

            # TRN rows: use mapping
            mapped_rows = []
            for m in interconnection_mapping:
                tech_rows = trn_rows[trn_rows["TECHNOLOGY"] == m["ref_tech"]]
                for _, row in tech_rows.iterrows():
                    mapped_rows.append(transform_trn_row(
                        row, m, ref_cc, ref_rr, new_cc, new_rr
                    ))

            if mapped_rows:
                mapped_df = pd.DataFrame(mapped_rows)
                new_rows = pd.concat([non_trn, mapped_df], ignore_index=True)
            else:
                new_rows = non_trn
        else:
            # Legacy: simple replacement for all rows
            new_rows = ref_rows.copy()
            for col in new_rows.columns:
                if new_rows[col].dtype == object:
                    new_rows[col] = new_rows[col].apply(
                        lambda x: replace_country_in_string(x, ref_cc, new_cc)
                    )

        if len(new_rows) > 0:
            templates[param] = new_rows

    # Parameters with FUEL column (no TRN handling needed)
    for param in PARAM_FILES_WITH_FUEL:
        df = load_csv(param)
        if df is None or len(df) == 0:
            continue

        ref_rows = get_country_rows(df, ref_cc, ["FUEL"])
        if len(ref_rows) == 0:
            continue

        new_rows = ref_rows.copy()
        for col in new_rows.columns:
            if new_rows[col].dtype == object:
                new_rows[col] = new_rows[col].apply(
                    lambda x: replace_country_in_string(x, ref_cc, new_cc)
                )

        templates[param] = new_rows

    # Parameters with STORAGE column (no TRN handling needed)
    for param in PARAM_FILES_WITH_STORAGE:
        df = load_csv(param)
        if df is None or len(df) == 0:
            continue

        search_cols = ["STORAGE"]
        if "TECHNOLOGY" in df.columns:
            search_cols.append("TECHNOLOGY")

        ref_rows = get_country_rows(df, ref_cc, search_cols)
        if len(ref_rows) == 0:
            continue

        new_rows = ref_rows.copy()
        for col in new_rows.columns:
            if new_rows[col].dtype == object:
                new_rows[col] = new_rows[col].apply(
                    lambda x: replace_country_in_string(x, ref_cc, new_cc)
                )

        templates[param] = new_rows

    return templates


def write_templates(templates, output_dir, mode="new_only"):
    """Write template DataFrames to CSV files.

    Args:
        templates: dict of {filename: DataFrame}
        output_dir: directory to write files to
        mode: "new_only" writes only new rows, "full" writes complete files
    """
    os.makedirs(output_dir, exist_ok=True)

    for filename, df in templates.items():
        filepath = os.path.join(output_dir, filename + ".csv")
        df.to_csv(filepath, index=False)
        print(f"  Written: {filename}.csv ({len(df)} rows)")


def generate_centerpoint_template(new_cc, new_rr, lat, lon, output_dir):
    """Generate a centerpoint CSV for the new country.

    Args:
        new_cc: 3-letter country code
        new_rr: 2-letter region code
        lat: latitude
        lon: longitude
        output_dir: directory to write the file to
    """
    region_code = f"{new_cc}{new_rr}"
    df = pd.DataFrame({"region": [region_code], "lat": [lat], "long": [lon]})
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, "centerpoint.csv")
    df.to_csv(filepath, index=False)
    print(f"  Written: centerpoint.csv ({region_code}: {lat}, {lon})")
    return df


def generate_merge_script(new_cc, output_dir):
    """Generate a helper script to merge templates into the main files."""
    script_path = os.path.join(output_dir, "merge_into_inputs.py")

    # Use forward slashes to avoid Unicode escape issues on Windows
    template_dir_str = os.path.abspath(output_dir).replace("\\", "/")
    input_dir_str = os.path.abspath(INPUT_DIR).replace("\\", "/")
    centerpoints_str = os.path.abspath(
        os.path.join(SCRIPT_DIR, "Miscellaneous", "centerpoints.csv")
    ).replace("\\", "/")
    rel_path_str = os.path.relpath(
        script_path, os.path.dirname(__file__)
    ).replace("\\", "/")

    lines = [
        '"""',
        f'Auto-generated script to merge {new_cc} template data into OG_csvs_inputs.',
        '',
        'IMPORTANT: Review the template CSV files BEFORE running this script.',
        f'Modify values as needed to reflect the actual data for {new_cc}.',
        '',
        'Usage:',
        '    cd t1_confection',
        f'    python {rel_path_str}',
        '"""',
        '',
        'import pandas as pd',
        'import os',
        'import shutil',
        'from datetime import datetime',
        '',
        f'TEMPLATE_DIR = "{template_dir_str}"',
        f'INPUT_DIR = "{input_dir_str}"',
        f'CENTERPOINTS_PATH = "{centerpoints_str}"',
        '',
        'backup_suffix = datetime.now().strftime("%Y%m%d_%H%M%S")',
        '',
        '',
        'def merge_file(filename):',
        '    """Merge a template file into the corresponding input file."""',
        '    template_path = os.path.join(TEMPLATE_DIR, filename + ".csv")',
        '    input_path = os.path.join(INPUT_DIR, filename + ".csv")',
        '',
        '    if not os.path.exists(template_path):',
        '        return',
        '',
        '    template_df = pd.read_csv(template_path)',
        '    if len(template_df) == 0:',
        '        return',
        '',
        '    if os.path.exists(input_path):',
        '        input_df = pd.read_csv(input_path)',
        '        backup_path = input_path.replace(".csv", f"_backup_{backup_suffix}.csv")',
        '        shutil.copy2(input_path, backup_path)',
        '        merged = pd.concat([input_df, template_df], ignore_index=True)',
        '        merged.to_csv(input_path, index=False)',
        '        print(f"  Merged {filename}.csv: {len(input_df)} + {len(template_df)} = {len(merged)} rows")',
        '    else:',
        '        template_df.to_csv(input_path, index=False)',
        '        print(f"  Created {filename}.csv: {len(template_df)} rows")',
        '',
        '',
        '# Sets',
        'for s in ["TECHNOLOGY", "FUEL", "EMISSION", "STORAGE"]:',
        '    merge_file(s)',
        '',
        '# Parameters',
        'params = [',
        '    "CapitalCost", "FixedCost", "VariableCost",',
        '    "ResidualCapacity", "CapacityFactor", "AvailabilityFactor",',
        '    "InputActivityRatio", "OutputActivityRatio", "EmissionActivityRatio",',
        '    "SpecifiedAnnualDemand", "SpecifiedDemandProfile",',
        '    "OperationalLife", "CapacityToActivityUnit",',
        '    "TotalAnnualMaxCapacity", "TotalAnnualMaxCapacityInvestment",',
        '    "TotalTechnologyAnnualActivityUpperLimit",',
        '    "ReserveMarginTagTechnology", "ReserveMarginTagFuel",',
        '    "CapitalCostStorage", "OperationalLifeStorage",',
        '    "StorageLevelStart", "ResidualStorageCapacity",',
        '    "TechnologyToStorage", "TechnologyFromStorage",',
        ']',
        '',
        'for p in params:',
        '    merge_file(p)',
        '',
        '# Centerpoint',
        'centerpoint_template = os.path.join(TEMPLATE_DIR, "centerpoint.csv")',
        'if os.path.exists(centerpoint_template):',
        '    cp_new = pd.read_csv(centerpoint_template)',
        '    if len(cp_new) > 0:',
        '        if os.path.exists(CENTERPOINTS_PATH):',
        '            cp_existing = pd.read_csv(CENTERPOINTS_PATH)',
        '            backup_path = CENTERPOINTS_PATH.replace(".csv", f"_backup_{backup_suffix}.csv")',
        '            shutil.copy2(CENTERPOINTS_PATH, backup_path)',
        '            # Remove existing entry for same region to avoid duplicates',
        '            new_regions = set(cp_new["region"].astype(str))',
        '            cp_existing = cp_existing[~cp_existing["region"].astype(str).isin(new_regions)]',
        '            merged = pd.concat([cp_existing, cp_new], ignore_index=True)',
        '            merged = merged.sort_values("region").reset_index(drop=True)',
        '            merged.to_csv(CENTERPOINTS_PATH, index=False)',
        '            print(f"  Merged centerpoint.csv into {CENTERPOINTS_PATH}")',
        '        else:',
        '            cp_new.to_csv(CENTERPOINTS_PATH, index=False)',
        '            print(f"  Created {CENTERPOINTS_PATH}")',
        '',
        'print("\\nDone! Backup files created with suffix: " + backup_suffix)',
        f'print("Run the validation script to verify: python Z_validate_country_data.py --country {new_cc}")',
    ]

    with open(script_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"  Written: merge_into_inputs.py (merge helper script)")


# ============================================================================
# Main
# ============================================================================

def load_yaml_config():
    """Load template_generation config from Config_country_codes.yaml.

    Returns a list of entry dicts. Supports both the legacy single-dict
    format and the new list format.
    """
    if not os.path.exists(CONFIG_PATH):
        return None
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    tg = cfg.get("template_generation")
    if tg is None:
        return None
    # Normalise to list: legacy single-dict → list of one
    if isinstance(tg, dict):
        return [tg]
    if isinstance(tg, list):
        return tg
    return None


def process_entry(new_cc, ref_cc, new_rr, ixn_raw, ixn_source,
                  cp_lat, cp_lon, output_dir):
    """Process a single country/region template entry."""

    if not os.path.isdir(INPUT_DIR):
        print(f"ERROR: Input directory not found: {INPUT_DIR}")
        return

    if len(new_cc) != 3:
        print(f"ERROR: Country code must be exactly 3 letters, got '{new_cc}'")
        return

    print(f"\nGenerating template for {new_cc}{new_rr} based on {ref_cc}")
    print(f"Output directory: {output_dir}")

    # Build interconnection mapping
    interconnection_mapping = None
    ref_rr = "XX"

    if ixn_raw is not None:
        new_neighbors = parse_interconnection_args(ixn_raw)
        ref_interconnections = extract_ref_interconnections(ref_cc)

        print(f"\n--- INTERCONNECTION MAPPING (source: {ixn_source}) ---")
        print(f"Reference {ref_cc} interconnections "
              f"({len(ref_interconnections)}):")
        for tech, nbr_cc, nbr_rr, pos in ref_interconnections:
            print(f"  {tech}  (neighbor: {nbr_cc}{nbr_rr})")

        print(f"\nNew {new_cc} requested interconnections "
              f"({len(new_neighbors)}):")
        for nbr_cc, nbr_rr in new_neighbors:
            print(f"  {new_cc}{new_rr} <-> {nbr_cc}{nbr_rr}")

        interconnection_mapping = build_interconnection_mapping(
            ref_cc, ref_rr, new_cc, new_rr,
            ref_interconnections, new_neighbors
        )

        print(f"\nMapping ({len(interconnection_mapping)} entries):")
        for m in interconnection_mapping:
            print(f"  {m['ref_tech']} -> {m['new_tech']}")

    print()

    # Generate centerpoint
    print("--- CENTERPOINT ---")
    if cp_lat is not None and cp_lon is not None:
        generate_centerpoint_template(new_cc, new_rr, cp_lat, cp_lon, output_dir)
    else:
        print("  WARNING: No centerpoint coordinates provided.")
        print("  Add 'centerpoint_lat' and 'centerpoint_lon' to Config_country_codes.yaml")
        print("  or pass --lat and --lon on the command line.")
        print("  The centerpoint will NOT be included in the merge script.")

    # Generate sets
    print("\n--- SETS ---")
    sets_templates = generate_sets_template(
        ref_cc, new_cc, interconnection_mapping
    )
    write_templates(sets_templates, output_dir)

    # Generate parameters
    print("\n--- PARAMETERS ---")
    param_templates = generate_param_templates(
        ref_cc, new_cc, interconnection_mapping, ref_rr, new_rr
    )
    write_templates(param_templates, output_dir)

    # Generate merge script
    print("\n--- HELPER SCRIPTS ---")
    generate_merge_script(new_cc, output_dir)

    # Summary
    total_files = len(sets_templates) + len(param_templates) + 1  # +1 for merge script
    if cp_lat is not None and cp_lon is not None:
        total_files += 1  # centerpoint.csv
    total_rows = sum(len(df) for df in sets_templates.values())
    total_rows += sum(len(df) for df in param_templates.values())

    print(f"\n{'=' * 50}")
    print(f"TEMPLATE GENERATION COMPLETE")
    print(f"{'=' * 50}")
    print(f"  Country: {new_cc}{new_rr} (based on {ref_cc})")
    if interconnection_mapping is not None:
        print(f"  Interconnections: {len(interconnection_mapping)}")
        for m in interconnection_mapping:
            print(f"    {m['new_tech']}")
    else:
        print(f"  Interconnections: copied from {ref_cc} (legacy mode)")
    if cp_lat is not None and cp_lon is not None:
        print(f"  Centerpoint: {new_cc}{new_rr} ({cp_lat}, {cp_lon})")
    else:
        print(f"  Centerpoint: NOT SET (add --lat/--lon or YAML config)")
    print(f"  Files created: {total_files}")
    print(f"  Total data rows: {total_rows}")
    print(f"  Output: {output_dir}")
    print()
    print("NEXT STEPS:")
    print(f"  1. Review and modify the CSV files in {output_dir}")
    print(f"  2. Update values to reflect actual data for {new_cc}")
    print(f"  3. Update Config_country_codes.yaml with {new_cc} entry")
    print(f"  4. Run: python {os.path.join(output_dir, 'merge_into_inputs.py')}")
    print(f"  5. Validate: python Z_validate_country_data.py --country {new_cc}")


def main():
    parser = argparse.ArgumentParser(
        description="Generate country template for OG_csvs_inputs. "
                    "Reads from Config_country_codes.yaml by default; "
                    "CLI args override the YAML values."
    )
    parser.add_argument(
        "--new", "-n",
        help="New country code (3 letters). Overrides YAML new_country."
    )
    parser.add_argument(
        "--ref", "-r",
        help="Reference country code. Overrides YAML reference_country."
    )
    parser.add_argument(
        "--output", "-o",
        help="Output directory (default: templates/<new_code>)"
    )
    parser.add_argument(
        "--interconnections", "-i", nargs="*",
        help="Neighbor codes for interconnections. Overrides YAML."
    )
    parser.add_argument(
        "--region",
        help="Region code for the new country. Overrides YAML region."
    )
    parser.add_argument(
        "--lat", type=float,
        help="Latitude for the country centerpoint. Overrides YAML."
    )
    parser.add_argument(
        "--lon", type=float,
        help="Longitude for the country centerpoint. Overrides YAML."
    )

    args = parser.parse_args()

    # --- Load YAML config ---
    yaml_entries = load_yaml_config()

    # --- CLI mode: single entry, CLI args override YAML first entry ---
    if args.new:
        entry = (yaml_entries[0] if yaml_entries else {}) if not args.ref else {}
        new_cc = args.new.upper()
        ref_cc = (args.ref or entry.get("reference_country", "ARG")).upper()
        new_rr = (args.region or entry.get("region", "XX")).upper()
        cp_lat = args.lat if args.lat is not None else entry.get("centerpoint_lat")
        cp_lon = args.lon if args.lon is not None else entry.get("centerpoint_lon")

        ixn_raw = None
        ixn_source = None
        if args.interconnections is not None:
            ixn_raw = args.interconnections
            ixn_source = "CLI"
        elif "interconnections" in entry:
            yaml_ixn = entry["interconnections"]
            ixn_raw = [] if yaml_ixn is None else [str(x) for x in yaml_ixn]
            ixn_source = "YAML"

        output_dir = args.output or os.path.join(SCRIPT_DIR, "templates", new_cc)
        process_entry(new_cc, ref_cc, new_rr, ixn_raw, ixn_source,
                      cp_lat, cp_lon, output_dir)
        return

    # --- YAML mode: process all entries in the list ---
    if not yaml_entries:
        print("ERROR: No configuration found.")
        print("  Set 'template_generation' in Config_country_codes.yaml")
        print("  or pass --new NCC on the command line.")
        sys.exit(1)

    print(f"Config loaded from: {os.path.basename(CONFIG_PATH)}")
    print(f"Found {len(yaml_entries)} entry/entries to process.\n")

    for i, entry in enumerate(yaml_entries, 1):
        new_cc = entry.get("new_country")
        if new_cc is None:
            print(f"WARNING: Entry {i} has no 'new_country', skipping.")
            continue
        new_cc = new_cc.upper()
        ref_cc = entry.get("reference_country", "ARG").upper()
        new_rr = entry.get("region", "XX").upper()
        cp_lat = entry.get("centerpoint_lat")
        cp_lon = entry.get("centerpoint_lon")

        ixn_raw = None
        ixn_source = None
        if "interconnections" in entry:
            yaml_ixn = entry["interconnections"]
            ixn_raw = [] if yaml_ixn is None else [str(x) for x in yaml_ixn]
            ixn_source = "YAML"

        # For multi-region countries, use CC+RR as subfolder
        subfolder = f"{new_cc}{new_rr}" if new_rr != "XX" else new_cc
        output_dir = os.path.join(SCRIPT_DIR, "templates", subfolder)

        if len(yaml_entries) > 1:
            print(f"\n{'#' * 50}")
            print(f"# Entry {i}/{len(yaml_entries)}: {new_cc}{new_rr}")
            print(f"{'#' * 50}")

        process_entry(new_cc, ref_cc, new_rr, ixn_raw, ixn_source,
                      cp_lat, cp_lon, output_dir)


if __name__ == "__main__":
    main()
