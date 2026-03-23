"""
Auto-generated script to merge MDV template data into OG_csvs_inputs.

IMPORTANT: Review the template CSV files BEFORE running this script.
Modify values as needed to reflect the actual data for MDV.

Usage:
    cd t1_confection
    python templates/MDV/merge_into_inputs.py

Author: Climate Lead Group, Andrey Salazar-Vargas
Date: 2025
"""

import pandas as pd
import os
import shutil
from datetime import datetime

TEMPLATE_DIR = "c:/users/climateleadgroup/desktop/clg_repositories/ostram/t1_confection/templates/MDV"
INPUT_DIR = "c:/users/climateleadgroup/desktop/clg_repositories/ostram/t1_confection/OG_csvs_inputs"
CENTERPOINTS_PATH = "c:/users/climateleadgroup/desktop/clg_repositories/ostram/t1_confection/Miscellaneous/centerpoints.csv"

backup_suffix = datetime.now().strftime("%Y%m%d_%H%M%S")


def merge_file(filename):
    """Merge a template file into the corresponding input file."""
    template_path = os.path.join(TEMPLATE_DIR, filename + ".csv")
    input_path = os.path.join(INPUT_DIR, filename + ".csv")

    if not os.path.exists(template_path):
        return

    template_df = pd.read_csv(template_path)
    if len(template_df) == 0:
        return

    if os.path.exists(input_path):
        input_df = pd.read_csv(input_path)
        backup_path = input_path.replace(".csv", f"_backup_{backup_suffix}.csv")
        shutil.copy2(input_path, backup_path)
        merged = pd.concat([input_df, template_df], ignore_index=True)
        merged.to_csv(input_path, index=False)
        print(f"  Merged {filename}.csv: {len(input_df)} + {len(template_df)} = {len(merged)} rows")
    else:
        template_df.to_csv(input_path, index=False)
        print(f"  Created {filename}.csv: {len(template_df)} rows")


# Sets
for s in ["TECHNOLOGY", "FUEL", "EMISSION", "STORAGE"]:
    merge_file(s)

# Parameters
params = [
    "CapitalCost", "FixedCost", "VariableCost",
    "ResidualCapacity", "CapacityFactor", "AvailabilityFactor",
    "InputActivityRatio", "OutputActivityRatio", "EmissionActivityRatio",
    "SpecifiedAnnualDemand", "SpecifiedDemandProfile",
    "OperationalLife", "CapacityToActivityUnit",
    "TotalAnnualMaxCapacity", "TotalAnnualMaxCapacityInvestment",
    "TotalTechnologyAnnualActivityUpperLimit",
    "ReserveMarginTagTechnology", "ReserveMarginTagFuel",
    "CapitalCostStorage", "OperationalLifeStorage",
    "StorageLevelStart", "ResidualStorageCapacity",
    "TechnologyToStorage", "TechnologyFromStorage",
]

for p in params:
    merge_file(p)

# Centerpoint
centerpoint_template = os.path.join(TEMPLATE_DIR, "centerpoint.csv")
if os.path.exists(centerpoint_template):
    cp_new = pd.read_csv(centerpoint_template)
    if len(cp_new) > 0:
        if os.path.exists(CENTERPOINTS_PATH):
            cp_existing = pd.read_csv(CENTERPOINTS_PATH)
            backup_path = CENTERPOINTS_PATH.replace(".csv", f"_backup_{backup_suffix}.csv")
            shutil.copy2(CENTERPOINTS_PATH, backup_path)
            # Remove existing entry for same region to avoid duplicates
            new_regions = set(cp_new["region"].astype(str))
            cp_existing = cp_existing[~cp_existing["region"].astype(str).isin(new_regions)]
            merged = pd.concat([cp_existing, cp_new], ignore_index=True)
            merged = merged.sort_values("region").reset_index(drop=True)
            merged.to_csv(CENTERPOINTS_PATH, index=False)
            print(f"  Merged centerpoint.csv into {CENTERPOINTS_PATH}")
        else:
            cp_new.to_csv(CENTERPOINTS_PATH, index=False)
            print(f"  Created {CENTERPOINTS_PATH}")

print("\nDone! Backup files created with suffix: " + backup_suffix)
print("Run the validation script to verify: python Z_validate_country_data.py --country MDV")
