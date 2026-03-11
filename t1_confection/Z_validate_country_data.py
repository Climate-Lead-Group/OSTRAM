"""
Validation script for OG_csvs_inputs data.
Validates that a country has all required data in the OSeMOSYS CSV input files.

Usage:
    python Z_validate_country_data.py                  # Validate all RELAC countries
    python Z_validate_country_data.py --country ARG    # Validate specific country
    python Z_validate_country_data.py --country NCC --report  # Generate detailed report

Author: Climate Lead Group, Andrey Salazar-Vargas
"""

import pandas as pd
import os
import sys
import argparse
from collections import defaultdict
import yaml
from Z_AUX_config_loader import get_countries
# ============================================================================
# Configuration
# ============================================================================

INPUT_DIR = os.path.join(os.path.dirname(__file__), "OG_csvs_inputs")

# Countries ISO-3 (read from Config_country_codes.yaml instead of hardcoded list)
OSTRAM_COUNTRIES  = get_countries()

# Sets that must contain entries for the country
REQUIRED_SETS = {
    "TECHNOLOGY": {"min_entries": 30, "description": "Technology definitions"},
    "FUEL": {"min_entries": 15, "description": "Fuel/commodity definitions"},
    "EMISSION": {"min_entries": 1, "description": "Emission type (CO2{CC})"},
    "STORAGE": {"min_entries": 2, "description": "Storage definitions (LDS + SDS)"},
}

# Parameters that MUST have data for every country
REQUIRED_PARAMS = [
    "CapitalCost",
    "FixedCost",
    "VariableCost",
    "ResidualCapacity",
    "CapacityFactor",
    "AvailabilityFactor",
    "InputActivityRatio",
    "OutputActivityRatio",
    "EmissionActivityRatio",
    "SpecifiedAnnualDemand",
    "SpecifiedDemandProfile",
    "OperationalLife",
    "CapacityToActivityUnit",
    "TotalAnnualMaxCapacity",
    "TotalAnnualMaxCapacityInvestment",
    "TotalTechnologyAnnualActivityUpperLimit",
    "ReserveMarginTagTechnology",
    "ReserveMarginTagFuel",
    "CapitalCostStorage",
    "OperationalLifeStorage",
    "StorageLevelStart",
    "TechnologyToStorage",
    "TechnologyFromStorage",
]

# Expected technology prefixes for a country
EXPECTED_TECH_PREFIXES = {
    "PWR": {"min": 15, "description": "Power generation technologies"},
    "MIN": {"min": 5, "description": "Mining/extraction technologies"},
    "RNW": {"min": 7, "description": "Renewable resource provisions"},
}

# Expected fuel patterns for a country
EXPECTED_FUEL_PATTERNS = {
    "ELC{CC}XX01": "Generated electricity",
    "ELC{CC}XX02": "Final demand electricity",
    "BIO{CC}XX": "Biomass",
    "HYD{CC}XX": "Hydroelectric resource",
    "SPV{CC}XX": "Solar PV resource",
    "WON{CC}XX": "Onshore wind resource",
    "COA{CC}": "Coal",
    "GAS{CC}": "Natural gas",
    "OIL{CC}": "Oil",
}

# Value range validations
VALUE_RANGES = {
    "CapitalCost": {"min": 0, "max": 15000, "description": "$/kW"},
    "FixedCost": {"min": 0, "max": 500, "description": "$/kW/year"},
    "VariableCost": {"min": 0, "max": 50, "description": "M$/PJ"},
    "ResidualCapacity": {"min": 0, "max": 200, "description": "GW"},
    "CapacityFactor": {"min": 0, "max": 1, "description": "fraction"},
    "AvailabilityFactor": {"min": 0, "max": 1, "description": "fraction"},
    "InputActivityRatio": {"min": 0.5, "max": 10, "description": "ratio"},
    "OutputActivityRatio": {"min": 0, "max": 5, "description": "ratio"},
    "EmissionActivityRatio": {"min": 0, "max": 1, "description": "Mt CO2/PJ"},
    "SpecifiedAnnualDemand": {"min": 0, "max": 10000, "description": "PJ"},
    "SpecifiedDemandProfile": {"min": 0, "max": 1, "description": "fraction"},
    "OperationalLife": {"min": 5, "max": 150, "description": "years"},
    "CapacityToActivityUnit": {"min": 31, "max": 32, "description": "PJ/GW/year"},
}


# ============================================================================
# Helper functions
# ============================================================================

def load_csv(filename):
    """Load a CSV file from the input directory."""
    filepath = os.path.join(INPUT_DIR, filename + ".csv")
    if not os.path.exists(filepath):
        return None
    return pd.read_csv(filepath)


def find_country_rows(df, country_code, columns=None):
    """Find rows in a DataFrame that contain the country code."""
    if columns is None:
        columns = [c for c in df.columns if c in
                   ["TECHNOLOGY", "FUEL", "STORAGE", "EMISSION"]]
    if not columns:
        return pd.DataFrame()

    mask = pd.Series(False, index=df.index)
    for col in columns:
        mask |= df[col].astype(str).str.contains(country_code, na=False)
    return df[mask]


# ============================================================================
# Validation checks
# ============================================================================

class ValidationResult:
    """Stores the result of a single validation check."""

    def __init__(self, check_name, status, message, details=None):
        self.check_name = check_name
        self.status = status  # "PASS", "FAIL", "WARN"
        self.message = message
        self.details = details or []

    def __str__(self):
        icon = {"PASS": "[OK]", "FAIL": "[FAIL]", "WARN": "[WARN]"}[self.status]
        return f"  {icon} {self.check_name}: {self.message}"


def validate_sets(country_code):
    """Validate that the country exists in all required sets."""
    results = []

    for set_name, config in REQUIRED_SETS.items():
        df = load_csv(set_name)
        if df is None:
            results.append(ValidationResult(
                f"SET:{set_name}", "FAIL",
                f"File {set_name}.csv not found"
            ))
            continue

        entries = [v for v in df["VALUE"].astype(str) if country_code in v]
        count = len(entries)
        min_expected = config["min_entries"]

        if count >= min_expected:
            results.append(ValidationResult(
                f"SET:{set_name}", "PASS",
                f"{count} entries found (min: {min_expected})",
                entries[:5]
            ))
        elif count > 0:
            results.append(ValidationResult(
                f"SET:{set_name}", "WARN",
                f"Only {count} entries (expected >= {min_expected})",
                entries
            ))
        else:
            results.append(ValidationResult(
                f"SET:{set_name}", "FAIL",
                f"No entries found for {country_code}",
            ))

    return results


def validate_required_params(country_code):
    """Validate that all required parameter files have data for the country."""
    results = []

    for param in REQUIRED_PARAMS:
        df = load_csv(param)
        if df is None:
            results.append(ValidationResult(
                f"PARAM:{param}", "FAIL",
                f"File {param}.csv not found"
            ))
            continue

        if len(df) == 0:
            results.append(ValidationResult(
                f"PARAM:{param}", "FAIL",
                f"File is empty (no data rows)"
            ))
            continue

        country_rows = find_country_rows(df, country_code)
        count = len(country_rows)

        if count > 0:
            results.append(ValidationResult(
                f"PARAM:{param}", "PASS",
                f"{count} rows found"
            ))
        else:
            results.append(ValidationResult(
                f"PARAM:{param}", "FAIL",
                f"No rows found for {country_code}"
            ))

    return results


def validate_value_ranges(country_code):
    """Validate that parameter values are within expected ranges."""
    results = []

    for param, ranges in VALUE_RANGES.items():
        df = load_csv(param)
        if df is None or len(df) == 0:
            continue

        country_rows = find_country_rows(df, country_code)
        if len(country_rows) == 0:
            continue

        values = country_rows["VALUE"]
        out_of_range = values[(values < ranges["min"]) | (values > ranges["max"])]

        if len(out_of_range) == 0:
            results.append(ValidationResult(
                f"RANGE:{param}", "PASS",
                f"All {len(values)} values in range "
                f"[{ranges['min']}, {ranges['max']}] {ranges['description']}"
            ))
        else:
            pct = len(out_of_range) / len(values) * 100
            results.append(ValidationResult(
                f"RANGE:{param}", "WARN",
                f"{len(out_of_range)}/{len(values)} values ({pct:.1f}%) "
                f"outside [{ranges['min']}, {ranges['max']}] {ranges['description']}",
                [f"  Min={values.min():.4f}, Max={values.max():.4f}"]
            ))

    return results


def validate_demand_profile(country_code):
    """Validate that SpecifiedDemandProfile sums to 1.0 per year."""
    results = []

    df = load_csv("SpecifiedDemandProfile")
    if df is None or len(df) == 0:
        results.append(ValidationResult(
            "PROFILE:DemandSum", "FAIL",
            "SpecifiedDemandProfile not found or empty"
        ))
        return results

    country_rows = find_country_rows(df, country_code, ["FUEL"])
    if len(country_rows) == 0:
        results.append(ValidationResult(
            "PROFILE:DemandSum", "FAIL",
            f"No demand profile data for {country_code}"
        ))
        return results

    # Check that profile sums to ~1.0 for each (FUEL, YEAR) combination
    grouped = country_rows.groupby(["FUEL", "YEAR"])["VALUE"].sum()
    bad_sums = grouped[abs(grouped - 1.0) > 0.001]

    if len(bad_sums) == 0:
        results.append(ValidationResult(
            "PROFILE:DemandSum", "PASS",
            f"All {len(grouped)} (FUEL, YEAR) groups sum to 1.0 (±0.001)"
        ))
    else:
        results.append(ValidationResult(
            "PROFILE:DemandSum", "FAIL",
            f"{len(bad_sums)}/{len(grouped)} groups DO NOT sum to 1.0",
            [f"  {idx}: sum={val:.6f}" for idx, val in bad_sums.head(5).items()]
        ))

    return results


def validate_fuel_completeness(country_code):
    """Validate that essential fuels exist for the country."""
    results = []

    df = load_csv("FUEL")
    if df is None:
        results.append(ValidationResult(
            "FUEL:Essential", "FAIL", "FUEL.csv not found"
        ))
        return results

    fuels = set(df["VALUE"].astype(str))

    # Check essential electricity fuels
    elc01 = f"ELC{country_code}XX01"
    elc02 = f"ELC{country_code}XX02"

    if elc01 in fuels and elc02 in fuels:
        results.append(ValidationResult(
            "FUEL:Electricity", "PASS",
            f"Both {elc01} and {elc02} exist"
        ))
    else:
        missing = []
        if elc01 not in fuels:
            missing.append(elc01)
        if elc02 not in fuels:
            missing.append(elc02)
        results.append(ValidationResult(
            "FUEL:Electricity", "FAIL",
            f"Missing electricity fuels: {missing}"
        ))

    # Check demand fuel in SpecifiedAnnualDemand
    sad = load_csv("SpecifiedAnnualDemand")
    if sad is not None and len(sad) > 0:
        demand_fuels = set(sad["FUEL"].astype(str))
        if elc02 in demand_fuels:
            results.append(ValidationResult(
                "FUEL:DemandDefined", "PASS",
                f"{elc02} has SpecifiedAnnualDemand entries"
            ))
        else:
            results.append(ValidationResult(
                "FUEL:DemandDefined", "FAIL",
                f"{elc02} NOT in SpecifiedAnnualDemand (country has no demand!)"
            ))

    return results


def validate_tech_types(country_code):
    """Validate that essential technology types exist."""
    results = []

    df = load_csv("TECHNOLOGY")
    if df is None:
        results.append(ValidationResult(
            "TECH:Types", "FAIL", "TECHNOLOGY.csv not found"
        ))
        return results

    techs = [t for t in df["VALUE"].astype(str) if country_code in t]

    for prefix, config in EXPECTED_TECH_PREFIXES.items():
        matching = [t for t in techs if t.startswith(prefix)]
        if len(matching) >= config["min"]:
            results.append(ValidationResult(
                f"TECH:{prefix}", "PASS",
                f"{len(matching)} {config['description']} found "
                f"(min: {config['min']})"
            ))
        elif len(matching) > 0:
            results.append(ValidationResult(
                f"TECH:{prefix}", "WARN",
                f"Only {len(matching)} {config['description']} "
                f"(expected >= {config['min']})",
                matching
            ))
        else:
            results.append(ValidationResult(
                f"TECH:{prefix}", "FAIL",
                f"No {config['description']} found"
            ))

    # Check backstop
    bck = [t for t in techs if "BCK" in t]
    if bck:
        results.append(ValidationResult(
            "TECH:Backstop", "PASS",
            f"Backstop technology exists: {bck[0]}"
        ))
    else:
        results.append(ValidationResult(
            "TECH:Backstop", "WARN",
            "No backstop (BCK) technology found — model may be infeasible"
        ))

    # Check intranational transmission
    trn = [t for t in techs if t.startswith("PWRTRN")]
    if trn:
        results.append(ValidationResult(
            "TECH:IntraTRN", "PASS",
            f"Intranational transmission exists: {trn[0]}"
        ))
    else:
        results.append(ValidationResult(
            "TECH:IntraTRN", "WARN",
            "No intranational transmission (PWRTRN) found"
        ))

    return results


def validate_referential_integrity(country_code):
    """Check that technologies in parameters exist in the TECHNOLOGY set."""
    results = []

    tech_df = load_csv("TECHNOLOGY")
    fuel_df = load_csv("FUEL")
    if tech_df is None or fuel_df is None:
        return results

    tech_set = set(tech_df["VALUE"].astype(str))
    fuel_set = set(fuel_df["VALUE"].astype(str))

    # Check a few key parameters for referential integrity
    for param in ["CapitalCost", "OperationalLife", "InputActivityRatio"]:
        df = load_csv(param)
        if df is None or len(df) == 0:
            continue

        country_rows = find_country_rows(df, country_code)

        if "TECHNOLOGY" in country_rows.columns:
            techs_in_param = set(country_rows["TECHNOLOGY"].astype(str))
            orphan_techs = techs_in_param - tech_set
            if orphan_techs:
                # Check if they match the consolidated pattern (without 00/01 suffix)
                truly_orphan = []
                for t in orphan_techs:
                    # Look for a match with suffix
                    if not any(s in tech_set for s in [t + "00", t + "01"]):
                        truly_orphan.append(t)

                if truly_orphan:
                    results.append(ValidationResult(
                        f"REF:{param}:TECH", "WARN",
                        f"{len(orphan_techs)} techs in {param} not in "
                        f"TECHNOLOGY.csv ({len(truly_orphan)} truly orphan)",
                        truly_orphan[:5]
                    ))
                else:
                    results.append(ValidationResult(
                        f"REF:{param}:TECH", "PASS",
                        f"All techs match (consolidated naming confirmed)"
                    ))

        if "FUEL" in country_rows.columns:
            fuels_in_param = set(country_rows["FUEL"].astype(str))
            orphan_fuels = fuels_in_param - fuel_set
            if orphan_fuels:
                results.append(ValidationResult(
                    f"REF:{param}:FUEL", "WARN",
                    f"{len(orphan_fuels)} fuels in {param} not in FUEL.csv",
                    list(orphan_fuels)[:5]
                ))

    return results


def validate_storage(country_code):
    """Validate storage configuration for the country."""
    results = []

    storage_df = load_csv("STORAGE")
    if storage_df is None:
        return results

    storages = [s for s in storage_df["VALUE"].astype(str)
                if country_code in s]

    if len(storages) >= 2:
        results.append(ValidationResult(
            "STORAGE:Defined", "PASS",
            f"{len(storages)} storages found: {storages}"
        ))
    elif len(storages) > 0:
        results.append(ValidationResult(
            "STORAGE:Defined", "WARN",
            f"Only {len(storages)} storage(s): {storages} (expected LDS + SDS)"
        ))
    else:
        results.append(ValidationResult(
            "STORAGE:Defined", "FAIL",
            "No storage defined"
        ))

    # Check storage links
    for param in ["TechnologyToStorage", "TechnologyFromStorage"]:
        df = load_csv(param)
        if df is None or len(df) == 0:
            continue
        country_rows = find_country_rows(df, country_code,
                                          ["TECHNOLOGY", "STORAGE"])
        if len(country_rows) > 0:
            results.append(ValidationResult(
                f"STORAGE:{param}", "PASS",
                f"{len(country_rows)} links found"
            ))
        else:
            results.append(ValidationResult(
                f"STORAGE:{param}", "WARN",
                "No storage links found"
            ))

    return results


# ============================================================================
# Main validation runner
# ============================================================================

def validate_country(country_code, verbose=True):
    """Run all validations for a country and return results."""
    all_results = []

    checks = [
        ("SET VALIDATION", validate_sets),
        ("TECHNOLOGY TYPES", validate_tech_types),
        ("FUEL COMPLETENESS", validate_fuel_completeness),
        ("REQUIRED PARAMETERS", validate_required_params),
        ("VALUE RANGES", validate_value_ranges),
        ("DEMAND PROFILE", validate_demand_profile),
        ("STORAGE", validate_storage),
        ("REFERENTIAL INTEGRITY", validate_referential_integrity),
    ]

    for section_name, check_fn in checks:
        results = check_fn(country_code)
        all_results.extend(results)

        if verbose:
            print(f"\n--- {section_name} ---")
            for r in results:
                print(r)
                if r.details and r.status != "PASS":
                    for d in r.details:
                        print(f"       {d}")

    return all_results


def generate_summary(country_code, results):
    """Generate a summary of validation results."""
    total = len(results)
    passed = sum(1 for r in results if r.status == "PASS")
    warned = sum(1 for r in results if r.status == "WARN")
    failed = sum(1 for r in results if r.status == "FAIL")

    print(f"\n{'=' * 60}")
    print(f"VALIDATION SUMMARY: {country_code}")
    print(f"{'=' * 60}")
    print(f"  Total checks: {total}")
    print(f"  Passed:       {passed} ({passed/total*100:.0f}%)")
    print(f"  Warnings:     {warned} ({warned/total*100:.0f}%)")
    print(f"  Failed:       {failed} ({failed/total*100:.0f}%)")
    print(f"{'=' * 60}")

    if failed > 0:
        print("\nFAILED CHECKS:")
        for r in results:
            if r.status == "FAIL":
                print(f"  - {r.check_name}: {r.message}")

    if warned > 0:
        print("\nWARNINGS:")
        for r in results:
            if r.status == "WARN":
                print(f"  - {r.check_name}: {r.message}")

    overall = "PASS" if failed == 0 else "FAIL"
    print(f"\nOverall: {overall}")
    return overall


def main():
    parser = argparse.ArgumentParser(
        description="Validate country data in OG_csvs_inputs"
    )
    parser.add_argument(
        "--country", "-c",
        help="Country code to validate (e.g., ARG). "
             "If not specified, validates all RELAC countries."
    )
    parser.add_argument(
        "--report", "-r",
        action="store_true",
        help="Generate detailed report"
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Only show summary"
    )

    args = parser.parse_args()

    if not os.path.isdir(INPUT_DIR):
        print(f"ERROR: Input directory not found: {INPUT_DIR}")
        sys.exit(1)

    countries = [args.country.upper()] if args.country else OSTRAM_COUNTRIES
    verbose = not args.quiet

    overall_results = {}

    for cc in countries:
        print(f"\n{'#' * 60}")
        print(f"# VALIDATING: {cc}")
        print(f"{'#' * 60}")

        results = validate_country(cc, verbose=verbose)
        overall = generate_summary(cc, results)
        overall_results[cc] = overall

    # Final summary for all countries
    if len(countries) > 1:
        print(f"\n\n{'=' * 60}")
        print("GLOBAL SUMMARY")
        print(f"{'=' * 60}")
        for cc, status in overall_results.items():
            icon = "[OK]  " if status == "PASS" else "[FAIL]"
            print(f"  {icon} {cc}")

        total_pass = sum(1 for s in overall_results.values() if s == "PASS")
        total_fail = sum(1 for s in overall_results.values() if s == "FAIL")
        print(f"\n  Passed: {total_pass}/{len(countries)}")
        print(f"  Failed: {total_fail}/{len(countries)}")

    sys.exit(0 if all(s == "PASS" for s in overall_results.values()) else 1)


if __name__ == "__main__":
    main()
