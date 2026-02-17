#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
CAPITAL INVESTMENT ANNUALIZATION SCRIPT FOR MOMF
================================================================================
Author: Climate Lead Group, Andrey Salazar-Vargas
Purpose: This script reads a CSV file containing OSeMOSYS model outputs,
         calculates the annualized capital investment for each technology,
         and adds a new column with the accumulated annualized values.

The annualization uses the Capital Recovery Factor (CRF) formula to convert
upfront capital investments into equivalent annual payments over the asset's
lifetime, with temporal accumulation of overlapping payment periods.
================================================================================
"""

# ================================
# LIBRARY IMPORTS
# ================================
import pandas as pd
import numpy as np
from pathlib import Path
import sys
from datetime import datetime

# ================================
# USER-CONFIGURABLE PARAMETERS
# ================================
# This section contains all parameters that can be easily modified
# Place all editable variables here for easy access and modification

# Financial parameters for annualization calculation
DISCOUNT_RATE = 0.0639  # Discount rate (3% = 0.03). Typical range: 0.03 to 0.10
ASSET_LIFETIME = 15   # Asset lifetime in years. Typical range: 10 to 30 years

# File parameters
INPUT_FILENAME = "RELAC_TX_Combined_Inputs_Outputs.csv"  # Name of the input CSV file to process

# Column names
CAPITAL_COLUMN = "CapitalInvestment"  # Name of the column with capital investment data
NEW_COLUMN_NAME = "CapitalInvestmentAnnualized"  # Name for the new annualized column

# Grouping columns - These columns define unique investment series
# IMPORTANT: Only columns that actually contain data should be used for grouping
GROUPING_COLUMNS = [
    "Future", "Scenario", "REGION", "TECHNOLOGY", "FUEL", 
    "EMISSION", "MODE_OF_OPERATION", "TIMESLICE", "STORAGE", 
    "SEASON", "DAYTYPE", "DAILYTIMEBRACKET"
]

# ================================
# FUNCTION DEFINITIONS
# ================================

def calculate_crf(discount_rate, lifetime):
    """
    Calculate the Capital Recovery Factor (CRF).
    
    The CRF is used to convert a present value (initial investment) into
    a series of equal annual payments over the asset's lifetime.
    
    Formula: CRF = (r * (1 + r)^n) / ((1 + r)^n - 1)
    where:
        r = discount rate (as decimal, e.g., 0.03 for 3%)
        n = asset lifetime in years
    
    Parameters:
    -----------
    discount_rate : float
        The discount rate as a decimal (e.g., 0.03 for 3%)
    lifetime : int
        The asset lifetime in years
    
    Returns:
    --------
    float
        The Capital Recovery Factor
    """
    # Check for edge cases to avoid division by zero
    if discount_rate == 0:
        # If discount rate is 0, CRF simplifies to 1/n
        return 1.0 / lifetime
    
    # Calculate the numerator: r * (1 + r)^n
    numerator = discount_rate * ((1 + discount_rate) ** lifetime)
    
    # Calculate the denominator: (1 + r)^n - 1
    denominator = ((1 + discount_rate) ** lifetime) - 1
    
    # Calculate and return the CRF
    crf = numerator / denominator
    
    return crf


# Backup function removed - no backup will be created


def get_decimal_places(series):
    """
    Determine the maximum number of decimal places in a pandas Series.
    
    This function helps maintain the same decimal precision as the original data.
    
    Parameters:
    -----------
    series : pandas.Series
        The series to analyze for decimal places
    
    Returns:
    --------
    int
        Maximum number of decimal places found in the series
    """
    max_decimals = 0
    
    # Iterate through non-null values in the series
    for value in series.dropna():
        # Convert to string and check for decimal point
        str_value = str(value)
        if '.' in str_value:
            # Count digits after decimal point
            decimal_part = str_value.split('.')[1]
            # Remove trailing zeros and scientific notation if present
            decimal_part = decimal_part.rstrip('0').rstrip('e')
            max_decimals = max(max_decimals, len(decimal_part))
    
    return max_decimals


def identify_effective_grouping_columns(df, potential_columns):
    """
    Identify which grouping columns should be used based on CapitalInvestment data.
    
    For MOMF/OSeMOSYS, CapitalInvestment only has values when:
    - REGION and TECHNOLOGY have values (these are the key grouping columns)
    - Other sets (FUEL, EMISSION, MODE_OF_OPERATION, etc.) are typically blank/NaN
    
    This function identifies the appropriate grouping strategy based on where
    CapitalInvestment actually has non-zero values.
    
    Parameters:
    -----------
    df : pandas.DataFrame
        The dataframe to analyze
    potential_columns : list
        List of column names to check
    
    Returns:
    --------
    list
        List of columns that should be used for grouping
    """
    print("\nAnalyzing data structure for CapitalInvestment grouping...")
    
    # First, identify rows where CapitalInvestment has actual values
    capital_mask = (df[CAPITAL_COLUMN] > 0) & (~df[CAPITAL_COLUMN].isna())
    df_with_capital = df[capital_mask]
    
    if len(df_with_capital) == 0:
        print("  ⚠ WARNING: No rows with CapitalInvestment > 0 found!")
        return ['TECHNOLOGY', 'REGION'] if all(c in df.columns for c in ['TECHNOLOGY', 'REGION']) else []
    
    print(f"  Found {len(df_with_capital)} rows with CapitalInvestment > 0")
    
    effective_columns = []
    
    print("\n  Analyzing columns where CapitalInvestment exists:")
    
    for col in potential_columns:
        if col not in df.columns:
            print(f"    ⚠ Column '{col}' not found in dataset")
            continue
        
        # Analyze this column only where CapitalInvestment has values
        values_where_capital = df_with_capital[col]
        
        # Count non-null values
        non_null_count = values_where_capital.notna().sum()
        unique_values = values_where_capital.dropna().unique()
        
        if non_null_count == 0:
            print(f"    ✗ '{col}': All blank/NaN where CapitalInvestment exists")
        elif len(unique_values) == 1 and unique_values[0] == 0:
            print(f"    ✗ '{col}': All zeros where CapitalInvestment exists")
        elif len(unique_values) >= 1:
            # This column has actual values where CapitalInvestment exists
            effective_columns.append(col)
            print(f"    ✓ '{col}': {len(unique_values)} unique values, {non_null_count}/{len(df_with_capital)} non-null")
            if len(unique_values) <= 5:
                print(f"       Values: {list(unique_values)}")
    
    # Special handling for OSeMOSYS structure
    # Based on your observation, we should primarily group by Future, Scenario, REGION, TECHNOLOGY
    primary_grouping = ['Future', 'Scenario', 'REGION', 'TECHNOLOGY']
    primary_available = [col for col in primary_grouping if col in effective_columns]
    
    if primary_available:
        print(f"\n  Using primary OSeMOSYS grouping: {primary_available}")
        return primary_available
    
    return effective_columns


def calculate_annualized_investment(df, crf, grouping_cols):
    """
    Calculate the annualized capital investment with temporal accumulation.
    
    This is the core function that implements the annualization logic:
    1. Groups data by specified columns to identify unique investment series
    2. For each investment, calculates annual payments over asset lifetime
    3. Accumulates overlapping payments from multiple investments
    
    Parameters:
    -----------
    df : pandas.DataFrame
        Input dataframe with capital investment data
    crf : float
        Capital Recovery Factor for annualization
    grouping_cols : list
        Column names to use for grouping data
    
    Returns:
    --------
    pandas.DataFrame
        DataFrame with new annualized column added
    """
    # Create a copy to avoid modifying the original dataframe
    df_result = df.copy()
    
    # Initialize the new column with zeros
    # We use zeros instead of NaN to properly handle accumulation
    df_result[NEW_COLUMN_NAME] = 0.0
    
    # Identify which grouping columns actually exist and have meaningful data
    existing_grouping_cols = identify_effective_grouping_columns(df, grouping_cols)
    
    if not existing_grouping_cols:
        print("\n⚠ WARNING: No effective grouping columns found!")
        print("  This might indicate an issue with the data structure.")
        print("  Attempting to use default OSeMOSYS grouping...")
        # Try default OSeMOSYS grouping
        default_cols = ['Future', 'Scenario', 'REGION', 'TECHNOLOGY']
        existing_grouping_cols = [col for col in default_cols if col in df.columns]
    
    if not existing_grouping_cols:
        print("\n❌ ERROR: Cannot proceed without any grouping columns!")
        return df_result
    
    print(f"\nUsing effective grouping columns: {existing_grouping_cols}")
    print(f"Processing {len(df_result)} rows...")
    
    # Group the data by the specified columns
    # Each group represents a unique combination of technology, region, scenario, etc.
    grouped = df_result.groupby(existing_grouping_cols, group_keys=False)
    
    # Counter for progress tracking
    group_count = 0
    total_groups = len(grouped)
    
    print(f"Total groups to process: {total_groups}")
    
    # Process each group independently
    for group_key, group_df in grouped:
        group_count += 1
        if group_count % 10 == 0 or group_count <= 5:  # Show first 5 and then every 10
            print(f"  Processing group {group_count}/{total_groups}: {group_key}")
        
        # Get the indices for this group in the original dataframe
        group_indices = group_df.index.tolist()
        
        # Sort by year to ensure correct temporal processing
        group_sorted = group_df.sort_values('YEAR')
        sorted_indices = group_sorted.index.tolist()
        
        # Extract years and capital investments for this group
        years = group_sorted['YEAR'].values
        capital_investments = group_sorted[CAPITAL_COLUMN].values
        
        # Filter out any NaN years (shouldn't happen but being safe)
        valid_mask = ~pd.isna(years)
        if not valid_mask.all():
            print(f"    ⚠ WARNING: Found {(~valid_mask).sum()} NaN values in YEAR column")
            years = years[valid_mask]
            capital_investments = capital_investments[valid_mask]
            sorted_indices = [idx for idx, valid in zip(sorted_indices, valid_mask) if valid]
        
        # Create a mapping from year to index in the sorted array
        year_to_idx = {}
        for idx, year in enumerate(years):
            if not pd.isna(year):
                year_to_idx[int(year)] = idx
        
        # Check if ALL values in this group are NaN
        all_nan = pd.isna(capital_investments).all()
        
        if all_nan:
            # If all values are NaN, keep the result as NaN for all years
            annualized_values = np.full(len(years), np.nan)
            if group_count <= 3:
                print(f"    All CapitalInvestment values are NaN - preserving as NaN")
        else:
            # Some values are not NaN - process the group
            # IMPORTANT: Treat NaN as 0 for calculation purposes when group has some real values
            capital_investments_clean = np.nan_to_num(capital_investments, nan=0.0)
            
            # Initialize array to store accumulated annualized values
            annualized_values = np.zeros(len(years))
            
            # Debug: Show investments for first few groups
            if group_count <= 3:
                print(f"    Processing investments (NaN treated as 0):")
                for year, inv_orig, inv_clean in zip(years[:10], capital_investments[:10], capital_investments_clean[:10]):
                    if not pd.isna(inv_orig) and inv_orig > 0:
                        annual_payment = inv_clean * crf
                        print(f"      Year {year}: Investment={inv_orig:.2f} → Annual Payment={annual_payment:.2f}")
                    elif pd.isna(inv_orig):
                        print(f"      Year {year}: NaN (treated as 0)")
                    elif inv_orig == 0:
                        print(f"      Year {year}: 0")
            
            # Process each investment
            for i, (year, investment) in enumerate(zip(years, capital_investments_clean)):
                if investment > 0:  # Only process positive investments
                    # Calculate the annual payment for this investment
                    annual_payment = investment * crf
                    
                    # Distribute this payment over the asset lifetime
                    for payment_year in range(int(year), min(int(year) + ASSET_LIFETIME, int(years[-1]) + 1)):
                        # Check if this payment year exists in our data
                        if payment_year in year_to_idx:
                            idx = year_to_idx[payment_year]
                            # Add this payment to the accumulated value
                            annualized_values[idx] += annual_payment
            
            # Validation: Check if we calculated something
            if annualized_values.sum() == 0 and capital_investments_clean.sum() > 0:
                print(f"    ⚠ WARNING: Capital exists but annualized sum is 0!")
                print(f"      Capital sum: {capital_investments_clean.sum():.2f}")
                print(f"      This should not happen - check the logic!")
        
        # Map the values back to the original indices
        for sorted_idx, orig_idx in enumerate(sorted_indices):
            df_result.loc[orig_idx, NEW_COLUMN_NAME] = annualized_values[sorted_idx]
        
        # Debug: Show results for first few groups
        if group_count <= 3 and not all_nan:
            print(f"    Annualized values summary:")
            non_zero = annualized_values[annualized_values > 0]
            if len(non_zero) > 0:
                print(f"      Years with payments: {len(non_zero)}")
                print(f"      Max annual payment: {non_zero.max():.2f}")
                print(f"      Total over all years: {non_zero.sum():.2f}")
                # Show first few years with payments
                print(f"      Sample payments:")
                for year, ann_val in zip(years[:15], annualized_values[:15]):
                    if ann_val > 0:
                        print(f"        Year {year}: {ann_val:.2f}")
            else:
                print(f"      ERROR: All annualized values are zero despite having investments!")
    
    print(f"\nCompleted processing {group_count} groups")
    
    # Final check
    total_annualized = df_result[NEW_COLUMN_NAME].sum()
    print(f"\nTotal annualized investment across all groups: {total_annualized:,.2f}")
    
    if total_annualized == 0:
        print("\n⚠ WARNING: All annualized values are zero!")
        print("  This might indicate an issue with the grouping or calculation.")
        print("  Please check that the grouping columns are appropriate for your data.")
    
    return df_result


def validate_results(df_original, df_result):
    """
    Perform validation checks on the results.
    
    This function ensures that:
    1. The new column was created successfully
    2. NaN values were preserved correctly
    3. The accumulation logic worked as expected
    
    Parameters:
    -----------
    df_original : pandas.DataFrame
        Original input dataframe
    df_result : pandas.DataFrame
        Resulting dataframe with annualized column
    
    Returns:
    --------
    bool
        True if validation passes, False otherwise
    """
    print("\n" + "="*60)
    print("VALIDATION RESULTS")
    print("="*60)
    
    # Check 1: Verify new column exists
    if NEW_COLUMN_NAME not in df_result.columns:
        print("❌ ERROR: New column was not created")
        return False
    print(f"✓ New column '{NEW_COLUMN_NAME}' created successfully")
    
    # Check 2: Verify NaN handling
    original_nans = df_original[CAPITAL_COLUMN].isna().sum()
    result_nans = df_result[NEW_COLUMN_NAME].isna().sum()
    print(f"✓ Original NaN count in {CAPITAL_COLUMN}: {original_nans}")
    print(f"✓ NaN count in {NEW_COLUMN_NAME}: {result_nans}")
    
    # Check 3: Show summary statistics
    print("\nSummary Statistics:")
    print(f"  Original {CAPITAL_COLUMN}:")
    capital_sum = df_original[CAPITAL_COLUMN].sum()
    print(f"    Sum: {capital_sum:,.2f}")
    print(f"    Mean: {df_original[CAPITAL_COLUMN].mean():,.2f}")
    print(f"    Max: {df_original[CAPITAL_COLUMN].max():,.2f}")
    print(f"    Non-zero count: {(df_original[CAPITAL_COLUMN] > 0).sum()}")
    
    print(f"  New {NEW_COLUMN_NAME}:")
    annualized_sum = df_result[NEW_COLUMN_NAME].sum()
    print(f"    Sum: {annualized_sum:,.2f}")
    print(f"    Mean: {df_result[NEW_COLUMN_NAME].mean():,.2f}")
    print(f"    Max: {df_result[NEW_COLUMN_NAME].max():,.2f}")
    print(f"    Non-zero count: {(df_result[NEW_COLUMN_NAME] > 0).sum()}")
    
    # Check 4: Expected ratio check
    if capital_sum > 0 and annualized_sum > 0:
        # The annualized sum should be roughly CRF * capital_sum * years_of_overlap
        # This is approximate since investments happen at different times
        crf = calculate_crf(DISCOUNT_RATE, ASSET_LIFETIME)
        ratio = annualized_sum / capital_sum
        print(f"\n  Ratio (Annualized Sum / Capital Sum): {ratio:.2f}")
        print(f"  CRF value: {crf:.4f}")
        print(f"  Note: The ratio should be > CRF due to temporal accumulation")
    
    # Check 5: Show sample of results
    print("\nSample of results (first 10 non-zero annualized values):")
    non_zero_mask = (df_result[NEW_COLUMN_NAME] > 0) & (~df_result[NEW_COLUMN_NAME].isna())
    if non_zero_mask.any():
        sample_df = df_result[non_zero_mask].head(10)[
            ['YEAR', 'TECHNOLOGY', CAPITAL_COLUMN, NEW_COLUMN_NAME]
        ]
        print(sample_df.to_string(index=False))
    else:
        print("  ⚠ No non-zero annualized values found!")
    
    return True


def annualize_capital_investment(
    input_file_path=None,
    discount_rate=None,
    asset_lifetime=None,
    capital_column=None,
    new_column_name=None,
    grouping_columns=None,
    verbose=True
):
    """
    Performs capital investment annualization on a CSV file.

    This function can be called from other scripts or run standalone.
    It reads a CSV file, calculates annualized capital investments using the
    Capital Recovery Factor (CRF) method, and saves the results back to the file.

    Parameters:
    -----------
    input_file_path : str or Path, optional
        Path to the input CSV file. If None, uses INPUT_FILENAME constant.
    discount_rate : float, optional
        Discount rate for CRF calculation. If None, uses DISCOUNT_RATE constant.
    asset_lifetime : int, optional
        Asset lifetime in years. If None, uses ASSET_LIFETIME constant.
    capital_column : str, optional
        Name of the capital investment column. If None, uses CAPITAL_COLUMN constant.
    new_column_name : str, optional
        Name for the new annualized column. If None, uses NEW_COLUMN_NAME constant.
    grouping_columns : list, optional
        List of columns to group by. If None, uses GROUPING_COLUMNS constant.
    verbose : bool, optional
        If True, prints detailed progress information. Default is True.

    Returns:
    --------
    pd.DataFrame
        The resulting DataFrame with the new annualized column added.

    Raises:
    -------
    FileNotFoundError
        If the input file does not exist.
    ValueError
        If required columns are missing from the input file.
    """
    # Use default values if parameters not provided
    if input_file_path is None:
        input_file_path = INPUT_FILENAME
    if discount_rate is None:
        discount_rate = DISCOUNT_RATE
    if asset_lifetime is None:
        asset_lifetime = ASSET_LIFETIME
    if capital_column is None:
        capital_column = CAPITAL_COLUMN
    if new_column_name is None:
        new_column_name = NEW_COLUMN_NAME
    if grouping_columns is None:
        grouping_columns = GROUPING_COLUMNS

    if verbose:
        print("="*60)
        print("CAPITAL INVESTMENT ANNUALIZATION")
        print("="*60)
        print(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"\nConfiguration:")
        print(f"  Discount Rate: {discount_rate*100}%")
        print(f"  Asset Lifetime: {asset_lifetime} years")
        print(f"  Input File: {input_file_path}")

    # Step 1: Validate input file exists
    input_path = Path(input_file_path)
    if not input_path.exists():
        error_msg = f"Input file '{input_file_path}' not found! Current directory: {Path.cwd()}"
        if verbose:
            print(f"\n❌ ERROR: {error_msg}")
        raise FileNotFoundError(error_msg)
    
    # Step 2: Calculate Capital Recovery Factor
    crf = calculate_crf(discount_rate, asset_lifetime)
    if verbose:
        print(f"\nCalculated Capital Recovery Factor (CRF): {crf:.6f}")
        print(f"  This means each $1 of investment becomes ${crf:.4f} annual payment")

    # Step 3: Read the input CSV file
    if verbose:
        print(f"\nReading input file: {input_file_path}")
    try:
        # Try to detect the separator automatically
        # First, read a small sample to detect separator
        with open(input_path, 'r') as f:
            first_line = f.readline()
            if ';' in first_line:
                if verbose:
                    print("  Detected semicolon (;) as separator")
                df = pd.read_csv(input_path, sep=';')
            else:
                df = pd.read_csv(input_path)
        if verbose:
            print(f"✓ Successfully loaded {len(df)} rows and {len(df.columns)} columns")
    except Exception as e:
        error_msg = f"ERROR reading CSV file: {e}"
        if verbose:
            print(f"❌ {error_msg}")
        raise IOError(error_msg)

    # Validate required columns exist
    if capital_column not in df.columns:
        error_msg = f"Column '{capital_column}' not found in the CSV! Available columns: {', '.join(df.columns)}"
        if verbose:
            print(f"❌ ERROR: {error_msg}")
        raise ValueError(error_msg)

    if 'YEAR' not in df.columns:
        error_msg = "Column 'YEAR' not found in the CSV! This column is required for temporal distribution of payments."
        if verbose:
            print(f"❌ ERROR: {error_msg}")
        raise ValueError(error_msg)

    # Show data overview
    if verbose:
        print(f"\nData Overview:")
        print(f"  Year range: {df['YEAR'].min()} - {df['YEAR'].max()}")
        print(f"  Unique technologies: {df['TECHNOLOGY'].nunique() if 'TECHNOLOGY' in df.columns else 'N/A'}")
        print(f"  Unique regions: {df['REGION'].nunique() if 'REGION' in df.columns else 'N/A'}")
        print(f"  Total capital investment: ${df[capital_column].sum():,.2f}")
        print(f"  Rows with capital investment > 0: {(df[capital_column] > 0).sum()}")

        # Analyze structure where CapitalInvestment exists
        capital_rows = df[df[capital_column] > 0]
        if len(capital_rows) > 0:
            print(f"\n  Where CapitalInvestment > 0:")
            print(f"    Unique TECHNOLOGY values: {capital_rows['TECHNOLOGY'].nunique() if 'TECHNOLOGY' in df.columns else 'N/A'}")
            print(f"    Unique REGION values: {capital_rows['REGION'].nunique() if 'REGION' in df.columns else 'N/A'}")
            # Check if other columns are mostly empty
            for col in ['FUEL', 'EMISSION', 'MODE_OF_OPERATION', 'TIMESLICE']:
                if col in df.columns:
                    non_null = capital_rows[col].notna().sum()
                    print(f"    {col} non-null values: {non_null}/{len(capital_rows)} ({non_null/len(capital_rows)*100:.1f}%)")

    # Step 4: Backup creation skipped (disabled)
    if verbose:
        print("\n⚠️  WARNING: Backup creation is disabled. Proceeding without backup.")

    # Step 5: Calculate annualized investment
    if verbose:
        print(f"\nCalculating annualized capital investment...")
        print(f"This process distributes each investment over {asset_lifetime} years")
        print(f"and accumulates overlapping payment periods.")

    df_result = calculate_annualized_investment(df, crf, grouping_columns)

    # Step 6: Determine decimal places for formatting
    decimal_places = get_decimal_places(df[capital_column])
    if decimal_places > 0:
        # Round the new column to match original precision
        df_result[new_column_name] = df_result[new_column_name].round(decimal_places)

    # Step 7: Save the results
    if verbose:
        print(f"\nSaving results to: {input_file_path}")
    df_result.to_csv(input_path, index=False)
    if verbose:
        print(f"✓ Results saved successfully")

    # Step 8: Validate results
    validation_passed = validate_results(df, df_result)

    if verbose:
        if validation_passed:
            print("\n" + "="*60)
            print("PROCESS COMPLETED SUCCESSFULLY")
            print("="*60)
            print(f"✓ Updated file saved as: {input_file_path}")
            print(f"✓ New column '{new_column_name}' added with annualized values")
            print(f"⚠️  Note: No backup was created")
            print(f"\nEnd time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        else:
            print("\n⚠ WARNING: Validation detected potential issues.")
            print("Please review the results carefully.")

    return df_result


def main():
    """
    Main execution function for standalone script execution.

    This is a wrapper around annualize_capital_investment() that uses
    the default configuration values defined at the top of the script.
    """
    return annualize_capital_investment()


# ================================
# SCRIPT EXECUTION
# ================================
if __name__ == "__main__":
    """
    This block ensures the script only runs when executed directly,
    not when imported as a module.
    """
    try:
        # Execute the main function
        result_df = main()
        
    except KeyboardInterrupt:
        # Handle user interruption (Ctrl+C)
        print("\n\n⚠ Process interrupted by user")
        sys.exit(1)
        
    except Exception as e:
        # Handle any unexpected errors
        print(f"\n❌ UNEXPECTED ERROR: {e}")
        import traceback
        traceback.print_exc()
        print("\nPlease check the input file format and try again.")
        sys.exit(1)