# -*- coding: utf-8 -*-
"""
Created on 2026
@author: Climate Lead Group, Luis Victor-Gallardo

Patches B1_Compiler.py to read 'System Parameters' sheet from
A-O_Parametrization and produce ReserveMargin.csv.

Two insertions:
  1) After the 'growth_formula' exclusion (~line 550):
     Exclude 'System Parameters' from param_sheets so the main
     tech-indexed loop does not crash on it.

  2) After the Conversions section (~line 1577, before the timing print):
     Read the 'System Parameters' sheet and build ReserveMargin rows
     following the same pattern as YearSplit.

Run with F5 in Spyder. Edit WORK_DIR below.
"""
import os

# ======================================================================
# USER CONFIGURATION
# ======================================================================
WORK_DIR = r'C:\Users\luisfernando\Desktop\OSeMOSYS\asia_ostram_refactored\t1_confection'
FILE_NAME = 'B1_Compiler.py'
# ======================================================================

file_path = os.path.join(WORK_DIR, FILE_NAME)

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# -----------------------------------------------------------------------
# PATCH 1: Exclude 'System Parameters' from the main param_sheets loop
# Insert right after the growth_formula exclusion block
# -----------------------------------------------------------------------
anchor_1 = "if 'growth_formula' in param_sheets:\n    param_sheets.remove('growth_formula')"

patch_1_addition = """if 'growth_formula' in param_sheets:
    param_sheets.remove('growth_formula')
if 'System Parameters' in param_sheets:
    param_sheets.remove('System Parameters')"""

if "param_sheets.remove('System Parameters')" in content:
    print('PATCH 1 already applied — skipping.')
else:
    if anchor_1 in content:
        content = content.replace(anchor_1, patch_1_addition)
        print('PATCH 1 applied — System Parameters excluded from param_sheets loop.')
    else:
        print('WARNING: Could not find anchor for PATCH 1. Apply manually:')
        print('  After the line: param_sheets.remove("growth_formula")')
        print('  Add: if "System Parameters" in param_sheets: param_sheets.remove("System Parameters")')

# -----------------------------------------------------------------------
# PATCH 2: Read System Parameters and build ReserveMargin rows
# Insert just before the timing print at the end of processing
# -----------------------------------------------------------------------
anchor_2 = "end_1 = time.time()"

patch_2_block = r"""#------------------------------------------------------------------------------
print('10 - System-level parameters (ReserveMargin).')
#
# Read the 'System Parameters' sheet from the Parametrization file.
# This sheet has rows like:  Parameter | Unit | 2023 | 2024 | ... | 2050
# Currently only ReserveMargin is expected.
#
if 'System Parameters' in Parametrization.sheet_names:
    sys_params_df = normalize_year_like_columns(Parametrization.parse('System Parameters'))
    accumulated_rows_sys = []
    for n in sys_params_df.index:
        this_param = sys_params_df.loc[n, 'Parameter']
        for y in range(len(time_range_vector)):
            yr_key = str(time_range_vector[y])
            this_value = sys_params_df.loc[n, yr_key]
            if pd.notna(this_value):
                accumulated_rows_sys.append({
                    'PARAMETER': this_param,
                    'Scenario': other_setup_params['Main_Scenario'],
                    'REGION': other_setup_params['Region'],
                    'YEAR': time_range_vector[y],
                    'Value': round(float(this_value), 4)
                })
    if accumulated_rows_sys:
        new_rows_sys_df = pd.DataFrame(accumulated_rows_sys)
        for param_name in new_rows_sys_df['PARAMETER'].unique():
            mask = new_rows_sys_df['PARAMETER'] == param_name
            overall_param_df_dict[param_name] = new_rows_sys_df.loc[mask].copy()
            overall_param_df_dict_ndp[param_name] = new_rows_sys_df.loc[mask].copy()
        print(f'   Loaded system parameters: {list(new_rows_sys_df["PARAMETER"].unique())}')
    else:
        print('   WARNING: System Parameters sheet found but no valid rows.')
else:
    print('   NOTE: No System Parameters sheet found — ReserveMargin not included.')
#
end_1 = time.time()"""

if "'10 - System-level parameters" in content:
    print('PATCH 2 already applied — skipping.')
else:
    if anchor_2 in content:
        content = content.replace(anchor_2, patch_2_block, 1)
        print('PATCH 2 applied — ReserveMargin reader added before timing print.')
    else:
        print('WARNING: Could not find anchor for PATCH 2. Apply manually.')

# -----------------------------------------------------------------------
# Write the patched file
# -----------------------------------------------------------------------
output_path = os.path.join(WORK_DIR, 'B1_Compiler.py')
backup_path = os.path.join(WORK_DIR, 'B1_Compiler_backup.py')

# Keep a backup
if not os.path.exists(backup_path):
    with open(file_path, 'r', encoding='utf-8') as f_orig:
        original = f_orig.read()
    with open(backup_path, 'w', encoding='utf-8') as f_bak:
        f_bak.write(original)
    print(f'Backup saved: {backup_path}')

with open(output_path, 'w', encoding='utf-8') as f:
    f.write(content)

print(f'Patched file written: {output_path}')
