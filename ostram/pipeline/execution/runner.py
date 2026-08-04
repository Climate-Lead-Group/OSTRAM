# -*- coding: utf-8 -*-
"""
Created on 2025

@author: Climate Lead Group, Andrey Salazar-Vargas
"""

import argparse
import os
import pandas as pd
import yaml
import subprocess
import sys
import platform
import shutil
import time
from datetime import date, datetime
import multiprocessing as mp
import math
from typing import List, Any
from pathlib import Path
import numpy as np

from ostram.paths import resolve_paths

from . import orchestrator as b2_orchestrator

########################################################################################
def _python_module_command(script_path, *arguments):
    path = Path(script_path).expanduser().resolve()
    project = resolve_paths()
    relative = path.relative_to(project.project_root)
    module = ".".join(relative.with_suffix("").parts)
    return [sys.executable, "-B", "-m", module, *[str(arg) for arg in arguments]]


def _run_stage_command(command):
    return subprocess.run(
        [str(token) for token in command],
        cwd=str(resolve_paths().stage_workspace("execution", create=True)),
        capture_output=True,
        text=True,
    )


def _require_stage_success(result, command, stage_name, scenario_name):
    """Propagate one required B2 child failure with its captured diagnostics."""
    if result.returncode == 0:
        return
    print(
        f"[ERROR] {stage_name} exited with code {result.returncode} "
        f"for scenario '{scenario_name}'"
    )
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)
    raise subprocess.CalledProcessError(
        result.returncode,
        [str(token) for token in command],
        output=result.stdout,
        stderr=result.stderr,
    )


def _resolve_model_input_path(value):
    """Resolve one B2 model input through the canonical project model root."""
    path = Path(str(value)).expanduser()
    resolved = (
        path
        if path.is_absolute()
        else (resolve_paths().model_root / path).resolve()
    )
    if not resolved.is_file():
        raise FileNotFoundError(f"B2 model input file not found: {resolved}")
    return resolved


def _resolve_storage_delay_model_output_path(value, scenario_name):
    """Route the generated model to the canonical scenario execution workspace."""
    project = resolve_paths()
    filename = Path(str(value)).name
    if not filename or filename in {".", ".."}:
        raise ValueError(f"invalid storage-delay model output name: {value!r}")
    resolved = (project.executables / f"{scenario_name}_0" / filename).resolve()
    for protected_root in (project.package_root, project.model_root):
        try:
            resolved.relative_to(protected_root)
        except ValueError:
            continue
        raise ValueError(
            "storage-delay model output must remain outside maintained source "
            f"directories: {resolved}"
        )
    return resolved


def ensure_env_tool_paths():
    """Expose the active Python environment's executable folders to subprocesses."""
    env_root = Path(sys.executable).resolve().parent
    candidate_dirs = [
        env_root / "Scripts",
        env_root / "Library" / "bin",
        env_root / "bin",
    ]
    current_path = os.environ.get("PATH", "")
    path_entries = current_path.split(os.pathsep) if current_path else []
    for candidate in candidate_dirs:
        candidate_str = str(candidate)
        if candidate.exists() and candidate_str not in path_entries:
            os.environ["PATH"] = candidate_str + os.pathsep + os.environ.get("PATH", "")
            path_entries.insert(0, candidate_str)


def get_env_executable(executable_name):
    """Return the full path to an executable inside the active environment when available."""
    ensure_env_tool_paths()
    env_root = Path(sys.executable).resolve().parent
    suffix = ".exe" if platform.system() == "Windows" else ""
    candidate_dirs = [
        env_root / "Scripts",
        env_root / "Library" / "bin",
        env_root / "bin",
    ]
    for candidate_dir in candidate_dirs:
        candidate = candidate_dir / f"{executable_name}{suffix}"
        if candidate.exists():
            return str(candidate)
    return executable_name


def sort_csv_files_in_folder(folder_path):
    if not os.path.isdir(folder_path):
        print(f"Invalid path: {folder_path}")
        return
    print('################################################################')
    print('Sort csv files.')
    for filename in sorted(os.listdir(folder_path)):
        if filename.endswith(".csv"):
            file_path = os.path.join(folder_path, filename)
            print(f"Processing: {filename}")
            try:
                # Read the CSV preserving the header
                df = pd.read_csv(file_path)

                # Sort using all columns
                df_sorted = df.sort_values(by=list(df.columns))

                # Overwrite the original file
                df_sorted.to_csv(file_path, index=False)
            except Exception as e:
                print(f"Error processing {filename}: {e}")

    print("✅ All files were sorted.")
    print('################################################################\n')

def process_scenario_folder(base_input_path, template_path, base_output_path, scenario_name):
    """
    Process a scenario folder: read its CSV files, align them with the template structure,
    map 'Value' to 'VALUE', exclude specific columns and save the results to the output.
    Also ensures that VALUE is int() for certain template files.
    """

    # Step 1: Define the scenario input path
    scenario_input_path = os.path.join(base_input_path, scenario_name)

    # Step 2: Skip if not a directory or is 'Default'
    if not os.path.isdir(scenario_input_path) or scenario_name == 'Default':
        return

    # Step 3: Read and clean the scenario CSVs
    scenario_files = {}
    for f in sorted(os.listdir(scenario_input_path)):
        if f.endswith('.csv'):
            df = pd.read_csv(os.path.join(scenario_input_path, f))

            # Remove unwanted columns
            df = df.drop(columns=[col for col in ['PARAMETERT', 'Scenario'] if col in df.columns])
            df = df.dropna(axis=1, how='all')

            # Rename 'Value' to 'VALUE'
            if 'Value' in df.columns:
                df = df.rename(columns={'Value': 'VALUE'})

            scenario_files[f] = df

    # Step 4: Read template files
    template_files = {
        f: pd.read_csv(os.path.join(template_path, f))
        for f in sorted(os.listdir(template_path))
        if f.endswith('.csv')
    }

    # Step 5: Create the output path
    scenario_output_path = os.path.join(base_output_path, scenario_name)
    os.makedirs(scenario_output_path, exist_ok=True)
    
    # Step 6: Fill templates with scenario data
    for template_name, template_df in template_files.items():
        output_file_path = os.path.join(scenario_output_path, template_name)
        
        if template_name in scenario_files:
            input_df = scenario_files[template_name]
            common_columns = [col for col in template_df.columns if col in input_df.columns]
            filled_df = template_df.copy()
            filled_df[common_columns] = input_df[common_columns]

            # Step 7: Convert VALUE to int if necessary
            if template_name in [
                'DAYTYPE.csv', 'DAILYTIMEBRACKET.csv', 'SEASON.csv',
                'MODE_OF_OPERATION.csv', 'YEAR.csv', 'EMISSION.csv',
                'FUEL.csv', 'REGION.csv', 'STORAGE.csv', 'TECHNOLOGY.csv',
                'TIMESLICE.csv', 'Conversionls.csv'
            ]:
                if 'VALUE' in filled_df.columns:
                    # Remove rows with NaN or empty string (including whitespace-only)
                    filled_df = filled_df[filled_df['VALUE'].notna() & (filled_df['VALUE'].astype(str).str.strip() != '')]
            
                    # Convert to int if necessary
                    if template_name in [
                        'DAYTYPE.csv', 'DAILYTIMEBRACKET.csv', 'SEASON.csv',
                        'MODE_OF_OPERATION.csv', 'YEAR.csv'
                    ]:
                        filled_df['VALUE'] = filled_df['VALUE'].astype(int)

            filled_df.to_csv(output_file_path, index=False)
        else:
            template_df.to_csv(output_file_path, index=False)
            
    folder_to_sort = os.path.join(base_output_path,scenario_name)
    sort_csv_files_in_folder(folder_to_sort)

    print(f"✅ Scenario '{scenario_name}': templates completed and saved successfully.\n")
    print('#------------------------------------------------------------------------------#')

def run_otoole_conversion(base_output_path, scenario_name, params):
    """
    Run the corrected 'otoole convert csv datafile' command for a given scenario.

    Parameters:
        base_output_path (str): Path where the scenario CSV files are stored.
        scenario_name (str): The scenario name.
        params (dict): Dictionary loaded from the YAML file with the required paths.
    """
    # Step 1: Define paths
    input_folder = os.path.join(base_output_path, scenario_name)
    scenario_exec_dir = os.path.join(HERE, params['executables'], scenario_name + '_0')
    output_file = os.path.join(scenario_exec_dir, f"{scenario_name}_0.txt")
    config_file = os.path.join(HERE, params['Miscellaneous'], params['otoole_config'])

    # Step 2: Ensure the scenario executable folder exists
    os.makedirs(scenario_exec_dir, exist_ok=True)

    # Step 3: Build the command
    otoole_exe = get_env_executable('otoole')
    command = [
        otoole_exe, 'convert', 'csv', 'datafile',
        input_folder,
        output_file,
        config_file
    ]

    print(f"Running command: {' '.join(command)}")

    # Step 4: Run the command
    result = _run_stage_command(command)
    _require_stage_success(result, command, "otoole conversion", scenario_name)

    # Step 5: Handle output
    print(f"✅ Scenario '{scenario_name}' converted successfully.\n{result.stdout}")
    print('#------------------------------------------------------------------------------#')
    return True

def run_days_in_day_type_patcher(params, scenario_name):
    """
    Runs inject_DaysInDayType.py against the preprocessed datafile to fix
    the empty DaysInDayType block (which would otherwise default to 7,
    breaking storage cycling vs energy balance scaling).
    Must run AFTER run_preprocessing_script, BEFORE the solve.
    """
    # Anchor to B2's own directory so it works regardless of how B2 is invoked
    script_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'patches',
        'days_in_day_type.py',
    )
    target_file = os.path.join(
        params['executables'],
        scenario_name + '_0',
        f"{params['preprocess_data_name']}{scenario_name}_0.txt",
    )
    command = _python_module_command(script_path, target_file)
    print(f"Patching DaysInDayType for '{scenario_name}_0':")
    print(' '.join(command))
    result = _run_stage_command(command)
    _require_stage_success(result, command, "DaysInDayType patcher", scenario_name)
    print(result.stdout)
    print('#------------------------------------------------------------------------------#')

def run_strip_storage_patcher(params, scenario_name):
    """
    OPTIONAL diagnostic step: strips selected storage facilities (and their
    feeding PWR techs) from the preprocessed datafile, writing a SIBLING file
    (e.g. Pre_processed_BAU_0_NoStorage.txt). Original .txt is never modified.

    Controlled by params['strip_storage_active'] (default False = no-op).

    YAML keys consumed:
        strip_storage_active:  bool   -- master switch (default False)
        strip_storage_mode:    str    -- "tech" | "class" | "all"  (default "all")
        strip_storage_targets: list   -- facility names (tech) or prefixes (class)
        strip_storage_suffix:  str    -- filename suffix (default "NoStorage")

    When active, main_executer redirects data_file/output_file to use the
    suffixed sibling, so the solver builds and solves the patched LP and
    writes outputs alongside the originals.
    """
    if not params.get('strip_storage_active', False):
        return  # No-op when disabled

    mode = params.get('strip_storage_mode', 'all')
    targets = params.get('strip_storage_targets') or []
    suffix = params.get('strip_storage_suffix', 'NoStorage')

    # Anchor strip_storage.py to B2's own directory (same pattern as DaysInDayType).
    script_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'patches',
        'strip_storage.py',
    )

    base = f"{params['preprocess_data_name']}{scenario_name}_0"
    in_file = os.path.join(params['executables'], scenario_name + '_0', f"{base}.txt")
    out_file = os.path.join(params['executables'], scenario_name + '_0', f"{base}_{suffix}.txt")

    command = _python_module_command(
        script_path, in_file, '-o', out_file, '--mode', mode
    )
    if mode != 'all' and targets:
        command += ['--targets'] + list(targets)

    print(f"Stripping storage for '{scenario_name}_0' (mode={mode}, suffix={suffix}):")
    print(' '.join(command))
    result = _run_stage_command(command)
    _require_stage_success(result, command, "strip_storage patcher", scenario_name)
    print(result.stdout)
    print('#------------------------------------------------------------------------------#')

def run_storage_delay_patcher(params, scenario_name):
    """
    OPTIONAL storage-delay step: keeps storage in the model but blocks storage
    builds for the first N model years, then reopens the linked PWRLDS*/PWRSDS*
    caps in later years. Writes SIBLING files (datafile + patched OSeMOSYS
    model); originals are never modified.

    Controlled by params['storage_delay_active'] (default False = no-op).

    YAML keys consumed:
        storage_delay_active:           bool  -- master switch (default False)
        storage_delay_first_n_years:    int   -- years to block (default 5)
        storage_delay_storage_prefixes: list  -- e.g. ["SDS", "LDS"]
        storage_delay_storages:         list  -- exact storage names (optional, overrides prefixes)
        storage_delay_allowed_value:    str   -- PWR cap value in open years (default "-1")
        storage_delay_suffix:           str   -- chained filename suffix (default "StorageDelayN5")
        storage_delay_model_input:      str   -- source OSeMOSYS model file (default params['osemosys_model'])
        storage_delay_model_output:     str   -- patched model file written next to B2 (default "osemosys_fast_preprocessed_storage_delay.txt")

    Mutually exclusive with strip_storage. When storage_delay_active is True,
    main_executer's __main__ already disables strip_storage and switches the
    solver model to the patched output produced here.
    """
    if not params.get('storage_delay_active', False):
        return  # No-op when disabled

    suffix = params.get('storage_delay_suffix', 'StorageDelayN5')
    first_n_years = params.get('storage_delay_first_n_years', 5)
    allowed_value = params.get('storage_delay_allowed_value', '-1')
    storage_prefixes = params.get('storage_delay_storage_prefixes', ['SDS', 'LDS'])
    exact_storages = params.get('storage_delay_storages', [])

    here = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(here, 'patches', 'storage_delay.py')

    base = f"{params['preprocess_data_name']}{scenario_name}_0"
    in_file = os.path.join(params['executables'], scenario_name + '_0', f"{base}.txt")
    out_file = os.path.join(params['executables'], scenario_name + '_0', f"{base}_{suffix}.txt")

    model_input = str(
        _resolve_model_input_path(
            params.get('storage_delay_model_input', params['osemosys_model'])
        )
    )
    model_output = str(
        _resolve_storage_delay_model_output_path(
            params.get(
                'storage_delay_model_output',
                'osemosys_fast_preprocessed_storage_delay.txt',
            ),
            scenario_name,
        )
    )

    command = _python_module_command(
        script_path,
        in_file,
        '-o',
        out_file,
        '--model-input',
        model_input,
        '--model-output',
        model_output,
        '--first-n-years',
        str(first_n_years),
        '--allowed-value',
        str(allowed_value),
    )
    if exact_storages:
        command += ['--storages'] + list(exact_storages)
    elif storage_prefixes:
        command += ['--storage-prefixes'] + list(storage_prefixes)

    print(f"Applying storage-delay patch for '{scenario_name}_0' (N={first_n_years}, suffix={suffix}):")
    print(' '.join(command))
    result = _run_stage_command(command)
    _require_stage_success(result, command, "storage_delay patcher", scenario_name)
    params['storage_delay_model_output'] = model_output
    params['osemosys_model'] = model_output
    print(result.stdout)
    if result.stderr:
        print(result.stderr)
    print('#------------------------------------------------------------------------------#')

def run_open_pwrbck_patcher(params, scenario_name):
    """
    OPTIONAL diagnostic step: opens PWRBCK* (backstop) caps in
    TotalAnnualMaxCapacity and TotalAnnualMaxCapacityInvestment by rewriting
    any 0-value cells to params['open_pwrbck_value'] (default 9999). Reads
    from the strip_storage output if that step is active, otherwise from the
    vanilla preprocessed datafile. Writes a sibling file with the OpenBCK
    suffix chained on; original is never modified.

    Controlled by params['open_pwrbck_active'] (default False = no-op).

    YAML keys consumed:
        open_pwrbck_active:  bool  -- master switch (default False)
        open_pwrbck_value:   int   -- replacement for 0 cells (default 9999)
        open_pwrbck_pattern: str   -- tech-name substring (default "PWRBCK")
        open_pwrbck_suffix:  str   -- filename suffix to chain (default "OpenBCK")

    When active, main_executer chains OpenBCK on top of any active strip
    suffix, so the solver builds and solves the patched LP and writes outputs
    alongside the originals.
    """
    if not params.get('open_pwrbck_active', False):
        return  # No-op when disabled

    value   = params.get('open_pwrbck_value',   9999)
    pattern = params.get('open_pwrbck_pattern', 'PWRBCK')
    suffix  = params.get('open_pwrbck_suffix',  'OpenBCK')

    # Anchor open_pwrbck_caps.py to B2's own directory (same pattern as the
    # strip and DaysInDayType patchers).
    script_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'patches',
        'open_pwrbck_caps.py',
    )

    base = f"{params['preprocess_data_name']}{scenario_name}_0"
    # Input is the previous patcher's output (storage_delay or strip) if active;
    # otherwise the vanilla file.
    _chain = []
    if params.get('storage_delay_active', False):
        _chain.append(params.get('storage_delay_suffix', 'StorageDelayN5'))
    if params.get('strip_storage_active', False):
        _chain.append(params.get('strip_storage_suffix', 'NoStorage'))
    in_base = f"{base}_{'_'.join(_chain)}" if _chain else base
    out_base = f"{in_base}_{suffix}"

    in_file  = os.path.join(params['executables'], scenario_name + '_0', f"{in_base}.txt")
    out_file = os.path.join(params['executables'], scenario_name + '_0', f"{out_base}.txt")

    command = _python_module_command(
        script_path, in_file, '-o', out_file,
        '--pattern', pattern, '--value', str(value)
    )

    print(f"Opening PWRBCK caps for '{scenario_name}_0' (pattern={pattern}, value={value}):")
    print(' '.join(command))
    result = _run_stage_command(command)
    _require_stage_success(result, command, "open_pwrbck patcher", scenario_name)
    print(result.stdout)
    print('#------------------------------------------------------------------------------#')

def run_reserve_margin_repair_patcher(params, scenario_name):
    """
    OPTIONAL diagnostic/final-ish step: patches ReserveMarginTagTechnology and
    opens selected firm capacity caps in the preprocessed datafile.

    Controlled by params['reserve_margin_repair_active'] (default False = no-op).

    YAML keys consumed:
        reserve_margin_repair_active: bool  -- master switch (default False)
        reserve_margin_repair_suffix: str   -- chained filename suffix (default "RMRepair")
        reserve_margin_backstop_credit: num -- PWRBCK reserve credit (default 1.0)
        reserve_margin_ccs_credit: num      -- PWRCCS reserve credit (default 0.9)
        reserve_margin_open_capacity_value: num -- cap value for selected techs (default 9999)
        reserve_margin_open_capacity_prefixes: list -- default ["PWRPET", "PWROIL", "PWRNGS"]
        reserve_margin_patch_backstop: bool -- set PWRBCK tags (default True)
        reserve_margin_patch_ccs: bool      -- set PWRCCS tags (default True)
        reserve_margin_open_capacity: bool  -- open selected caps (default True)

    This chains after strip_storage/open_pwrbck when those patchers are active.
    """
    if not params.get('reserve_margin_repair_active', False):
        return  # No-op when disabled

    suffix = params.get('reserve_margin_repair_suffix', 'RMRepair')
    script_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'patches',
        'reserve_margin_repair.py',
    )

    base = f"{params['preprocess_data_name']}{scenario_name}_0"
    chain_parts = []
    if params.get('storage_delay_active', False):
        chain_parts.append(params.get('storage_delay_suffix', 'StorageDelayN5'))
    if params.get('strip_storage_active', False):
        chain_parts.append(params.get('strip_storage_suffix', 'NoStorage'))
    if params.get('open_pwrbck_active', False):
        chain_parts.append(params.get('open_pwrbck_suffix', 'OpenBCK'))

    in_base = f"{base}_{'_'.join(chain_parts)}" if chain_parts else base
    out_base = f"{in_base}_{suffix}"

    in_file = os.path.join(params['executables'], scenario_name + '_0', f"{in_base}.txt")
    out_file = os.path.join(params['executables'], scenario_name + '_0', f"{out_base}.txt")

    command = _python_module_command(
        script_path,
        in_file,
        '-o',
        out_file,
        '--backstop-credit',
        str(params.get('reserve_margin_backstop_credit', 1.0)),
        '--ccs-credit',
        str(params.get('reserve_margin_ccs_credit', 0.9)),
        '--open-capacity-value',
        str(params.get('reserve_margin_open_capacity_value', 9999)),
    )

    open_prefixes = params.get(
        'reserve_margin_open_capacity_prefixes',
        ['PWRPET', 'PWROIL', 'PWRNGS'],
    )
    command += ['--open-capacity-prefixes'] + list(open_prefixes)

    if not params.get('reserve_margin_patch_backstop', True):
        command.append('--skip-backstop-credit')
    if not params.get('reserve_margin_patch_ccs', True):
        command.append('--skip-ccs-credit')
    if not params.get('reserve_margin_open_capacity', True):
        command.append('--skip-capacity-opening')

    print(f"Repairing reserve margin data for '{scenario_name}_0' (suffix={suffix}):")
    print(' '.join(command))
    result = _run_stage_command(command)
    _require_stage_success(
        result, command, "reserve_margin_repair patcher", scenario_name
    )
    print(result.stdout)
    print('#------------------------------------------------------------------------------#')

def run_reserve_margin_xlsx_patcher(params, scenario_name):
    """
    OPTIONAL careful reserve-margin repair step using an XLSX fallback workbook.

    Controlled by params['reserve_margin_xlsx_active'] (default False = no-op).

    YAML keys consumed:
        reserve_margin_xlsx_active: bool  -- master switch (default False)
        reserve_margin_xlsx_suffix: str   -- chained suffix (default "RMCarefulXLSX")
        reserve_margin_xlsx_workbook: str -- workbook path, relative to this B2 file if not absolute
        reserve_margin_xlsx_sheet: str    -- optional worksheet name
        reserve_margin_xlsx_backstop_credit: num -- PWRBCK reserve credit (default 1.0)
        reserve_margin_xlsx_ccs_credit: num      -- PWRCCS reserve credit (default 0.9)
        reserve_margin_xlsx_target_prefixes: list -- default ["PWRPET", "PWROIL", "PWRNGS"]
        reserve_margin_xlsx_sentinel_values: list -- default [0, 9999]

    This chains after strip_storage/open_pwrbck and also after the older
    reserve_margin_repair patch if that older patch is active.
    """
    if not params.get('reserve_margin_xlsx_active', False):
        return

    suffix = params.get('reserve_margin_xlsx_suffix', 'RMCarefulXLSX')
    here = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(here, 'patches', 'reserve_margin_repair_xlsx.py')

    workbook = params.get(
        'reserve_margin_xlsx_workbook',
        'firm_capacity_fallbacks_by_cr_0p5.xlsx',
    )
    if not os.path.isabs(workbook):
        workbook = os.path.join(here, workbook)

    base = f"{params['preprocess_data_name']}{scenario_name}_0"
    chain_parts = []
    if params.get('storage_delay_active', False):
        chain_parts.append(params.get('storage_delay_suffix', 'StorageDelayN5'))
    if params.get('strip_storage_active', False):
        chain_parts.append(params.get('strip_storage_suffix', 'NoStorage'))
    if params.get('open_pwrbck_active', False):
        chain_parts.append(params.get('open_pwrbck_suffix', 'OpenBCK'))
    if params.get('reserve_margin_repair_active', False):
        chain_parts.append(params.get('reserve_margin_repair_suffix', 'RMRepair'))

    in_base = f"{base}_{'_'.join(chain_parts)}" if chain_parts else base
    out_base = f"{in_base}_{suffix}"

    in_file = os.path.join(params['executables'], scenario_name + '_0', f"{in_base}.txt")
    out_file = os.path.join(params['executables'], scenario_name + '_0', f"{out_base}.txt")
    warnings_file = os.path.join(params['executables'], scenario_name + '_0', f"{out_base}.warnings.txt")

    command = _python_module_command(
        script_path,
        in_file,
        '-o',
        out_file,
        '--fallback-xlsx',
        workbook,
        '--backstop-credit',
        str(params.get('reserve_margin_xlsx_backstop_credit', 1.0)),
        '--ccs-credit',
        str(params.get('reserve_margin_xlsx_ccs_credit', 0.9)),
        '--warnings-file',
        warnings_file,
    )

    xlsx_sheet = params.get('reserve_margin_xlsx_sheet')
    if xlsx_sheet:
        command += ['--xlsx-sheet', str(xlsx_sheet)]

    target_prefixes = params.get(
        'reserve_margin_xlsx_target_prefixes',
        ['PWRPET', 'PWROIL', 'PWRNGS'],
    )
    command += ['--target-prefixes'] + list(target_prefixes)

    sentinel_values = params.get('reserve_margin_xlsx_sentinel_values', [0, 9999])
    command += ['--sentinel-values'] + [str(value) for value in sentinel_values]

    if not params.get('reserve_margin_xlsx_patch_backstop', True):
        command.append('--skip-backstop-credit')
    if not params.get('reserve_margin_xlsx_patch_ccs', True):
        command.append('--skip-ccs-credit')

    print(f"Repairing reserve margin data from XLSX for '{scenario_name}_0' (suffix={suffix}):")
    print(' '.join(command))
    result = _run_stage_command(command)
    _require_stage_success(
        result, command, "reserve_margin_xlsx patcher", scenario_name
    )
    print(result.stdout)
    if result.stderr:
        print(result.stderr)
    print('#------------------------------------------------------------------------------#')

def run_preprocessing_script(params, scenario_name):
    """
    Run the Python preprocessing script specified in the YAML parameters file for a given scenario.

    Parameters:
        params (dict): Parameters loaded from the YAML file.
        scenario_name (str): The name of the scenario to preprocess.
    """
    # Step 1: Define paths
    script_path = os.path.join(params['Miscellaneous'], params['preprocess_data'])
    input_file = os.path.join(params['executables'], scenario_name + '_0', f"{scenario_name}_0.txt")
    output_file = os.path.join(params['executables'], scenario_name + '_0', f"{params['preprocess_data_name']}{scenario_name}_0.txt")

    # Step 2: Build command
    command = _python_module_command(script_path, input_file, output_file)

    print(f"Running preprocessing script for scenario '{scenario_name}_0':")
    print(' '.join(command))

    # Step 3: Run the script
    result = _run_stage_command(command)
    _require_stage_success(result, command, "preprocessing", scenario_name)

    # Step 4: Output result
    print(f"✅ Preprocessing completed for scenario '{scenario_name}':\n{result.stdout}")
    print('#------------------------------------------------------------------------------#')

def check_enviro_variables(solver_command):
    ensure_env_tool_paths()
    # Determine the command according to the operating system
    command = 'where' if platform.system() == 'Windows' else 'which'

    # Run the appropriate command
    where_solver = subprocess.run([command, solver_command], capture_output=True, text=True)
    paths = where_solver.stdout.splitlines()

    if paths:  # Ensure at least one path was found
        path_solver = paths[0]

        # Check whether the path is already in the PATH environment variable
        if path_solver not in os.environ["PATH"]:
            # If it is not in PATH, add it
            os.environ["PATH"] += os.pathsep + path_solver
            print("Path added:", path_solver)
    else:
        print(f"'{solver_command}' was not found on the system.")
    #

def main_executer(params, scenario_name, HERE):
    execution_dependencies = b2_orchestrator.ScenarioExecutionDependencies(
        run_process=subprocess.run,
        check_environment=check_enviro_variables,
        get_executable=get_env_executable,
        path_exists=os.path.exists,
        remove_file=os.remove,
        python_executable=sys.executable,
    )
    return b2_orchestrator.execute_scenario(
        params,
        scenario_name,
        HERE,
        execution_dependencies,
    )
def delete_files(file, data_file, solver):
    # Delete files
    if file and os.path.exists(file):
        shutil.os.remove(file)
    if data_file and os.path.exists(data_file):
        shutil.os.remove(data_file)
    
    # Check whether the .sol file exists and is empty
    log_file = file.replace('.sol', '.log')
    if os.path.exists(log_file) and os.path.getsize(log_file) == 0:
        if os.path.exists(log_file):
            os.remove(log_file)
    
    if solver == 'glpk':
        glp_file = file.replace('sol', 'glp')
        if os.path.exists(glp_file):
            shutil.os.remove(glp_file)
    else:
        lp_file = file.replace('sol', 'lp')
        if os.path.exists(lp_file):
            shutil.os.remove(lp_file)
    
    # Delete log files when the solver is 'cplex' and del_files is True
    if solver == 'cplex':
        for filename in ['cplex.log', 'clone1.log', 'clone2.log']:
            if os.path.exists(filename):
                os.remove(filename)

    # Delete log files when the solver is 'gurobi' and del_files is True
    if solver == 'gurobi':
        if os.path.exists('gurobi.log'):
            os.remove('gurobi.log')

def read_csv_files(input_dir):
    """Read all CSV files in the given directory and return a dictionary of DataFrames."""
    data_dict = {}
    for filename in sorted(os.listdir(input_dir)):
        if filename.endswith(".csv"):
            file_path = os.path.join(input_dir, filename)
            df = pd.read_csv(file_path)
            key = os.path.splitext(filename)[0]
            data_dict[key] = df
    return data_dict

def generate_combined_input_file(input_folder, output_folder, scenario_name):
    """
    Read CSVs from input_folder, filter metadata keys, rename VALUE columns by key,
    concatenate all non-empty DataFrames, sort columns and save the result to a CSV file.
    """
    keys_sets_delete = ['REGION', 'YEAR', 'TECHNOLOGY', 'FUEL', 'EMISSION', 'MODE_OF_OPERATION',
                        'TIMESLICE', 'STORAGE', 'SEASON', 'DAYTYPE', 'DAILYTIMEBRACKET']

    inputs_dataframes = []
    print(input_folder)
    print(sorted(os.listdir(input_folder)))
    for filename in sorted(os.listdir(input_folder)):
        if not filename.endswith(".csv"):
            continue
        key = filename.replace(".csv", "")
        if key in keys_sets_delete:
            continue
        path = os.path.join(input_folder, filename)
        df = pd.read_csv(path)
        if df.empty or 'VALUE' not in df.columns:
            continue
        df = df.rename(columns={'VALUE': key})
        inputs_dataframes.append(df)

    if not inputs_dataframes:
        print("[Warning] No valid dataframes found to concatenate.")
        return None, None

    # Concatenate all non-empty dataframes
    inputs_data = pd.concat(inputs_dataframes, ignore_index=True, sort=True)  # Sort for deterministic column order

    # Reorder columns
    present_keys = [col for col in keys_sets_delete if col in inputs_data.columns]
    other_columns = sorted([col for col in inputs_data.columns if col not in present_keys])
    inputs_data = inputs_data[present_keys + other_columns]

    # Save to CSV
    os.makedirs(output_folder, exist_ok=True)
    output_path = os.path.join(output_folder, f"{scenario_name}_Input.csv")
    inputs_data.to_csv(output_path, index=False)

    print(f'✅ Inputs concatenated to {scenario_name}_Input.csv successfully.')
    print('\n#------------------------------------------------------------------------------#')

    return output_path, inputs_data.head()


def export_root_datafile(here, params, scenario_name, export_name=None):
    """
    Copy the preprocessed main-scenario datafile to the project root so the
    user has a single easy-to-find model datafile.

    When patchers (storage_delay, strip_storage, open_pwrbck, reserve_margin_*)
    are active, exports the final patched sibling — not the vanilla preprocessed
    file — so the root datafile matches what the solver actually consumed.
    """
    if export_name is None:
        if params.get('storage_delay_active', False):
            export_name = params.get('storage_delay_root_datafile', 'OSTRAM_data_storage_delay.txt')
        else:
            export_name = 'OSTRAM_data.txt'

    repo_root = Path(here).parent
    base = f"{params['preprocess_data_name']}{scenario_name}_0"
    chain_parts = []
    if params.get('storage_delay_active', False):
        chain_parts.append(params.get('storage_delay_suffix', 'StorageDelayN5'))
    if params.get('strip_storage_active', False):
        chain_parts.append(params.get('strip_storage_suffix', 'NoStorage'))
    if params.get('open_pwrbck_active', False):
        chain_parts.append(params.get('open_pwrbck_suffix', 'OpenBCK'))
    if params.get('reserve_margin_repair_active', False):
        chain_parts.append(params.get('reserve_margin_repair_suffix', 'RMRepair'))
    if params.get('reserve_margin_xlsx_active', False):
        chain_parts.append(params.get('reserve_margin_xlsx_suffix', 'RMCarefulXLSX'))

    source_name = f"{base}_{'_'.join(chain_parts)}.txt" if chain_parts else f"{base}.txt"
    source_path = (
        Path(here)
        / params['executables']
        / f"{scenario_name}_0"
        / source_name
    )
    target_path = repo_root / export_name

    if not source_path.exists():
        print(f"[WARN] Root datafile export skipped because source was not found: {source_path}")
        return None

    shutil.copy2(source_path, target_path)
    print(f"✅ Datafile exported to repository root: {target_path}")
    print('#------------------------------------------------------------------------------#')
    return target_path


def active_output_csv_candidates(params, scenario_future_name):
    """
    Return output CSV names in the same suffix order used by main_executer.

    The solver/otoole path can become, for example:
      Pre_processed_BAU_0_NoStorage_OpenBCK_RMCarefulXLSX_output.csv

    The final scenario concatenator used to look only for:
      Pre_processed_BAU_0_Output.csv

    Keep the active chained name first, with legacy fallbacks after it.
    """
    base = f"{params['preprocess_data_name']}{scenario_future_name}"
    chain_parts = []

    if params.get('storage_delay_active', False):
        chain_parts.append(params.get('storage_delay_suffix', 'StorageDelayN5'))
    if params.get('strip_storage_active', False):
        chain_parts.append(params.get('strip_storage_suffix', 'NoStorage'))
    if params.get('open_pwrbck_active', False):
        chain_parts.append(params.get('open_pwrbck_suffix', 'OpenBCK'))
    if params.get('reserve_margin_repair_active', False):
        chain_parts.append(params.get('reserve_margin_repair_suffix', 'RMRepair'))
    if params.get('reserve_margin_xlsx_active', False):
        chain_parts.append(params.get('reserve_margin_xlsx_suffix', 'RMCarefulXLSX'))

    candidates = []
    if chain_parts:
        candidates.append(f"{base}_{'_'.join(chain_parts)}{params['output_files']}.csv")

    candidates.extend([
        f"{base}{params['output_files']}.csv",
        f"{base}_Output.csv",
    ])

    return candidates



def concatenate_all_scenarios(HERE, params):
    """
    Iterate over all scenario folders in `base_input_path` (excluding 'Default'),
    read *_Input.csv and *_Output.csv files, add scenario metadata columns, concatenate them
    into single CSV files for inputs, outputs and combined, and return their paths.

    Args:
        params (dict):
          - executables (str): Path to the base directory containing the scenario folders.
          - prefix_final_files (str): Folder/path where to save the results.
          - inputs_file (str): Base name for the inputs CSV.
          - outputs_file (str): Base name for the outputs CSV.
          - combined_file (str, optional): Base name for the combined inputs+outputs CSV.
    Returns:
        tuple: (input_csv_path, output_csv_path, combined_csv_path)
    """
    # Metadata columns that we move to the front
    keys_sets_delete = [
        'REGION','YEAR','TECHNOLOGY','FUEL','EMISSION','MODE_OF_OPERATION',
        'TIMESLICE','STORAGE','SEASON','DAYTYPE','DAILYTIMEBRACKET'
    ]

    combined_inputs = []
    combined_outputs = []
    combined_inputs_outputs = []
    base_input_path = params['executables']

    for scenario_future_name in sorted(os.listdir(base_input_path)):
        if scenario_future_name.lower() in ['default', '__pycache__', 'local_dataset_creator_0.py']:
            continue

        scenario_path = os.path.join(HERE, base_input_path, scenario_future_name)
        parts = scenario_future_name.rsplit("_", 1)
        scenario = parts[0]
        future = parts[1]

        input_file = os.path.join(scenario_path, f"{scenario_future_name}_Input.csv")
        output_file = None
        for output_name in active_output_csv_candidates(params, scenario_future_name):
            candidate = os.path.join(scenario_path, output_name)
            if os.path.exists(candidate):
                output_file = candidate
                break

        if os.path.exists(input_file):
            df_in = pd.read_csv(input_file, low_memory=False)
            df_in.insert(0, "Future", future)
            df_in.insert(1, "Scenario", scenario)
            combined_inputs.append(df_in)
            combined_inputs_outputs.append(df_in)

        if output_file and os.path.exists(output_file):
            df_out = pd.read_csv(output_file, low_memory=False)
            df_out.insert(0, "Future", future)
            df_out.insert(1, "Scenario", scenario)
            combined_outputs.append(df_out)
            combined_inputs_outputs.append(df_out)

    # Concatenate inputs and outputs separately
    df_inputs_all = pd.concat(combined_inputs, ignore_index=True) if combined_inputs else pd.DataFrame()
    df_outputs_all = pd.concat(combined_outputs, ignore_index=True) if combined_outputs else pd.DataFrame()
    # df_inputs_outputs_all = pd.concat(combined_inputs_outputs, ignore_index=True) if combined_inputs_outputs else pd.DataFrame()
    # df_list = []
    # df_list.append(combined_inputs)
    # df_list.append(combined_outputs)
    df_inputs_outputs_all = pd.concat([df_inputs_all,df_outputs_all], ignore_index=True, sort=True)  # Sort for deterministic column order
    

    # Function to reorder columns: metadata first, then alphabetical
    def reorder_columns(df):
        front = ['Future','Scenario'] + [c for c in keys_sets_delete if c in df.columns]
        rest = sorted([c for c in df.columns if c not in front])
        return df[front + rest]

    today = date.today().isoformat()  # 'YYYY-MM-DD'

    # 1) Save inputs
    if not df_inputs_all.empty:
        df_inputs_all = reorder_columns(df_inputs_all)
        # Sort rows for deterministic output
        sort_cols = [c for c in ['Future', 'Scenario', 'REGION', 'TECHNOLOGY', 'YEAR'] if c in df_inputs_all.columns]
        if sort_cols:
            df_inputs_all = df_inputs_all.sort_values(by=sort_cols).reset_index(drop=True)
        path_in = os.path.join(HERE,params['prefix_final_files'] + params['inputs_file'])
        df_inputs_all.to_csv(path_in, index=False)
        dated = path_in.replace('.csv', f'_{today}.csv')
        df_inputs_all.to_csv(dated, index=False)
    else:
        path_in = None

    # 2) Save outputs
    if not df_outputs_all.empty:
        df_outputs_all = reorder_columns(df_outputs_all)
        # Sort rows for deterministic output
        sort_cols = [c for c in ['Future', 'Scenario', 'REGION', 'TECHNOLOGY', 'YEAR'] if c in df_outputs_all.columns]
        if sort_cols:
            df_outputs_all = df_outputs_all.sort_values(by=sort_cols).reset_index(drop=True)
        path_out = os.path.join(HERE,params['prefix_final_files'] + params['outputs_file'])
        df_outputs_all.to_csv(path_out, index=False)
        dated = path_out.replace('.csv', f'_{today}.csv')
        df_outputs_all.to_csv(dated, index=False)
    else:
        path_out = None

    # 3) Again, combine both DataFrames into a single one and save it
    combined_name = params.get('combined_file', 'Combined_Inputs_Outputs.csv')
    if not df_inputs_outputs_all.empty and not df_outputs_all.empty:
        # df_combined = pd.concat([df_inputs_all, df_outputs_all],
        #                         ignore_index=True, sort=False)
        df_combined = reorder_columns(df_inputs_outputs_all)
        # Sort rows for deterministic output
        sort_cols = [c for c in ['Future', 'Scenario', 'REGION', 'TECHNOLOGY', 'YEAR'] if c in df_combined.columns]
        if sort_cols:
            df_combined = df_combined.sort_values(by=sort_cols).reset_index(drop=True)
        
        
        #########################################################################################
        # Calculate AccumulatedTotalAnnualMinCapacityInvestment
        # Must group by (Future, Scenario, TECHNOLOGY) and accumulate within each group
        if "TotalAnnualMinCapacityInvestment" in df_combined.columns:
            df = df_combined.copy()

            # Initialize the accumulated column with NaN
            df['AccumulatedTotalAnnualMinCapacityInvestment'] = np.nan

            # Define grouping columns (exclude YEAR since we accumulate over years)
            group_cols = ['Future', 'Scenario', 'TECHNOLOGY']
            group_cols = [c for c in group_cols if c in df.columns]

            if group_cols:
                # Sort by group columns + YEAR to ensure correct order for cumsum
                sort_cols = group_cols + ['YEAR']
                df = df.sort_values(by=sort_cols).reset_index(drop=True)

                # Calculate cumulative sum within each group
                # Only for rows that have a value in TotalAnnualMinCapacityInvestment
                mask = df['TotalAnnualMinCapacityInvestment'].notna()
                df.loc[mask, 'AccumulatedTotalAnnualMinCapacityInvestment'] = (
                    df.loc[mask]
                    .groupby(group_cols, sort=False)['TotalAnnualMinCapacityInvestment']
                    .cumsum()
                )
            else:
                # Fallback: if there are no group columns, do a simple cumsum
                mask = df['TotalAnnualMinCapacityInvestment'].notna()
                df.loc[mask, 'AccumulatedTotalAnnualMinCapacityInvestment'] = (
                    df.loc[mask, 'TotalAnnualMinCapacityInvestment'].cumsum()
                )

            df_combined = df
        #########################################################################################
        
        
        path_comb = os.path.join(HERE,params['prefix_final_files'] + combined_name)
        df_combined.to_csv(path_comb, index=False)
        # Note: the dated copy with annualized data will be created after annualization (if enabled)
    else:
        path_comb = None

    return path_in, path_out, path_comb






def chunk_scenarios(
    scenarios: List[Any],
    max_x_per_iter: int,
) -> List[List[Any]]:
    """
    Split the input list ``scenarios`` into chunks of size ``max_x_per_iter``.

    Parameters
    ----------
    scenarios : List[Any]
        The list containing all scenario values.
    max_x_per_iter : int
        Maximum number of elements allowed in each chunk.

    Returns
    -------
    List[List[Any]]
        A list where each element is a sub-list of ``scenarios`` with length
        up to ``max_x_per_iter``.
    """
    if max_x_per_iter <= 0:
        raise ValueError("max_x_per_iter must be a positive integer")

    # Build the chunks using slicing in a comprehension
    scenarios_list_max_per_iter: List[List[Any]] = [
        scenarios[i : i + max_x_per_iter]  # noqa: E203 (spacing around :)
        for i in range(0, len(scenarios), max_x_per_iter)
    ]
    return scenarios_list_max_per_iter

########################################################################################
def _set_here(value):
    global HERE
    HERE = value


def _load_annualizer():
    from .annualization import annualize_capital_investment

    return annualize_capital_investment


def _orchestration_dependencies():
    return b2_orchestrator.B2Dependencies(
        process_scenario_folder=process_scenario_folder,
        run_otoole_conversion=run_otoole_conversion,
        run_preprocessing_script=run_preprocessing_script,
        run_days_in_day_type_patcher=run_days_in_day_type_patcher,
        run_storage_delay_patcher=run_storage_delay_patcher,
        run_strip_storage_patcher=run_strip_storage_patcher,
        run_open_pwrbck_patcher=run_open_pwrbck_patcher,
        run_reserve_margin_repair_patcher=run_reserve_margin_repair_patcher,
        run_reserve_margin_xlsx_patcher=run_reserve_margin_xlsx_patcher,
        generate_combined_input_file=generate_combined_input_file,
        export_root_datafile=export_root_datafile,
        main_executer=main_executer,
        chunk_scenarios=chunk_scenarios,
        delete_files=delete_files,
        concatenate_all_scenarios=concatenate_all_scenarios,
        load_annualizer=_load_annualizer,
        yaml_safe_load=yaml.safe_load,
        mp_module=mp,
    )


def main(argv=None):
    return b2_orchestrator.orchestrate_b2(
        _orchestration_dependencies,
        argv=argv,
        set_here=_set_here,
    )


if __name__ == "__main__":
    main()
