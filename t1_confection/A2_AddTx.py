# -*- coding: utf-8 -*-
"""
Created on 2025

@author: ClimateLeadGroup, Javier Monge-Matamoros
"""

import argparse
import sys
import os
import yaml
import pandas as pd
from typing import List
from pathlib import Path
from Z_AUX_config_loader import get_renewable_fuels, get_iso_country_map

# Country and technology mappings from centralized config
RENEWABLE_FUELS = get_renewable_fuels()
iso_country_map = get_iso_country_map()

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
def load_country_region_pairs(yaml_path):
    """Return list of (country, region) tuples from the 'countries' key in YAML.

    3‑letter codes ⇒ region='XX'
    5‑letter codes ⇒ last 2 letters are region
    """
    with open(yaml_path, 'r', encoding='utf-8') as fh:
        data = yaml.safe_load(fh)

    codes = data.get('countries', [])
    if not isinstance(codes, list):
        print("⚠️  'countries' key in YAML is not a list", file=sys.stderr)
        return []

    pairs = []
    for code in codes:
        if not isinstance(code, str):
            continue
        code = code.strip().upper()
        if len(code) == 3:
            pairs.append((code, "XX"))
        elif len(code) == 5:
            pairs.append((code[:3], code[3:]))
        else:
            print(f"⚠️  Skipping unrecognised code '{code}'", file=sys.stderr)
    # Remove duplicates while preserving order
    seen = set()
    ordered = []
    for c,r in pairs:
        if (c,r) not in seen:
            ordered.append((c,r))
            seen.add((c,r))
    return ordered

def parse_pwr_code(tech_code):  
    remainder = tech_code[3:]
    fuel     = remainder[:3]
    country  = remainder[3:6]
    region   = remainder[6:8] if len(remainder) >= 8 else "XX"
    return fuel, country, region

def ensure_columns(df, cols):
    """Make sure DataFrame contains each column in *cols* (creates if absent)."""
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    return df

# ---------------------------------------------------------------------------
# 1. Process A-O_AR_Model_Base_Year.xlsx
# ---------------------------------------------------------------------------
def process_base_year(path, pairs):
    print(f"Processing '{path}' …")
    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        # ---------- Secondary sheet ----------
        sec = pd.read_excel(path, sheet_name='Secondary', engine='openpyxl')
        sec = ensure_columns(sec, ['Fuel.O','Fuel.O.Name'])
        mask_pwr = (
            sec['Tech'].str.startswith('PWR', na=False)            # begins with PWR …
            & ~sec['Tech'].str.startswith(('PWRSDS', 'PWRLDS'),    # … but NOT these 2
                                  na=False)
            )
        sto_mode1 = ((sec['Mode.Operation'] == 1) & (sec['Tech'].str.startswith(('PWRLDS','PWRSDS'), na=False))) # Storage mode 1
        sto_mode2 = ((sec['Mode.Operation'] == 2) & (sec['Tech'].str.startswith(('PWRLDS','PWRSDS'), na=False))) # Storage mode 2
        print(f"  Found {mask_pwr.sum()} power plant output techs in Secondary sheet.")
        for idx in sec[mask_pwr].index:
            tech = sec.at[idx,'Tech']
            try:
                fuel, country, region = parse_pwr_code(tech)
                countryname = iso_country_map.get(country, f"Unknown ({country})")
            except ValueError:
                continue
            if fuel in RENEWABLE_FUELS:
                sec.at[idx,'Fuel.O'] = f"ELC{country}{region}00"
                sec.at[idx,'Fuel.O.Name'] = f"Electricity, {countryname}, Region {region}, renewable power plant output"
            else:
                sec.at[idx,'Fuel.O'] = f"ELC{country}{region}01"
                sec.at[idx,'Fuel.O.Name'] = f"Electricity, {countryname}, Region {region}, NO renewable power plant output"        
        sec.to_excel(writer, sheet_name='Secondary', index=False)
        for idx in sec[sto_mode1].index:
            tech = sec.at[idx,'Tech']
            try:
                fuel, country, region = parse_pwr_code(tech)
                countryname = iso_country_map.get(country, f"Unknown ({country})")
            except ValueError:
                continue
            sec.at[idx,'Fuel.I'] = f"ELC{country}{region}00"
            sec.at[idx,'Fuel.I.Name'] = f"Electricity, {countryname}, Region {region}, renewable power plant output"
        sec.to_excel(writer, sheet_name='Secondary', index=False)
        for idx in sec[sto_mode2].index:
            tech = sec.at[idx,'Tech']
            try:
                fuel, country, region = parse_pwr_code(tech)
                countryname = iso_country_map.get(country, f"Unknown ({country})")
            except ValueError:
                continue
            sec.at[idx,'Fuel.O'] = f"ELC{country}{region}00"
            sec.at[idx,'Fuel.O.Name'] = f"Electricity, {countryname}, Region {region}, renewable power plant output"    
        sec.to_excel(writer, sheet_name='Secondary', index=False)    
        # ---------- Demand Techs sheet ----------
        dtech = pd.read_excel(path, sheet_name='Demand Techs', engine='openpyxl')
        header = list(dtech.columns)
        dtech = dtech.iloc[0:0]  # clear rows

        def add_row(lst):
            lst.append({k:v for k,v in row.items()})

        rows = []
        for country, region in pairs:
            countryname = iso_country_map.get(country, f"Unknown ({country})")
            ren_in  = f"ELC{country}{region}00"
            nor_in  = f"ELC{country}{region}01"
            line_out= f"ELC{country}{region}02"
            entries = [
                ('RNWTRN', ren_in,  'renewable'),
                ('RNWRPO', ren_in,  'renewable'),
                ('RNWNLI', ren_in,  'renewable'),
                ('PWRTRN', nor_in,  'NO renewable'),
                ('TRNRPO', nor_in,  'NO renewable'),
                ('TRNNLI', nor_in,  'NO renewable'),
            ]
            for tech_prefix, fuel_in, label in entries:
                tech = f"{tech_prefix}{country}{region}"
                row = {
                    'Mode.Operation': 1,
                    'Fuel.I': fuel_in,
                    'Fuel.I.Name': f"Electricity from {label} power plants, {countryname}, Region {region}",
                    'Value.Fuel.I': 1,
                    'Unit.Fuel.I': '',
                    'Tech': tech,
                    'Tech.Name': (
                        'Existing' if tech_prefix in ('RNWTRN','PWRTRN') else
                        'Repower'  if tech_prefix in ('RNWRPO','TRNRPO') else
                        'New line'
                    ) + f" transmission technology from {label} power plants, {countryname}, Region {region}",
                    'Fuel.O': line_out,
                    'Fuel.O.Name': f"Electricity, {countryname}, Region {region}, transmission line output",
                    'Value.Fuel.O': 1,
                    'Unit.Fuel.O': ''
                }
                rows.append(row)

        dtech = pd.DataFrame(rows, columns=header)
        dtech.to_excel(writer, sheet_name='Demand Techs', index=False)

    print("✔ Base‑year file updated.")

# ---------------------------------------------------------------------------
# 2. Process A-O_AR_Projections.xlsx
# ---------------------------------------------------------------------------
def process_projections(path, pairs):
    print(f"Processing '{path}' …")
    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        # ---------- Secondary ----------
        sec = pd.read_excel(path, sheet_name='Secondary', engine='openpyxl')
        sec = ensure_columns(sec, ['Fuel','Fuel.Name'])
        mask = sec['Tech'].str.startswith('PWR', na=False) & (sec.get('Direction','')=='Output')
        masksto = sec['Tech'].str.startswith(('PWRLDS','PWRSDS'), na=False) 
        for idx in sec[mask].index:
            tech = sec.at[idx,'Tech']
            try:
                fuel, country, region = parse_pwr_code(tech)
                countryname = iso_country_map.get(country, f"Unknown ({country})")
            except ValueError:
                continue
            if fuel in RENEWABLE_FUELS:
                sec.at[idx,'Fuel'] = f"ELC{country}{region}00"
                sec.at[idx,'Fuel.Name'] = f"Electricity, {countryname}, Region {region}, renewable power plant output"
            else:
                sec.at[idx,'Fuel'] = f"ELC{country}{region}01"
                sec.at[idx,'Fuel.Name'] = f"Electricity, {countryname}, Region {region}, NO renewable power plant output"
        sec.to_excel(writer, sheet_name='Secondary', index=False)
        for idx in sec[masksto].index:
            tech = sec.at[idx,'Tech']
            try:
                fuel, country, region = parse_pwr_code(tech)
                countryname = iso_country_map.get(country, f"Unknown ({country})")
            except ValueError:
                continue
            sec.at[idx,'Fuel'] = f"ELC{country}{region}00"
            sec.at[idx,'Fuel.Name'] = f"Electricity, {countryname}, Region {region}, renewable power plant output"
        sec.to_excel(writer, sheet_name='Secondary', index=False)    

        # ---------- Demand Techs ----------
        dtech = pd.read_excel(path, sheet_name='Demand Techs', engine='openpyxl')
        header = list(dtech.columns)
        # Identify year columns (numeric headers from column index >=8)
        year_cols = [c for c in header if str(c).isdigit()]
        dtech = dtech.iloc[0:0]

        rows = []
        for country, region in pairs:
            ren_in  = f"ELC{country}{region}00"
            nor_in  = f"ELC{country}{region}01"
            line_out= f"ELC{country}{region}02"
            tech_entries = [
                ('RNWTRN', ren_in,  'renewable'),
                ('RNWRPO', ren_in,  'renewable'),
                ('RNWNLI', ren_in,  'renewable'),
                ('PWRTRN', nor_in,  'NO renewable'),
                ('TRNRPO', nor_in,  'NO renewable'),
                ('TRNNLI', nor_in,  'NO renewable'),
            ]
            for tech_prefix, fuel_in, label in tech_entries:
                tech = f"{tech_prefix}{country}{region}"
                countryname = iso_country_map.get(country, f"Unknown ({country})")
                # input row
                rows.append({
                    'Mode.Operation': 1,
                    'Tech': tech,
                    'Tech.Name': (
                        'Existing' if tech_prefix in ('RNWTRN','PWRTRN') else
                        'Repower'  if tech_prefix in ('RNWRPO','TRNRPO') else
                        'New line'
                    ) + f" transmission technology from {label} power plants, {countryname}, Region {region}",
                    'Fuel': fuel_in,
                    'Fuel.Name': f"Electricity from {label} power plants, {countryname}, Region {region}",
                    'Direction': 'Input',
                    'Projection.Mode': 'User defined',
                    'Projection.Parameter': 0,
                    **{yr:1 for yr in year_cols}
                })
                # output row
                rows.append({
                    'Mode.Operation': 1,
                    'Tech': tech,
                    'Tech.Name': (
                        'Existing' if tech_prefix in ('RNWTRN','PWRTRN') else
                        'Repower'  if tech_prefix in ('RNWRPO','TRNRPO') else
                        'New line'
                    ) + f" transmission technology from {label} power plants, {countryname}, Region {region}",
                    'Fuel': line_out,
                    'Fuel.Name': f"Electricity, {countryname}, Region {region}, transmission line output",
                    'Direction': 'Output',
                    'Projection.Mode': 'User defined',
                    'Projection.Parameter': 0,
                    **{yr:1 for yr in year_cols}
                })

        dtech = pd.DataFrame(rows, columns=header)
        dtech.to_excel(writer, sheet_name='Demand Techs', index=False)
    print("✔ Projections file updated.")

# ---------------------------------------------------------------------------
# 3. Process A-O_Parametrization.xlsx
# ---------------------------------------------------------------------------
PARAM_LIST = [
    'CapitalCost','FixedCost','ResidualCapacity','TotalAnnualMinCapacityInvestment', 'TotalAnnualMaxCapacity'
]

def process_parametrization(path, pairs, yaml_data):
    print(f"Processing '{path}' …")

    # ───── 1. Load sheets ───────────────────────────────────────────────────
    fhp   = pd.read_excel(path, sheet_name='Fixed Horizon Parameters',
                          engine='openpyxl')
    dtech = pd.read_excel(path, sheet_name='Demand Techs',
                          engine='openpyxl')


    # Quick map Tech → Tech.ID for already existing techs (only for NON-transmission techs)
    # Transmission technologies will have sequential IDs starting from 1
    transmission_prefixes = ('RNWTRN', 'RNWRPO', 'RNWNLI', 'TRNRPO', 'TRNNLI', 'PWRTRN')
    existing_ids = {}
    for _, row in fhp.iterrows():
        tech = row.get('Tech', '')
        tech_id = row.get('Tech.ID', 0)
        if tech and not any(tech.startswith(p) for p in transmission_prefixes):
            existing_ids[tech] = tech_id

    new_rows_fhp   = []          # new (or missing) rows for FHP
    new_rows_dtech = []          # all rows to be added to Demand Techs

    # Tech.ID counter for transmission technologies (starting from 1)
    tx_tech_id = 0

    # ───── 2. Generate / update technologies ─────────────────────────────
    for country, region in pairs:
        for tech_prefix in transmission_prefixes:
            tech_code = f"{tech_prefix}{country}{region}"
            countryname = iso_country_map.get(country, f"Unknown ({country})")
            # 2.1 Tech.ID: assign sequential ID for transmission technologies
            tx_tech_id += 1
            tech_id = tx_tech_id
            existing_ids[tech_code] = tech_id

            # 2.2 Descriptive name
            tech_name = (
                'Existing' if tech_prefix in ('RNWTRN','PWRTRN') else
                'Repower'  if tech_prefix in ('RNWRPO','TRNRPO') else
                'New line'
            ) + (' transmission technology from renewable power plants, '
                 if tech_prefix.startswith('RNW') else
                 ' transmission technology from NO renewable power plants, ') \
              + f"{countryname}, Region {region}"

            # 2.3 YAML configuration
            cfg = yaml_data.get(tech_prefix, {})
            cap_to_act       = cfg.get('CapacityToActivityUnit', '')
            operational_life = cfg.get('OperationalLife', '')

            # ── A) FIXED HORIZON PARAMETERS ────────────────────────────────
            # Update (or create if missing) CapacityToActivityUnit and OperationalLife
            for par_id, (par_name, par_val) in enumerate(
                    [('CapacityToActivityUnit', cap_to_act),
                     ('OperationalLife',       operational_life)], start=1):
                mask = (fhp['Tech'] == tech_code) & (fhp['Parameter'] == par_name)
                if mask.any():
                    fhp.loc[mask, 'Value'] = par_val             # just update
                else:
                    new_rows_fhp.append({
                        'Tech.Type'  : 'Demand',
                        'Tech.ID'    : tech_id,
                        'Tech'       : tech_code,
                        'Tech.Name'  : tech_name,
                        'Parameter.ID': par_id,
                        'Parameter'  : par_name,
                        'Unit'       : '',
                        'Value'      : par_val
                    })

            # ── B) DEMAND TECHS ────────────────────────────────────────────
            # 1) Remove any previous rows (to avoid duplicates)
            dtech = dtech[dtech['Tech'] != tech_code]

            # 2) Add the block of 12 parameters with the correct Tech.ID
            years = [c for c in dtech.columns if str(c).isdigit()]
            base_row = {
                'Tech.ID'            : tech_id,
                'Tech'               : tech_code,
                'Tech.Name'          : tech_name,
                'Unit'               : '',
                'Projection.Parameter': 0
            }

            for p_id, param in enumerate(PARAM_LIST, start=1):
                row = base_row.copy()
                row['Parameter.ID'] = p_id
                row['Parameter']    = param
                value_cfg = cfg.get(param, None)

                if isinstance(value_cfg, dict):                 # year-by-year values
                    row['Projection.Mode'] = 'User defined'
                    for yr in years:
                        row[yr] = value_cfg.get(yr, '')
                elif value_cfg is not None:                     # constant value
                    row['Projection.Mode'] = 'User defined'
                    for yr in years:
                        row[yr] = value_cfg
                else:                                           # no data
                    row['Projection.Mode'] = 'EMPTY'
                    for yr in years:
                        row[yr] = ''

                new_rows_dtech.append(row)

    # ───── 3. Combinar y guardar ───────────────────────────────────────────
    if new_rows_fhp:
        fhp = pd.concat([fhp, pd.DataFrame(new_rows_fhp)],
                        ignore_index=True)

    # Concatenate all rows (new or recreated) of Demand Techs
    dtech = pd.concat([dtech, pd.DataFrame(new_rows_dtech)],
                      ignore_index=True)

    with pd.ExcelWriter(path, engine='openpyxl',
                        mode='a', if_sheet_exists='replace') as writer:
        fhp.to_excel  (writer, sheet_name='Fixed Horizon Parameters', index=False)
        dtech.to_excel(writer, sheet_name='Demand Techs',            index=False)

    print("✔ Parametrization file updated.")

def list_scenario_suffixes(base_dir: Path) -> List[str]:
    """Return list like ['BAU_NoRPO','NDC','NDC+ELC'] from folders 'A1_Outputs_*'."""
    suffixes: List[str] = []
    for item in sorted(base_dir.iterdir()):
        if item.is_dir() and item.name.startswith("A1_Outputs_"):
            suffix = item.name.split("A1_Outputs_", 1)[1]
            if suffix:  # Ensure non-empty
                suffixes.append(suffix)
    return suffixes

# ---------------------------------------------------------------------------
# CLI glue
# ---------------------------------------------------------------------------
def main():

    script_dir = Path(__file__).resolve().parent
    OUTPUT_FOLDER = script_dir / "A1_Outputs"
    scenario_suffixes = list_scenario_suffixes(OUTPUT_FOLDER)
    for scen in scenario_suffixes:
    
    
        defaults = {
            "yaml": str(script_dir / "Config_country_codes.yaml"),
            "base": str(script_dir / f"A1_Outputs/A1_Outputs_{scen}/A-O_AR_Model_Base_Year.xlsx"),
            "proj": str(script_dir / f"A1_Outputs/A1_Outputs_{scen}/A-O_AR_Projections.xlsx"),
            "param": str(script_dir / f"A1_Outputs/A1_Outputs_{scen}/A-O_Parametrization.xlsx")
        }
        ap = argparse.ArgumentParser(description='Process CLG model spreadsheets.')
        ap.add_argument('--yaml', help='Config_country_codes.yaml')
        ap.add_argument('--base', help='A-O_AR_Model_Base_Year.xlsx')
        ap.add_argument('--proj', help='A-O_AR_Projections.xlsx')
        ap.add_argument('--param', help='A-O_Parametrization.xlsx')
        ap.set_defaults(**defaults)
        args = ap.parse_args()
       
        pairs = load_country_region_pairs(args.yaml)
        if not pairs:
            sys.exit('No valid country/region codes found in YAML.')
    
        # Optional detailed YAML per-tech/parameter values
        with open(args.yaml,'r',encoding='utf-8') as fh:
            yaml_data = yaml.safe_load(fh)
            if not isinstance(yaml_data, dict):
                yaml_data = {}
    
        process_base_year(args.base, pairs)
        process_projections(args.proj, pairs)
        process_parametrization(args.param, pairs, yaml_data)

    print("\n  All done!")

if __name__ == '__main__':
    main()
