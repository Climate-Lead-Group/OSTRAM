# -*- coding: utf-8 -*-
"""
Created on 2025

@author: Climate Lead Group, Javier Monge-Matamoros, Andrey Salazar-Vargas
"""

import argparse
import sys
import os
import re
import yaml
import pandas as pd
from typing import List
from pathlib import Path
from Z_AUX_config_loader import get_renewable_fuels, get_iso_country_map

# Mapeos de países y tecnologías desde configuración centralizada
RENEWABLE_FUELS = get_renewable_fuels()
iso_country_map = get_iso_country_map()

# Patrón para detectar códigos de interconexión TRN de 13 caracteres (ej. TRNBGDXXNPLXX)
# NO coincide con PWRTRN, TRNRPO, TRNNLI (esos son <=11 caracteres)
TRN_INTERCONNECTION = re.compile(r'^TRN[A-Z]{5}[A-Z]{5}$')

# ---------------------------------------------------------------------------
# Funciones auxiliares
# ---------------------------------------------------------------------------
def load_country_region_pairs(yaml_path):
    """Retorna lista de tuplas (país, región) desde la clave 'countries' en YAML.

    Códigos de 3 letras ⇒ region='XX'
    Códigos de 5 letras ⇒ últimas 2 letras son la región
    """
    with open(yaml_path, 'r', encoding='utf-8') as fh:
        data = yaml.safe_load(fh)

    codes = data.get('countries', [])
    if not isinstance(codes, list):
        print("⚠️  La clave 'countries' en YAML no es una lista", file=sys.stderr)
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
            print(f"⚠️  Omitiendo código no reconocido '{code}'", file=sys.stderr)
    # Eliminar duplicados preservando el orden
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
    """Asegurar que el DataFrame contenga cada columna en *cols* (crea si no existe)."""
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    return df

# ---------------------------------------------------------------------------
# 1. Procesar A-O_AR_Model_Base_Year.xlsx
# ---------------------------------------------------------------------------
def process_base_year(path, pairs, enable_dsptrn=False):
    print(f"Procesando '{path}' …")
    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        # ---------- Hoja Secondary ----------
        sec = pd.read_excel(path, sheet_name='Secondary', engine='openpyxl')
        sec = ensure_columns(sec, ['Fuel.O','Fuel.O.Name'])
        mask_pwr = (
            sec['Tech'].str.startswith('PWR', na=False)            # begins with PWR …
            & ~sec['Tech'].str.startswith(('PWRSDS', 'PWRLDS'),    # … but NOT these 2
                                  na=False)
            )
        sto_mode1 = ((sec['Mode.Operation'] == 1) & (sec['Tech'].str.startswith(('PWRLDS','PWRSDS'), na=False))) # Storage mode 1
        sto_mode2 = ((sec['Mode.Operation'] == 2) & (sec['Tech'].str.startswith(('PWRLDS','PWRSDS'), na=False))) # Storage mode 2
        print(f"  Se encontraron {mask_pwr.sum()} techs de salida de planta en hoja Secondary.")
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
        # ---------- Interconexión TRN: actualizar códigos de combustible 02->03, 01->03 ----------
        if enable_dsptrn:
            mask_trn = sec['Tech'].apply(lambda x: bool(TRN_INTERCONNECTION.match(str(x))))
            for idx in sec[mask_trn].index:
                # --- Fuel.I: reescribir a ELC..03 (listo para despacho de interconexión) ---
                fuel_i = str(sec.at[idx, 'Fuel.I'])
                if fuel_i.startswith('ELC') and fuel_i.endswith('02'):
                    sec.at[idx, 'Fuel.I'] = fuel_i[:-2] + '03'
                    sec.at[idx, 'Fuel.I.Name'] = str(sec.at[idx, 'Fuel.I.Name']).replace(
                        'transmission line output', 'dispatch-ready for interconnection')
                elif fuel_i.startswith('ELC') and fuel_i.endswith('01'):
                    sec.at[idx, 'Fuel.I'] = fuel_i[:-2] + '03'
                    sec.at[idx, 'Fuel.I.Name'] = str(sec.at[idx, 'Fuel.I.Name']).replace(
                        'NO renewable power plant output', 'dispatch-ready for interconnection')
                # --- Fuel.O: reescribir a ELC..04 (electricidad importada) ---
                fuel_o = str(sec.at[idx, 'Fuel.O'])
                if fuel_o.startswith('ELC') and fuel_o.endswith('02'):
                    sec.at[idx, 'Fuel.O'] = fuel_o[:-2] + '04'
                    sec.at[idx, 'Fuel.O.Name'] = str(sec.at[idx, 'Fuel.O.Name']).replace(
                        'transmission line output', 'imported electricity')
                elif fuel_o.startswith('ELC') and fuel_o.endswith('01'):
                    sec.at[idx, 'Fuel.O'] = fuel_o[:-2] + '04'
                    sec.at[idx, 'Fuel.O.Name'] = str(sec.at[idx, 'Fuel.O.Name']).replace(
                        'NO renewable power plant output', 'imported electricity')
                elif fuel_o.startswith('ELC') and fuel_o.endswith('03'):
                    sec.at[idx, 'Fuel.O'] = fuel_o[:-2] + '04'
                    sec.at[idx, 'Fuel.O.Name'] = str(sec.at[idx, 'Fuel.O.Name']).replace(
                        'dispatch-ready for interconnection', 'imported electricity')
        sec.to_excel(writer, sheet_name='Secondary', index=False)
        # ---------- Hoja Demand Techs ----------
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

            # --- DSPTRN: Tecnología de despacho (2 modos de operación) ---
            if enable_dsptrn:
                dsp_tech = f"DSPTRN{country}{region}"
                dsp_02   = f"ELC{country}{region}02"
                dsp_03   = f"ELC{country}{region}03"
                dsp_04   = f"ELC{country}{region}04"
                dsp_name = f"Dispatch technology, {countryname}, Region {region}"
                # Modo 1: ELC...02 -> ELC...03 (despacho a interconexión)
                rows.append({
                    'Mode.Operation': 1,
                    'Fuel.I': dsp_02,
                    'Fuel.I.Name': f"Electricity, {countryname}, Region {region}, transmission line output",
                    'Value.Fuel.I': 1,
                    'Unit.Fuel.I': '',
                    'Tech': dsp_tech,
                    'Tech.Name': dsp_name,
                    'Fuel.O': dsp_03,
                    'Fuel.O.Name': f"Electricity, {countryname}, Region {region}, dispatch-ready for interconnection",
                    'Value.Fuel.O': 1,
                    'Unit.Fuel.O': ''
                })
                # Modo 2: ELC...04 -> ELC...03 (recibir de interconexión)
                rows.append({
                    'Mode.Operation': 2,
                    'Fuel.I': dsp_04,
                    'Fuel.I.Name': f"Electricity, {countryname}, Region {region}, imported electricity",
                    'Value.Fuel.I': 1,
                    'Unit.Fuel.I': '',
                    'Tech': dsp_tech,
                    'Tech.Name': dsp_name,
                    'Fuel.O': dsp_03,
                    'Fuel.O.Name': f"Electricity, {countryname}, Region {region}, dispatch-ready for interconnection",
                    'Value.Fuel.O': 1,
                    'Unit.Fuel.O': ''
                })

        dtech = pd.DataFrame(rows, columns=header)
        dtech.to_excel(writer, sheet_name='Demand Techs', index=False)

    print("✔ Archivo de año base actualizado.")

# ---------------------------------------------------------------------------
# 2. Procesar A-O_AR_Projections.xlsx
# ---------------------------------------------------------------------------
def process_projections(path, pairs, enable_dsptrn=False):
    print(f"Procesando '{path}' …")
    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        # ---------- Hoja Secondary ----------
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
        # ---------- Interconexión TRN: actualizar códigos de combustible (Input->03, Output->04) ----------
        if enable_dsptrn:
            mask_trn = sec['Tech'].apply(lambda x: bool(TRN_INTERCONNECTION.match(str(x))))
            for idx in sec[mask_trn].index:
                fuel = str(sec.at[idx, 'Fuel'])
                name = str(sec.at[idx, 'Fuel.Name'])
                direction = str(sec.at[idx, 'Direction']) if 'Direction' in sec.columns else ''
                if direction == 'Input':
                    # Lado Input: reescribir a ELC..03 (listo para despacho de interconexión)
                    if fuel.startswith('ELC') and fuel.endswith('02'):
                        sec.at[idx, 'Fuel'] = fuel[:-2] + '03'
                        sec.at[idx, 'Fuel.Name'] = name.replace(
                            'transmission line output', 'dispatch-ready for interconnection')
                    elif fuel.startswith('ELC') and fuel.endswith('01'):
                        sec.at[idx, 'Fuel'] = fuel[:-2] + '03'
                        sec.at[idx, 'Fuel.Name'] = name.replace(
                            'NO renewable power plant output', 'dispatch-ready for interconnection')
                elif direction == 'Output':
                    # Lado Output: reescribir a ELC..04 (electricidad importada)
                    if fuel.startswith('ELC') and fuel.endswith('02'):
                        sec.at[idx, 'Fuel'] = fuel[:-2] + '04'
                        sec.at[idx, 'Fuel.Name'] = name.replace(
                            'transmission line output', 'imported electricity')
                    elif fuel.startswith('ELC') and fuel.endswith('01'):
                        sec.at[idx, 'Fuel'] = fuel[:-2] + '04'
                        sec.at[idx, 'Fuel.Name'] = name.replace(
                            'NO renewable power plant output', 'imported electricity')
                    elif fuel.startswith('ELC') and fuel.endswith('03'):
                        sec.at[idx, 'Fuel'] = fuel[:-2] + '04'
                        sec.at[idx, 'Fuel.Name'] = name.replace(
                            'dispatch-ready for interconnection', 'imported electricity')
        sec.to_excel(writer, sheet_name='Secondary', index=False)

        # ---------- Hoja Demand Techs ----------
        dtech = pd.read_excel(path, sheet_name='Demand Techs', engine='openpyxl')
        header = list(dtech.columns)
        # Identificar columnas de años (encabezados numéricos desde índice de columna >=8)
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
                # fila de entrada
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
                # fila de salida
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

            # --- DSPTRN: Tecnología de despacho (2 modos de operación) ---
            if enable_dsptrn:
                dsp_tech = f"DSPTRN{country}{region}"
                dsp_02   = f"ELC{country}{region}02"
                dsp_03   = f"ELC{country}{region}03"
                dsp_04   = f"ELC{country}{region}04"
                dsp_name = f"Dispatch technology, {countryname}, Region {region}"
                # Modo 1 Input: ELC...02
                rows.append({
                    'Mode.Operation': 1,
                    'Tech': dsp_tech,
                    'Tech.Name': dsp_name,
                    'Fuel': dsp_02,
                    'Fuel.Name': f"Electricity, {countryname}, Region {region}, transmission line output",
                    'Direction': 'Input',
                    'Projection.Mode': 'User defined',
                    'Projection.Parameter': 0,
                    **{yr: 1 for yr in year_cols}
                })
                # Modo 1 Output: ELC...03
                rows.append({
                    'Mode.Operation': 1,
                    'Tech': dsp_tech,
                    'Tech.Name': dsp_name,
                    'Fuel': dsp_03,
                    'Fuel.Name': f"Electricity, {countryname}, Region {region}, dispatch-ready for interconnection",
                    'Direction': 'Output',
                    'Projection.Mode': 'User defined',
                    'Projection.Parameter': 0,
                    **{yr: 1 for yr in year_cols}
                })
                # Modo 2 Input: ELC...04 (electricidad importada)
                rows.append({
                    'Mode.Operation': 2,
                    'Tech': dsp_tech,
                    'Tech.Name': dsp_name,
                    'Fuel': dsp_04,
                    'Fuel.Name': f"Electricity, {countryname}, Region {region}, imported electricity",
                    'Direction': 'Input',
                    'Projection.Mode': 'User defined',
                    'Projection.Parameter': 0,
                    **{yr: 1 for yr in year_cols}
                })
                # Modo 2 Output: ELC...03 (listo para despacho de interconexión)
                rows.append({
                    'Mode.Operation': 2,
                    'Tech': dsp_tech,
                    'Tech.Name': dsp_name,
                    'Fuel': dsp_03,
                    'Fuel.Name': f"Electricity, {countryname}, Region {region}, dispatch-ready for interconnection",
                    'Direction': 'Output',
                    'Projection.Mode': 'User defined',
                    'Projection.Parameter': 0,
                    **{yr: 1 for yr in year_cols}
                })

        dtech = pd.DataFrame(rows, columns=header)
        dtech.to_excel(writer, sheet_name='Demand Techs', index=False)
    print("✔ Archivo de proyecciones actualizado.")

# ---------------------------------------------------------------------------
# 3. Procesar A-O_Parametrization.xlsx
# ---------------------------------------------------------------------------
PARAM_LIST = [
    'CapitalCost','FixedCost','ResidualCapacity','TotalAnnualMinCapacityInvestment', 'TotalAnnualMaxCapacity'
]

def process_parametrization(path, pairs, yaml_data, enable_dsptrn=False):
    print(f"Procesando '{path}' …")

    # ───── 1. Cargar hojas ───────────────────────────────────────────────────
    fhp   = pd.read_excel(path, sheet_name='Fixed Horizon Parameters',
                          engine='openpyxl')
    dtech = pd.read_excel(path, sheet_name='Demand Techs',
                          engine='openpyxl')


    # Mapa rápido Tech → Tech.ID para techs ya existentes (solo para techs que NO son de transmisión)
    # Las tecnologías de transmisión tendrán IDs secuenciales comenzando desde 1
    transmission_prefixes = ('RNWTRN', 'RNWRPO', 'RNWNLI', 'TRNRPO', 'TRNNLI', 'PWRTRN')
    if enable_dsptrn:
        transmission_prefixes = transmission_prefixes + ('DSPTRN',)
    existing_ids = {}
    for _, row in fhp.iterrows():
        tech = row.get('Tech', '')
        tech_id = row.get('Tech.ID', 0)
        if tech and not any(tech.startswith(p) for p in transmission_prefixes):
            existing_ids[tech] = tech_id

    new_rows_fhp   = []          # filas nuevas (o faltantes) para FHP
    new_rows_dtech = []          # todas las filas a agregar a Demand Techs

    # Contador de Tech.ID para tecnologías de transmisión (comenzando desde 1)
    tx_tech_id = 0

    # ───── 2. Generar / actualizar tecnologías ─────────────────────────────
    for country, region in pairs:
        for tech_prefix in transmission_prefixes:
            tech_code = f"{tech_prefix}{country}{region}"
            countryname = iso_country_map.get(country, f"Unknown ({country})")
            # 2.1 Tech.ID: asignar ID secuencial para tecnologías de transmisión
            tx_tech_id += 1
            tech_id = tx_tech_id
            existing_ids[tech_code] = tech_id

            # 2.2 Nombre descriptivo
            if tech_prefix == 'DSPTRN':
                tech_name = f"Dispatch technology, {countryname}, Region {region}"
            else:
                tech_name = (
                    'Existing' if tech_prefix in ('RNWTRN','PWRTRN') else
                    'Repower'  if tech_prefix in ('RNWRPO','TRNRPO') else
                    'New line'
                ) + (' transmission technology from renewable power plants, '
                     if tech_prefix.startswith('RNW') else
                     ' transmission technology from NO renewable power plants, ') \
                  + f"{countryname}, Region {region}"

            # 2.3 Configuración YAML
            cfg = yaml_data.get(tech_prefix, {})
            cap_to_act       = cfg.get('CapacityToActivityUnit', '')
            operational_life = cfg.get('OperationalLife', '')

            # ── A) PARÁMETROS DE HORIZONTE FIJO ────────────────────────────────
            # Actualizar (o crear si falta) CapacityToActivityUnit y OperationalLife
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
            # 1) Eliminar filas previas (para evitar duplicados)
            dtech = dtech[dtech['Tech'] != tech_code]

            # 2) Agregar el bloque de 12 parámetros con el Tech.ID correcto
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

                if isinstance(value_cfg, dict):                 # valores año por año
                    row['Projection.Mode'] = 'User defined'
                    for yr in years:
                        row[yr] = value_cfg.get(yr, '')
                elif value_cfg is not None:                     # valor constante
                    row['Projection.Mode'] = 'User defined'
                    for yr in years:
                        row[yr] = value_cfg
                else:                                           # sin datos
                    row['Projection.Mode'] = 'EMPTY'
                    for yr in years:
                        row[yr] = ''

                new_rows_dtech.append(row)

    # ───── 3. Combinar y guardar ───────────────────────────────────────────
    if new_rows_fhp:
        fhp = pd.concat([fhp, pd.DataFrame(new_rows_fhp)],
                        ignore_index=True)

    # Concatenar todas las filas (nuevas o recreadas) de Demand Techs
    dtech = pd.concat([dtech, pd.DataFrame(new_rows_dtech)],
                      ignore_index=True)

    with pd.ExcelWriter(path, engine='openpyxl',
                        mode='a', if_sheet_exists='replace') as writer:
        fhp.to_excel  (writer, sheet_name='Fixed Horizon Parameters', index=False)
        dtech.to_excel(writer, sheet_name='Demand Techs',            index=False)

    print("✔ Archivo de parametrización actualizado.")

# ---------------------------------------------------------------------------
# 4. Procesar A-O_Demand.xlsx
# ---------------------------------------------------------------------------
def process_demand(path, enable_dsptrn=False):
    """Reescribe ELC…02 → ELC…03 en el archivo Demand cuando el despacho DSPTRN está habilitado.

    Con la capa de despacho activa, el punto de demanda se mueve de ELC…02
    (salida de línea de transmisión) a ELC…03 (electricidad lista para despacho).
    Ambas hojas (Demand_Projection y Profiles) se actualizan in-place.
    """
    if not enable_dsptrn:
        return

    print(f"Procesando '{path}' (Demand – ELC02→ELC03) …")

    elc02 = re.compile(r'^(ELC[A-Z]{5})02$')

    with pd.ExcelWriter(path, engine='openpyxl', mode='a',
                         if_sheet_exists='overlay') as writer:

        # --- Hoja Demand_Projection (Fuel/Tech en col B) ---
        dp = pd.read_excel(path, sheet_name='Demand_Projection', engine='openpyxl')
        fuel_col = 'Fuel/Tech'
        name_col = 'Name'
        updated = 0
        for idx in dp.index:
            fuel = str(dp.at[idx, fuel_col]) if pd.notna(dp.at[idx, fuel_col]) else ''
            m = elc02.match(fuel)
            if m:
                dp.at[idx, fuel_col] = m.group(1) + '03'
                if name_col in dp.columns and pd.notna(dp.at[idx, name_col]):
                    dp.at[idx, name_col] = (
                        str(dp.at[idx, name_col])
                        .replace('transmission line', 'dispatch-ready')
                        .replace('transmission lines', 'dispatch-ready')
                    )
                updated += 1
        dp.to_excel(writer, sheet_name='Demand_Projection', index=False)
        print(f"  Demand_Projection: {updated} códigos de combustible actualizados.")

        # --- Hoja Profiles (Fuel/Tech en col C) ---
        pf = pd.read_excel(path, sheet_name='Profiles', engine='openpyxl')
        updated = 0
        for idx in pf.index:
            fuel = str(pf.at[idx, fuel_col]) if pd.notna(pf.at[idx, fuel_col]) else ''
            m = elc02.match(fuel)
            if m:
                pf.at[idx, fuel_col] = m.group(1) + '03'
                if name_col in pf.columns and pd.notna(pf.at[idx, name_col]):
                    pf.at[idx, name_col] = (
                        str(pf.at[idx, name_col])
                        .replace('transmission line', 'dispatch-ready')
                        .replace('transmission lines', 'dispatch-ready')
                    )
                updated += 1
        pf.to_excel(writer, sheet_name='Profiles', index=False)
        print(f"  Profiles: {updated} códigos de combustible actualizados.")

    print("✔ Archivo de demanda actualizado.")


def list_scenario_suffixes(base_dir: Path) -> List[str]:
    """Retorna lista como ['BAU_NoRPO','NDC','NDC+ELC'] desde carpetas 'A1_Outputs_*'."""
    suffixes: List[str] = []
    for item in sorted(base_dir.iterdir()):
        if item.is_dir() and item.name.startswith("A1_Outputs_"):
            suffix = item.name.split("A1_Outputs_", 1)[1]
            if suffix:  # Ensure non-empty
                suffixes.append(suffix)
    return suffixes

# ---------------------------------------------------------------------------
# Interfaz CLI
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
            "param": str(script_dir / f"A1_Outputs/A1_Outputs_{scen}/A-O_Parametrization.xlsx"),
            "demand": str(script_dir / f"A1_Outputs/A1_Outputs_{scen}/A-O_Demand.xlsx"),
        }
        ap = argparse.ArgumentParser(description='Procesar hojas de cálculo del modelo CLG.')
        ap.add_argument('--yaml', help='Config_country_codes.yaml')
        ap.add_argument('--base', help='A-O_AR_Model_Base_Year.xlsx')
        ap.add_argument('--proj', help='A-O_AR_Projections.xlsx')
        ap.add_argument('--param', help='A-O_Parametrization.xlsx')
        ap.add_argument('--demand', help='A-O_Demand.xlsx')
        ap.set_defaults(**defaults)
        args = ap.parse_args()

        pairs = load_country_region_pairs(args.yaml)
        if not pairs:
            sys.exit('No se encontraron códigos de país/región válidos en el YAML.')

        # Valores opcionales detallados del YAML por tech/parámetro
        with open(args.yaml,'r',encoding='utf-8') as fh:
            yaml_data = yaml.safe_load(fh)
            if not isinstance(yaml_data, dict):
                yaml_data = {}

        enable_dsptrn = yaml_data.get('enable_dsptrn', False)
        process_base_year(args.base, pairs, enable_dsptrn=enable_dsptrn)
        process_projections(args.proj, pairs, enable_dsptrn=enable_dsptrn)
        process_parametrization(args.param, pairs, yaml_data, enable_dsptrn=enable_dsptrn)
        process_demand(args.demand, enable_dsptrn=enable_dsptrn)

    print("\n  ¡Todo listo!")

if __name__ == '__main__':
    main()
