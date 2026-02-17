# -*- coding: utf-8 -*-
"""
Created on Tue Jul 29 11:20:29 2025

@author: ClimateLeadGroup
"""

import pandas as pd
import os
import re

parametrization = False
demand = True
storage = False

folder = 'BAU'



if parametrization:
    # Cargar archivo Excel
    file_path = os.path.join(f"A1_Outputs_{folder}","A-O_Parametrization.xlsx")  # Cambia esto si el archivo está en otra ubicación
    xls = pd.ExcelFile(file_path)
    
    # Leer la hoja
    df_fixed = xls.parse("Fixed Horizon Parameters")
    
    # --- Parte 1: Reemplazar BRACN por BRAXX ---
    mask_bracn = df_fixed["Tech"].str.contains("BRACN", na=False)
    df_fixed.loc[mask_bracn, "Tech"] = df_fixed.loc[mask_bracn, "Tech"].str.replace("BRACN", "BRAXX", regex=False)
    df_fixed.loc[mask_bracn, "Tech.Name"] = df_fixed.loc[mask_bracn, "Tech.Name"].str.replace("CN", "XX", regex=False)
    
    # --- Parte 2: Eliminar combinaciones regionales no deseadas de BRA ---
    brazil_bad_regions = ["BRANW", "BRANE", "BRACW", "BRASO", "BRASE", "BRAWE"]
    mask_bad_regions = df_fixed["Tech"].str[6:11].isin(brazil_bad_regions)
    df_fixed = df_fixed[~mask_bad_regions]
    
    # --- Parte 3: Eliminar tecnologías de longitud 13 donde "BRA" aparece dos veces ---
    is_length_13 = df_fixed["Tech"].str.len() == 13
    has_two_bra  = df_fixed["Tech"].str.count("BRA") > 1
    df_fixed = df_fixed[~(is_length_13 & has_two_bra)]
    
    # --- Parte 4: Unificar SOLO las interconexiones BRA<->otro país y cambiar región BRA a XX ---
    
    # 4.1) Mascara para filtrar solo TRN… que contienen BRA
    mask_bra_trn = (
        df_fixed["Tech"].str.startswith("TRN") &
        df_fixed["Tech"].str.contains("BRA") &
        (df_fixed["Tech"].str.len() == 13)
    )
    df_inter = df_fixed[mask_bra_trn].copy()
    
    # 4.2) Funciones de normalización
    def normalize_interconnection(code):
        p1, p2 = code[3:8], code[8:13]
        n1 = "BRAXX" if "BRA" in p1 else p1
        n2 = "BRAXX" if "BRA" in p2 else p2
        return "TRN" + "".join(sorted([n1, n2]))
    
    def update_bra_region(code):
        p1, p2 = code[3:8], code[8:13]
        if "BRA" in p1: p1 = "BRAXX"
        if "BRA" in p2: p2 = "BRAXX"
        return "TRN" + p1 + p2
    
    # 4.3) Agrupar por interconexión normalizada y por parámetro
    df_inter["NormKey"] = df_inter["Tech"].apply(normalize_interconnection)
    new_trn_rows = []
    for (norm_key, parameter), group in df_inter.groupby(["NormKey", "Parameter"]):
        base = group.iloc[0].copy()
        # Generar el nuevo código y nombre
        base["Tech"]      = update_bra_region(base["Tech"])
        base["Tech.Name"] = re.sub(
            r"Brazil, region [A-Z]{2}",
            "Brazil, region XX",
            base["Tech.Name"],
            flags=re.IGNORECASE
        )
        # Mantener el resto de columnas tal cual (Parameter.ID, Unit, años, etc.)
        new_trn_rows.append(base)
    
    # 4.4) Reconstruir df_fixed_final: quitar los BRA-TRN originales y añadir los nuevos
    df_fixed_final = pd.concat([
        df_fixed[~mask_bra_trn],
        pd.DataFrame(new_trn_rows).drop(columns=["NormKey"])
    ], ignore_index=True)
    
    # --- Reasignar Tech.ID agrupado por Tech ---
    unique_fixed = pd.unique(df_fixed_final["Tech"])
    fixed_id_map = {tech: i+1 for i, tech in enumerate(unique_fixed)}
    df_fixed_final["Tech.ID"] = df_fixed_final["Tech"].map(fixed_id_map)
    
    print("Hoja 'Fixed Horizon Parameters' procesada correctamente.")
    
    
    
    df_sec = xls.parse("Secondary Techs")
    
    # 2) Quitar duplicados internos BRA–BRA
    mask_len13     = df_sec["Tech"].str.len() == 13
    mask_bra_twice = df_sec["Tech"].str.count("BRA") > 1
    df_sec = df_sec[~(mask_len13 & mask_bra_twice)].copy()
    
    # 3) Definir regiones y parámetros
    brazil_regions = ["BRACN","BRANW","BRANE","BRACW","BRASO","BRASE","BRAWE"]
    parameters_avg = [
        "CapitalCost","FixedCost","AvailabilityFactor",
        "ReserveMarginTagFuel","ReserveMarginTagTechnology"
    ]
    parameters_sum = [
        "ResidualCapacity","TotalAnnualMaxCapacity",
        "TotalTechnologyAnnualActivityUpperLimit",
        "TotalTechnologyAnnualActivityLowerLimit",
        "TotalAnnualMinCapacityInvestment",
        "TotalAnnualMaxCapacityInvestment"
    ]
    
    # 4) Columnas de años 2021–2050
    year_cols = [
        c for c in df_sec.columns
        if (isinstance(c,int)   and 2021 <= c <= 2050)
           or (isinstance(c,str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]
    
    # 5) Identificar filas usadas para generar BRAXX
    mask_bra_pwr = (
        df_sec["Tech"].str.startswith("PWR") &
        df_sec["Tech"].str.contains("BRA") &
        df_sec["Tech"].str[-2:].isin(["00","01"]) &
        df_sec["Tech"].str[6:11].isin(brazil_regions)
    )
    mask_trn = (
        df_sec["Tech"].str.startswith("TRN") &
        df_sec["Tech"].str.contains("BRA") &
        (df_sec["Tech"].str.len() == 13)
    )
    mask_elc = (
        df_sec["Tech"].str.startswith("ELC") &
        df_sec["Tech"].str.contains("BRA") &
        df_sec["Tech"].str.endswith("01") &
        df_sec["Tech"].str[3:8].isin(brazil_regions)
    )
    mask_bra_bck = (
        df_sec["Tech"].str.startswith("PWRBCK") &
        df_sec["Tech"].str.contains("BRA") &
        df_sec["Tech"].str[6:11].isin(brazil_regions)
    )
    used_mask = mask_bra_pwr | mask_trn | mask_elc | mask_bra_bck
    
    # 6) Conservar sólo filas que NO se usaron para calcular BRAXX
    df_original = df_sec[~used_mask].copy()
    
    new_rows = []
    
    # 7.1) PWR...BRA...00/01 → BRAXX
    df_bra_pwr = df_sec[mask_bra_pwr].copy()
    df_bra_pwr["TechKey"] = df_bra_pwr["Tech"].str[:6] + df_bra_pwr["Tech"].str[-2:]
    for (tech_key, parameter), group in df_bra_pwr.groupby(["TechKey","Parameter"]):
        base = group.iloc[0].copy()
        base["Tech"] = base["Tech"][:6] + "BRAXX" + base["Tech"][-2:]
        base["Tech.Name"] = re.sub(r"Brazil, region [A-Z]{2}", "Brazil, region XX", base["Tech.Name"])
        if parameter in parameters_avg:
            vals = group[year_cols].astype(float).mean()
        else:
            vals = group[year_cols].astype(float).sum(min_count=1)
        base[year_cols] = vals
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_rows.append(base)
    
    # 7.2) TRN…BRA interconexiones → BRAXX
    df_trn = df_sec[mask_trn].copy()
    def normalize_trn(code):
        p1,p2 = code[3:8], code[8:13]
        n1 = "BRAXX" if "BRA" in p1 else p1
        n2 = "BRAXX" if "BRA" in p2 else p2
        return "TRN" + "".join(sorted([n1,n2]))
    def update_trn(code):
        p1,p2 = code[3:8], code[8:13]
        if "BRA" in p1: p1="BRAXX"
        if "BRA" in p2: p2="BRAXX"
        return "TRN"+p1+p2
    df_trn["NormKey"] = df_trn["Tech"].apply(normalize_trn)
    for (norm_key, parameter), group in df_trn.groupby(["NormKey","Parameter"]):
        base = group.iloc[0].copy()
        base["Tech"] = update_trn(base["Tech"])
        base["Tech.Name"] = re.sub(r"Brazil, region [A-Z]{2}", "Brazil, region XX", base["Tech.Name"])
        if parameter in parameters_avg:
            vals = group[year_cols].astype(float).mean()
        else:
            vals = group[year_cols].astype(float).sum(min_count=1)
        base[year_cols] = vals
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_rows.append(base)
    
    # 7.3) ELC…BRA…01 → BRAXX
    df_elc = df_sec[mask_elc].copy()
    df_elc["ElcKey"] = df_elc["Tech"].str[:3] + df_elc["Tech"].str[-2:]
    for (elc_key, parameter), group in df_elc.groupby(["ElcKey","Parameter"]):
        base = group.iloc[0].copy()
        base["Tech"] = elc_key[:3] + "BRAXX" + elc_key[-2:]
        base["Tech.Name"] = re.sub(r"Brazil, region [A-Z]{2}", "Brazil, region XX", base["Tech.Name"])
        if parameter in parameters_avg:
            vals = group[year_cols].astype(float).mean()
        else:
            vals = group[year_cols].astype(float).sum(min_count=1)
        base[year_cols] = vals
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_rows.append(base)
    
    # 7.4) PWRBCK…BRA… → BRAXX
    df_bra_bck = df_sec[mask_bra_bck].copy()
    df_bra_bck["TechKey"] = df_bra_bck["Tech"].str[:6]
    for (tech_key, parameter), group in df_bra_bck.groupby(["TechKey","Parameter"]):
        base = group.iloc[0].copy()
        base["Tech"] = tech_key + "BRAXX"
        base["Tech.Name"] = re.sub(r"Brazil, region [A-Z]{2}", "Brazil, region XX", base["Tech.Name"])
        if parameter in parameters_avg:
            vals = group[year_cols].astype(float).mean()
        else:
            vals = group[year_cols].astype(float).sum(min_count=1)
        base[year_cols] = vals
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_rows.append(base)
    
    # 8) Concatenar originales + nuevas filas BRAXX
    df_sec_final = pd.concat([df_original, pd.DataFrame(new_rows)], ignore_index=True)
    
    # 9) Eliminar columnas auxiliares
    df_sec_final.drop(columns=["TechKey", "NormKey", "ElcKey"], errors="ignore", inplace=True)
    
    # --- Reasignar Tech.ID agrupado por Tech ---
    unique_sec = pd.unique(df_sec_final["Tech"])
    sec_id_map = {tech: i+1 for i, tech in enumerate(unique_sec)}
    df_sec_final["Tech.ID"] = df_sec_final["Tech"].map(sec_id_map)
    
    print("Hoja 'Secondary Techs' procesada correctamente.")
    # ... tras calcular df_fixed_final y df_sec_final ...
    
    
    
    
    
    df_dem = xls.parse("Demand Techs")
    
    # 2) Definir parámetros y operaciones
    parameters_avg = ["CapitalCost", "FixedCost"]
    parameters_sum = [
        "ResidualCapacity",
        "TotalAnnualMinCapacityInvestment",
        "TotalAnnualMaxCapacity"
    ]
    
    # 3) Detectar columnas de años (2021–2050), sean int o str
    year_cols = [
        c for c in df_dem.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]
    
    # 4) Detectar filas BRA…RR con expresión regular
    #     - .{6} fija cualquier carácter en posiciones 0–5
    #     - BRA en posiciones 6–8
    #     - uno de los sufijos de región en 9–10
    brazil_pattern = r"^.{6}BRA(?:CN|NW|NE|CW|SO|SE|WE)$"
    mask_bra = df_dem["Tech"].str.contains(
        brazil_pattern,
        regex=True,
        na=False,
        flags=re.IGNORECASE
    )
    
    # 5) Conservar filas que NO participan en la consolidación
    df_original = df_dem[~mask_bra].copy()
    
    # 6) Extraer y agrupar las filas BRA
    df_bra = df_dem[mask_bra].copy()
    # Clave base: primeros 6 caracteres (identificador sin región)
    df_bra["TechKey"] = df_bra["Tech"].str[:6]
    
    # 7) Construir nuevas filas con BRAXX
    new_rows = []
    for (tech_key, parameter), group in df_bra.groupby(["TechKey", "Parameter"]):
        base = group.iloc[0].copy()
        # Nuevo código: TechKey + "XX"
        base["Tech"] = tech_key + "BRAXX"
        # Actualizar Tech.Name (insensible a mayúsculas)
        base["Tech.Name"] = re.sub(
            r"Brazil, region [A-Z]{2}",
            "Brazil, region XX",
            base["Tech.Name"],
            flags=re.IGNORECASE
        )
        # Calcular valores 2021–2050
        if parameter in parameters_avg:
            vals = group[year_cols].astype(float).mean()
        else:  # parameters_sum
            vals = group[year_cols].astype(float).sum(min_count=1)
        base[year_cols] = vals
        # Projection.Mode según presencia de datos numéricos
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_rows.append(base)
    
    # 8) Combinar originales + nuevas filas
    df_dem_final = pd.concat([df_original, pd.DataFrame(new_rows)], ignore_index=True)
    
    # 9) Eliminar columnas auxiliares
    df_sec_final.drop(columns=["TechKey"], errors="ignore", inplace=True)
    
    # 9) Reasignar Tech.ID agrupado por Tech
    unique_dem = pd.unique(df_dem_final["Tech"])
    id_map = {tech: i + 1 for i, tech in enumerate(unique_dem)}
    df_dem_final["Tech.ID"] = df_dem_final["Tech"].map(id_map)
    
    print("Hoja 'Demand Techs' procesada correctamente.")
    
    
    
    
    
    
    
    
    
    
    
    # ——— Procesamiento de la hoja “Capacities” ———
    
    # Leer la hoja
    df_cap = xls.parse("Capacities")
    
    # 1) Parámetro único y operación (promedio)
    parameters_avg = ["CapacityFactor"]
    
    # 2) Columnas de años (2021–2050), int o str
    year_cols = [
        c for c in df_cap.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]
    
    # 3) Identificar filas PWR…BRA… (región BRA en pos. 6–9, región válida en 9–11)
    brazil_regions = ["CN", "NW", "NE", "CW", "SO", "SE", "WE"]
    mask_bra = (
        df_cap["Tech"].str.startswith("PWR") &
        df_cap["Tech"].str[6:9].eq("BRA") &
        df_cap["Tech"].str[9:11].isin(brazil_regions)
    )
    
    # 4) Conservar filas que NO participan
    df_cap_orig = df_cap[~mask_bra].copy()
    
    # 5) Extraer y agrupar las filas BRA
    df_cap_bra = df_cap[mask_bra].copy()
    df_cap_bra["TechKey"] = df_cap_bra["Tech"].str[:6]
    
    # 6) Crear nuevas filas BRAXX por TechKey + Timeslices + Parameter
    new_cap_rows = []
    for (tech_key, timeslice, parameter), group in df_cap_bra.groupby(
            ["TechKey", "Timeslices", "Parameter"]
        ):
        base = group.iloc[0].copy()
        # Extraer los dos últimos caracteres originales
        suffix = base["Tech"][-2:]
        # Nuevo Tech: TechKey + "BRAXX" + suffix
        base["Tech"] = f"{tech_key}BRAXX{suffix}"
        # Actualizar Tech.Name si existe
        tn = base.get("Tech.Name", "")
        if pd.notna(tn):
            base["Tech.Name"] = re.sub(
                r"Brazil, region [A-Z]{2}",
                "Brazil, region XX",
                str(tn),
                flags=re.IGNORECASE
            )
        # Promedio de CapacityFactor en años
        base[year_cols] = group[year_cols].astype(float).mean()
        # Projection.Mode
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_cap_rows.append(base)
    
    # 7) Unir originales + nuevas filas
    df_cap_final = pd.concat([df_cap_orig, pd.DataFrame(new_cap_rows)], ignore_index=True)
    
    # 8) Eliminar columna auxiliar
    df_cap_final.drop(columns=["TechKey"], errors="ignore", inplace=True)
    
    # 9) Reasignar Tech.ID agrupado por Tech
    unique_caps = pd.unique(df_cap_final["Tech"])
    id_map = {tech: i + 1 for i, tech in enumerate(unique_caps)}
    df_cap_final["Tech.ID"] = df_cap_final["Tech"].map(id_map)
    
    print("Hoja 'Capacities' procesada correctamente.")
    # ——— Fin de Capacities ———
    
    
    
    
    # ——— Procesamiento de la hoja “VariableCost” ———
    
    # Leer la hoja
    df_var = xls.parse("VariableCost")
    
    # 0) Eliminar filas que tengan 'BRA' dos veces en Tech
    mask_two_bra = df_var["Tech"].str.count("BRA") > 1
    df_var = df_var[~mask_two_bra].copy()
    
    # 1) Parámetro único y operación (promedio)
    parameters_avg = ["VariableCost"]
    
    # 2) Columnas de años (2021–2050), int o str
    year_cols = [
        c for c in df_var.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]
    
    # 3) Máscaras para las tres estructuras
    brazil_regions = ["CN","NW","NE","CW","SO","SE","WE"]
    
    # PWR...BRA...XX## (backstop y general PWR)
    mask_pwr_bra = (
        df_var["Tech"].str.startswith("PWR") &
        df_var["Tech"].str[6:9].eq("BRA") &
        df_var["Tech"].str[9:11].isin(brazil_regions) &
        ~df_var["Tech"].str.contains("BCK")
    )
    
    # TRN interconexiones
    mask_trn = (
        df_var["Tech"].str.startswith("TRN") &
        df_var["Tech"].str.contains("BRA") &
        (df_var["Tech"].str.len() == 13)
    )
    
    # PWRBCK...BRA... (backstop)
    mask_bra_bck = (
        df_var["Tech"].str.startswith("PWRBCK") &
        df_var["Tech"].str[6:11].isin([f"BRA{r}" for r in brazil_regions])
    )
    
    used_mask = mask_pwr_bra | mask_trn | mask_bra_bck
    
    # 4) Conservar filas que NO participan
    df_orig = df_var[~used_mask].copy()
    
    new_rows = []
    
    # 5.1) PWR...BRA... → consolidación BRAXX## por Mode.Operation + Parameter
    df_p = df_var[mask_pwr_bra].copy()
    df_p["TechKey"] = df_p["Tech"].str[:6]
    for (tech_key, mode_op, parameter), group in df_p.groupby(["TechKey","Mode.Operation","Parameter"]):
        base = group.iloc[0].copy()
        suffix = base["Tech"][-2:]
        base["Tech"] = f"{tech_key}BRAXX{suffix}"
        if pd.notna(base["Tech.Name"]):
            base["Tech.Name"] = re.sub(
                r"Brazil, region [A-Z]{2}",
                "Brazil, region XX",
                str(base["Tech.Name"]),
                flags=re.IGNORECASE
            )
        base[year_cols] = group[year_cols].astype(float).mean()
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_rows.append(base)
    
    # 5.2) TRN…BRA interconexiones → consolidación BRAXX por NormalizedTech + Mode.Operation + Parameter
    df_t = df_var[mask_trn].copy()
    def normalize_interconnection(code):
        p1, p2 = code[3:8], code[8:13]
        n1 = "BRAXX" if "BRA" in p1 else p1
        n2 = "BRAXX" if "BRA" in p2 else p2
        return "TRN" + "".join(sorted([n1, n2]))
    def update_bra_region(code):
        p1, p2 = code[3:8], code[8:13]
        if "BRA" in p1: p1 = "BRAXX"
        if "BRA" in p2: p2 = "BRAXX"
        return "TRN" + p1 + p2
    
    df_t["NormKey"] = df_t["Tech"].apply(normalize_interconnection)
    for (norm_key, mode_op, parameter), group in df_t.groupby(["NormKey","Mode.Operation","Parameter"]):
        base = group.iloc[0].copy()
        base["Tech"] = update_bra_region(base["Tech"])
        if pd.notna(base["Tech.Name"]):
            base["Tech.Name"] = re.sub(
                r"Brazil, region [A-Z]{2}",
                "Brazil, region XX",
                str(base["Tech.Name"]),
                flags=re.IGNORECASE
            )
        base[year_cols] = group[year_cols].astype(float).mean()
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_rows.append(base)
    
    # 5.3) PWRBCK…BRA… → consolidación BRAXX por TechKey + Mode.Operation + Parameter
    df_b = df_var[mask_bra_bck].copy()
    df_b["TechKey"] = df_b["Tech"].str[:6]
    for (tech_key, mode_op, parameter), group in df_b.groupby(["TechKey","Mode.Operation","Parameter"]):
        base = group.iloc[0].copy()
        # Aquí quitamos el sufijo completamente:
        base["Tech"] = f"{tech_key}BRAXX"
        if pd.notna(base["Tech.Name"]):
            base["Tech.Name"] = re.sub(
                r"Brazil, region [A-Z]{2}",
                "Brazil, region XX",
                str(base["Tech.Name"]),
                flags=re.IGNORECASE
            )
        base[year_cols] = group[year_cols].astype(float).mean()
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_rows.append(base)
    
    # 6) Unir originales + nuevas filas
    df_var_final = pd.concat([df_orig, pd.DataFrame(new_rows)], ignore_index=True)
    
    # 7) Eliminar columnas auxiliares
    df_var_final.drop(columns=["TechKey","NormKey"], errors="ignore", inplace=True)
    
    # 8) Reasignar Tech.ID agrupado por Tech
    unique_vars = pd.unique(df_var_final["Tech"])
    id_map = {tech: i+1 for i, tech in enumerate(unique_vars)}
    df_var_final["Tech.ID"] = df_var_final["Tech"].map(id_map)
    
    print("Hoja 'VariableCost' procesada correctamente.")
    
    # ——— Fin de VariableCost ———
    
    
    
    #  (1) Cargar lista de todas las hojas originales
    all_sheets = xls.sheet_names  # hereda tu xls = pd.ExcelFile(...) del principio
    
    #  (2) Escribir un nuevo archivo, iterando sobre cada hoja
    output_file = os.path.join(f"A1_Outputs_{folder}","A-O_Parametrization_cleaned.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet in all_sheets:
            if sheet == "Fixed Horizon Parameters":
                df_fixed_final.to_excel(writer, sheet_name=sheet, index=False)
            elif sheet == "Secondary Techs":
                df_sec_final.to_excel(writer, sheet_name=sheet, index=False)
            elif sheet == "Demand Techs":
                df_dem_final.to_excel(writer, sheet_name=sheet, index=False)
            elif sheet == "Capacities":
                df_cap_final.to_excel(writer, sheet_name=sheet, index=False)
            elif sheet == "VariableCost":
                df_var_final.to_excel(writer, sheet_name=sheet, index=False)
            else:
                # Carga y escribe tal cual la hoja original
                df_orig = xls.parse(sheet)
                df_orig.to_excel(writer, sheet_name=sheet, index=False)
    
    print(f"✔ Archivo guardado en {output_file}")



if demand:
    # Ruta del archivo original
    file_path = os.path.join(f"A1_Outputs_{folder}","A-O_Demand.xlsx")
    
    # Carga todas las hojas
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    
    # Procesar únicamente “Demand_Projection”
    df_proj = xls.parse("Demand_Projection")
    
    # 1) Detectar columnas de años (2021–2050)
    year_cols = [
        c for c in df_proj.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]
    
    # 2) Mascara para líneas BRA…rr## en Fuel/Tech
    pattern = r"^.{3}BRA(?:CN|NW|NE|CW|SO|SE|WE)\d{2}$"
    mask_bra = df_proj["Fuel/Tech"].str.contains(pattern, regex=True, na=False)
    
    # 3) Separar originales y brasileños
    df_orig = df_proj[~mask_bra].copy()
    df_bra  = df_proj[mask_bra].copy()
    
    # 4) Preparar clave de agrupación
    df_bra["TechKey"] = df_bra["Fuel/Tech"].str[:3]    # prefijo (p.ej. "ELC")
    df_bra["Suffix"]  = df_bra["Fuel/Tech"].str[-2:]   # sufijo numérico
    
    # 5) Agrupar y sumar años, crear BRAXX
    new_rows = []
    for (tk, suf), group in df_bra.groupby(["TechKey","Suffix"]):
        base = group.iloc[0].copy()
        # Nuevo código Fuel/Tech
        base["Fuel/Tech"] = f"{tk}BRAXX{suf}"
        # Limpiar Name (quitar “, region XX”)
        base["Name"] = re.sub(r",\s*region\s+[A-Z]{2}$", "", base["Name"], flags=re.IGNORECASE)
        # Sumar todos los year_cols
        base[year_cols] = group[year_cols].astype(float).sum()
        new_rows.append(base)
    
    # 6) Unir y limpiar
    df_proj_clean = pd.concat([df_orig, pd.DataFrame(new_rows)], ignore_index=True)
    df_proj_clean.drop(columns=["TechKey","Suffix"], errors="ignore", inplace=True)
    
    
    print("Hoja 'Demand_Projection' procesada correctamente.")
    
    # ——— Procesamiento de la hoja “Profiles” ———
    
    # 1) Leer la hoja
    df_profiles = xls.parse("Profiles")
    
    # 2) Detectar columnas de años 2021–2050 (int o str)
    year_cols = [
        c for c in df_profiles.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]
    
    # 3) Mascara para filas BRA…rr## en Fuel/Tech
    #    ^.{3}    cualquier prefijo de 3 chars
    #    BRA      literal BRA
    #    (CN|NW|…)
    #    \d{2}$   sufijo numérico de 2 dígitos
    pattern = r"^.{3}BRA(?:CN|NW|NE|CW|SO|SE|WE)\d{2}$"
    mask_bra = df_profiles["Fuel/Tech"].str.contains(pattern, regex=True, na=False)
    
    # 4) Separar filas no brasileñas y brasileñas
    df_orig  = df_profiles[~mask_bra].copy()
    df_bra   = df_profiles[mask_bra].copy()
    
    # 5) Extraer claves de agrupación
    df_bra["TechKey"] = df_bra["Fuel/Tech"].str[:3]     # prefijo
    df_bra["Suffix"]  = df_bra["Fuel/Tech"].str[-2:]    # código final
    
    # 6) Agrupar por TechKey, Suffix y Timeslices
    new_rows = []
    for (tk, suf, ts), group in df_bra.groupby(["TechKey", "Suffix", "Timeslices"]):
        base = group.iloc[0].copy()
        # Nuevo Fuel/Tech consolidado
        base["Fuel/Tech"] = f"{tk}BRAXX{suf}"
        # Limpiar Name: quitar ", region XX"
        base["Name"] = re.sub(
            r",\s*region\s+[A-Z]{2}$",
            "",
            str(base["Name"]),
            flags=re.IGNORECASE
        )
        # Sumar los valores de los años
        base[year_cols] = group[year_cols].astype(float).mean()
        new_rows.append(base)
    
    # 7) Concatenar originales + filas consolidadas
    df_profiles_clean = pd.concat([df_orig, pd.DataFrame(new_rows)], ignore_index=True)
    
    # 8) Eliminar columnas auxiliares
    df_profiles_clean.drop(columns=["TechKey", "Suffix"], errors="ignore", inplace=True)
    
    # 9) Reasignar Tech.ID por valor único de Fuel/Tech
    unique_codes = pd.unique(df_profiles_clean["Fuel/Tech"])
    id_map = {code: i+1 for i, code in enumerate(unique_codes)}
    df_profiles_clean["Tech.ID"] = df_profiles_clean["Fuel/Tech"].map(id_map)
    
    print("Hoja 'Profiles' procesada correctamente.")
    
    # ——— Fin de procesamiento de “Profiles” ———
    
    # 10) Guardar todas las hojas en un nuevo libro
    output_file = os.path.join(f"A1_Outputs_{folder}","A-O_Demand.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet in sheet_names:
            if sheet == "Profiles":
                df_profiles_clean.to_excel(writer, sheet_name=sheet, index=False)
            elif sheet == "Demand_Projection":
                # Asumimos que df_proj_clean ya está preparado
                df_proj_clean.to_excel(writer, sheet_name=sheet, index=False)
            else:
                xls.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)
    
    print(f"✔ Archivo guardado en {output_file}")



if storage:
    # Ruta del archivo
    file_path = os.path.join("A2_Extra_Inputs","A-Xtra_Storage.xlsx")
    xls = pd.ExcelFile(file_path)
    
    # Función auxiliar para detectar columnas de años 2021–2050
    def get_year_cols(df):
        return [
            c for c in df.columns
            if (isinstance(c, int) and 2021 <= c <= 2050)
                or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
        ]
    
    # Lista de regiones de Brasil
    brazil_regions = ["CN","NW","NE","CW","SO","SE","WE"]
    bad_regions = ["BRANW","BRANE","BRACW","BRASO","BRASE","BRAWE"]
    
    # ===== 1) Fixed Horizon Parameters =====
    df_fhp = xls.parse("Fixed Horizon Parameters")
    
    # 1.1) Eliminar duplicados internos BRA–BRA
    mask_len13   = df_fhp["STORAGE"].str.len() == 13
    mask_two_bra = df_fhp["STORAGE"].str.count("BRA") > 1
    df_fhp = df_fhp[~(mask_len13 & mask_two_bra)].copy()
    
    # 1.2) Reemplazar BRACN→BRAXX y ajustar STORAGE.Name
    mask_bracn = df_fhp["STORAGE"].str.contains("BRACN", na=False)
    df_fhp.loc[mask_bracn, "STORAGE"] = df_fhp.loc[mask_bracn, "STORAGE"].str.replace("BRACN","BRAXX", regex=False)
    df_fhp.loc[mask_bracn, "STORAGE.Name"] = df_fhp.loc[mask_bracn, "STORAGE.Name"].apply(
        lambda tn: re.sub(r"CN$", "XX", tn) if pd.notna(tn) else tn
    )
    
    # 1.3) Eliminar demás regiones brasileñas
    df_fhp = df_fhp[~df_fhp["STORAGE"].str[3:8].isin(bad_regions)].copy()
    
    # 1.4) Unificar interconexiones TRN…BRA…XX
    def normalize_trn(code):
        p1,p2 = code[3:8], code[8:13]
        n1 = "BRAXX" if "BRA" in p1 else p1
        n2 = "BRAXX" if "BRA" in p2 else p2
        return "TRN" + "".join(sorted([n1, n2]))
    
    def update_trn(code):
        p1,p2 = code[3:8], code[8:13]
        if "BRA" in p1: p1="BRAXX"
        if "BRA" in p2: p2="BRAXX"
        return "TRN" + p1 + p2
    
    df_inter = df_fhp[df_fhp["STORAGE"].str.startswith("TRN") & df_fhp["STORAGE"].str.len()==13].copy()
    df_inter["NormKey"] = df_inter["STORAGE"].apply(normalize_trn)
    df_inter_dedup = df_inter.drop_duplicates(subset=["NormKey"]).copy()
    df_inter_dedup["STORAGE"] = df_inter_dedup["STORAGE"].apply(update_trn)
    df_inter_dedup["STORAGE.Name"] = df_inter_dedup["STORAGE.Name"].apply(
        lambda tn: re.sub(r"Brazil, region [A-Z]{2}", "Brazil, region XX", tn, flags=re.IGNORECASE)
    )
    
    mask_trn = df_fhp["STORAGE"].str.startswith("TRN") & df_fhp["STORAGE"].str.len()==13
    df_fhp = pd.concat([df_fhp[~mask_trn], df_inter_dedup.drop(columns=["NormKey"])], ignore_index=True)
    
    # ——— Reasignar STORAGE.ID en Fixed Horizon Parameters ———
    unique_fhp = pd.unique(df_fhp["STORAGE"])
    fhp_id_map = {stor: i+1 for i, stor in enumerate(unique_fhp)}
    df_fhp["STORAGE.ID"] = df_fhp["STORAGE"].map(fhp_id_map)
    
    print("Hoja 'Fixed Horizon Parameters' procesada correctamente.")
    
    # ——— 2) CapitalCostStorage ———
    df_ccs = xls.parse("CapitalCostStorage")
    year_cols = [
        c for c in df_ccs.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
            or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]
    
    # Detectar sólo las filas brasileñas (no backstop)
    mask_ccs_bra = (
        df_ccs["STORAGE"].str[3:6].eq("BRA") &
        df_ccs["STORAGE"].str[6:8].isin(brazil_regions)
    )
    
    # Separar originales y brasileñas
    orig_ccs = df_ccs[~mask_ccs_bra].copy()
    bra_ccs  = df_ccs[mask_ccs_bra].copy()
    bra_ccs["Key"] = bra_ccs["STORAGE"].str[:3]
    
    new_ccs = []
    for (key, param), g in bra_ccs.groupby(["Key","Parameter"]):
        base = g.iloc[0].copy()
        suffix = base["STORAGE"][-2:]
        base["STORAGE"] = f"{key}BRAXX{suffix}"
        # Actualizar STORAGE.Name
        tn = base.get("STORAGE.Name", "")
        if pd.notna(tn):
            base["STORAGE.Name"] = re.sub(
                r"Brazil, region [A-Z]{2}",
                "Brazil, region XX",
                str(tn),
                flags=re.IGNORECASE
            )
        # Calcular valores de 2021–2050
        if param.lower().startswith("capital"):
            vals = g[year_cols].astype(float).mean()
        else:
            vals = g[year_cols].astype(float).sum(min_count=1)
        base[year_cols] = vals
        # Projection.Mode: User defined si hay al menos un valor no-NaN
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_ccs.append(base)
    
    df_ccs_clean = pd.concat([orig_ccs, pd.DataFrame(new_ccs)], ignore_index=True)
    df_ccs_clean.drop(columns=["Key"], errors="ignore", inplace=True)
    
    # ——— Reasignar STORAGE.ID en CapitalCostStorage ———
    unique_ccs = pd.unique(df_ccs_clean["STORAGE"])
    ccs_id_map = {stor: i+1 for i, stor in enumerate(unique_ccs)}
    df_ccs_clean["STORAGE.ID"] = df_ccs_clean["STORAGE"].map(ccs_id_map)
    
    print("Hoja 'CapitalCostStorage' procesada correctamente.")
    
    # ——— 3) TechnologyStorage ———
    df_ts = xls.parse("TechnologyStorage")
    year_cols = [c for c in df_ts.columns if (isinstance(c,int) and 2021<=c<=2050) or (isinstance(c,str) and c.isdigit() and 2021<=int(c)<=2050)]
    
    # 3.1) Quitar duplicados BRA–BRA
    mask_dupl = (df_ts["TECHNOLOGY"].str.len()==13) & (df_ts["TECHNOLOGY"].str.count("BRA")>1)
    df_ts = df_ts[~mask_dupl].copy()
    
    # 3.2) Detectar filas BRA sin backstop
    mask_ts_bra = (
        df_ts["TECHNOLOGY"].str[6:9].eq("BRA") &
        df_ts["TECHNOLOGY"].str[9:11].isin(brazil_regions)
    )
    
    orig_ts = df_ts[~mask_ts_bra].copy()
    bra_ts  = df_ts[mask_ts_bra].copy()
    
    # **Clave = primeros 6 caracteres** (p. ej. "PWRLDS", "PWRSDS")
    bra_ts["TechKey"] = bra_ts["TECHNOLOGY"].str[:6]
    
    # Define operaciones
    parameters_avg = ["TechnologyToStorage","TechnologyFromStorage"]
    parameters_sum = []  # añade aquí los que deban sumarse
    
    new_ts = []
    for (tk, mode, param), g in bra_ts.groupby(["TechKey","MODE_OF_OPERATION","Parameter"]):
        base = g.iloc[0].copy()
        suffix = base["TECHNOLOGY"][-2:]
        base["TECHNOLOGY"] = f"{tk}BRAXX{suffix}"  # p. ej. "PWRLDSBRAXX01"
        if pd.notna(base["TECHNOLOGY.Name"]):
            base["TECHNOLOGY.Name"] = re.sub(
                r"Brazil, region [A-Z]{2}",
                "Brazil, region XX",
                str(base["TECHNOLOGY.Name"]),
                flags=re.IGNORECASE
            )
        if param in parameters_avg:
            vals = g[year_cols].astype(float).mean()
        else:
            vals = g[year_cols].astype(float).sum(min_count=1)
        base[year_cols] = vals
        new_ts.append(base)
    
    df_ts_clean = pd.concat([orig_ts, pd.DataFrame(new_ts)], ignore_index=True)
    df_ts_clean.drop(columns=["TechKey"], errors="ignore", inplace=True)
    print("Hoja 'TechnologyStorage' procesada correctamente.")
    
    # ===== Guardar todas las hojas en nuevo libro =====
    output_file = os.path.join("A2_Extra_Inputs","A-Xtra_Storage.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_fhp.to_excel(writer, sheet_name="Fixed Horizon Parameters", index=False)
        df_ccs_clean.to_excel(writer, sheet_name="CapitalCostStorage", index=False)
        df_ts_clean.to_excel(writer, sheet_name="TechnologyStorage", index=False)
    
    print(f"✔ Archivo guardado en {output_file}")

