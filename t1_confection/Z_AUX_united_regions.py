# -*- coding: utf-8 -*-
"""
Created on Tue Jul 29 11:20:29 2025

@author: Climate Lead Group, Andrey Salazar-Vargas
"""

import pandas as pd
import os
import re

parametrization = False
demand = True
storage = False

folder = 'BAU'



if parametrization:
    # Load Excel file
    file_path = os.path.join(f"A1_Outputs_{folder}","A-O_Parametrization.xlsx")  # Change this if the file is in another location
    xls = pd.ExcelFile(file_path)

    # Read the sheet
    df_fixed = xls.parse("Fixed Horizon Parameters")

    # --- Part 1: Replace BRACN with BRAXX ---
    mask_bracn = df_fixed["Tech"].str.contains("BRACN", na=False)
    df_fixed.loc[mask_bracn, "Tech"] = df_fixed.loc[mask_bracn, "Tech"].str.replace("BRACN", "BRAXX", regex=False)
    df_fixed.loc[mask_bracn, "Tech.Name"] = df_fixed.loc[mask_bracn, "Tech.Name"].str.replace("CN", "XX", regex=False)

    # --- Part 2: Remove unwanted BRA regional combinations ---
    brazil_bad_regions = ["BRANW", "BRANE", "BRACW", "BRASO", "BRASE", "BRAWE"]
    mask_bad_regions = df_fixed["Tech"].str[6:11].isin(brazil_bad_regions)
    df_fixed = df_fixed[~mask_bad_regions]

    # --- Part 3: Remove technologies of length 13 where "BRA" appears twice ---
    is_length_13 = df_fixed["Tech"].str.len() == 13
    has_two_bra  = df_fixed["Tech"].str.count("BRA") > 1
    df_fixed = df_fixed[~(is_length_13 & has_two_bra)]

    # --- Part 4: Unify ONLY the BRA<->other country interconnections and change BRA region to XX ---

    # 4.1) Mask to filter only TRN... that contain BRA
    mask_bra_trn = (
        df_fixed["Tech"].str.startswith("TRN") &
        df_fixed["Tech"].str.contains("BRA") &
        (df_fixed["Tech"].str.len() == 13)
    )
    df_inter = df_fixed[mask_bra_trn].copy()

    # 4.2) Normalization functions
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

    # 4.3) Group by normalized interconnection and by parameter
    df_inter["NormKey"] = df_inter["Tech"].apply(normalize_interconnection)
    new_trn_rows = []
    for (norm_key, parameter), group in df_inter.groupby(["NormKey", "Parameter"]):
        base = group.iloc[0].copy()
        # Generate the new code and name
        base["Tech"]      = update_bra_region(base["Tech"])
        base["Tech.Name"] = re.sub(
            r"Brazil, region [A-Z]{2}",
            "Brazil, region XX",
            base["Tech.Name"],
            flags=re.IGNORECASE
        )
        # Keep the rest of the columns as is (Parameter.ID, Unit, years, etc.)
        new_trn_rows.append(base)

    # 4.4) Rebuild df_fixed_final: remove the original BRA-TRN rows and add the new ones
    df_fixed_final = pd.concat([
        df_fixed[~mask_bra_trn],
        pd.DataFrame(new_trn_rows).drop(columns=["NormKey"])
    ], ignore_index=True)

    # --- Reassign Tech.ID grouped by Tech ---
    unique_fixed = pd.unique(df_fixed_final["Tech"])
    fixed_id_map = {tech: i+1 for i, tech in enumerate(unique_fixed)}
    df_fixed_final["Tech.ID"] = df_fixed_final["Tech"].map(fixed_id_map)

    print("Sheet 'Fixed Horizon Parameters' processed successfully.")



    df_sec = xls.parse("Secondary Techs")

    # 2) Remove internal BRA-BRA duplicates
    mask_len13     = df_sec["Tech"].str.len() == 13
    mask_bra_twice = df_sec["Tech"].str.count("BRA") > 1
    df_sec = df_sec[~(mask_len13 & mask_bra_twice)].copy()

    # 3) Define regions and parameters
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

    # 4) Year columns 2021-2050
    year_cols = [
        c for c in df_sec.columns
        if (isinstance(c,int)   and 2021 <= c <= 2050)
           or (isinstance(c,str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]

    # 5) Identify rows used to generate BRAXX
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

    # 6) Keep only rows that were NOT used to calculate BRAXX
    df_original = df_sec[~used_mask].copy()

    new_rows = []

    # 7.1) PWR...BRA...00/01 -> BRAXX
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

    # 7.2) TRN...BRA interconnections -> BRAXX
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

    # 7.3) ELC...BRA...01 -> BRAXX
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

    # 7.4) PWRBCK...BRA... -> BRAXX
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

    # 8) Concatenate originals + new BRAXX rows
    df_sec_final = pd.concat([df_original, pd.DataFrame(new_rows)], ignore_index=True)

    # 9) Remove auxiliary columns
    df_sec_final.drop(columns=["TechKey", "NormKey", "ElcKey"], errors="ignore", inplace=True)

    # --- Reassign Tech.ID grouped by Tech ---
    unique_sec = pd.unique(df_sec_final["Tech"])
    sec_id_map = {tech: i+1 for i, tech in enumerate(unique_sec)}
    df_sec_final["Tech.ID"] = df_sec_final["Tech"].map(sec_id_map)

    print("Sheet 'Secondary Techs' processed successfully.")
    # ... after calculating df_fixed_final and df_sec_final ...





    df_dem = xls.parse("Demand Techs")

    # 2) Define parameters and operations
    parameters_avg = ["CapitalCost", "FixedCost"]
    parameters_sum = [
        "ResidualCapacity",
        "TotalAnnualMinCapacityInvestment",
        "TotalAnnualMaxCapacity"
    ]

    # 3) Detect year columns (2021-2050), whether int or str
    year_cols = [
        c for c in df_dem.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]

    # 4) Detect BRA...RR rows with regular expression
    #     - .{6} matches any character in positions 0-5
    #     - BRA in positions 6-8
    #     - one of the region suffixes in 9-10
    brazil_pattern = r"^.{6}BRA(?:CN|NW|NE|CW|SO|SE|WE)$"
    mask_bra = df_dem["Tech"].str.contains(
        brazil_pattern,
        regex=True,
        na=False,
        flags=re.IGNORECASE
    )

    # 5) Keep rows that do NOT participate in the consolidation
    df_original = df_dem[~mask_bra].copy()

    # 6) Extract and group the BRA rows
    df_bra = df_dem[mask_bra].copy()
    # Base key: first 6 characters (identifier without region)
    df_bra["TechKey"] = df_bra["Tech"].str[:6]

    # 7) Build new rows with BRAXX
    new_rows = []
    for (tech_key, parameter), group in df_bra.groupby(["TechKey", "Parameter"]):
        base = group.iloc[0].copy()
        # New code: TechKey + "XX"
        base["Tech"] = tech_key + "BRAXX"
        # Update Tech.Name (case insensitive)
        base["Tech.Name"] = re.sub(
            r"Brazil, region [A-Z]{2}",
            "Brazil, region XX",
            base["Tech.Name"],
            flags=re.IGNORECASE
        )
        # Calculate values 2021-2050
        if parameter in parameters_avg:
            vals = group[year_cols].astype(float).mean()
        else:  # parameters_sum
            vals = group[year_cols].astype(float).sum(min_count=1)
        base[year_cols] = vals
        # Projection.Mode based on presence of numeric data
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_rows.append(base)

    # 8) Combine originals + new rows
    df_dem_final = pd.concat([df_original, pd.DataFrame(new_rows)], ignore_index=True)

    # 9) Remove auxiliary columns
    df_sec_final.drop(columns=["TechKey"], errors="ignore", inplace=True)

    # 9) Reassign Tech.ID grouped by Tech
    unique_dem = pd.unique(df_dem_final["Tech"])
    id_map = {tech: i + 1 for i, tech in enumerate(unique_dem)}
    df_dem_final["Tech.ID"] = df_dem_final["Tech"].map(id_map)

    print("Sheet 'Demand Techs' processed successfully.")











    # --- Processing the "Capacities" sheet ---

    # Read the sheet
    df_cap = xls.parse("Capacities")

    # 1) Single parameter and operation (average)
    parameters_avg = ["CapacityFactor"]

    # 2) Year columns (2021-2050), int or str
    year_cols = [
        c for c in df_cap.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]

    # 3) Identify PWR...BRA... rows (BRA region in pos. 6-9, valid region in 9-11)
    brazil_regions = ["CN", "NW", "NE", "CW", "SO", "SE", "WE"]
    mask_bra = (
        df_cap["Tech"].str.startswith("PWR") &
        df_cap["Tech"].str[6:9].eq("BRA") &
        df_cap["Tech"].str[9:11].isin(brazil_regions)
    )

    # 4) Keep rows that do NOT participate
    df_cap_orig = df_cap[~mask_bra].copy()

    # 5) Extract and group the BRA rows
    df_cap_bra = df_cap[mask_bra].copy()
    df_cap_bra["TechKey"] = df_cap_bra["Tech"].str[:6]

    # 6) Create new BRAXX rows by TechKey + Timeslices + Parameter
    new_cap_rows = []
    for (tech_key, timeslice, parameter), group in df_cap_bra.groupby(
            ["TechKey", "Timeslices", "Parameter"]
        ):
        base = group.iloc[0].copy()
        # Extract the last two original characters
        suffix = base["Tech"][-2:]
        # New Tech: TechKey + "BRAXX" + suffix
        base["Tech"] = f"{tech_key}BRAXX{suffix}"
        # Update Tech.Name if it exists
        tn = base.get("Tech.Name", "")
        if pd.notna(tn):
            base["Tech.Name"] = re.sub(
                r"Brazil, region [A-Z]{2}",
                "Brazil, region XX",
                str(tn),
                flags=re.IGNORECASE
            )
        # Average of CapacityFactor across years
        base[year_cols] = group[year_cols].astype(float).mean()
        # Projection.Mode
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_cap_rows.append(base)

    # 7) Merge originals + new rows
    df_cap_final = pd.concat([df_cap_orig, pd.DataFrame(new_cap_rows)], ignore_index=True)

    # 8) Remove auxiliary column
    df_cap_final.drop(columns=["TechKey"], errors="ignore", inplace=True)

    # 9) Reassign Tech.ID grouped by Tech
    unique_caps = pd.unique(df_cap_final["Tech"])
    id_map = {tech: i + 1 for i, tech in enumerate(unique_caps)}
    df_cap_final["Tech.ID"] = df_cap_final["Tech"].map(id_map)

    print("Sheet 'Capacities' processed successfully.")
    # --- End of Capacities ---




    # --- Processing the "VariableCost" sheet ---

    # Read the sheet
    df_var = xls.parse("VariableCost")

    # 0) Remove rows that have 'BRA' twice in Tech
    mask_two_bra = df_var["Tech"].str.count("BRA") > 1
    df_var = df_var[~mask_two_bra].copy()

    # 1) Single parameter and operation (average)
    parameters_avg = ["VariableCost"]

    # 2) Year columns (2021-2050), int or str
    year_cols = [
        c for c in df_var.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]

    # 3) Masks for the three structures
    brazil_regions = ["CN","NW","NE","CW","SO","SE","WE"]

    # PWR...BRA...XX## (backstop and general PWR)
    mask_pwr_bra = (
        df_var["Tech"].str.startswith("PWR") &
        df_var["Tech"].str[6:9].eq("BRA") &
        df_var["Tech"].str[9:11].isin(brazil_regions) &
        ~df_var["Tech"].str.contains("BCK")
    )

    # TRN interconnections
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

    # 4) Keep rows that do NOT participate
    df_orig = df_var[~used_mask].copy()

    new_rows = []

    # 5.1) PWR...BRA... -> consolidation BRAXX## by Mode.Operation + Parameter
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

    # 5.2) TRN...BRA interconnections -> consolidation BRAXX by NormalizedTech + Mode.Operation + Parameter
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

    # 5.3) PWRBCK...BRA... -> consolidation BRAXX by TechKey + Mode.Operation + Parameter
    df_b = df_var[mask_bra_bck].copy()
    df_b["TechKey"] = df_b["Tech"].str[:6]
    for (tech_key, mode_op, parameter), group in df_b.groupby(["TechKey","Mode.Operation","Parameter"]):
        base = group.iloc[0].copy()
        # Here we remove the suffix completely:
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

    # 6) Merge originals + new rows
    df_var_final = pd.concat([df_orig, pd.DataFrame(new_rows)], ignore_index=True)

    # 7) Remove auxiliary columns
    df_var_final.drop(columns=["TechKey","NormKey"], errors="ignore", inplace=True)

    # 8) Reassign Tech.ID grouped by Tech
    unique_vars = pd.unique(df_var_final["Tech"])
    id_map = {tech: i+1 for i, tech in enumerate(unique_vars)}
    df_var_final["Tech.ID"] = df_var_final["Tech"].map(id_map)

    print("Sheet 'VariableCost' processed successfully.")

    # --- End of VariableCost ---



    #  (1) Load list of all original sheets
    all_sheets = xls.sheet_names  # inherits your xls = pd.ExcelFile(...) from the beginning

    #  (2) Write a new file, iterating over each sheet
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
                # Load and write the original sheet as is
                df_orig = xls.parse(sheet)
                df_orig.to_excel(writer, sheet_name=sheet, index=False)

    print(f"File saved to {output_file}")



if demand:
    # Path to the original file
    file_path = os.path.join(f"A1_Outputs_{folder}","A-O_Demand.xlsx")

    # Load all sheets
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names

    # Process only "Demand_Projection"
    df_proj = xls.parse("Demand_Projection")

    # 1) Detect year columns (2021-2050)
    year_cols = [
        c for c in df_proj.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]

    # 2) Mask for BRA...rr## lines in Fuel/Tech
    pattern = r"^.{3}BRA(?:CN|NW|NE|CW|SO|SE|WE)\d{2}$"
    mask_bra = df_proj["Fuel/Tech"].str.contains(pattern, regex=True, na=False)

    # 3) Separate original and Brazilian rows
    df_orig = df_proj[~mask_bra].copy()
    df_bra  = df_proj[mask_bra].copy()

    # 4) Prepare grouping key
    df_bra["TechKey"] = df_bra["Fuel/Tech"].str[:3]    # prefix (e.g. "ELC")
    df_bra["Suffix"]  = df_bra["Fuel/Tech"].str[-2:]   # numeric suffix

    # 5) Group and sum years, create BRAXX
    new_rows = []
    for (tk, suf), group in df_bra.groupby(["TechKey","Suffix"]):
        base = group.iloc[0].copy()
        # New Fuel/Tech code
        base["Fuel/Tech"] = f"{tk}BRAXX{suf}"
        # Clean Name (remove ", region XX")
        base["Name"] = re.sub(r",\s*region\s+[A-Z]{2}$", "", base["Name"], flags=re.IGNORECASE)
        # Sum all year_cols
        base[year_cols] = group[year_cols].astype(float).sum()
        new_rows.append(base)

    # 6) Merge and clean
    df_proj_clean = pd.concat([df_orig, pd.DataFrame(new_rows)], ignore_index=True)
    df_proj_clean.drop(columns=["TechKey","Suffix"], errors="ignore", inplace=True)


    print("Sheet 'Demand_Projection' processed successfully.")

    # --- Processing the "Profiles" sheet ---

    # 1) Read the sheet
    df_profiles = xls.parse("Profiles")

    # 2) Detect year columns 2021-2050 (int or str)
    year_cols = [
        c for c in df_profiles.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
           or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]

    # 3) Mask for BRA...rr## rows in Fuel/Tech
    #    ^.{3}    any 3-char prefix
    #    BRA      literal BRA
    #    (CN|NW|...)
    #    \d{2}$   2-digit numeric suffix
    pattern = r"^.{3}BRA(?:CN|NW|NE|CW|SO|SE|WE)\d{2}$"
    mask_bra = df_profiles["Fuel/Tech"].str.contains(pattern, regex=True, na=False)

    # 4) Separate non-Brazilian and Brazilian rows
    df_orig  = df_profiles[~mask_bra].copy()
    df_bra   = df_profiles[mask_bra].copy()

    # 5) Extract grouping keys
    df_bra["TechKey"] = df_bra["Fuel/Tech"].str[:3]     # prefix
    df_bra["Suffix"]  = df_bra["Fuel/Tech"].str[-2:]    # final code

    # 6) Group by TechKey, Suffix and Timeslices
    new_rows = []
    for (tk, suf, ts), group in df_bra.groupby(["TechKey", "Suffix", "Timeslices"]):
        base = group.iloc[0].copy()
        # New consolidated Fuel/Tech
        base["Fuel/Tech"] = f"{tk}BRAXX{suf}"
        # Clean Name: remove ", region XX"
        base["Name"] = re.sub(
            r",\s*region\s+[A-Z]{2}$",
            "",
            str(base["Name"]),
            flags=re.IGNORECASE
        )
        # Average the year values
        base[year_cols] = group[year_cols].astype(float).mean()
        new_rows.append(base)

    # 7) Concatenate originals + consolidated rows
    df_profiles_clean = pd.concat([df_orig, pd.DataFrame(new_rows)], ignore_index=True)

    # 8) Remove auxiliary columns
    df_profiles_clean.drop(columns=["TechKey", "Suffix"], errors="ignore", inplace=True)

    # 9) Reassign Tech.ID by unique Fuel/Tech value
    unique_codes = pd.unique(df_profiles_clean["Fuel/Tech"])
    id_map = {code: i+1 for i, code in enumerate(unique_codes)}
    df_profiles_clean["Tech.ID"] = df_profiles_clean["Fuel/Tech"].map(id_map)

    print("Sheet 'Profiles' processed successfully.")

    # --- End of "Profiles" processing ---

    # 10) Save all sheets to a new workbook
    output_file = os.path.join(f"A1_Outputs_{folder}","A-O_Demand.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet in sheet_names:
            if sheet == "Profiles":
                df_profiles_clean.to_excel(writer, sheet_name=sheet, index=False)
            elif sheet == "Demand_Projection":
                # We assume df_proj_clean is already prepared
                df_proj_clean.to_excel(writer, sheet_name=sheet, index=False)
            else:
                xls.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)

    print(f"File saved to {output_file}")



if storage:
    # File path
    file_path = os.path.join("A2_Extra_Inputs","A-Xtra_Storage.xlsx")
    xls = pd.ExcelFile(file_path)

    # Helper function to detect year columns 2021-2050
    def get_year_cols(df):
        return [
            c for c in df.columns
            if (isinstance(c, int) and 2021 <= c <= 2050)
                or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
        ]

    # List of Brazil regions
    brazil_regions = ["CN","NW","NE","CW","SO","SE","WE"]
    bad_regions = ["BRANW","BRANE","BRACW","BRASO","BRASE","BRAWE"]

    # ===== 1) Fixed Horizon Parameters =====
    df_fhp = xls.parse("Fixed Horizon Parameters")

    # 1.1) Remove internal BRA-BRA duplicates
    mask_len13   = df_fhp["STORAGE"].str.len() == 13
    mask_two_bra = df_fhp["STORAGE"].str.count("BRA") > 1
    df_fhp = df_fhp[~(mask_len13 & mask_two_bra)].copy()

    # 1.2) Replace BRACN->BRAXX and adjust STORAGE.Name
    mask_bracn = df_fhp["STORAGE"].str.contains("BRACN", na=False)
    df_fhp.loc[mask_bracn, "STORAGE"] = df_fhp.loc[mask_bracn, "STORAGE"].str.replace("BRACN","BRAXX", regex=False)
    df_fhp.loc[mask_bracn, "STORAGE.Name"] = df_fhp.loc[mask_bracn, "STORAGE.Name"].apply(
        lambda tn: re.sub(r"CN$", "XX", tn) if pd.notna(tn) else tn
    )

    # 1.3) Remove remaining Brazilian regions
    df_fhp = df_fhp[~df_fhp["STORAGE"].str[3:8].isin(bad_regions)].copy()

    # 1.4) Unify TRN...BRA...XX interconnections
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

    # --- Reassign STORAGE.ID in Fixed Horizon Parameters ---
    unique_fhp = pd.unique(df_fhp["STORAGE"])
    fhp_id_map = {stor: i+1 for i, stor in enumerate(unique_fhp)}
    df_fhp["STORAGE.ID"] = df_fhp["STORAGE"].map(fhp_id_map)

    print("Sheet 'Fixed Horizon Parameters' processed successfully.")

    # --- 2) CapitalCostStorage ---
    df_ccs = xls.parse("CapitalCostStorage")
    year_cols = [
        c for c in df_ccs.columns
        if (isinstance(c, int) and 2021 <= c <= 2050)
            or (isinstance(c, str) and c.isdigit() and 2021 <= int(c) <= 2050)
    ]

    # Detect only the Brazilian rows (not backstop)
    mask_ccs_bra = (
        df_ccs["STORAGE"].str[3:6].eq("BRA") &
        df_ccs["STORAGE"].str[6:8].isin(brazil_regions)
    )

    # Separate original and Brazilian rows
    orig_ccs = df_ccs[~mask_ccs_bra].copy()
    bra_ccs  = df_ccs[mask_ccs_bra].copy()
    bra_ccs["Key"] = bra_ccs["STORAGE"].str[:3]

    new_ccs = []
    for (key, param), g in bra_ccs.groupby(["Key","Parameter"]):
        base = g.iloc[0].copy()
        suffix = base["STORAGE"][-2:]
        base["STORAGE"] = f"{key}BRAXX{suffix}"
        # Update STORAGE.Name
        tn = base.get("STORAGE.Name", "")
        if pd.notna(tn):
            base["STORAGE.Name"] = re.sub(
                r"Brazil, region [A-Z]{2}",
                "Brazil, region XX",
                str(tn),
                flags=re.IGNORECASE
            )
        # Calculate values for 2021-2050
        if param.lower().startswith("capital"):
            vals = g[year_cols].astype(float).mean()
        else:
            vals = g[year_cols].astype(float).sum(min_count=1)
        base[year_cols] = vals
        # Projection.Mode: User defined if there is at least one non-NaN value
        base["Projection.Mode"] = "User defined" if base[year_cols].notna().any() else "EMPTY"
        new_ccs.append(base)

    df_ccs_clean = pd.concat([orig_ccs, pd.DataFrame(new_ccs)], ignore_index=True)
    df_ccs_clean.drop(columns=["Key"], errors="ignore", inplace=True)

    # --- Reassign STORAGE.ID in CapitalCostStorage ---
    unique_ccs = pd.unique(df_ccs_clean["STORAGE"])
    ccs_id_map = {stor: i+1 for i, stor in enumerate(unique_ccs)}
    df_ccs_clean["STORAGE.ID"] = df_ccs_clean["STORAGE"].map(ccs_id_map)

    print("Sheet 'CapitalCostStorage' processed successfully.")

    # --- 3) TechnologyStorage ---
    df_ts = xls.parse("TechnologyStorage")
    year_cols = [c for c in df_ts.columns if (isinstance(c,int) and 2021<=c<=2050) or (isinstance(c,str) and c.isdigit() and 2021<=int(c)<=2050)]

    # 3.1) Remove BRA-BRA duplicates
    mask_dupl = (df_ts["TECHNOLOGY"].str.len()==13) & (df_ts["TECHNOLOGY"].str.count("BRA")>1)
    df_ts = df_ts[~mask_dupl].copy()

    # 3.2) Detect BRA rows without backstop
    mask_ts_bra = (
        df_ts["TECHNOLOGY"].str[6:9].eq("BRA") &
        df_ts["TECHNOLOGY"].str[9:11].isin(brazil_regions)
    )

    orig_ts = df_ts[~mask_ts_bra].copy()
    bra_ts  = df_ts[mask_ts_bra].copy()

    # **Key = first 6 characters** (e.g. "PWRLDS", "PWRSDS")
    bra_ts["TechKey"] = bra_ts["TECHNOLOGY"].str[:6]

    # Define operations
    parameters_avg = ["TechnologyToStorage","TechnologyFromStorage"]
    parameters_sum = []  # add here those that should be summed

    new_ts = []
    for (tk, mode, param), g in bra_ts.groupby(["TechKey","MODE_OF_OPERATION","Parameter"]):
        base = g.iloc[0].copy()
        suffix = base["TECHNOLOGY"][-2:]
        base["TECHNOLOGY"] = f"{tk}BRAXX{suffix}"  # e.g. "PWRLDSBRAXX01"
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
    print("Sheet 'TechnologyStorage' processed successfully.")

    # ===== Save all sheets to new workbook =====
    output_file = os.path.join("A2_Extra_Inputs","A-Xtra_Storage.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_fhp.to_excel(writer, sheet_name="Fixed Horizon Parameters", index=False)
        df_ccs_clean.to_excel(writer, sheet_name="CapitalCostStorage", index=False)
        df_ts_clean.to_excel(writer, sheet_name="TechnologyStorage", index=False)

    print(f"File saved to {output_file}")
