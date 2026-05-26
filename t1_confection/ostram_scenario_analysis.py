#!/usr/bin/env python3
"""
OSTRAM Multi-Scenario Analysis & Visualization
================================================
Loads the combined inputs/outputs CSV and the pre-processed .txt files.
Produces diagnostic plots for:
  1. Fleet-mix capacity by scenario (stacked area)
  2. New capacity additions by tech family
  3. Binding MaxCapInv constraint diagnosis
  4. MinCapInv floor vs actual build
  5. Cross-scenario capacity deltas vs BAU
  6. Generation mix (ProductionByTechnologyAnnual)
  7. Interconnector capacity & flows
  8. Internal transmission (PWRTRN) capacity
  9. Backstop (PWRBCK) activation
 10. Reserve margin headroom (from .txt files)
 11. Constraint parameter comparison across .txt scenarios
 12. Per-country capacity breakdown

Usage:
    python ostram_scenario_analysis.py --csv OSTRAM_Combined_Inputs_Outputs.csv
                                       [--txt-dir DIR_WITH_TXT_FILES]
                                       [--output-dir plots]
                                       [--scenarios BAU,A_Calibrated_BAU,...]
                                       [--countries INDNO,INDEA,...]

All flags optional; defaults use current directory.
"""

import argparse
import os
import re
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D

# ──────────────────────────────────────────────────────────────
#  CONFIGURATION
# ──────────────────────────────────────────────────────────────

SCENARIO_ORDER = [
    "BAU",
    "A_Calibrated_BAU",
    "B_Optimised_VRE",
    "C_Target_VRE",
]
SCENARIO_SHORT = {
    "BAU": "BAU₀",
    "A_Calibrated_BAU": "A-CalBAU",
    "B_Optimised_VRE": "B-OptVRE",
    "C_Target_VRE": "C-TgtVRE",
}

# Technology family mapping (prefix → label + color)
TECH_FAMILIES = {
    "PWRCOA": ("Coal",          "#4a4a4a"),
    "PWRNGS": ("Gas",           "#d4a017"),
    "PWROIL": ("Oil",           "#8b4513"),
    "PWRPET": ("Petroleum",     "#a0522d"),
    "PWRURN": ("Nuclear",       "#e74c3c"),
    "PWRHYD": ("Hydro",         "#2980b9"),
    "PWRSHP": ("Small Hydro",   "#5dade2"),
    "PWRSPV": ("Solar PV",      "#f39c12"),
    "PWRWON": ("Wind Onshore",  "#27ae60"),
    "PWRWOF": ("Wind Offshore", "#1abc9c"),
    "PWRCSP": ("CSP",           "#e67e22"),
    "PWRBIO": ("Biomass",       "#7d3c98"),
    "PWRWAS": ("Waste",         "#95a5a6"),
    "PWRGEO": ("Geothermal",    "#c0392b"),
    "PWROTH": ("Other",         "#bdc3c7"),
    "PWRCCS": ("CCS",           "#566573"),
}
BACKSTOP_PREFIX = "PWRBCK"
BACKSTOP_COLOR = "#ff1493"

STORAGE_PREFIXES = ("PWRLDS", "PWRSDS")
INTERNAL_TRN_PREFIX = "PWRTRN"
DISPATCH_TRN_PREFIX = "DSPTRN"

# Cross-border interconnectors: TRN{FROM}{TO} but NOT TRNNLI/TRNRPO
INTERCONNECTOR_COUNTRY_PREFIXES = ("TRNBGD", "TRNBTN", "TRNIND", "TRNLKA",
                                   "TRNMDV", "TRNNPL")
INTERCONNECTOR_INFRA_PREFIXES = ("TRNNLI", "TRNRPO")

COUNTRY_CODES = {
    "BGDXX": "Bangladesh",
    "BTNXX": "Bhutan",
    "INDEA": "India-East",
    "INDNE": "India-NE",
    "INDNO": "India-North",
    "INDSO": "India-South",
    "INDWE": "India-West",
    "LKAXX": "Sri Lanka",
    "MDVXX": "Maldives",
    "NPLXX": "Nepal",
}

COUNTRY_COLORS = {
    "Bangladesh":   "#1f77b4",
    "Bhutan":       "#aec7e8",
    "India-East":   "#ff7f0e",
    "India-NE":     "#ffbb78",
    "India-North":  "#2ca02c",
    "India-South":  "#98df8a",
    "India-West":   "#d62728",
    "Sri Lanka":    "#9467bd",
    "Maldives":     "#c5b0d5",
    "Nepal":        "#8c564b",
}


# ──────────────────────────────────────────────────────────────
#  HELPERS
# ──────────────────────────────────────────────────────────────

def tech_family(tech: str) -> str:
    """Map a technology code to its family label."""
    for pfx, (label, _) in TECH_FAMILIES.items():
        if tech.startswith(pfx):
            return label
    if tech.startswith(BACKSTOP_PREFIX):
        return "Backstop"
    return None


def tech_color(family: str) -> str:
    """Color for a tech family label."""
    for pfx, (label, color) in TECH_FAMILIES.items():
        if label == family:
            return color
    if family == "Backstop":
        return BACKSTOP_COLOR
    return "#cccccc"


def tech_country(tech: str) -> str:
    """Extract country from a PWR* tech code (last 5 chars)."""
    suffix = tech[-5:]
    return COUNTRY_CODES.get(suffix, suffix)


def is_generation_tech(tech: str) -> bool:
    """True for PWR* generation (not TRN, not DSP, not storage, not backstop)."""
    if not tech.startswith("PWR"):
        return False
    if tech.startswith(INTERNAL_TRN_PREFIX):
        return False
    if tech.startswith(BACKSTOP_PREFIX):
        return False
    if any(tech.startswith(p) for p in STORAGE_PREFIXES):
        return False
    return True


def is_interconnector(tech: str) -> bool:
    return any(tech.startswith(p) for p in
               INTERCONNECTOR_COUNTRY_PREFIXES + INTERCONNECTOR_INFRA_PREFIXES)


def ordered_families():
    """Return family labels in stacking order (thermal bottom, RE top)."""
    order = ["Coal", "Petroleum", "Oil", "Gas", "Nuclear", "Other", "CCS",
             "Waste", "Biomass", "Geothermal", "Hydro", "Small Hydro",
             "CSP", "Solar PV", "Wind Onshore", "Wind Offshore", "Backstop"]
    return order


def save_fig(fig, outdir, name):
    path = os.path.join(outdir, f"{name}.png")
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  ✓  {path}")


# ──────────────────────────────────────────────────────────────
#  DATA LOADING
# ──────────────────────────────────────────────────────────────

def load_csv(csv_path: str) -> pd.DataFrame:
    """Load the combined inputs/outputs CSV with only needed columns."""
    cols = [
        "Scenario", "YEAR", "TECHNOLOGY", "FUEL",
        "AccumulatedNewCapacity",
        "NewCapacity",
        "TotalCapacityAnnual",
        "ProductionByTechnologyAnnual",
        "TotalTechnologyAnnualActivity",
        "TotalAnnualMaxCapacityInvestment",
        "TotalAnnualMinCapacityInvestment",
        "ResidualCapacity",
        "CapitalCost",
        "CapacityToActivityUnit",
        "CapitalInvestment",
        "DiscountedCostByTechnology",
        "ReserveMargin",
        "ReserveMarginTagFuel",
        "ReserveMarginTagTechnology",
    ]
    print(f"Loading {csv_path} ...")
    df = pd.read_csv(csv_path, usecols=cols, low_memory=False)
    df["YEAR"] = df["YEAR"].astype("Int64")
    print(f"  {len(df):,} rows, scenarios: {list(df['Scenario'].unique())}")
    return df


def load_txt_param(txt_path: str, param_name: str) -> pd.DataFrame:
    """
    Parse a single param block from an otoole .txt file.
    Returns DataFrame with columns depending on param dimensionality.
    """
    rows = []
    inside = False
    with open(txt_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith(f"param") and f": {param_name} :=" in line:
                inside = True
                continue
            if inside:
                if line == ";":
                    break
                parts = line.split()
                if parts:
                    rows.append(parts)
    if not rows:
        return pd.DataFrame()
    # Infer columns from number of fields
    ncols = len(rows[0])
    if ncols == 4:
        return pd.DataFrame(rows, columns=["REGION", "TECHNOLOGY", "YEAR", "VALUE"])
    elif ncols == 3:
        return pd.DataFrame(rows, columns=["REGION", "TECHNOLOGY", "VALUE"])
    elif ncols == 5:
        return pd.DataFrame(rows, columns=["REGION", "TECHNOLOGY", "FUEL_OR_MODE", "YEAR", "VALUE"])
    else:
        cols_auto = [f"col{i}" for i in range(ncols)]
        return pd.DataFrame(rows, columns=cols_auto)


# ──────────────────────────────────────────────────────────────
#  EXTRACTION: get clean per-variable tables
# ──────────────────────────────────────────────────────────────

def extract_capacity_table(df: pd.DataFrame, scenarios=None) -> pd.DataFrame:
    """Unique (Scenario, YEAR, TECHNOLOGY, TotalCapacityAnnual) for gen techs."""
    mask = df["TotalCapacityAnnual"].notna()
    sub = df.loc[mask, ["Scenario", "YEAR", "TECHNOLOGY", "TotalCapacityAnnual"]].drop_duplicates()
    sub["family"] = sub["TECHNOLOGY"].map(tech_family)
    sub["country"] = sub["TECHNOLOGY"].map(tech_country)
    if scenarios:
        sub = sub[sub["Scenario"].isin(scenarios)]
    return sub


def extract_new_capacity(df: pd.DataFrame, scenarios=None) -> pd.DataFrame:
    mask = df["NewCapacity"].notna() & (df["NewCapacity"] != 0)
    sub = df.loc[mask, ["Scenario", "YEAR", "TECHNOLOGY", "NewCapacity"]].drop_duplicates()
    sub["family"] = sub["TECHNOLOGY"].map(tech_family)
    sub["country"] = sub["TECHNOLOGY"].map(tech_country)
    if scenarios:
        sub = sub[sub["Scenario"].isin(scenarios)]
    return sub


def extract_production(df: pd.DataFrame, scenarios=None) -> pd.DataFrame:
    mask = df["ProductionByTechnologyAnnual"].notna() & (df["ProductionByTechnologyAnnual"] != 0)
    sub = df.loc[mask, ["Scenario", "YEAR", "TECHNOLOGY", "ProductionByTechnologyAnnual"]].drop_duplicates()
    sub["family"] = sub["TECHNOLOGY"].map(tech_family)
    sub["country"] = sub["TECHNOLOGY"].map(tech_country)
    if scenarios:
        sub = sub[sub["Scenario"].isin(scenarios)]
    return sub


def extract_constraints(df: pd.DataFrame, scenarios=None) -> pd.DataFrame:
    """Merge NewCapacity with MaxCapInv and MinCapInv for constraint analysis."""
    nc = df.loc[df["NewCapacity"].notna(),
                ["Scenario", "YEAR", "TECHNOLOGY", "NewCapacity"]].drop_duplicates()
    mx = df.loc[df["TotalAnnualMaxCapacityInvestment"].notna(),
                ["Scenario", "YEAR", "TECHNOLOGY", "TotalAnnualMaxCapacityInvestment"]].drop_duplicates()
    mn = df.loc[df["TotalAnnualMinCapacityInvestment"].notna(),
                ["Scenario", "YEAR", "TECHNOLOGY", "TotalAnnualMinCapacityInvestment"]].drop_duplicates()
    out = nc.merge(mx, on=["Scenario", "YEAR", "TECHNOLOGY"], how="left")
    out = out.merge(mn, on=["Scenario", "YEAR", "TECHNOLOGY"], how="left")
    out["family"] = out["TECHNOLOGY"].map(tech_family)
    out["country"] = out["TECHNOLOGY"].map(tech_country)
    out["max_gap"] = out["TotalAnnualMaxCapacityInvestment"] - out["NewCapacity"]
    out["max_binding"] = out["max_gap"].abs() < 0.001
    out["min_gap"] = out["NewCapacity"] - out["TotalAnnualMinCapacityInvestment"]
    out["min_binding"] = (out["TotalAnnualMinCapacityInvestment"].notna()
                          & (out["min_gap"].abs() < 0.001)
                          & (out["NewCapacity"] > 0))
    if scenarios:
        out = out[out["Scenario"].isin(scenarios)]
    return out


# ──────────────────────────────────────────────────────────────
#  PLOT 1: Fleet-mix capacity – stacked area per scenario
# ──────────────────────────────────────────────────────────────

def plot_fleet_mix(cap_df: pd.DataFrame, outdir: str, scenarios: list):
    """One subplot per scenario – stacked area of TotalCapacityAnnual by family."""
    fam_order = [f for f in ordered_families() if f in cap_df["family"].unique()]
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(6 * n, 5), sharey=True)
    if n == 1:
        axes = [axes]
    for ax, sc in zip(axes, scenarios):
        sub = cap_df[(cap_df["Scenario"] == sc) & cap_df["family"].notna()]
        piv = sub.pivot_table(index="YEAR", columns="family",
                              values="TotalCapacityAnnual", aggfunc="sum").fillna(0)
        piv = piv.reindex(columns=fam_order, fill_value=0)
        years = piv.index.values
        colors = [tech_color(f) for f in piv.columns]
        ax.stackplot(years, *[piv[c].values for c in piv.columns],
                     labels=piv.columns, colors=colors, alpha=0.85)
        ax.set_title(SCENARIO_SHORT.get(sc, sc), fontsize=11, fontweight="bold")
        ax.set_xlabel("Year")
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    axes[0].set_ylabel("Installed capacity (GW)")
    handles, labels = axes[-1].get_legend_handles_labels()
    fig.legend(handles[::-1], labels[::-1], loc="center left",
               bbox_to_anchor=(1.0, 0.5), fontsize=8)
    fig.suptitle("Fleet-Mix Capacity by Scenario", fontsize=13, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "01_fleet_mix_capacity")


# ──────────────────────────────────────────────────────────────
#  PLOT 2: New capacity additions by family – grouped bars
# ──────────────────────────────────────────────────────────────

def plot_new_capacity_bars(nc_df: pd.DataFrame, outdir: str, scenarios: list):
    """Cumulative new capacity 2023-2050 by family, bar per scenario."""
    sub = nc_df[nc_df["family"].notna()].copy()
    agg = sub.groupby(["Scenario", "family"])["NewCapacity"].sum().reset_index()
    fam_order = [f for f in ordered_families() if f in agg["family"].unique()]
    fig, ax = plt.subplots(figsize=(14, 6))
    x = np.arange(len(fam_order))
    w = 0.8 / len(scenarios)
    for i, sc in enumerate(scenarios):
        vals = []
        for f in fam_order:
            row = agg[(agg["Scenario"] == sc) & (agg["family"] == f)]
            vals.append(row["NewCapacity"].sum() if len(row) else 0)
        ax.bar(x + i * w, vals, w, label=SCENARIO_SHORT.get(sc, sc),
               color=[tech_color(f) for f in fam_order], alpha=0.6 + 0.1 * i,
               edgecolor="white", linewidth=0.5)
    ax.set_xticks(x + w * (len(scenarios) - 1) / 2)
    ax.set_xticklabels(fam_order, rotation=45, ha="right", fontsize=9)
    ax.set_ylabel("Cumulative new capacity 2023-2050 (GW)")
    ax.legend(fontsize=9)
    ax.set_title("Cumulative New Capacity Additions by Technology Family", fontweight="bold")
    fig.tight_layout()
    save_fig(fig, outdir, "02_new_capacity_cumulative")


# ──────────────────────────────────────────────────────────────
#  PLOT 3: Binding MaxCapInv constraints – heatmap
# ──────────────────────────────────────────────────────────────

def plot_binding_maxcapinv(constr_df: pd.DataFrame, outdir: str, scenarios: list):
    """Per scenario: count of binding MaxCapInv constraints by family and year."""
    for sc in scenarios:
        sub = constr_df[(constr_df["Scenario"] == sc) & constr_df["max_binding"]
                        & constr_df["family"].notna()]
        if sub.empty:
            continue
        piv = sub.pivot_table(index="family", columns="YEAR",
                              values="max_binding", aggfunc="sum").fillna(0)
        fig, ax = plt.subplots(figsize=(16, max(4, len(piv) * 0.45)))
        im = ax.imshow(piv.values, aspect="auto", cmap="YlOrRd", interpolation="nearest")
        ax.set_yticks(range(len(piv.index)))
        ax.set_yticklabels(piv.index, fontsize=9)
        ax.set_xticks(range(len(piv.columns)))
        ax.set_xticklabels([int(y) for y in piv.columns], rotation=90, fontsize=8)
        ax.set_title(f"Binding MaxCapInv Constraints — {SCENARIO_SHORT.get(sc, sc)}",
                     fontweight="bold")
        fig.colorbar(im, ax=ax, label="# techs hitting lid")
        fig.tight_layout()
        save_fig(fig, outdir, f"03_binding_maxcapinv_{sc}")


# ──────────────────────────────────────────────────────────────
#  PLOT 4: MinCapInv floor analysis
# ──────────────────────────────────────────────────────────────

def plot_mincapinv_floor(constr_df: pd.DataFrame, outdir: str, scenarios: list):
    """For each scenario, show techs where NewCapacity == MinCapInv (floor-only builds)."""
    for sc in scenarios:
        sub = constr_df[(constr_df["Scenario"] == sc)
                        & constr_df["min_binding"]
                        & constr_df["family"].notna()]
        if sub.empty:
            continue
        piv = sub.pivot_table(index="TECHNOLOGY", columns="YEAR",
                              values="NewCapacity", aggfunc="sum").fillna(0)
        # Keep only techs with at least some floor-only years
        if piv.empty:
            continue
        fig, ax = plt.subplots(figsize=(16, max(4, len(piv) * 0.35)))
        im = ax.imshow(piv.values, aspect="auto", cmap="Blues", interpolation="nearest")
        ax.set_yticks(range(len(piv.index)))
        ax.set_yticklabels(piv.index, fontsize=7)
        ax.set_xticks(range(len(piv.columns)))
        ax.set_xticklabels([int(y) for y in piv.columns], rotation=90, fontsize=8)
        ax.set_title(f"Floor-Only Builds (NewCap = MinCapInv) — {SCENARIO_SHORT.get(sc, sc)}",
                     fontweight="bold")
        fig.colorbar(im, ax=ax, label="NewCapacity (GW)")
        fig.tight_layout()
        save_fig(fig, outdir, f"04_mincapinv_floor_{sc}")


# ──────────────────────────────────────────────────────────────
#  PLOT 5: Cross-scenario capacity delta vs BAU
# ──────────────────────────────────────────────────────────────

def plot_capacity_delta_vs_bau(cap_df: pd.DataFrame, outdir: str, scenarios: list):
    """For each non-BAU scenario, show delta TotalCapacityAnnual by family vs BAU."""
    bau_name = "BAU"
    if bau_name not in cap_df["Scenario"].unique():
        print("  ⚠  No BAU scenario found; skipping delta plot.")
        return
    bau = cap_df[cap_df["Scenario"] == bau_name].copy()
    bau_piv = bau.pivot_table(index="YEAR", columns="family",
                               values="TotalCapacityAnnual", aggfunc="sum").fillna(0)
    others = [s for s in scenarios if s != bau_name]
    if not others:
        return
    fig, axes = plt.subplots(1, len(others), figsize=(6 * len(others), 5), sharey=True)
    if len(others) == 1:
        axes = [axes]
    fam_order = [f for f in ordered_families() if f in cap_df["family"].unique()]
    for ax, sc in zip(axes, others):
        sc_piv = (cap_df[cap_df["Scenario"] == sc]
                  .pivot_table(index="YEAR", columns="family",
                               values="TotalCapacityAnnual", aggfunc="sum").fillna(0))
        # Align columns
        all_fam = sorted(set(bau_piv.columns) | set(sc_piv.columns))
        for f in all_fam:
            if f not in bau_piv.columns:
                bau_piv[f] = 0
            if f not in sc_piv.columns:
                sc_piv[f] = 0
        years = sorted(set(bau_piv.index) & set(sc_piv.index))
        delta = sc_piv.loc[years] - bau_piv.loc[years]
        for f in fam_order:
            if f in delta.columns and delta[f].abs().max() > 0.01:
                ax.plot(years, delta[f], label=f, color=tech_color(f), linewidth=1.5)
        ax.axhline(0, color="grey", linewidth=0.5, linestyle="--")
        ax.set_title(f"Δ vs BAU₀ — {SCENARIO_SHORT.get(sc, sc)}", fontweight="bold")
        ax.set_xlabel("Year")
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    axes[0].set_ylabel("ΔCapacity (GW)")
    axes[-1].legend(fontsize=7, loc="upper left")
    fig.suptitle("Capacity Delta vs BAU₀ by Family", fontsize=13, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "05_capacity_delta_vs_bau")


# ──────────────────────────────────────────────────────────────
#  PLOT 6: Generation mix – stacked area
# ──────────────────────────────────────────────────────────────

def plot_generation_mix(prod_df: pd.DataFrame, outdir: str, scenarios: list):
    """Stacked area of ProductionByTechnologyAnnual (PJ) by family."""
    gen = prod_df[prod_df["family"].notna()].copy()
    fam_order = [f for f in ordered_families() if f in gen["family"].unique()]
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(6 * n, 5), sharey=True)
    if n == 1:
        axes = [axes]
    for ax, sc in zip(axes, scenarios):
        sub = gen[gen["Scenario"] == sc]
        piv = sub.pivot_table(index="YEAR", columns="family",
                              values="ProductionByTechnologyAnnual",
                              aggfunc="sum").fillna(0)
        piv = piv.reindex(columns=fam_order, fill_value=0)
        years = piv.index.values
        colors = [tech_color(f) for f in piv.columns]
        ax.stackplot(years, *[piv[c].values for c in piv.columns],
                     labels=piv.columns, colors=colors, alpha=0.85)
        ax.set_title(SCENARIO_SHORT.get(sc, sc), fontsize=11, fontweight="bold")
        ax.set_xlabel("Year")
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    axes[0].set_ylabel("Annual generation (PJ)")
    handles, labels = axes[-1].get_legend_handles_labels()
    fig.legend(handles[::-1], labels[::-1], loc="center left",
               bbox_to_anchor=(1.0, 0.5), fontsize=8)
    fig.suptitle("Generation Mix by Scenario", fontsize=13, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "06_generation_mix")


# ──────────────────────────────────────────────────────────────
#  PLOT 7: Interconnector capacity
# ──────────────────────────────────────────────────────────────

def plot_interconnectors(cap_df_full: pd.DataFrame, outdir: str, scenarios: list):
    """Show TotalCapacityAnnual for cross-border interconnector techs."""
    ic_mask = cap_df_full["TECHNOLOGY"].apply(is_interconnector)
    ic = cap_df_full[ic_mask & cap_df_full["TotalCapacityAnnual"].notna()].copy()
    if ic.empty:
        print("  ⚠  No interconnector capacity data.")
        return
    ic = ic[["Scenario", "YEAR", "TECHNOLOGY", "TotalCapacityAnnual"]].drop_duplicates()
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(6 * n, 6), sharey=True)
    if n == 1:
        axes = [axes]
    for ax, sc in zip(axes, scenarios):
        sub = ic[ic["Scenario"] == sc]
        for tech in sorted(sub["TECHNOLOGY"].unique()):
            t = sub[sub["TECHNOLOGY"] == tech].sort_values("YEAR")
            ax.plot(t["YEAR"], t["TotalCapacityAnnual"], label=tech, linewidth=1)
        ax.set_title(SCENARIO_SHORT.get(sc, sc), fontsize=11, fontweight="bold")
        ax.set_xlabel("Year")
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    axes[0].set_ylabel("Interconnector capacity (GW)")
    axes[-1].legend(fontsize=5, loc="upper left", ncol=2)
    fig.suptitle("Cross-Border Interconnector Capacity", fontsize=13, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "07_interconnector_capacity")


# ──────────────────────────────────────────────────────────────
#  PLOT 8: Internal transmission (PWRTRN)
# ──────────────────────────────────────────────────────────────

def plot_internal_transmission(cap_df_full: pd.DataFrame, outdir: str, scenarios: list):
    trn_mask = cap_df_full["TECHNOLOGY"].str.startswith(INTERNAL_TRN_PREFIX)
    trn = cap_df_full[trn_mask & cap_df_full["TotalCapacityAnnual"].notna()].copy()
    trn = trn[["Scenario", "YEAR", "TECHNOLOGY", "TotalCapacityAnnual"]].drop_duplicates()
    if trn.empty:
        print("  ⚠  No PWRTRN capacity data.")
        return
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(6 * n, 5), sharey=True)
    if n == 1:
        axes = [axes]
    for ax, sc in zip(axes, scenarios):
        sub = trn[trn["Scenario"] == sc]
        for tech in sorted(sub["TECHNOLOGY"].unique()):
            t = sub[sub["TECHNOLOGY"] == tech].sort_values("YEAR")
            cty = tech_country(tech)
            ax.plot(t["YEAR"], t["TotalCapacityAnnual"], label=cty,
                    color=COUNTRY_COLORS.get(cty, None), linewidth=1.5)
        ax.set_title(SCENARIO_SHORT.get(sc, sc), fontweight="bold")
        ax.set_xlabel("Year")
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    axes[0].set_ylabel("Internal TRN capacity (GW)")
    axes[-1].legend(fontsize=8)
    fig.suptitle("Internal Transmission Capacity (PWRTRN)", fontsize=13, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "08_internal_transmission")


# ──────────────────────────────────────────────────────────────
#  PLOT 9: Backstop (PWRBCK) activation
# ──────────────────────────────────────────────────────────────

def plot_backstop(cap_df_full: pd.DataFrame, prod_df: pd.DataFrame,
                  outdir: str, scenarios: list):
    bck_cap = cap_df_full[cap_df_full["TECHNOLOGY"].str.startswith(BACKSTOP_PREFIX)
                          & cap_df_full["TotalCapacityAnnual"].notna()].copy()
    bck_cap = bck_cap[["Scenario", "YEAR", "TECHNOLOGY", "TotalCapacityAnnual"]].drop_duplicates()
    bck_prod = prod_df[prod_df["TECHNOLOGY"].str.startswith(BACKSTOP_PREFIX)].copy()

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
    for sc in scenarios:
        sub = bck_cap[bck_cap["Scenario"] == sc]
        agg = sub.groupby("YEAR")["TotalCapacityAnnual"].sum()
        ax1.plot(agg.index, agg.values, label=SCENARIO_SHORT.get(sc, sc), linewidth=1.5)
    ax1.set_title("Backstop Capacity (PWRBCK)", fontweight="bold")
    ax1.set_ylabel("GW")
    ax1.legend(fontsize=9)
    ax1.xaxis.set_major_locator(mticker.MultipleLocator(5))

    for sc in scenarios:
        sub = bck_prod[bck_prod["Scenario"] == sc]
        agg = sub.groupby("YEAR")["ProductionByTechnologyAnnual"].sum()
        if not agg.empty:
            ax2.plot(agg.index, agg.values, label=SCENARIO_SHORT.get(sc, sc), linewidth=1.5)
    ax2.set_title("Backstop Generation (PWRBCK)", fontweight="bold")
    ax2.set_ylabel("PJ")
    ax2.legend(fontsize=9)
    ax2.xaxis.set_major_locator(mticker.MultipleLocator(5))
    fig.tight_layout()
    save_fig(fig, outdir, "09_backstop_activation")


# ──────────────────────────────────────────────────────────────
#  PLOT 10: Per-country capacity breakdown (one per scenario)
# ──────────────────────────────────────────────────────────────

def plot_country_capacity(cap_df: pd.DataFrame, outdir: str, scenarios: list):
    """Stacked area of total gen capacity by country, per scenario."""
    gen = cap_df[cap_df["family"].notna()].copy()
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(6 * n, 5), sharey=True)
    if n == 1:
        axes = [axes]
    cty_order = list(COUNTRY_CODES.values())
    for ax, sc in zip(axes, scenarios):
        sub = gen[gen["Scenario"] == sc]
        piv = sub.pivot_table(index="YEAR", columns="country",
                              values="TotalCapacityAnnual", aggfunc="sum").fillna(0)
        cols = [c for c in cty_order if c in piv.columns]
        piv = piv[cols]
        years = piv.index.values
        colors = [COUNTRY_COLORS.get(c, "#cccccc") for c in cols]
        ax.stackplot(years, *[piv[c].values for c in cols],
                     labels=cols, colors=colors, alpha=0.85)
        ax.set_title(SCENARIO_SHORT.get(sc, sc), fontweight="bold")
        ax.set_xlabel("Year")
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    axes[0].set_ylabel("Generation capacity (GW)")
    handles, labels = axes[-1].get_legend_handles_labels()
    fig.legend(handles[::-1], labels[::-1], loc="center left",
               bbox_to_anchor=(1.0, 0.5), fontsize=8)
    fig.suptitle("Generation Capacity by Country", fontsize=13, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "10_country_capacity")


# ──────────────────────────────────────────────────────────────
#  PLOT 11: Constraint comparison across TXT files
# ──────────────────────────────────────────────────────────────

def plot_txt_constraint_comparison(txt_files: dict, outdir: str):
    """
    Compare TotalAnnualMaxCapacityInvestment across .txt scenarios
    for a selection of important techs.
    """
    if not txt_files:
        print("  ⚠  No .txt files provided; skipping constraint comparison.")
        return

    spotlight_techs = [
        "PWRSPVINDNO", "PWRSPVINDSO", "PWRSPVINDWE",
        "PWRWONINDNO", "PWRWONINDSO",
        "PWRCOAINDNO", "PWRCOAINDEA",
        "PWRNGSINDNO", "PWRNGSINDEA",
        "PWRHYDINDNO", "PWRHYDBGDXX",
    ]

    all_data = {}
    for label, fpath in txt_files.items():
        df = load_txt_param(fpath, "TotalAnnualMaxCapacityInvestment")
        if df.empty:
            continue
        df["YEAR"] = df["YEAR"].astype(int)
        df["VALUE"] = df["VALUE"].astype(float)
        all_data[label] = df

    if not all_data:
        print("  ⚠  Could not parse MaxCapInv from .txt files.")
        return

    # Filter to spotlight techs that exist
    avail_techs = set()
    for df in all_data.values():
        avail_techs.update(df["TECHNOLOGY"].unique())
    spot = [t for t in spotlight_techs if t in avail_techs]
    if not spot:
        spot = sorted(avail_techs)[:12]

    ncols = 3
    nrows = (len(spot) + ncols - 1) // ncols
    fig, axes = plt.subplots(nrows, ncols, figsize=(5 * ncols, 3.5 * nrows), sharey=False)
    axes = axes.flatten() if nrows > 1 else (axes if ncols > 1 else [axes])

    for i, tech in enumerate(spot):
        ax = axes[i]
        for label, df in all_data.items():
            sub = df[df["TECHNOLOGY"] == tech].sort_values("YEAR")
            if not sub.empty:
                ax.plot(sub["YEAR"], sub["VALUE"], label=label, linewidth=1.2)
        ax.set_title(tech, fontsize=9, fontweight="bold")
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.tick_params(labelsize=7)
    for j in range(i + 1, len(axes)):
        axes[j].set_visible(False)
    axes[0].legend(fontsize=6, loc="best")
    fig.suptitle("MaxCapInv Comparison Across Pre-processed Scenarios (.txt)",
                 fontsize=12, fontweight="bold", y=1.01)
    fig.tight_layout()
    save_fig(fig, outdir, "11_txt_maxcapinv_comparison")


# ──────────────────────────────────────────────────────────────
#  PLOT 12: Summary constraint tightness table
# ──────────────────────────────────────────────────────────────

def plot_constraint_summary(constr_df: pd.DataFrame, outdir: str, scenarios: list):
    """Bar chart: # of binding max vs min constraints by scenario."""
    rows = []
    for sc in scenarios:
        sub = constr_df[constr_df["Scenario"] == sc]
        gen = sub[sub["family"].notna()]
        n_max_bind = gen["max_binding"].sum()
        n_min_bind = gen["min_binding"].sum()
        n_total = len(gen)
        n_unconstrained = n_total - n_max_bind - n_min_bind
        rows.append({
            "Scenario": SCENARIO_SHORT.get(sc, sc),
            "Max-binding (lid)": n_max_bind,
            "Min-binding (floor)": n_min_bind,
            "Unconstrained": max(0, n_unconstrained),
        })
    sdf = pd.DataFrame(rows).set_index("Scenario")
    fig, ax = plt.subplots(figsize=(10, 5))
    sdf.plot(kind="bar", stacked=True, ax=ax,
             color=["#e74c3c", "#3498db", "#95a5a6"], edgecolor="white")
    ax.set_ylabel("# (tech × year) entries")
    ax.set_title("Constraint Tightness Summary Across Scenarios", fontweight="bold")
    ax.legend(fontsize=9)
    plt.xticks(rotation=0)
    fig.tight_layout()
    save_fig(fig, outdir, "12_constraint_summary")


# ──────────────────────────────────────────────────────────────
#  PLOT 13: Year-over-year new capacity profile per family
# ──────────────────────────────────────────────────────────────

def plot_new_capacity_timeseries(nc_df: pd.DataFrame, outdir: str, scenarios: list):
    """Line chart of annual new capacity by family, one subplot per family."""
    gen = nc_df[nc_df["family"].notna()].copy()
    families = [f for f in ordered_families() if f in gen["family"].unique()]
    ncols = 3
    nrows = (len(families) + ncols - 1) // ncols
    fig, axes = plt.subplots(nrows, ncols, figsize=(5 * ncols, 3.5 * nrows))
    axes = axes.flatten()
    for i, fam in enumerate(families):
        ax = axes[i]
        sub = gen[gen["family"] == fam]
        for sc in scenarios:
            s = sub[sub["Scenario"] == sc].groupby("YEAR")["NewCapacity"].sum()
            if not s.empty:
                ax.plot(s.index, s.values, label=SCENARIO_SHORT.get(sc, sc), linewidth=1.2)
        ax.set_title(fam, fontsize=9, fontweight="bold", color=tech_color(fam))
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.tick_params(labelsize=7)
    for j in range(i + 1, len(axes)):
        axes[j].set_visible(False)
    axes[0].legend(fontsize=6)
    fig.suptitle("Annual New Capacity by Family & Scenario",
                 fontsize=12, fontweight="bold", y=1.01)
    fig.tight_layout()
    save_fig(fig, outdir, "13_new_capacity_timeseries")


# ──────────────────────────────────────────────────────────────
#  PLOT 14: Capacity factor utilisation (gen / cap × C2AU)
# ──────────────────────────────────────────────────────────────

def plot_utilisation(df: pd.DataFrame, outdir: str, scenarios: list):
    """Effective utilisation = Production / (Capacity × C2AU × YearFraction)."""
    # Get capacity
    cap = df.loc[df["TotalCapacityAnnual"].notna(),
                 ["Scenario","YEAR","TECHNOLOGY","TotalCapacityAnnual"]].drop_duplicates()
    prod = df.loc[df["ProductionByTechnologyAnnual"].notna(),
                  ["Scenario","YEAR","TECHNOLOGY","ProductionByTechnologyAnnual"]].drop_duplicates()
    c2a = df.loc[df["CapacityToActivityUnit"].notna(),
                 ["Scenario","YEAR","TECHNOLOGY","CapacityToActivityUnit"]].drop_duplicates()

    m = cap.merge(prod, on=["Scenario","YEAR","TECHNOLOGY"], how="inner")
    m = m.merge(c2a, on=["Scenario","YEAR","TECHNOLOGY"], how="left")
    m["CapacityToActivityUnit"] = m["CapacityToActivityUnit"].fillna(31.536)
    m["max_gen"] = m["TotalCapacityAnnual"] * m["CapacityToActivityUnit"]
    m["util"] = m["ProductionByTechnologyAnnual"] / m["max_gen"]
    m["util"] = m["util"].clip(0, 1.5)  # cap outliers for display
    m["family"] = m["TECHNOLOGY"].map(tech_family)
    m = m[m["family"].notna()]

    families = [f for f in ordered_families() if f in m["family"].unique()]
    fig, ax = plt.subplots(figsize=(14, 6))
    x = np.arange(len(families))
    w = 0.8 / len(scenarios)
    for i, sc in enumerate(scenarios):
        vals = []
        for fam in families:
            sub = m[(m["Scenario"]==sc) & (m["family"]==fam) & (m["YEAR"]==2050)]
            vals.append(sub["util"].mean() if len(sub) else 0)
        ax.bar(x + i*w, vals, w, label=SCENARIO_SHORT.get(sc,sc), alpha=0.7+0.08*i)
    ax.axhline(1.0, color="red", linewidth=0.5, linestyle="--", label="100% util")
    ax.set_xticks(x + w*(len(scenarios)-1)/2)
    ax.set_xticklabels(families, rotation=45, ha="right", fontsize=9)
    ax.set_ylabel("Avg utilisation @ 2050")
    ax.set_title("Technology Utilisation (Production / MaxOutput) @ 2050", fontweight="bold")
    ax.legend(fontsize=8)
    fig.tight_layout()
    save_fig(fig, outdir, "14_utilisation_2050")


# ──────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="OSTRAM multi-scenario analysis")
    parser.add_argument("--csv", required=True, help="Path to OSTRAM_Combined_Inputs_Outputs.csv")
    parser.add_argument("--txt-dir", default=None,
                        help="Directory containing Pre_processed_*.txt files")
    parser.add_argument("--output-dir", default="ostram_plots",
                        help="Output directory for PNG plots (default: ostram_plots)")
    parser.add_argument("--scenarios", default=None,
                        help="Comma-separated scenario names to include")
    args = parser.parse_args()

    outdir = args.output_dir
    os.makedirs(outdir, exist_ok=True)

    # ── Load CSV ──
    df = load_csv(args.csv)
    scenarios = args.scenarios.split(",") if args.scenarios else \
        [s for s in SCENARIO_ORDER if s in df["Scenario"].unique()]
    print(f"  Analysing scenarios: {scenarios}")

    # ── Load .txt files ──
    txt_files = {}
    if args.txt_dir:
        for fname in sorted(os.listdir(args.txt_dir)):
            if fname.startswith("Pre_processed") and fname.endswith(".txt"):
                label = fname.replace("Pre_processed_", "").replace(
                    "_0_NoStorage_OpenBCK_RMCarefulXLSX.txt", "")
                txt_files[label] = os.path.join(args.txt_dir, fname)
        print(f"  Found {len(txt_files)} .txt files: {list(txt_files.keys())}")

    # ── Extract tables ──
    print("\nExtracting capacity data...")
    cap_all = df.loc[df["TotalCapacityAnnual"].notna(),
                     ["Scenario","YEAR","TECHNOLOGY","TotalCapacityAnnual"]].drop_duplicates()
    cap_gen = extract_capacity_table(df, scenarios)

    print("Extracting new capacity...")
    nc = extract_new_capacity(df, scenarios)

    print("Extracting production...")
    prod = extract_production(df, scenarios)

    print("Extracting constraints...")
    constr = extract_constraints(df, scenarios)

    # ── Generate plots ──
    print("\n" + "=" * 60)
    print("  GENERATING PLOTS")
    print("=" * 60)

    print("\n[1/14] Fleet-mix capacity...")
    plot_fleet_mix(cap_gen, outdir, scenarios)

    print("[2/14] New capacity cumulative bars...")
    plot_new_capacity_bars(nc, outdir, scenarios)

    print("[3/14] Binding MaxCapInv heatmaps...")
    plot_binding_maxcapinv(constr, outdir, scenarios)

    print("[4/14] MinCapInv floor analysis...")
    plot_mincapinv_floor(constr, outdir, scenarios)

    print("[5/14] Capacity delta vs BAU...")
    plot_capacity_delta_vs_bau(cap_gen, outdir, scenarios)

    print("[6/14] Generation mix...")
    plot_generation_mix(prod, outdir, scenarios)

    print("[7/14] Interconnector capacity...")
    plot_interconnectors(cap_all, outdir, scenarios)

    print("[8/14] Internal transmission...")
    plot_internal_transmission(cap_all, outdir, scenarios)

    print("[9/14] Backstop activation...")
    plot_backstop(cap_all, prod, outdir, scenarios)

    print("[10/14] Country capacity breakdown...")
    plot_country_capacity(cap_gen, outdir, scenarios)

    print("[11/14] TXT constraint comparison...")
    plot_txt_constraint_comparison(txt_files, outdir)

    print("[12/14] Constraint summary...")
    plot_constraint_summary(constr, outdir, scenarios)

    print("[13/14] New capacity time-series...")
    plot_new_capacity_timeseries(nc, outdir, scenarios)

    print("[14/14] Utilisation @ 2050...")
    plot_utilisation(df, outdir, scenarios)

    print("\n" + "=" * 60)
    print(f"  DONE — {len(os.listdir(outdir))} plots saved to {outdir}/")
    print("=" * 60)


if __name__ == "__main__":
    main()
