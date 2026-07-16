#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reproduce_A1_A6.py
==================
Regenerate summary figures A1-A6 from the OSTRAM combined long-format dump
(OSTRAM_Combined_Inputs_Outputs.csv) for scenarios A_Calibrated_BAU (A-CalBAU)
and B_Optimised_VRE (B-OptVRE).

    A1  Annual discounted cost          (US$ B)
    A2  Cumulative discounted cost      (US$ B)
    A3  Coal share of generation        (%)
    A4  VRE share (solar + wind)        (%)
    A5  Petroleum + Oil capacity        (GW)
    A6  Storage capacity SDS / LDS      (GW)

Design notes
------------
* Config-only knobs live in the CONFIG block below. No CLI args.
* Non-destructive: reads the source CSV, writes PNGs (+ an audit CSV) into OUT_DIR.
* Duplicate-safe: each metric is de-duplicated on its natural OSeMOSYS index
  before aggregation, so a dupe-afflicted export cannot double-count.
* Family is derived from the PWR<XXX> technology prefix. If a TECH_TYPES.csv is
  present next to the data it is NOT required here (prefix mapping is sufficient
  and was validated to reproduce Set A exactly).
"""

import os
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
HERE     = str(Path(__file__).resolve().parents[2] / "t1_confection")
CSV_PATH = os.path.join(HERE, "OSTRAM_Combined_Inputs_Outputs.csv")
OUT_DIR  = os.path.join(HERE, "figs_A1_A6")

# scenario id in the file  ->  short label used on the plots
SCENARIOS = {
    "A_Calibrated_BAU": "A-CalBAU",
    "B_Optimised_VRE":  "B-OptVRE",
}

# PWR<prefix> family groupings (prefix = TECHNOLOGY characters 4-6)
COAL_FAMS    = ["COA"]                      # A3 numerator
VRE_FAMS     = ["SPV", "WON", "WOF"]        # A4 numerator (solar PV + wind on/offshore; CSP excluded)
PETOIL_FAMS  = ["PET", "OIL"]              # A5
SDS_FAMS     = ["SDS"]                      # A6 short-duration storage
LDS_FAMS     = ["LDS"]                      # A6 long-duration storage

# generation families used as the share DENOMINATOR (A3, A4).
# Core generation set that reproduces Set A; storage / transmission / backstop excluded.
GEN_FAMS = ["COA", "PET", "OIL", "NGS", "URN", "WAS", "BIO",
            "HYD", "SHP", "CSP", "SPV", "WON", "WOF"]

COST_TO_BILLION = 1.0 / 1000.0             # TotalDiscountedCost is in US$ M -> US$ B

# colours (approx. Set A palette)
C_A   = "#1f4e6b"   # A-CalBAU  (dark navy)
C_B   = "#2bb6a3"   # B-OptVRE  (teal-green)
C_ALDS = "#7fbfe0"  # A-CalBAU  LDS (light blue)
C_BLDS = "#1d7a4d"  # B-OptVRE  LDS (dark green)

ORDER = list(SCENARIOS.values())           # consistent A, B ordering


# ----------------------------------------------------------------------------
# LOAD
# ----------------------------------------------------------------------------
def load():
    cols = ["Scenario", "REGION", "YEAR", "TECHNOLOGY", "FUEL",
            "TotalDiscountedCost", "TotalCapacityAnnual",
            "ProductionByTechnologyAnnual"]
    df = pd.read_csv(CSV_PATH, usecols=cols, low_memory=False)
    df["YEAR"] = pd.to_numeric(df["YEAR"], errors="coerce").astype("Int64")
    df = df[df["Scenario"].isin(SCENARIOS)].copy()
    df["S"]   = df["Scenario"].map(SCENARIOS)
    df["fam"] = df["TECHNOLOGY"].str[3:6].where(df["TECHNOLOGY"].str.startswith("PWR"))
    return df


# ----------------------------------------------------------------------------
# METRIC BUILDERS  (each returns a wide DataFrame indexed by YEAR, columns=labels)
# ----------------------------------------------------------------------------
def capacity_by_fams(df, fams, years=None):
    """Sum TotalCapacityAnnual (GW) over the given families, per scenario-year.
    Reindexed onto `years` (pre-build years -> 0) so all curves share an x-axis."""
    s = df[df["TotalCapacityAnnual"].notna() & df["fam"].isin(fams)]
    s = s.drop_duplicates(subset=["S", "TECHNOLOGY", "YEAR"])          # dupe-safe
    w = s.groupby(["YEAR", "S"])["TotalCapacityAnnual"].sum().unstack("S")
    w = w.reindex(columns=ORDER)
    if years is not None:
        w = w.reindex(years).fillna(0.0)
    return w


def generation(df):
    """Generation (PJ) by scenario-year-family using ProductionByTechnologyAnnual."""
    s = df[df["ProductionByTechnologyAnnual"].notna() & df["fam"].isin(GEN_FAMS)]
    s = s.drop_duplicates(subset=["S", "TECHNOLOGY", "FUEL", "YEAR"])  # dupe-safe
    return s.groupby(["YEAR", "S", "fam"])["ProductionByTechnologyAnnual"].sum().reset_index()


def share(gen_long, fams):
    """Percentage share of `fams` in total generation, per scenario-year."""
    tot = gen_long.groupby(["YEAR", "S"])["ProductionByTechnologyAnnual"].sum()
    num = (gen_long[gen_long["fam"].isin(fams)]
           .groupby(["YEAR", "S"])["ProductionByTechnologyAnnual"].sum())
    w = (100.0 * num / tot).unstack("S")
    return w.reindex(columns=ORDER)


def cost(df):
    """Annual discounted cost (US$ B), per scenario-year."""
    s = df[df["TotalDiscountedCost"].notna()]
    s = s.drop_duplicates(subset=["S", "REGION", "YEAR"])             # dupe-safe
    w = s.groupby(["YEAR", "S"])["TotalDiscountedCost"].sum().unstack("S") * COST_TO_BILLION
    return w.reindex(columns=ORDER)


# ----------------------------------------------------------------------------
# PLOTTING
# ----------------------------------------------------------------------------
def _style_axes(ax, title, ylabel):
    ax.set_title(title, fontsize=15, pad=12, weight="bold", color="#243b53")
    ax.set_ylabel(ylabel, fontsize=11)
    ax.grid(axis="y", color="#e6e9ee", lw=1, zorder=0)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    ax.tick_params(axis="x", rotation=45, labelsize=8)
    ax.margins(x=0.01)


def line_two_series(df_wide, title, ylabel, fname, colors=(C_A, C_B)):
    fig, ax = plt.subplots(figsize=(9.6, 6.6))
    for lab, col in zip(ORDER, colors):
        if lab in df_wide.columns:
            ax.plot(df_wide.index.astype(int), df_wide[lab], marker="o", ms=5,
                    lw=2.6, color=col, label=lab, zorder=3)
    _style_axes(ax, title, ylabel)
    ax.set_ylim(bottom=0)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=2,
              frameon=False, fontsize=11)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, fname), dpi=150, bbox_inches="tight")
    plt.close(fig)


def line_storage(sds, lds, title, ylabel, fname):
    fig, ax = plt.subplots(figsize=(10.0, 6.6))
    yrs = sds.index.astype(int)
    series = [
        ("A-CalBAU \u2014 SDS", sds.get("A-CalBAU"), C_A),
        ("A-CalBAU \u2014 LDS", lds.get("A-CalBAU"), C_ALDS),
        ("B-OptVRE \u2014 SDS", sds.get("B-OptVRE"), C_B),
        ("B-OptVRE \u2014 LDS", lds.get("B-OptVRE"), C_BLDS),
    ]
    for lab, y, col in series:
        if y is not None:
            ax.plot(yrs, y, marker="o", ms=4.5, lw=2.4, color=col, label=lab, zorder=3)
    _style_axes(ax, title, ylabel)
    ax.set_ylim(bottom=0)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=4,
              frameon=False, fontsize=10)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, fname), dpi=150, bbox_inches="tight")
    plt.close(fig)


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    df = load()
    YEARS = list(range(int(df["YEAR"].min()), int(df["YEAR"].max()) + 1))

    # --- build series ---
    annual_cost = cost(df)                         # A1
    cum_cost    = annual_cost.cumsum()             # A2
    gen_long    = generation(df)
    coal_share  = share(gen_long, COAL_FAMS)       # A3
    vre_share   = share(gen_long, VRE_FAMS)        # A4
    petoil      = capacity_by_fams(df, PETOIL_FAMS, YEARS)  # A5
    sds         = capacity_by_fams(df, SDS_FAMS, YEARS)     # A6
    lds         = capacity_by_fams(df, LDS_FAMS, YEARS)     # A6

    # --- plots ---
    line_two_series(annual_cost, "Annual discounted cost (US$ B)",
                    "US$ B", "A1_annual_discounted_cost.png")
    line_two_series(cum_cost, "Cumulative discounted cost (US$ B)",
                    "US$ B", "A2_cumulative_discounted_cost.png")
    line_two_series(coal_share, "Coal share of generation (%)",
                    "%", "A3_coal_share.png")
    line_two_series(vre_share, "VRE share (solar + wind, %)",
                    "%", "A4_vre_share.png")
    line_two_series(petoil, "Petroleum + Oil capacity (GW)",
                    "GW", "A5_petroleum_oil_capacity.png")
    line_storage(sds, lds, "Storage capacity (GW)",
                 "GW", "A6_storage_capacity.png")

    # --- audit dump (one tidy CSV with every plotted series) ---
    def melt(w, metric):
        return (w.reset_index().melt("YEAR", var_name="Scenario", value_name="value")
                  .assign(metric=metric))
    audit = pd.concat([
        melt(annual_cost, "A1_annual_cost_USD_B"),
        melt(cum_cost,    "A2_cumulative_cost_USD_B"),
        melt(coal_share,  "A3_coal_share_pct"),
        melt(vre_share,   "A4_vre_share_pct"),
        melt(petoil,      "A5_petroleum_oil_GW"),
        melt(sds,         "A6_storage_SDS_GW"),
        melt(lds,         "A6_storage_LDS_GW"),
    ], ignore_index=True).dropna(subset=["value"])
    audit.to_csv(os.path.join(OUT_DIR, "A1_A6_series_audit.csv"), index=False)

    print("Done. Outputs in:", OUT_DIR)
    for f in sorted(os.listdir(OUT_DIR)):
        print("  ", f)


if __name__ == "__main__":
    main()
