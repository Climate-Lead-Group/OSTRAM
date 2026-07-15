#!/usr/bin/env python3
"""
OSTRAM Interconnector & Trade Analysis Plotter
===============================================
Generates capacity + trade flow plots for bilateral interconnectors,
internal transmission, and RPO. Includes per-country trade breakdowns.

Usage:
    python ostram_trn_plotter.py --csv OSTRAM_Combined_Inputs_Outputs.csv
                                 [--output-dir trn_plots]
                                 [--scenarios BAU,A_Calibrated_BAU,B_Optimised_VRE]
"""

import argparse
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd

# ──────────────────────────────────────────────────────────────
#  CONFIGURATION
# ──────────────────────────────────────────────────────────────

DEFAULT_SCENARIOS = ["BAU", "A_Calibrated_BAU", "B_Optimised_VRE"]

SCENARIO_SHORT = {
    "BAU": "BAU₀",
    "A_Calibrated_BAU": "A-CalBAU",
    "B_Optimised_VRE": "B-OptVRE",
    "C_Target_VRE": "C-TgtVRE",
}
SC_COLORS = {
    "BAU": "#2c3e50",
    "A_Calibrated_BAU": "#e67e22",
    "B_Optimised_VRE": "#27ae60",
    "C_Target_VRE": "#c0392b",
}

COUNTRY_CODES = {
    "BGDXX": "Bangladesh", "BTNXX": "Bhutan",
    "INDEA": "India-East", "INDNE": "India-NE", "INDNO": "India-North",
    "INDSO": "India-South", "INDWE": "India-West",
    "LKAXX": "Sri Lanka", "MDVXX": "Maldives", "NPLXX": "Nepal",
}

COUNTRY_COLORS = {
    "Bangladesh": "#1f77b4", "Bhutan": "#aec7e8",
    "India-East": "#ff7f0e", "India-NE": "#ffbb78",
    "India-North": "#2ca02c", "India-South": "#98df8a",
    "India-West": "#d62728", "Sri Lanka": "#9467bd",
    "Maldives": "#c5b0d5", "Nepal": "#8c564b",
}

CTY_ORDER = list(COUNTRY_CODES.values())

CORRIDOR_NAMES = {
    "TRNBGDXXINDEA": "BGD↔IND-E", "TRNBGDXXINDNE": "BGD↔IND-NE",
    "TRNBTNXXBGDXX": "BTN↔BGD", "TRNBTNXXINDEA": "BTN↔IND-E",
    "TRNBTNXXINDNE": "BTN↔IND-NE", "TRNINDEAINDNE": "IND-E↔IND-NE",
    "TRNINDEAINDNO": "IND-E↔IND-N", "TRNINDEAINDSO": "IND-E↔IND-S",
    "TRNINDEAINDWE": "IND-E↔IND-W", "TRNINDEANPLXX": "IND-E↔NPL",
    "TRNINDNEINDNO": "IND-NE↔IND-N", "TRNINDNOINDWE": "IND-N↔IND-W",
    "TRNINDNONPLXX": "IND-N↔NPL", "TRNINDSOINDWE": "IND-S↔IND-W",
    "TRNINDSOLKAXX": "IND-S↔LKA", "TRNLKAXXMDVXX": "LKA↔MDV",
    "TRNMDVXXINDSO": "MDV↔IND-S", "TRNNPLXXBGDXX": "NPL↔BGD",
}

INDIA_INTERNAL = {"TRNINDEAINDNE", "TRNINDEAINDNO", "TRNINDEAINDSO",
                  "TRNINDEAINDWE", "TRNINDNEINDNO", "TRNINDNOINDWE",
                  "TRNINDSOINDWE"}


# ──────────────────────────────────────────────────────────────
#  HELPERS
# ──────────────────────────────────────────────────────────────

def ic_category(tech):
    if tech.startswith("TRNNLI"): return "NLI"
    if tech.startswith("TRNRPO"): return "RPO"
    if tech in INDIA_INTERNAL: return "India-Internal"
    return "Cross-Border"

def rpo_country(tech):
    return COUNTRY_CODES.get(tech.replace("TRNRPO",""), tech[-5:])

def get_sender(tech, receiver_code):
    a, b = tech[3:8], tech[8:13]
    if receiver_code == a: return COUNTRY_CODES.get(b, b)
    if receiver_code == b: return COUNTRY_CODES.get(a, a)
    return "Unknown"

def save_fig(fig, outdir, name):
    path = os.path.join(outdir, f"{name}.png")
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  ✓  {path}")

def sc_label(sc):
    return SCENARIO_SHORT.get(sc, sc)


def label_line_end(ax, x, y, color, fmt="{:.1f}", dx=5, fontsize=8):
    """Annotate the final (rightmost) value at the end of a line.
    Works with pandas Series, Index, or numpy arrays."""
    if len(x) == 0:
        return
    xl = x.iloc[-1] if hasattr(x, "iloc") else x[-1]
    yl = y.iloc[-1] if hasattr(y, "iloc") else y[-1]
    if pd.isna(xl) or pd.isna(yl):
        return
    ax.annotate(fmt.format(yl), xy=(xl, yl),
                xytext=(dx, 0), textcoords="offset points",
                fontsize=fontsize, va="center", fontweight="bold", color=color)


# ──────────────────────────────────────────────────────────────
#  DATA LOADING
# ──────────────────────────────────────────────────────────────

def load_data(csv_path, scenarios):
    print(f"Loading {csv_path} ...")
    df = pd.read_csv(csv_path,
        usecols=["Scenario","YEAR","TECHNOLOGY","FUEL",
                 "TotalCapacityAnnual","NewCapacity",
                 "ProductionByTechnologyAnnual",
                 "TotalTechnologyAnnualActivity"],
        low_memory=False)
    df["YEAR"] = df["YEAR"].astype("Int64")
    avail = [s for s in scenarios if s in df["Scenario"].unique()]
    print(f"  Scenarios: {avail}")
    df = df[df["Scenario"].isin(avail)]
    return df, avail


def extract_capacity(df):
    """Bilateral + RPO + NLI + PWRTRN capacity."""
    prefixes = ("TRNBGD","TRNBTN","TRNIND","TRNLKA","TRNMDV","TRNNPL",
                "TRNNLI","TRNRPO","PWRTRN")
    mask = df["TECHNOLOGY"].str.startswith(prefixes) & df["TotalCapacityAnnual"].notna()
    cap = df.loc[mask, ["Scenario","YEAR","TECHNOLOGY","TotalCapacityAnnual"]].drop_duplicates()
    cap["category"] = cap["TECHNOLOGY"].map(ic_category)
    cap.loc[cap["TECHNOLOGY"].str.startswith("PWRTRN"), "category"] = "PWRTRN"
    return cap


def extract_trade(df):
    """Bilateral production on 04 fuels = delivered power with direction."""
    bil_pfx = ("TRNBGD","TRNBTN","TRNIND","TRNLKA","TRNMDV","TRNNPL")
    mask = (df["TECHNOLOGY"].str.startswith(bil_pfx) &
            df["ProductionByTechnologyAnnual"].notna() &
            (df["ProductionByTechnologyAnnual"] != 0) &
            df["FUEL"].str.endswith("04"))
    tr = df.loc[mask, ["Scenario","YEAR","TECHNOLOGY","FUEL",
                       "ProductionByTechnologyAnnual"]].drop_duplicates()
    tr["receiver"] = tr["FUEL"].str[3:8]
    tr["receiver_name"] = tr["receiver"].map(COUNTRY_CODES)
    tr["sender_name"] = tr.apply(lambda r: get_sender(r["TECHNOLOGY"], r["receiver"]), axis=1)
    return tr


# ──────────────────────────────────────────────────────────────
#  PLOT FUNCTIONS
# ──────────────────────────────────────────────────────────────

def plot_T1_category_overview(cap, outdir, scenarios):
    """Aggregate transmission capacity by category."""
    cats = ["Cross-Border", "India-Internal", "RPO", "NLI", "PWRTRN"]
    cat_colors = {"Cross-Border": "#e74c3c", "India-Internal": "#3498db",
                  "RPO": "#2ecc71", "NLI": "#9b59b6", "PWRTRN": "#f39c12"}
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(5.5*n, 5), sharey=True)
    if n == 1: axes = [axes]
    for ax, sc in zip(axes, scenarios):
        agg = cap[cap["Scenario"]==sc].groupby(["YEAR","category"])["TotalCapacityAnnual"].sum().reset_index()
        for cat in cats:
            c = agg[agg["category"]==cat].sort_values("YEAR")
            if not c.empty:
                line, = ax.plot(c["YEAR"], c["TotalCapacityAnnual"], label=cat,
                                color=cat_colors[cat], linewidth=2)
                label_line_end(ax, c["YEAR"], c["TotalCapacityAnnual"], line.get_color(), fmt="{:.0f}")
        ax.set_title(sc_label(sc), fontweight="bold")
        ax.set_xlabel("Year"); ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.grid(alpha=0.2)
    axes[0].set_ylabel("Total capacity (GW)")
    axes[-1].legend(fontsize=8, loc="upper left")
    fig.suptitle("Transmission Capacity by Category", fontsize=13, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "T1_category_overview")


def plot_T2_cross_border(cap, outdir, scenarios):
    """Top cross-border bilateral corridors."""
    cross = cap[cap["category"]=="Cross-Border"]
    bau_2050 = cross[(cross["Scenario"]==scenarios[0]) & (cross["YEAR"]==2050)]
    tops = bau_2050.nlargest(6, "TotalCapacityAnnual")["TECHNOLOGY"].tolist()
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(5.5*n, 4.5), sharey=True)
    if n == 1: axes = [axes]
    colors = plt.cm.Set2(np.linspace(0, 1, len(tops)))
    for ax, sc in zip(axes, scenarios):
        sub = cross[cross["Scenario"]==sc]
        for i, tech in enumerate(tops):
            t = sub[sub["TECHNOLOGY"]==tech].sort_values("YEAR")
            if not t.empty:
                ax.plot(t["YEAR"], t["TotalCapacityAnnual"],
                        label=CORRIDOR_NAMES.get(tech, tech), color=colors[i], linewidth=1.8)
        ax.set_title(sc_label(sc), fontweight="bold")
        ax.set_xlabel("Year"); ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.grid(alpha=0.2)
    axes[0].set_ylabel("Capacity (GW)")
    axes[-1].legend(fontsize=7, loc="upper left")
    fig.suptitle("Cross-Border Interconnectors (Top 6)", fontsize=12, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "T2_cross_border")


def plot_T3_india_internal(cap, outdir, scenarios):
    """India-internal bilateral corridors."""
    india = cap[cap["category"]=="India-Internal"]
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(5.5*n, 4.5), sharey=True)
    if n == 1: axes = [axes]
    techs = sorted(india["TECHNOLOGY"].unique())
    colors = plt.cm.tab10(np.linspace(0, 1, len(techs)))
    for ax, sc in zip(axes, scenarios):
        sub = india[india["Scenario"]==sc]
        for i, tech in enumerate(techs):
            t = sub[sub["TECHNOLOGY"]==tech].sort_values("YEAR")
            if not t.empty:
                ax.plot(t["YEAR"], t["TotalCapacityAnnual"],
                        label=CORRIDOR_NAMES.get(tech, tech), color=colors[i], linewidth=1.8)
        ax.set_title(sc_label(sc), fontweight="bold")
        ax.set_xlabel("Year"); ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.grid(alpha=0.2)
    axes[0].set_ylabel("Capacity (GW)")
    axes[-1].legend(fontsize=7, loc="upper left")
    fig.suptitle("India-Internal Bilateral Interconnectors", fontsize=12, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "T3_india_internal")


def plot_T4_rpo(cap, outdir, scenarios):
    """RPO by destination country — stacked area."""
    rpo = cap[cap["category"]=="RPO"].copy()
    rpo["country"] = rpo["TECHNOLOGY"].map(rpo_country)
    cty_order = ["India-West","India-North","India-East","India-South",
                 "Bangladesh","Nepal","Sri Lanka","Bhutan","India-NE","Maldives"]
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(5.5*n, 5), sharey=True)
    if n == 1: axes = [axes]
    for ax, sc in zip(axes, scenarios):
        piv = rpo[rpo["Scenario"]==sc].pivot_table(
            index="YEAR", columns="country", values="TotalCapacityAnnual", aggfunc="sum").fillna(0)
        cols = [c for c in cty_order if c in piv.columns]
        piv = piv[cols]
        colors = [COUNTRY_COLORS.get(c, "#ccc") for c in cols]
        ax.stackplot(piv.index, *[piv[c].values for c in cols], labels=cols, colors=colors, alpha=0.85)
        ax.set_title(sc_label(sc), fontweight="bold")
        ax.set_xlabel("Year"); ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
    axes[0].set_ylabel("RPO capacity (GW)")
    h, l = axes[0].get_legend_handles_labels()
    fig.legend(h[::-1], l[::-1], loc="center left", bbox_to_anchor=(1.0, 0.5), fontsize=8)
    fig.suptitle("RPO (Repowering) Capacity by Country", fontsize=12, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "T4_rpo_by_country")


def plot_T5_pwrtrn(cap, outdir, scenarios):
    """Internal transmission (PWRTRN)."""
    trn = cap[cap["category"]=="PWRTRN"].copy()
    trn["country"] = trn["TECHNOLOGY"].map(lambda t: COUNTRY_CODES.get(t[-5:], t[-5:]))
    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(5.5*n, 4.5), sharey=True)
    if n == 1: axes = [axes]
    for ax, sc in zip(axes, scenarios):
        sub = trn[trn["Scenario"]==sc]
        for cty in sorted(sub["country"].unique()):
            t = sub[sub["country"]==cty].sort_values("YEAR")
            ax.plot(t["YEAR"], t["TotalCapacityAnnual"],
                    label=cty, color=COUNTRY_COLORS.get(cty, None), linewidth=1.5)
        ax.set_title(sc_label(sc), fontweight="bold")
        ax.set_xlabel("Year"); ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.grid(alpha=0.2)
    axes[0].set_ylabel("Internal TRN capacity (GW)")
    axes[-1].legend(fontsize=7, loc="upper left")
    fig.suptitle("Internal Transmission (PWRTRN)", fontsize=12, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "T5_internal_transmission")


def plot_T6_key_corridors(cap, outdir, scenarios):
    """Cross-scenario comparison for 6 key corridors."""
    keys = [("TRNBGDXXINDEA","Cross-Border"), ("TRNINDNOINDWE","India-Internal"),
            ("TRNINDSOINDWE","India-Internal"), ("TRNINDSOLKAXX","Cross-Border"),
            ("TRNBTNXXBGDXX","Cross-Border"), ("TRNNPLXXBGDXX","Cross-Border")]
    fig, axes = plt.subplots(2, 3, figsize=(16, 9))
    axes = axes.flatten()
    for i, (tech, cat) in enumerate(keys):
        ax = axes[i]
        for sc in scenarios:
            t = cap[(cap["Scenario"]==sc) & (cap["TECHNOLOGY"]==tech)].sort_values("YEAR")
            if not t.empty:
                line, = ax.plot(t["YEAR"], t["TotalCapacityAnnual"],
                                label=sc_label(sc), color=SC_COLORS[sc], linewidth=1.8)
                label_line_end(ax, t["YEAR"], t["TotalCapacityAnnual"], line.get_color(), fmt="{:.1f}", fontsize=7)
        ax.set_title(f"{CORRIDOR_NAMES.get(tech,tech)}\n({cat})", fontsize=10, fontweight="bold")
        ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.grid(alpha=0.2); ax.tick_params(labelsize=8)
    axes[0].legend(fontsize=8)
    axes[0].set_ylabel("Capacity (GW)"); axes[3].set_ylabel("Capacity (GW)")
    fig.suptitle("Key Corridors — Cross-Scenario", fontsize=13, fontweight="bold", y=1.01)
    fig.tight_layout()
    save_fig(fig, outdir, "T6_key_corridors")


def plot_T8_net_trade(trade, outdir, scenarios):
    """Net trade balance by country."""
    imp = trade.groupby(["Scenario","YEAR","receiver_name"])["ProductionByTechnologyAnnual"].sum().reset_index()
    imp.rename(columns={"receiver_name":"country","ProductionByTechnologyAnnual":"imp"}, inplace=True)
    exp = trade.groupby(["Scenario","YEAR","sender_name"])["ProductionByTechnologyAnnual"].sum().reset_index()
    exp.rename(columns={"sender_name":"country","ProductionByTechnologyAnnual":"exp"}, inplace=True)
    tb = imp.merge(exp, on=["Scenario","YEAR","country"], how="outer").fillna(0)
    tb["net"] = tb["imp"] - tb["exp"]

    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(5.5*n, 5), sharey=True)
    if n == 1: axes = [axes]
    for ax, sc in zip(axes, scenarios):
        sub = tb[tb["Scenario"]==sc]
        for cty in CTY_ORDER:
            c = sub[sub["country"]==cty].sort_values("YEAR")
            if not c.empty and c["net"].abs().max() > 1:
                ax.plot(c["YEAR"], c["net"], label=cty,
                        color=COUNTRY_COLORS.get(cty, None), linewidth=1.5)
        ax.axhline(0, color="gray", linewidth=0.5, linestyle="--")
        ax.set_title(sc_label(sc), fontweight="bold")
        ax.set_xlabel("Year"); ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.grid(alpha=0.2)
    axes[0].set_ylabel("Net import (PJ)")
    axes[0].legend(fontsize=7, loc="lower left")
    fig.suptitle("Net Trade Balance (positive = net importer)", fontsize=13, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, "T8_net_trade_balance")


def plot_T9_trade_matrix(trade, outdir, scenarios, year=2050):
    """Trade heatmap for a given year."""
    cf = trade.groupby(["Scenario","YEAR","sender_name","receiver_name"])[
        "ProductionByTechnologyAnnual"].sum().reset_index()
    countries_short = {"Bangladesh":"BGD","Bhutan":"BTN","India-East":"IND-E","India-NE":"IND-NE",
                       "India-North":"IND-N","India-South":"IND-S","India-West":"IND-W",
                       "Sri Lanka":"LKA","Maldives":"MDV","Nepal":"NPL"}
    labels = [countries_short[c] for c in CTY_ORDER]

    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(5*n, 5))
    if n == 1: axes = [axes]
    for ax, sc in zip(axes, scenarios):
        sub = cf[(cf["Scenario"]==sc) & (cf["YEAR"]==year)]
        mat = pd.DataFrame(0.0, index=CTY_ORDER, columns=CTY_ORDER)
        for _, r in sub.iterrows():
            if r["sender_name"] in CTY_ORDER and r["receiver_name"] in CTY_ORDER:
                mat.loc[r["sender_name"], r["receiver_name"]] += r["ProductionByTechnologyAnnual"]
        im = ax.imshow(mat.values, cmap="YlOrRd", aspect="auto", interpolation="nearest")
        ax.set_xticks(range(len(labels))); ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
        ax.set_yticks(range(len(labels))); ax.set_yticklabels(labels, fontsize=7)
        ax.set_xlabel("Receiver"); ax.set_ylabel("Sender")
        ax.set_title(f"{sc_label(sc)} — {year}", fontweight="bold")
        vmax = mat.values.max()
        for i in range(len(CTY_ORDER)):
            for j in range(len(CTY_ORDER)):
                v = mat.values[i, j]
                if v > 5:
                    ax.text(j, i, f"{v:.0f}", ha="center", va="center", fontsize=5,
                            color="white" if v > vmax*0.5 else "black")
        fig.colorbar(im, ax=ax, label="PJ", shrink=0.7)
    fig.suptitle(f"Bilateral Trade Matrix — {year} (Sender→Receiver, PJ)",
                 fontsize=13, fontweight="bold", y=1.02)
    fig.tight_layout()
    save_fig(fig, outdir, f"T9_trade_matrix_{year}")


def plot_country_trade(trade, outdir, scenarios, country):
    """Per-country trade breakdown: imports by source (stacked) + exports (dashed)."""
    imp = trade[trade["receiver_name"]==country].copy()
    exp = trade[trade["sender_name"]==country].copy()

    n = len(scenarios)
    fig, axes = plt.subplots(1, n, figsize=(5*n, 4), sharey=True)
    if n == 1: axes = [axes]
    for ax, sc in zip(axes, scenarios):
        # Stacked imports by source
        sub_imp = imp[imp["Scenario"]==sc]
        if not sub_imp.empty:
            sources = sub_imp.groupby(["YEAR","sender_name"])["ProductionByTechnologyAnnual"].sum().reset_index()
            piv = sources.pivot_table(index="YEAR", columns="sender_name",
                                      values="ProductionByTechnologyAnnual", aggfunc="sum").fillna(0)
            cols = [c for c in CTY_ORDER if c in piv.columns and c != country]
            if cols:
                piv = piv[cols]
                colors = [COUNTRY_COLORS.get(c, "#ccc") for c in cols]
                ax.stackplot(piv.index, *[piv[c].values for c in cols],
                             labels=cols, colors=colors, alpha=0.85)
        # Exports overlay
        sub_exp = exp[exp["Scenario"]==sc]
        if not sub_exp.empty:
            ea = sub_exp.groupby("YEAR")["ProductionByTechnologyAnnual"].sum()
            ax.plot(ea.index, -ea.values, color="black", linewidth=1.5,
                    linestyle="--", label="Exports (neg)")
        ax.axhline(0, color="gray", linewidth=0.5)
        ax.set_title(sc_label(sc), fontweight="bold", fontsize=10)
        ax.set_xlabel("Year"); ax.xaxis.set_major_locator(mticker.MultipleLocator(5))
        ax.grid(alpha=0.15)

    axes[0].set_ylabel("PJ")
    axes[0].legend(fontsize=6, loc="upper left")
    fig.suptitle(f"{country}: Imports by Source & Exports",
                 fontsize=12, fontweight="bold", y=1.02)
    fig.tight_layout()
    safe_name = country.replace(" ", "_").replace("-", "")
    save_fig(fig, outdir, f"TC_{safe_name}")


def plot_multi_country_summary(trade, outdir, scenarios):
    """Small-multiples grid: one row per country, columns = scenarios.
    Shows net import timeseries with colored fill."""
    imp = trade.groupby(["Scenario","YEAR","receiver_name"])["ProductionByTechnologyAnnual"].sum().reset_index()
    imp.rename(columns={"receiver_name":"country","ProductionByTechnologyAnnual":"imp"}, inplace=True)
    exp = trade.groupby(["Scenario","YEAR","sender_name"])["ProductionByTechnologyAnnual"].sum().reset_index()
    exp.rename(columns={"sender_name":"country","ProductionByTechnologyAnnual":"exp"}, inplace=True)
    tb = imp.merge(exp, on=["Scenario","YEAR","country"], how="outer").fillna(0)
    tb["net"] = tb["imp"] - tb["exp"]

    nrows = len(CTY_ORDER)
    ncols = len(scenarios)
    fig, axes = plt.subplots(nrows, ncols, figsize=(4*ncols, 2.2*nrows), sharex=True)

    for i, cty in enumerate(CTY_ORDER):
        for j, sc in enumerate(scenarios):
            ax = axes[i, j] if nrows > 1 else axes[j]
            sub = tb[(tb["Scenario"]==sc) & (tb["country"]==cty)].sort_values("YEAR")
            if not sub.empty:
                yrs = sub["YEAR"].values
                vals = sub["net"].values
                ax.fill_between(yrs, vals, 0, where=vals>=0,
                                color=COUNTRY_COLORS.get(cty, "#999"), alpha=0.4)
                ax.fill_between(yrs, vals, 0, where=vals<0,
                                color="#e74c3c", alpha=0.3)
                ax.plot(yrs, vals, color=COUNTRY_COLORS.get(cty, "#333"), linewidth=1.2)
            ax.axhline(0, color="gray", linewidth=0.3)
            ax.tick_params(labelsize=6)
            ax.xaxis.set_major_locator(mticker.MultipleLocator(10))

            if i == 0:
                ax.set_title(sc_label(sc), fontsize=9, fontweight="bold")
            if j == 0:
                ax.set_ylabel(cty, fontsize=8, fontweight="bold", rotation=0,
                              labelpad=60, ha="right", va="center")

    fig.suptitle("Net Bilateral Trade by Country × Scenario (PJ)\n"
                 "Green fill = net importer · Red fill = net exporter",
                 fontsize=12, fontweight="bold", y=1.01)
    fig.tight_layout()
    save_fig(fig, outdir, "TM_multi_country_summary")


# ──────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="OSTRAM Interconnector & Trade Plotter")
    parser.add_argument("--csv", required=True, help="OSTRAM_Combined_Inputs_Outputs.csv")
    parser.add_argument("--output-dir", default="trn_plots", help="Output directory")
    parser.add_argument("--scenarios", default=None,
                        help="Comma-separated scenario names (default: BAU,A_Calibrated_BAU,B_Optimised_VRE)")
    args = parser.parse_args()

    outdir = args.output_dir
    os.makedirs(outdir, exist_ok=True)
    scenarios = args.scenarios.split(",") if args.scenarios else DEFAULT_SCENARIOS

    # ── Load ──
    df, scenarios = load_data(args.csv, scenarios)

    print("\nExtracting capacity...")
    cap = extract_capacity(df)

    print("Extracting trade flows...")
    trade = extract_trade(df)

    # ── Capacity plots ──
    print("\n" + "="*60)
    print("  CAPACITY PLOTS")
    print("="*60)
    print("\n[T1] Category overview...")
    plot_T1_category_overview(cap, outdir, scenarios)
    print("[T2] Cross-border corridors...")
    plot_T2_cross_border(cap, outdir, scenarios)
    print("[T3] India-internal corridors...")
    plot_T3_india_internal(cap, outdir, scenarios)
    print("[T4] RPO by country...")
    plot_T4_rpo(cap, outdir, scenarios)
    print("[T5] Internal transmission...")
    plot_T5_pwrtrn(cap, outdir, scenarios)
    print("[T6] Key corridors cross-scenario...")
    plot_T6_key_corridors(cap, outdir, scenarios)

    # ── Trade plots ──
    print("\n" + "="*60)
    print("  TRADE FLOW PLOTS")
    print("="*60)
    print("\n[T8] Net trade balance...")
    plot_T8_net_trade(trade, outdir, scenarios)
    print("[T9] Trade matrix 2050...")
    plot_T9_trade_matrix(trade, outdir, scenarios, year=2050)
    print("[T9b] Trade matrix 2035...")
    plot_T9_trade_matrix(trade, outdir, scenarios, year=2035)

    # ── Per-country batch ──
    print("\n" + "="*60)
    print("  PER-COUNTRY TRADE BREAKDOWNS")
    print("="*60)
    for cty in CTY_ORDER:
        print(f"  [{cty}]...")
        plot_country_trade(trade, outdir, scenarios, cty)

    # ── Multi-country summary ──
    print("\n[TM] Multi-country summary grid...")
    plot_multi_country_summary(trade, outdir, scenarios)

    # ── Interactive HTML explorer ──
    print("\n" + "="*60)
    print("  INTERACTIVE HTML")
    print("="*60)
    print("\n[HTML] Trade explorer...")
    generate_html_explorer(trade, outdir, scenarios)

    print("\n" + "="*60)
    print(f"  DONE — {len(os.listdir(outdir))} files in {outdir}/")
    print("="*60)


# ──────────────────────────────────────────────────────────────
#  INTERACTIVE HTML EXPLORER
# ──────────────────────────────────────────────────────────────

def generate_html_explorer(trade, outdir, scenarios):
    """Self-contained HTML with year slider, animated trade matrices + balance bars."""
    import json

    # ── Prepare compact JSON ──
    # Flows: scenario × year × sender × receiver → PJ
    cf = trade.groupby(["Scenario","YEAR","sender_name","receiver_name"])[
        "ProductionByTechnologyAnnual"].sum().reset_index()
    flows = []
    for _, r in cf.iterrows():
        flows.append({"s": r["Scenario"], "y": int(r["YEAR"]),
                       "from": r["sender_name"], "to": r["receiver_name"],
                       "pj": round(r["ProductionByTechnologyAnnual"], 1)})

    # Balances: scenario × year × country → net PJ
    imp = trade.groupby(["Scenario","YEAR","receiver_name"])[
        "ProductionByTechnologyAnnual"].sum().reset_index()
    imp.rename(columns={"receiver_name":"country","ProductionByTechnologyAnnual":"imp"}, inplace=True)
    exp = trade.groupby(["Scenario","YEAR","sender_name"])[
        "ProductionByTechnologyAnnual"].sum().reset_index()
    exp.rename(columns={"sender_name":"country","ProductionByTechnologyAnnual":"exp"}, inplace=True)
    tb = imp.merge(exp, on=["Scenario","YEAR","country"], how="outer").fillna(0)
    tb["net"] = (tb["imp"] - tb["exp"]).round(1)
    balances = []
    for _, r in tb.iterrows():
        balances.append({"s": r["Scenario"], "y": int(r["YEAR"]),
                          "c": r["country"], "net": round(r["net"], 1)})

    flows_json = json.dumps(flows)
    bal_json = json.dumps(balances)

    # ── Scenario JS arrays ──
    sc_list_js = json.dumps(scenarios)
    sc_labels_js = json.dumps({s: sc_label(s) for s in scenarios})
    sc_colors_js = json.dumps({s: SC_COLORS.get(s, "#999") for s in scenarios})
    ctys_js = json.dumps(CTY_ORDER)
    cty_short_js = json.dumps({"Bangladesh":"BGD","Bhutan":"BTN","India-East":"IND-E",
        "India-NE":"IND-NE","India-North":"IND-N","India-South":"IND-S",
        "India-West":"IND-W","Sri Lanka":"LKA","Maldives":"MDV","Nepal":"NPL"})
    cty_colors_js = json.dumps(COUNTRY_COLORS)

    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OSTRAM · Trade Flow Explorer</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root {{
  --bg: #141B2A; --card: #1C2333; --border: #1E2A3D; --border2: #2A3548;
  --text: #E8ECF1; --text2: #C5CDD8; --text3: #8895A7; --text4: #6B7785;
  --accent: #D4623B;
}}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'DM Sans',system-ui,sans-serif; background:var(--bg); color:var(--text); padding:24px 28px; min-height:100vh; }}
.tag {{ font-size:10px; color:var(--accent); letter-spacing:0.15em; text-transform:uppercase; font-weight:700; margin-bottom:4px; }}
h1 {{ font-size:22px; font-weight:800; color:#F1F4F8; line-height:1.2; }}
.sub {{ font-size:12px; color:var(--text4); margin-top:4px; margin-bottom:6px; }}

.slider-row {{ display:flex; align-items:center; gap:16px; margin:20px 0 10px; }}
.slider-row label {{ font-size:13px; color:var(--text3); font-weight:600; }}
.slider-row input[type=range] {{ flex:1; accent-color:var(--accent); cursor:pointer; height:6px; }}
.year-display {{ font-size:28px; font-weight:800; color:var(--accent); font-variant-numeric:tabular-nums; min-width:56px; }}

.play-btn {{ background:var(--card); border:1px solid var(--border2); color:var(--accent); padding:6px 16px;
  border-radius:6px; cursor:pointer; font-family:inherit; font-weight:700; font-size:12px; transition:all 0.15s; }}
.play-btn:hover {{ background:rgba(212,98,59,0.15); }}

.speed-btn {{ background:var(--card); border:1px solid var(--border2); color:var(--text3); padding:4px 10px;
  border-radius:6px; cursor:pointer; font-family:inherit; font-weight:600; font-size:11px; transition:all 0.15s; }}
.speed-btn.active {{ border-color:var(--accent); color:var(--accent); background:rgba(212,98,59,0.1); }}

.grid {{ display:grid; grid-template-columns:repeat({len(scenarios)},1fr); gap:14px; margin-top:16px; }}
.panel {{ background:var(--card); border-radius:10px; padding:16px; position:relative; }}
.panel-title {{ font-size:11px; color:var(--text3); letter-spacing:0.06em; text-transform:uppercase; margin-bottom:8px; font-weight:700; }}
.sc-label {{ font-size:13px; font-weight:700; margin-bottom:6px; }}

.heatmap {{ display:grid; gap:1px; }}
.hm-cell {{ aspect-ratio:1; display:flex; align-items:center; justify-content:center; font-size:8px;
  font-weight:600; border-radius:2px; transition:background 0.25s; font-variant-numeric:tabular-nums; }}
.hm-label {{ font-size:8px; color:var(--text3); font-weight:600; display:flex; align-items:center; justify-content:center; }}

.bal-row {{ display:flex; align-items:center; gap:6px; margin-bottom:3px; }}
.bal-label {{ font-size:9px; color:var(--text3); width:56px; text-align:right; font-weight:600; }}
.bal-bar-wrap {{ flex:1; height:14px; position:relative; background:var(--border); border-radius:2px; overflow:hidden; }}
.bal-bar {{ height:100%; position:absolute; top:0; border-radius:2px; transition:width 0.25s, left 0.25s; }}
.bal-val {{ font-size:9px; color:var(--text2); width:52px; font-variant-numeric:tabular-nums; font-weight:600; }}

.insight {{ background:rgba(212,98,59,0.08); border-radius:8px; padding:12px 14px; font-size:11px;
  color:var(--text3); line-height:1.6; margin-top:14px; grid-column:1/-1; }}
.insight strong {{ color:var(--accent); }}

.totals {{ display:grid; grid-template-columns:repeat({len(scenarios)},1fr); gap:14px; margin-top:14px; }}
.total-card {{ background:var(--card); border-radius:8px; padding:10px 14px; }}
.total-card .label {{ font-size:10px; color:var(--text3); text-transform:uppercase; letter-spacing:0.06em; font-weight:700; }}
.total-card .val {{ font-size:20px; font-weight:800; font-variant-numeric:tabular-nums; margin-top:2px; }}
.total-card .detail {{ font-size:10px; color:var(--text4); margin-top:2px; }}

@media (max-width:900px) {{ .grid,.totals {{ grid-template-columns:1fr; }} body {{ padding:16px 12px; }} }}
</style>
</head>
<body>

<div class="tag">OSTRAM · Interconnector Trade Explorer</div>
<h1>Bilateral Energy Flows by Year</h1>
<div class="sub">Drag the slider or hit Play to watch trade patterns evolve · 10 countries × {len(scenarios)} scenarios × 2023–2050</div>

<div class="slider-row">
  <label>Year</label>
  <span class="year-display" id="yr-display">2023</span>
  <input type="range" id="yr-slider" min="2023" max="2050" value="2023" step="1">
  <button class="play-btn" id="play-btn">▶ Play</button>
  <button class="speed-btn" data-ms="600" id="sp1">1×</button>
  <button class="speed-btn active" data-ms="400" id="sp2">2×</button>
  <button class="speed-btn" data-ms="200" id="sp3">4×</button>
</div>

<div class="totals" id="totals"></div>
<div class="grid" id="panels"></div>
<div class="grid"><div class="insight" id="insight-box"></div></div>

<script>
const FLOWS = {flows_json};
const BALS = {bal_json};
const SCENARIOS = {sc_list_js};
const SC_LABELS = {sc_labels_js};
const SC_COLORS = {sc_colors_js};
const CTYS = {ctys_js};
const CTY_SHORT = {cty_short_js};
const CTY_COLORS = {cty_colors_js};

let speed = 400;

function heatColor(val, mx) {{
  if (val < 0.5 || mx < 1) return 'transparent';
  const t = Math.min(val / Math.max(mx, 1), 1);
  const r = Math.round(80 + 175 * t);
  const g = Math.round(60 + 40 * (1-t));
  const b = Math.round(40 + 10 * (1-t));
  return `rgba(${{r}},${{g}},${{b}},${{0.3 + 0.7*t}})`;
}}

function buildPanels() {{
  const container = document.getElementById('panels');
  container.innerHTML = '';
  SCENARIOS.forEach(sc => {{
    const p = document.createElement('div');
    p.className = 'panel';
    p.innerHTML = `<div class="sc-label" style="color:${{SC_COLORS[sc]}}">${{SC_LABELS[sc]}}</div>
      <div class="panel-title">Trade matrix (PJ) — sender ↓ receiver →</div>
      <div id="hm-${{sc}}" style="margin-bottom:14px"></div>
      <div class="panel-title">Net balance (PJ) — importer ▸ | ◂ exporter</div>
      <div id="bal-${{sc}}"></div>`;
    container.appendChild(p);
  }});
}}

function renderHeatmap(sc, year) {{
  const el = document.getElementById(`hm-${{sc}}`);
  const sf = FLOWS.filter(d => d.s === sc && d.y === year);
  const n = CTYS.length;
  const mat = Array.from({{length:n}}, () => Array(n).fill(0));
  let mx = 0;
  sf.forEach(d => {{
    const i = CTYS.indexOf(d.from), j = CTYS.indexOf(d.to);
    if (i >= 0 && j >= 0) {{ mat[i][j] += d.pj; mx = Math.max(mx, mat[i][j]); }}
  }});
  let html = `<div class="heatmap" style="grid-template-columns:32px repeat(${{n}},1fr)">`;
  html += '<div></div>';
  CTYS.forEach(c => html += `<div class="hm-label">${{CTY_SHORT[c]}}</div>`);
  for (let i = 0; i < n; i++) {{
    html += `<div class="hm-label">${{CTY_SHORT[CTYS[i]]}}</div>`;
    for (let j = 0; j < n; j++) {{
      const v = mat[i][j];
      const bg = heatColor(v, mx);
      const txt = v > 5 ? Math.round(v) : '';
      const fc = v > mx*0.5 ? '#fff' : (v > 5 ? '#ccc' : 'transparent');
      html += `<div class="hm-cell" style="background:${{bg}};color:${{fc}}">${{txt}}</div>`;
    }}
  }}
  html += '</div>';
  el.innerHTML = html;
}}

function renderBalance(sc, year) {{
  const el = document.getElementById(`bal-${{sc}}`);
  const bf = BALS.filter(d => d.s === sc && d.y === year);
  const vals = {{}};
  bf.forEach(d => vals[d.c] = d.net);
  const maxAbs = Math.max(1, ...CTYS.map(c => Math.abs(vals[c] || 0)));
  let html = '';
  CTYS.forEach(c => {{
    const v = vals[c] || 0;
    const pct = Math.abs(v) / maxAbs * 45;
    const isPos = v >= 0;
    const left = isPos ? 50 : 50 - pct;
    const color = isPos ? CTY_COLORS[c] : '#e74c3c';
    html += `<div class="bal-row">
      <div class="bal-label">${{CTY_SHORT[c]}}</div>
      <div class="bal-bar-wrap">
        <div class="bal-bar" style="left:${{left}}%;width:${{pct}}%;background:${{color}};opacity:0.7"></div>
        <div style="position:absolute;left:50%;top:0;bottom:0;width:1px;background:var(--text4);opacity:0.3"></div>
      </div>
      <div class="bal-val" style="color:${{isPos?'#8bc':'#e88'}}">${{v>0?'+':''}}${{Math.round(v)}}</div>
    </div>`;
  }});
  el.innerHTML = html;
}}

function renderTotals(year) {{
  const el = document.getElementById('totals');
  let html = '';
  SCENARIOS.forEach(sc => {{
    const sf = FLOWS.filter(d => d.s === sc && d.y === year);
    const total = sf.reduce((a,d) => a + d.pj, 0);
    const bf = BALS.filter(d => d.s === sc && d.y === year);
    const topImp = bf.reduce((a,b) => (a.net||0) > (b.net||0) ? a : b, {{c:'—',net:0}});
    const topExp = bf.reduce((a,b) => (a.net||0) < (b.net||0) ? a : b, {{c:'—',net:0}});
    html += `<div class="total-card" style="border-left:3px solid ${{SC_COLORS[sc]}}">
      <div class="label">${{SC_LABELS[sc]}}</div>
      <div class="val" style="color:${{SC_COLORS[sc]}}">${{Math.round(total).toLocaleString()}} PJ</div>
      <div class="detail">Top importer: ${{CTY_SHORT[topImp.c]||'—'}} (+${{Math.round(topImp.net||0)}}) · Exporter: ${{CTY_SHORT[topExp.c]||'—'}} (${{Math.round(topExp.net||0)}})</div>
    </div>`;
  }});
  el.innerHTML = html;
}}

function renderInsight(year) {{
  const el = document.getElementById('insight-box');
  const lines = SCENARIOS.map(sc => {{
    const sf = FLOWS.filter(d => d.s === sc && d.y === year);
    if (!sf.length) return '';
    const top = sf.reduce((a,b) => a.pj > b.pj ? a : b);
    const bf = BALS.filter(d => d.s === sc && d.y === year);
    const nCorr = new Set(sf.filter(d=>d.pj>1).map(d=>d.from+d.to)).size;
    return `<strong>${{SC_LABELS[sc]}}</strong>: ${{nCorr}} active corridors · largest: ${{CTY_SHORT[top.from]}}→${{CTY_SHORT[top.to]}} (${{Math.round(top.pj)}} PJ)`;
  }});
  el.innerHTML = lines.join('<br>');
}}

function update(year) {{
  document.getElementById('yr-display').textContent = year;
  document.getElementById('yr-slider').value = year;
  renderTotals(year);
  SCENARIOS.forEach(sc => {{
    renderHeatmap(sc, year);
    renderBalance(sc, year);
  }});
  renderInsight(year);
}}

document.getElementById('yr-slider').addEventListener('input', e => update(+e.target.value));

let playing = false, timer = null;
document.getElementById('play-btn').addEventListener('click', () => {{
  if (playing) {{
    clearInterval(timer); playing = false;
    document.getElementById('play-btn').textContent = '▶ Play';
  }} else {{
    playing = true;
    document.getElementById('play-btn').textContent = '⏸ Pause';
    let yr = +document.getElementById('yr-slider').value;
    if (yr >= 2050) yr = 2023;
    timer = setInterval(() => {{
      yr++;
      if (yr > 2050) {{ clearInterval(timer); playing = false; document.getElementById('play-btn').textContent = '▶ Play'; return; }}
      update(yr);
    }}, speed);
  }}
}});

document.querySelectorAll('.speed-btn').forEach(btn => {{
  btn.addEventListener('click', () => {{
    speed = +btn.dataset.ms;
    document.querySelectorAll('.speed-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    if (playing) {{
      clearInterval(timer);
      let yr = +document.getElementById('yr-slider').value;
      timer = setInterval(() => {{
        yr++;
        if (yr > 2050) {{ clearInterval(timer); playing = false; document.getElementById('play-btn').textContent = '▶ Play'; return; }}
        update(yr);
      }}, speed);
    }}
  }});
}});

buildPanels();
update(2023);
</script>
</body>
</html>'''

    path = os.path.join(outdir, "OSTRAM_Trade_Explorer.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✓  {path}")


if __name__ == "__main__":
    main()
