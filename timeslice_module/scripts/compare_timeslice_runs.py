#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
compare_timeslice_runs.py
=========================

Compare two solved OSTRAM runs that differ in their timeslice fabric, at the
highest temporal resolution the results carry: per-timeslice dispatch, demand,
plus a summary table of cost and new capacity by tech class.
Storage dispatch plots were removed deliberately: NetChargeWithinDay is empty
at training-model scale, so storage differences appear in NewCapacity (see the
summary table), not in intra-day cycling.

Reads (per run):
    <outputs_dir>/ProductionByTechnology.csv     (REGION,TIMESLICE,TECHNOLOGY,FUEL,YEAR,VALUE)
    <outputs_dir>/Demand.csv
    <outputs_dir>/NewCapacity.csv
    <outputs_dir>/TotalDiscountedCost.csv
    <fabric_dir>/DaySplit.csv                    (bracket -> fraction of day)
    <fabric_dir>/YearSplit.csv                   (timeslice -> fraction of year)

Every plotted mark traces to a results row. The only derivation is the
conversion of per-timeslice ENERGY into AVERAGE POWER within the bracket
(value / hours-in-bracket-over-the-year), and the placement of brackets on
the 24h axis, which assumes brackets are contiguous from midnight in bracket
order (true for the OSTRAM fabric: D1 starts at 00:00).

Non-destructive: reads inputs, writes figures and a CSV table into --out-dir.

USAGE
    python compare_timeslice_runs.py \
        --run <runA>/Outputs <runA_fabric> "20 ts (5dp)" \
        --run <runB>/Outputs <runB_fabric> "16 ts (4x6h)" \
        --run <runC>/Outputs <runC_fabric> "12 ts (3x8h)" \
        --year 2035 --season 3 --out-dir ./ts_compare_out

REQUIREMENTS: pandas, matplotlib
"""
from __future__ import annotations
import argparse, os, re
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# House palette (OSTRAM report)
NAVY, TEAL, GREY, BORDER = "#0C2340", "#15A39B", "#565F6B", "#D2D8DE"
TECH_COLORS = {
    "COA": "#3A4756", "SPV": "#F2B705", "WON": TEAL, "WOF": "#0F7F79",
    "HYD": "#2E6E8E", "SHP": "#4E8EAE", "URN": "#7E57C2", "NGS": "#9AA5B1",
    "OIL": "#C0504D", "HFO": "#C0504D", "HSD": "#A0403D", "BIO": "#8FB56B",
    "WAS": "#6E9550", "GEO": "#B07C4F", "SDS": "#D98CB3", "LDS": "#8C6BB1",
    "TRN": "#B8C0C9", "DSP": "#D2D8DE",
}
TS_RE = re.compile(r"S(\d+)D(\d+)")


def tech_class(t: str) -> str:
    m = re.match(r"(?:PWR|DSP)?([A-Z]{3})", str(t))
    return m.group(1) if m else str(t)[:3]


def load_run(outputs: str, fabric: str):
    r = {}
    r["prod"] = pd.read_csv(os.path.join(outputs, "ProductionByTechnology.csv"))
    r["dem"] = pd.read_csv(os.path.join(outputs, "Demand.csv"))
    ncw = os.path.join(outputs, "NetChargeWithinDay.csv")
    r["charge"] = pd.read_csv(ncw) if os.path.exists(ncw) else None
    r["newcap"] = pd.read_csv(os.path.join(outputs, "NewCapacity.csv"))
    r["cost"] = pd.read_csv(os.path.join(outputs, "TotalDiscountedCost.csv"))
    def read_param(name, param):
        df = pd.read_csv(os.path.join(fabric, name))
        if "VALUE" not in df.columns:  # wide precursor schema: PARAMETER,...,Value
            df = df[df["PARAMETER"] == param].rename(columns={"Value": "VALUE"})
        return df
    ds = read_param("DaySplit.csv", "DaySplit")
    ys = read_param("YearSplit.csv", "YearSplit")
    y0 = ds["YEAR"].min()
    ds = ds[ds["YEAR"] == y0]
    hours = {int(b): float(v) * 24.0 for b, v in zip(ds["DAILYTIMEBRACKET"], ds["VALUE"])}
    # bracket start hour, assuming contiguity from midnight in bracket order
    starts, cursor = {}, 0.0
    for b in sorted(hours):
        starts[b] = cursor
        cursor += hours[b]
    assert abs(cursor - 24.0) < 1e-6, f"bracket hours sum to {cursor}, not 24"
    r["hours"], r["starts"] = hours, starts
    ys0 = ys[ys["YEAR"] == ys["YEAR"].min()]
    r["yearsplit"] = {t: float(v) for t, v in zip(ys0["TIMESLICE"], ys0["VALUE"])}
    return r


def ts_profile(df: pd.DataFrame, run: dict, year: int, season: int,
               value_col="VALUE", group_tech=True):
    """Per-bracket average power (PJ/yr equivalent rate) for one season+year.

    energy-in-ts / (yearsplit-of-ts * 8760h) = average rate during the ts.
    Returned per bracket with x-extent [start, start+hours] for step plotting.
    """
    d = df[df["YEAR"] == year].copy()
    d["S"] = d["TIMESLICE"].str.extract(r"S(\d+)D\d+").astype(float)
    d["B"] = d["TIMESLICE"].str.extract(r"S\d+D(\d+)").astype(float)
    d = d[d["S"] == season]
    keys = ["B", "TECHCLASS"] if group_tech else ["B"]
    if group_tech:
        d["TECHCLASS"] = d["TECHNOLOGY"].map(tech_class)
    g = d.groupby(keys)[value_col].sum().reset_index()
    out = {}
    for _, row in g.iterrows():
        b = int(row["B"])
        ts = f"S{season}D{b}"
        ysf = run["yearsplit"].get(ts)
        if not ysf:
            continue
        rate = row[value_col] / (ysf * 8760.0) * 277.7777778  # PJ per ts-hour -> GW average
        key = row["TECHCLASS"] if group_tech else "total"
        out.setdefault(key, {})[b] = rate
    return out


def plot_dispatch(runs, labels, year, season, out_path):
    n = len(runs)
    fig, axes = plt.subplots(1, n, figsize=(6.5 * n, 5), sharey=True)
    if n == 1:
        axes = [axes]
    for ax, run, label in zip(axes, runs, labels):
        prod = run["prod"]
        gen = prod[prod["TECHNOLOGY"].str.startswith(("PWR",), na=False)]
        prof = ts_profile(gen, run, year, season)
        brackets = sorted(run["hours"])
        order = sorted(prof, key=lambda c: -sum(prof[c].values()))
        bottoms = {b: 0.0 for b in brackets}
        for cls in order:
            xs, hs, ys, bots = [], [], [], []
            for b in brackets:
                v = prof[cls].get(b, 0.0)
                xs.append(run["starts"][b]); hs.append(run["hours"][b])
                ys.append(v); bots.append(bottoms[b]); bottoms[b] += v
            ax.bar(xs, ys, width=hs, bottom=bots, align="edge",
                   color=TECH_COLORS.get(cls, "#B8C0C9"), label=cls,
                   edgecolor="white", linewidth=0.4)
        for b in brackets:
            ax.axvline(run["starts"][b], color=BORDER, linewidth=0.6, zorder=0)
        ax.set_xlim(0, 24); ax.set_xticks(range(0, 25, 3))
        ax.set_xlabel("hour of day", color=GREY, fontsize=9)
        ax.set_title(label, color=NAVY, fontsize=11, fontweight="bold", loc="left")
        ax.tick_params(colors=GREY, labelsize=8)
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
    axes[0].set_ylabel("average output within bracket (GW)", color=GREY, fontsize=9)
    handles, lab = axes[0].get_legend_handles_labels()
    fig.legend(handles, lab, ncol=min(len(lab), 8), fontsize=8, frameon=False,
               loc="lower center", bbox_to_anchor=(0.5, -0.04))
    fig.suptitle(f"Intra-day dispatch by technology - season {season}, {year}",
                 color=NAVY, fontsize=13, fontweight="bold", x=0.07, ha="left")
    fig.tight_layout(rect=(0, 0.05, 1, 0.95))
    fig.savefig(out_path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def plot_demand(runs, labels, year, season, out_path):
    fig, ax = plt.subplots(figsize=(9, 4.5))
    colors = [NAVY, TEAL, "#C0504D", "#F2B705", "#7E57C2"]
    for run, label, col in zip(runs, labels, colors):
        prof = ts_profile(run["dem"], run, year, season, group_tech=False)
        pts = prof.get("total", {})
        xs, ys = [], []
        for b in sorted(run["hours"]):
            x0 = run["starts"][b]; h = run["hours"][b]
            v = pts.get(b, 0.0)
            xs += [x0, x0 + h]; ys += [v, v]
        ax.plot(xs, ys, color=col, linewidth=2.2, label=label)
    ax.set_xlim(0, 24); ax.set_xticks(range(0, 25, 3))
    ax.set_xlabel("hour of day", color=GREY, fontsize=9)
    ax.set_ylabel("average demand within bracket (GW)", color=GREY, fontsize=9)
    ax.tick_params(colors=GREY, labelsize=8)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.legend(frameon=False, fontsize=9)
    ax.set_title(f"Demand profile as the fabric sees it - season {season}, {year}",
                 color=NAVY, fontsize=12, fontweight="bold", loc="left")
    fig.tight_layout()
    fig.savefig(out_path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def summary_table(runs, labels, out_path):
    rows = []
    for run, label in zip(runs, labels):
        cost = float(run["cost"]["VALUE"].sum())
        nc = run["newcap"]
        by = nc.assign(C=nc["TECHNOLOGY"].map(tech_class)).groupby("C")["VALUE"].sum()
        row = {"run": label, "sum_TotalDiscountedCost_csv": round(cost, 2)}
        for c in sorted(by.index):
            row[f"newcap_{c}"] = round(float(by[c]), 3)
        rows.append(row)
    df = pd.DataFrame(rows)
    df.to_csv(out_path, index=False)
    return df


def main():
    ap = argparse.ArgumentParser(
        description="Compare N timeslice-fabric runs (demand, dispatch, summary).")
    ap.add_argument("--run", nargs=3, action="append", required=True,
                    metavar=("OUTPUTS_DIR", "FABRIC_DIR", "LABEL"),
                    help="repeatable: results Outputs dir, otoole fabric dir, plot label")
    ap.add_argument("--year", type=int, default=2035)
    ap.add_argument("--season", type=int, default=3)
    ap.add_argument("--out-dir", default="./ts_compare_out")
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    runs = [load_run(o, f) for o, f, _ in a.run]
    labels = [lab for _, _, lab in a.run]
    plot_dispatch(runs, labels, a.year, a.season, os.path.join(a.out_dir, "dispatch_intraday.png"))
    plot_demand(runs, labels, a.year, a.season, os.path.join(a.out_dir, "demand_profile.png"))
    df = summary_table(runs, labels, os.path.join(a.out_dir, "summary.csv"))
    print(df.to_string(index=False))
    print(f"\nfigures + summary in {a.out_dir}")


if __name__ == "__main__":
    main()
