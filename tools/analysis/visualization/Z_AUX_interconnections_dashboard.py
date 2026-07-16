"""Build an interactive HTML dashboard comparing cross-border power
interconnection flows across OSTRAM scenarios.

Inputs:
  t1_confection/OSTRAM_Combined_Inputs_Outputs.csv

Output:
  t1_confection/interconnections_dashboard.html  (self-contained, Plotly inline)

Interconnection technologies are TRN[ORIGIN][DESTINATION] where ORIGIN and
DESTINATION are 5-character node codes (e.g. BGDXX, INDEA, NPLXX), plus
TRNNLI[DEST] for non-local imports (treated as an external "NLI" node).
"""

from __future__ import annotations

import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

HERE = Path(__file__).resolve().parents[3] / "t1_confection"
CSV_PATH = HERE / "OSTRAM_Combined_Inputs_Outputs.csv"
OUT_DIR = HERE / "Figures"
OUT_PATH = OUT_DIR / "interconnections_dashboard.html"

BASELINE = "BAU"
SCENARIO_ORDER = ["BAU", "A_Calibrated_BAU", "B_Optimised_VRE", "C_Target_VRE"]
SCENARIO_COLORS = {
    "BAU": "#7f8c8d",
    "A_Calibrated_BAU": "#2980b9",
    "B_Optimised_VRE": "#27ae60",
    "C_Target_VRE": "#e67e22",
}

NODE_NAMES = {
    "BGDXX": "Bangladesh",
    "BTNXX": "Bhutan",
    "INDEA": "India East",
    "INDNE": "India NorthEast",
    "INDNO": "India North",
    "INDSO": "India South",
    "INDWE": "India West",
    "LKAXX": "Sri Lanka",
    "MDVXX": "Maldives",
    "NPLXX": "Nepal",
}

# Highlight years used for snapshot views (Sankey, season detail, delta bars).
SNAPSHOT_YEARS = [2030, 2040, 2050]
DEFAULT_SNAPSHOT_YEAR = 2030

# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

USECOLS = [
    "Scenario", "YEAR", "TECHNOLOGY", "TIMESLICE",
    "ProductionByTechnologyAnnual", "TotalCapacityAnnual",
    "RateOfActivity", "YearSplit",
]


def parse_line(tech: str) -> tuple[str, str] | None:
    """Return (origin, destination) node codes or None if not a line tech.

    TRNNLI* (Nuevas Lineas Indicativas / unplanned candidate lines) are
    excluded — they represent a modelling artifact, not real corridors.
    """
    if not tech.startswith("TRN"):
        return None
    rest = tech[3:]
    if rest.startswith("NLI"):
        return None
    if len(rest) == 10:
        return rest[:5], rest[5:]
    return None


def node_label(code: str) -> str:
    return NODE_NAMES.get(code, code)


def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    print(f"Loading {CSV_PATH.name} ...")
    df = pd.read_csv(CSV_PATH, usecols=USECOLS, low_memory=False)
    print(f"  total rows: {len(df):,}")

    # Year split: independent of TECHNOLOGY, keyed by Scenario, YEAR, TIMESLICE
    ys = (
        df.dropna(subset=["YearSplit", "TIMESLICE"])
        .groupby(["Scenario", "YEAR", "TIMESLICE"], dropna=False)["YearSplit"]
        .first()
        .reset_index()
    )
    print(f"  YearSplit rows: {len(ys):,}")

    # Detect interconnection lines
    techs = df["TECHNOLOGY"].dropna().unique()
    line_map = {t: parse_line(t) for t in techs}
    lines = {t: pair for t, pair in line_map.items() if pair is not None}
    print(f"  interconnection lines: {len(lines)}")

    df_l = df[df["TECHNOLOGY"].isin(lines)].copy()
    df_l["ORIGIN"] = df_l["TECHNOLOGY"].map(lambda t: lines[t][0])
    df_l["DEST"] = df_l["TECHNOLOGY"].map(lambda t: lines[t][1])
    df_l["PAIR"] = df_l["ORIGIN"] + " -> " + df_l["DEST"]

    return df_l, ys, pd.DataFrame(
        [(t, o, d) for t, (o, d) in lines.items()], columns=["TECHNOLOGY", "ORIGIN", "DEST"]
    )


# ---------------------------------------------------------------------------
# Metric extraction
# ---------------------------------------------------------------------------

def build_annual(df_l: pd.DataFrame) -> pd.DataFrame:
    """One row per (Scenario, YEAR, line) with annual energy and capacity."""
    prod = (
        df_l.dropna(subset=["ProductionByTechnologyAnnual"])
        .groupby(["Scenario", "YEAR", "TECHNOLOGY", "ORIGIN", "DEST", "PAIR"], dropna=False)
        ["ProductionByTechnologyAnnual"].sum()
        .reset_index()
        .rename(columns={"ProductionByTechnologyAnnual": "energy_pj"})
    )
    cap = (
        df_l.dropna(subset=["TotalCapacityAnnual"])
        .groupby(["Scenario", "YEAR", "TECHNOLOGY", "ORIGIN", "DEST", "PAIR"], dropna=False)
        ["TotalCapacityAnnual"].sum()
        .reset_index()
        .rename(columns={"TotalCapacityAnnual": "capacity_gw"})
    )
    out = prod.merge(
        cap, on=["Scenario", "YEAR", "TECHNOLOGY", "ORIGIN", "DEST", "PAIR"], how="outer"
    )
    return out


def build_seasonal(df_l: pd.DataFrame, ys: pd.DataFrame) -> pd.DataFrame:
    """One row per (Scenario, YEAR, line, SEASON) with seasonal energy.

    seasonal_energy = sum_over_timeslices( RateOfActivity * YearSplit )
    where SEASON = first two chars of TIMESLICE (S1..S4).
    """
    roa = df_l.dropna(subset=["RateOfActivity", "TIMESLICE"]).copy()
    if roa.empty:
        return pd.DataFrame(
            columns=["Scenario", "YEAR", "TECHNOLOGY", "ORIGIN", "DEST", "PAIR",
                     "SEASON", "energy_pj"]
        )
    roa = roa.drop(columns=["YearSplit"], errors="ignore")
    roa = roa.merge(ys, on=["Scenario", "YEAR", "TIMESLICE"], how="left")
    roa["YearSplit"] = roa["YearSplit"].fillna(0.0)
    roa["energy_pj"] = roa["RateOfActivity"] * roa["YearSplit"]
    roa["SEASON"] = roa["TIMESLICE"].str.extract(r"^(S\d)", expand=False)
    agg = (
        roa.dropna(subset=["SEASON"])
        .groupby(
            ["Scenario", "YEAR", "TECHNOLOGY", "ORIGIN", "DEST", "PAIR", "SEASON"],
            dropna=False,
        )["energy_pj"].sum()
        .reset_index()
    )
    return agg


# ---------------------------------------------------------------------------
# Units
# ---------------------------------------------------------------------------

# Internal storage is PJ (the model's native unit). Display can be toggled to
# TWh via the UI button. 1 PJ = 1/3.6 TWh.
UNIT_FACTOR = {"PJ": 1.0, "TWh": 1.0 / 3.6}


def scale_col(df: pd.DataFrame, col: str, unit: str) -> pd.DataFrame:
    out = df.copy()
    out[col] = out[col] * UNIT_FACTOR[unit]
    return out


# ---------------------------------------------------------------------------
# Figure builders
# ---------------------------------------------------------------------------

def fig_kpi_cards(annual: pd.DataFrame, unit: str = "PJ") -> go.Figure:
    """Bar chart with total energy exchanged across all years per scenario."""
    annual = scale_col(annual, "energy_pj", unit)
    totals = (
        annual.groupby("Scenario")["energy_pj"].sum().reindex(SCENARIO_ORDER).dropna()
    )
    fig = go.Figure(
        go.Bar(
            x=totals.index.tolist(),
            y=totals.values,
            text=[f"{v:,.0f} {unit}" for v in totals.values],
            textposition="outside",
            marker_color=[SCENARIO_COLORS[s] for s in totals.index],
            hovertemplate=f"<b>%{{x}}</b><br>Total: %{{y:,.1f}} {unit}<extra></extra>",
        )
    )
    fig.update_layout(
        title=f"Total cross-border energy exchanged (sum 2023-2050) — {unit}",
        yaxis_title=unit,
        showlegend=False,
        height=380,
        margin=dict(l=60, r=20, t=60, b=40),
        template="plotly_white",
    )
    return fig


def fig_annual_trend(annual: pd.DataFrame, unit: str = "PJ") -> go.Figure:
    """Total interconnection energy per scenario over time."""
    annual = scale_col(annual, "energy_pj", unit)
    grp = (
        annual.groupby(["Scenario", "YEAR"])["energy_pj"].sum().reset_index()
    )
    fig = go.Figure()
    for sc in SCENARIO_ORDER:
        sub = grp[grp["Scenario"] == sc].sort_values("YEAR")
        if sub.empty:
            continue
        fig.add_trace(go.Scatter(
            x=sub["YEAR"], y=sub["energy_pj"],
            mode="lines+markers", name=sc,
            line=dict(color=SCENARIO_COLORS[sc], width=3),
            hovertemplate=f"<b>{sc}</b><br>%{{x}}: %{{y:,.1f}} {unit}<extra></extra>",
        ))
    fig.update_layout(
        title=f"Annual cross-border flows by scenario — {unit}",
        xaxis_title="Year", yaxis_title=unit,
        height=440, template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=60, r=20, t=70, b=50),
    )
    return fig


def fig_heatmap_per_scenario(annual: pd.DataFrame, unit: str = "PJ") -> go.Figure:
    """Heatmap of energy per line per year, one subplot per scenario."""
    annual = scale_col(annual, "energy_pj", unit)
    scenarios = [s for s in SCENARIO_ORDER if s in annual["Scenario"].unique()]
    fig = make_subplots(
        rows=1, cols=len(scenarios),
        subplot_titles=scenarios,
        shared_yaxes=True,
        horizontal_spacing=0.02,
    )
    pairs = sorted(annual["PAIR"].unique())
    years = sorted(annual["YEAR"].dropna().unique())
    vmax = annual["energy_pj"].max()
    for i, sc in enumerate(scenarios, start=1):
        sub = annual[annual["Scenario"] == sc]
        pivot = (
            sub.pivot_table(
                index="PAIR", columns="YEAR", values="energy_pj", aggfunc="sum"
            )
            .reindex(index=pairs, columns=years)
        )
        fig.add_trace(
            go.Heatmap(
                z=pivot.values,
                x=[int(y) for y in pivot.columns],
                y=pivot.index,
                colorscale="Viridis",
                zmin=0, zmax=vmax,
                colorbar=dict(title=unit) if i == len(scenarios) else None,
                showscale=(i == len(scenarios)),
                hovertemplate=f"<b>%{{y}}</b><br>%{{x}}: %{{z:,.2f}} {unit}<extra></extra>",
            ),
            row=1, col=i,
        )
    fig.update_layout(
        title=f"Annual flow per line by scenario ({unit})",
        height=max(420, 22 * len(pairs)),
        template="plotly_white",
        margin=dict(l=140, r=40, t=80, b=60),
    )
    return fig


def fig_sankey(annual: pd.DataFrame, year: int, unit: str = "PJ") -> go.Figure:
    """Sankey for a single year, with a dropdown to switch scenarios."""
    annual = scale_col(annual, "energy_pj", unit)
    nodes = sorted(set(annual["ORIGIN"]).union(annual["DEST"]))
    idx = {n: i for i, n in enumerate(nodes)}
    labels = [node_label(n) for n in nodes]
    # Soft palette per node (consistent across scenarios)
    palette = (
        ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
         "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf", "#aec7e8"]
        * 3
    )
    node_colors = [palette[i % len(palette)] for i in range(len(nodes))]

    traces = []
    buttons = []
    scenarios = [s for s in SCENARIO_ORDER if s in annual["Scenario"].unique()]
    for sc in scenarios:
        sub = annual[(annual["Scenario"] == sc) & (annual["YEAR"] == year)]
        sub = sub[sub["energy_pj"].fillna(0) > 0]
        traces.append(go.Sankey(
            arrangement="snap",
            node=dict(
                label=labels, pad=14, thickness=18,
                color=node_colors,
                line=dict(color="rgba(0,0,0,0.3)", width=0.5),
            ),
            link=dict(
                source=[idx[o] for o in sub["ORIGIN"]],
                target=[idx[d] for d in sub["DEST"]],
                value=sub["energy_pj"].tolist(),
                customdata=np.stack(
                    [sub["ORIGIN"].map(node_label), sub["DEST"].map(node_label)], axis=-1
                ) if not sub.empty else None,
                hovertemplate=f"%{{customdata[0]}} → %{{customdata[1]}}<br>%{{value:,.2f}} {unit}<extra></extra>",
                color="rgba(120,120,120,0.35)",
            ),
            visible=(sc == scenarios[0]),
        ))
    for i, sc in enumerate(scenarios):
        buttons.append(dict(
            label=sc, method="update",
            args=[{"visible": [j == i for j in range(len(scenarios))]},
                  {"title": f"Sankey of inter-node flows — {sc} — {year} ({unit})"}],
        ))
    fig = go.Figure(data=traces)
    fig.update_layout(
        title=f"Sankey of inter-node flows — {scenarios[0]} — {year} ({unit})",
        height=560, template="plotly_white",
        margin=dict(l=20, r=20, t=80, b=20),
        updatemenus=[dict(
            type="dropdown", x=0.0, y=1.12, xanchor="left", yanchor="top",
            buttons=buttons,
        )],
    )
    return fig


def fig_delta_vs_baseline(annual: pd.DataFrame, year: int, unit: str = "PJ") -> go.Figure:
    """Delta per line vs BAU, grouped by scenario."""
    annual = scale_col(annual, "energy_pj", unit)
    sub = annual[annual["YEAR"] == year]
    if sub.empty or BASELINE not in sub["Scenario"].unique():
        return go.Figure().update_layout(
            title=f"Delta vs {BASELINE} — no data for {year}"
        )
    base = sub[sub["Scenario"] == BASELINE].set_index("PAIR")["energy_pj"]
    pairs = sorted(sub["PAIR"].unique())
    fig = go.Figure()
    for sc in [s for s in SCENARIO_ORDER if s != BASELINE]:
        sc_data = sub[sub["Scenario"] == sc].set_index("PAIR")["energy_pj"]
        delta = (sc_data.reindex(pairs).fillna(0) - base.reindex(pairs).fillna(0))
        fig.add_trace(go.Bar(
            x=pairs, y=delta.values, name=sc,
            marker_color=SCENARIO_COLORS[sc],
            hovertemplate=f"<b>{sc}</b><br>%{{x}}<br>Δ: %{{y:+.2f}} {unit}<extra></extra>",
        ))
    fig.update_layout(
        title=f"Δ vs {BASELINE} in {year} — {unit} per line",
        barmode="group", height=460, template="plotly_white",
        yaxis_title=unit, xaxis_tickangle=-45,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=60, r=20, t=70, b=150),
    )
    return fig


def fig_net_flow_per_node(annual: pd.DataFrame, year: int, unit: str = "PJ") -> go.Figure:
    """Net flow (exports - imports) per node, grouped by scenario."""
    annual = scale_col(annual, "energy_pj", unit)
    sub = annual[annual["YEAR"] == year].copy()
    out = sub.groupby(["Scenario", "ORIGIN"])["energy_pj"].sum().rename("out").reset_index()
    out = out.rename(columns={"ORIGIN": "NODE"})
    inn = sub.groupby(["Scenario", "DEST"])["energy_pj"].sum().rename("in").reset_index()
    inn = inn.rename(columns={"DEST": "NODE"})
    net = out.merge(inn, on=["Scenario", "NODE"], how="outer").fillna(0.0)
    net["net"] = net["out"] - net["in"]
    nodes = [n for n in net["NODE"].unique() if n != "NLI"]
    nodes = sorted(nodes, key=lambda n: node_label(n))
    fig = go.Figure()
    for sc in [s for s in SCENARIO_ORDER if s in net["Scenario"].unique()]:
        sub_sc = net[net["Scenario"] == sc].set_index("NODE").reindex(nodes)
        fig.add_trace(go.Bar(
            x=[node_label(n) for n in nodes],
            y=sub_sc["net"].values,
            name=sc, marker_color=SCENARIO_COLORS[sc],
            hovertemplate=f"<b>{sc}</b><br>%{{x}}<br>Net: %{{y:+.2f}} {unit}<extra></extra>",
        ))
    fig.add_hline(y=0, line=dict(color="black", width=1))
    fig.update_layout(
        title=f"Net flow per node in {year} (positive = net exporter) — {unit}",
        barmode="group", height=460, template="plotly_white",
        yaxis_title=f"{unit} (exports - imports)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=60, r=20, t=70, b=80),
    )
    return fig


def fig_seasonal(seasonal: pd.DataFrame, year: int, unit: str = "PJ") -> go.Figure:
    """Seasonal breakdown: total energy per SEASON per scenario for one year."""
    seasonal = scale_col(seasonal, "energy_pj", unit)
    sub = seasonal[seasonal["YEAR"] == year]
    if sub.empty:
        return go.Figure().update_layout(title=f"Seasonal breakdown — no data for {year}")
    grp = sub.groupby(["Scenario", "SEASON"])["energy_pj"].sum().reset_index()
    seasons = sorted(grp["SEASON"].unique())
    fig = go.Figure()
    for sc in [s for s in SCENARIO_ORDER if s in grp["Scenario"].unique()]:
        sub_sc = grp[grp["Scenario"] == sc].set_index("SEASON").reindex(seasons)
        fig.add_trace(go.Bar(
            x=seasons, y=sub_sc["energy_pj"].values,
            name=sc, marker_color=SCENARIO_COLORS[sc],
            hovertemplate=f"<b>{sc}</b><br>%{{x}}<br>%{{y:,.2f}} {unit}<extra></extra>",
        ))
    fig.update_layout(
        title=f"Total cross-border flow per season — {year} ({unit})",
        barmode="group", height=420, template="plotly_white",
        yaxis_title=unit, xaxis_title="Season",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=60, r=20, t=70, b=60),
    )
    return fig


def fig_seasonal_heatmap(seasonal: pd.DataFrame, unit: str = "PJ") -> go.Figure:
    """Heatmap (line × season) for a chosen year, one subplot per scenario."""
    seasonal = scale_col(seasonal, "energy_pj", unit)
    year = DEFAULT_SNAPSHOT_YEAR
    sub = seasonal[seasonal["YEAR"] == year]
    scenarios = [s for s in SCENARIO_ORDER if s in sub["Scenario"].unique()]
    if not scenarios:
        return go.Figure().update_layout(title=f"Seasonal heatmap — no data for {year}")
    fig = make_subplots(
        rows=1, cols=len(scenarios), subplot_titles=scenarios,
        shared_yaxes=True, horizontal_spacing=0.03,
    )
    pairs = sorted(sub["PAIR"].unique())
    seasons = sorted(sub["SEASON"].unique())
    vmax = sub["energy_pj"].max()
    for i, sc in enumerate(scenarios, start=1):
        piv = (
            sub[sub["Scenario"] == sc]
            .pivot_table(index="PAIR", columns="SEASON", values="energy_pj", aggfunc="sum")
            .reindex(index=pairs, columns=seasons)
        )
        fig.add_trace(go.Heatmap(
            z=piv.values, x=piv.columns, y=piv.index,
            colorscale="Viridis", zmin=0, zmax=vmax,
            showscale=(i == len(scenarios)),
            colorbar=dict(title=unit) if i == len(scenarios) else None,
            hovertemplate=f"<b>%{{y}}</b><br>%{{x}}: %{{z:,.3f}} {unit}<extra></extra>",
        ), row=1, col=i)
    fig.update_layout(
        title=f"Seasonal flow per line — {year} ({unit})",
        height=max(420, 22 * len(pairs)),
        template="plotly_white",
        margin=dict(l=140, r=40, t=80, b=60),
    )
    return fig


def fig_capacity_trend(annual: pd.DataFrame) -> go.Figure:
    """Total interconnection capacity per scenario over time."""
    grp = annual.groupby(["Scenario", "YEAR"])["capacity_gw"].sum().reset_index()
    fig = go.Figure()
    for sc in SCENARIO_ORDER:
        sub = grp[grp["Scenario"] == sc].sort_values("YEAR")
        if sub.empty:
            continue
        fig.add_trace(go.Scatter(
            x=sub["YEAR"], y=sub["capacity_gw"],
            mode="lines+markers", name=sc,
            line=dict(color=SCENARIO_COLORS[sc], width=3, dash="solid"),
            hovertemplate=f"<b>{sc}</b><br>%{{x}}: %{{y:,.1f}} GW<extra></extra>",
        ))
    fig.update_layout(
        title="Total interconnection capacity over time",
        xaxis_title="Year", yaxis_title="GW",
        height=400, template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=60, r=20, t=70, b=50),
    )
    return fig


# ---------------------------------------------------------------------------
# HTML assembly
# ---------------------------------------------------------------------------

CSS = """
<style>
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     background:#f5f6fa;color:#2c3e50;margin:0;padding:0;}
.header{background:linear-gradient(135deg,#1e3c72,#2a5298);color:white;
        padding:36px 48px;box-shadow:0 2px 12px rgba(0,0,0,0.1);
        display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:24px;}
.header-text h1{margin:0 0 6px 0;font-size:28px;font-weight:600;}
.header-text p{margin:0;opacity:0.85;font-size:14px;}
.unit-toggle{background:rgba(255,255,255,0.12);border-radius:8px;padding:6px;display:flex;gap:4px;
             align-items:center;align-self:center;}
.unit-toggle .lbl{font-size:11px;color:rgba(255,255,255,0.7);text-transform:uppercase;
                  letter-spacing:0.6px;padding:0 8px;}
.unit-toggle button{background:transparent;color:white;border:none;padding:6px 16px;
                    border-radius:6px;font-size:13px;font-weight:600;cursor:pointer;
                    transition:background 0.15s;}
.unit-toggle button:hover{background:rgba(255,255,255,0.18);}
.unit-toggle button.active{background:white;color:#1e3c72;}
.container{max-width:1500px;margin:0 auto;padding:24px 32px;}
.section{background:white;border-radius:10px;padding:18px 24px;margin:20px 0;
         box-shadow:0 1px 4px rgba(0,0,0,0.06);}
.section h2{margin:0 0 4px 0;font-size:18px;color:#1e3c72;font-weight:600;}
.section .lede{color:#7f8c8d;font-size:13px;margin:0 0 14px 0;}
.kpi-row{display:flex;gap:18px;flex-wrap:wrap;margin:16px 0;}
.kpi-card{flex:1;min-width:200px;background:white;border-left:5px solid #2a5298;
         border-radius:6px;padding:14px 18px;box-shadow:0 1px 3px rgba(0,0,0,0.07);}
.kpi-card.bau{border-color:#7f8c8d;}
.kpi-card.cal{border-color:#2980b9;}
.kpi-card.opt{border-color:#27ae60;}
.kpi-card.tgt{border-color:#e67e22;}
.kpi-label{font-size:12px;color:#7f8c8d;text-transform:uppercase;letter-spacing:0.5px;}
.kpi-value{font-size:22px;font-weight:700;color:#2c3e50;margin:4px 0 2px;}
.kpi-delta{font-size:12px;color:#95a5a6;}
.kpi-delta.pos{color:#27ae60;}
.kpi-delta.neg{color:#c0392b;}
.unit-block[data-unit="TWh"]{display:none;}
body.unit-twh .unit-block[data-unit="PJ"]{display:none;}
body.unit-twh .unit-block[data-unit="TWh"]{display:block;}
footer{text-align:center;padding:24px;color:#95a5a6;font-size:12px;}
</style>
"""

TOGGLE_JS = """
<script>
function setUnit(u){
  if(u==='TWh'){document.body.classList.add('unit-twh');}
  else{document.body.classList.remove('unit-twh');}
  document.querySelectorAll('.unit-toggle button').forEach(function(b){
    b.classList.toggle('active', b.dataset.unit===u);
  });
  // resize plotly figures that became visible
  setTimeout(function(){
    document.querySelectorAll('.unit-block:not([style*="display: none"]) .plotly-graph-div').forEach(function(el){
      if(window.Plotly){ try{ Plotly.Plots.resize(el); }catch(e){} }
    });
  }, 50);
}
document.addEventListener('DOMContentLoaded', function(){ setUnit('PJ'); });
</script>
"""


def render_kpi_cards(annual: pd.DataFrame, unit: str = "PJ") -> str:
    annual = scale_col(annual, "energy_pj", unit)
    totals = annual.groupby("Scenario")["energy_pj"].sum()
    if BASELINE not in totals.index:
        return ""
    base = totals[BASELINE]
    cards = []
    css_class_map = {
        "BAU": "bau", "A_Calibrated_BAU": "cal",
        "B_Optimised_VRE": "opt", "C_Target_VRE": "tgt",
    }
    for sc in SCENARIO_ORDER:
        if sc not in totals.index:
            continue
        v = totals[sc]
        delta = v - base
        delta_pct = (delta / base * 100) if base else 0
        sign = "+" if delta >= 0 else ""
        klass = "pos" if delta > 0 else ("neg" if delta < 0 else "")
        delta_html = (
            "Baseline" if sc == BASELINE
            else f'<span class="kpi-delta {klass}">{sign}{delta:,.0f} {unit} ({sign}{delta_pct:,.1f}%) vs {BASELINE}</span>'
        )
        cards.append(
            f'<div class="kpi-card {css_class_map.get(sc,"")}">'
            f'<div class="kpi-label">{sc}</div>'
            f'<div class="kpi-value">{v:,.0f} {unit}</div>'
            f'{delta_html}'
            f'</div>'
        )
    return f'<div class="kpi-row">{"".join(cards)}</div>'


def fig_html(fig: go.Figure, include_js: bool = False) -> str:
    return fig.to_html(
        full_html=False,
        include_plotlyjs="inline" if include_js else False,
        config={"displaylogo": False, "responsive": True},
    )


def build_html(annual: pd.DataFrame, seasonal: pd.DataFrame) -> str:
    years_available = sorted([int(y) for y in annual["YEAR"].dropna().unique()])
    snap_year = DEFAULT_SNAPSHOT_YEAR if DEFAULT_SNAPSHOT_YEAR in years_available else years_available[len(years_available) // 2]

    # (title, lede, figure_builder_fn, is_energy_fig, include_js_once)
    fig_specs = [
        ("Annual energy traded by scenario",
         "Total cross-border electricity (sum of all line flows) traded each year.",
         lambda u: fig_annual_trend(annual, unit=u), True, True),
        ("Total exchanged 2023-2050",
         "Cumulative energy exchanged across all interconnections.",
         lambda u: fig_kpi_cards(annual, unit=u), True, False),
        ("Heatmap — flow per line, per year",
         "Each row is a directional line, each column a year. Brighter = more energy. Compare panels side-by-side to spot which lines change.",
         lambda u: fig_heatmap_per_scenario(annual, unit=u), True, False),
        (f"Sankey snapshot — {snap_year}",
         "Switch scenarios using the dropdown above the diagram. Link width is proportional to energy exchanged that year.",
         lambda u: fig_sankey(annual, snap_year, unit=u), True, False),
        (f"Δ vs {BASELINE} in {snap_year}",
         "Per-line change relative to BAU. Positive = scenario uses that link more; negative = less.",
         lambda u: fig_delta_vs_baseline(annual, snap_year, unit=u), True, False),
        (f"Net flow per node — {snap_year}",
         "Exports minus imports per node. Positive bars = net exporter; negative = net importer.",
         lambda u: fig_net_flow_per_node(annual, snap_year, unit=u), True, False),
        ("Interconnection capacity over time",
         "Sum of installed transfer capacity (GW). Useful to see whether the model is building new corridors. Unit is GW regardless of the toggle.",
         lambda u: fig_capacity_trend(annual), False, False),
        (f"Seasonal totals — {snap_year}",
         "Total flow per season (S1-S4) for the snapshot year. Shows seasonality of interregional trade.",
         lambda u: fig_seasonal(seasonal, snap_year, unit=u), True, False),
        (f"Seasonal heatmap — {snap_year}",
         "Line-by-season heatmap. Useful to see whether certain lines are season-specific (e.g., hydro-rich seasons drive Bhutan/Nepal exports).",
         lambda u: fig_seasonal_heatmap(seasonal, unit=u), True, False),
    ]

    sections = []
    js_emitted = False
    for title, lede, builder, is_energy, include_js in fig_specs:
        if is_energy:
            html_pj = fig_html(builder("PJ"), include_js=(include_js and not js_emitted))
            if include_js and not js_emitted:
                js_emitted = True
            html_twh = fig_html(builder("TWh"), include_js=False)
            fig_block = (
                f'<div class="unit-block" data-unit="PJ">{html_pj}</div>'
                f'<div class="unit-block" data-unit="TWh">{html_twh}</div>'
            )
        else:
            html_one = fig_html(builder("PJ"), include_js=(include_js and not js_emitted))
            if include_js and not js_emitted:
                js_emitted = True
            fig_block = html_one
        sections.append(
            f'<div class="section"><h2>{title}</h2><p class="lede">{lede}</p>{fig_block}</div>'
        )

    kpi_pj = render_kpi_cards(annual, unit="PJ")
    kpi_twh = render_kpi_cards(annual, unit="TWh")
    kpi_block = (
        f'<div class="unit-block" data-unit="PJ">{kpi_pj}</div>'
        f'<div class="unit-block" data-unit="TWh">{kpi_twh}</div>'
    )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>OSTRAM — Cross-border interconnections dashboard</title>
{CSS}
{TOGGLE_JS}
</head>
<body>
<div class="header">
  <div class="header-text">
    <h1>OSTRAM — Cross-border power interconnections</h1>
    <p>Comparing flows, capacity and seasonality across {len(SCENARIO_ORDER)} scenarios &middot; {min(years_available)}-{max(years_available)} &middot; {annual['PAIR'].nunique()} directional lines &middot; baseline: {BASELINE}</p>
  </div>
  <div class="unit-toggle">
    <span class="lbl">Units</span>
    <button data-unit="PJ" class="active" onclick="setUnit('PJ')">PJ</button>
    <button data-unit="TWh" onclick="setUnit('TWh')">TWh</button>
  </div>
</div>
<div class="container">
  <div class="section">
    <h2>Headline numbers</h2>
    <p class="lede">Cumulative energy exchanged across all interconnections, 2023-2050, with delta vs the {BASELINE} baseline. Toggle units in the top-right.</p>
    {kpi_block}
  </div>
  {''.join(sections)}
</div>
<footer>Generated from OSTRAM_Combined_Inputs_Outputs.csv &middot; Native unit is PJ (1 PJ = 1/3.6 TWh).</footer>
</body>
</html>
"""
    return html


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    df_l, ys, lines = load_data()
    print(f"Lines parsed: {len(lines)} (sample {lines.head().to_dict('records')})")

    annual = build_annual(df_l)
    print(f"annual rows: {len(annual):,}")
    seasonal = build_seasonal(df_l, ys)
    print(f"seasonal rows: {len(seasonal):,}")

    html = build_html(annual, seasonal)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(html, encoding="utf-8")
    size_mb = OUT_PATH.stat().st_size / (1024 * 1024)
    print(f"Wrote {OUT_PATH} ({size_mb:.2f} MB)")


if __name__ == "__main__":
    main()
