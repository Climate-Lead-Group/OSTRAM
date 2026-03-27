# -*- coding: utf-8 -*-
"""
Transmission Maps Generator for OSTRAM

Generates a standalone interactive HTML file with three map views:
  - Transmission Capacity (GW/TW)
  - Transmission Flow (PJ/TWh/GWh)
  - Load-Capacity Ratio (ratio/%)

Features:
  - Year selector (all years discovered from data)
  - Scenario selector (all scenarios discovered from data)
  - Unit toggle (GW↔TW for capacity; PJ↔TWh↔GWh for flow; ratio↔% for load-capacity)
  - Interactive Plotly Scattergeo maps with hover info
  - Color-coded lines for load-capacity ratio (green→red by utilization)
  - PNG download button
  - Fully standalone (no server required)

No countries, technologies, years, or scenarios are hardcoded.
Everything is discovered dynamically from the input data.

Author: Climate Lead Group, Andrey Salazar-Vargas
"""

import pandas as pd
import json
import os
import re
import glob
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
INTERCONNECTION_PATTERN = re.compile(r'^TRN[A-Z]{5}[A-Z]{5}$')

CAPACITY_COL = 'TotalCapacityAnnual'
FLOW_COL = 'ProductionByTechnologyAnnual'
PRODUCTION_BY_TIMESLICE_COL = 'ProductionByTechnology'
CAPACITY_TO_ACTIVITY_COL = 'CapacityToActivityUnit'
YEAR_SPLIT_COL = 'YearSplit'


def find_combined_csv(script_dir):
    """Find the combined inputs/outputs CSV by glob pattern."""
    pattern = str(script_dir / '*_Combined_Inputs_Outputs.csv')
    matches = glob.glob(pattern)
    if matches:
        return Path(matches[0])
    return None
def load_data(csv_path):
    """Load the combined CSV and filter to interconnection technologies.

    Also extracts the global YearSplit lookup (TIMESLICE -> fraction)
    from the full CSV before filtering, since YearSplit is not
    technology-specific.

    Returns (filtered_df, year_split_dict).
    """
    print(f"  Loading data from {csv_path} ...")
    df = pd.read_csv(csv_path, low_memory=False)

    # Extract YearSplit before filtering (it is a global parameter)
    ys_rows = df[df[YEAR_SPLIT_COL].notna()][['TIMESLICE', YEAR_SPLIT_COL]].drop_duplicates()
    year_split = dict(zip(ys_rows['TIMESLICE'], ys_rows[YEAR_SPLIT_COL]))

    mask = df['TECHNOLOGY'].astype(str).apply(
        lambda t: bool(INTERCONNECTION_PATTERN.match(t))
    )
    df = df[mask].copy()
    print(f"  Found {len(df):,} rows with interconnection technologies")
    print(f"  Unique interconnections: {df['TECHNOLOGY'].nunique()}")
    return df, year_split


def load_centerpoints(csv_path):
    """Load centerpoints CSV into a dict: region -> {lat, long}."""
    print(f"  Loading centerpoints from {csv_path} ...")
    cp = pd.read_csv(csv_path)
    return {
        row['region']: {'lat': row['lat'], 'long': row['long']}
        for _, row in cp.iterrows()
    }


def extract_from_to(tech):
    """Extract FROM and TO region codes from an interconnection technology name.

    Pattern: TRN{FROM_5}{TO_5}  e.g. TRNBGDXXINDEA
    FROM = tech[3:8], TO = tech[8:13]
    """
    return tech[3:8], tech[8:13]


def classify_flow_direction(fuel, tech_from, tech_to):
    """Classify flow direction from FUEL code suffix '04'.

    For TRN{A}{B}, output fuel ELC{B}04 means flow A→B,
    and ELC{A}04 means flow B→A.
    """
    if not isinstance(fuel, str) or not fuel.endswith('04'):
        return None
    fuel_region = fuel[3:-2]  # 'ELCINDEA04' -> 'INDEA'
    if fuel_region == tech_to:
        return 'a_to_b'
    elif fuel_region == tech_from:
        return 'b_to_a'
    return None


def prepare_json_data(df, centerpoints, year_split):
    """Prepare all data as nested dicts for JSON embedding in HTML.

    Returns three dicts (capacity_data, flow_data, ratio_data) structured as:
      { scenario: { year: [ {from, to, from_lat, from_lon, to_lat, to_lon, value}, ... ] } }
    """
    df = df.copy()
    df['YEAR'] = pd.to_numeric(df['YEAR'], errors='coerce')
    df = df.dropna(subset=['YEAR'])
    df['YEAR'] = df['YEAR'].astype(int)

    df['FROM'] = df['TECHNOLOGY'].apply(lambda t: extract_from_to(t)[0])
    df['TO'] = df['TECHNOLOGY'].apply(lambda t: extract_from_to(t)[1])

    # Filter to nodes that exist in centerpoints
    known_regions = set(centerpoints.keys())
    df = df[df['FROM'].isin(known_regions) & df['TO'].isin(known_regions)]

    if df.empty:
        print("  WARNING: No interconnections match known centerpoints!")
        return {}, {}, {}

    # Add coordinates
    df['from_lat'] = df['FROM'].map(lambda r: centerpoints[r]['lat'])
    df['from_lon'] = df['FROM'].map(lambda r: centerpoints[r]['long'])
    df['to_lat'] = df['TO'].map(lambda r: centerpoints[r]['lat'])
    df['to_lon'] = df['TO'].map(lambda r: centerpoints[r]['long'])

    # Build CapacityToActivityUnit lookup: TECHNOLOGY -> value
    cta_rows = df[df[CAPACITY_TO_ACTIVITY_COL].notna()][['TECHNOLOGY', CAPACITY_TO_ACTIVITY_COL]].drop_duplicates()
    cta_map = dict(zip(cta_rows['TECHNOLOGY'], cta_rows[CAPACITY_TO_ACTIVITY_COL]))

    scenarios = sorted(df['Scenario'].dropna().unique().tolist())
    years = sorted(df['YEAR'].unique().tolist())

    capacity_data = {}
    flow_data = {}
    ratio_data = {}

    for scenario in scenarios:
        capacity_data[scenario] = {}
        flow_data[scenario] = {}
        ratio_data[scenario] = {}
        sdf = df[df['Scenario'] == scenario]

        for year in years:
            ydf = sdf[sdf['YEAR'] == year]

            # --- Capacity ---
            cap = ydf[ydf[CAPACITY_COL].notna() & (ydf[CAPACITY_COL] != 0)]
            cap_agg = (
                cap.groupby(['FROM', 'TO', 'from_lat', 'from_lon', 'to_lat', 'to_lon'])
                [CAPACITY_COL].sum().reset_index()
            )
            capacity_data[scenario][str(year)] = [
                {
                    'from': row['FROM'], 'to': row['TO'],
                    'from_lat': row['from_lat'], 'from_lon': row['from_lon'],
                    'to_lat': row['to_lat'], 'to_lon': row['to_lon'],
                    'value': round(float(row[CAPACITY_COL]), 4)
                }
                for _, row in cap_agg.iterrows()
            ]

            # --- Flow (directional) ---
            flw = ydf[ydf[FLOW_COL].notna() & (ydf[FLOW_COL] != 0)].copy()
            flw['DIRECTION'] = flw.apply(
                lambda r: classify_flow_direction(
                    str(r.get('FUEL', '')), r['FROM'], r['TO']
                ), axis=1
            )
            flw = flw[flw['DIRECTION'].notna()]

            if not flw.empty:
                flw_agg = (
                    flw.groupby(['FROM', 'TO', 'from_lat', 'from_lon', 'to_lat', 'to_lon', 'DIRECTION'])
                    [FLOW_COL].sum().reset_index()
                )
                pivot = flw_agg.pivot_table(
                    index=['FROM', 'TO', 'from_lat', 'from_lon', 'to_lat', 'to_lon'],
                    columns='DIRECTION', values=FLOW_COL, fill_value=0
                ).reset_index()
                # Flatten MultiIndex columns if present
                if hasattr(pivot.columns, 'levels'):
                    pivot.columns = [
                        col[1] if col[1] else col[0]
                        for col in pivot.columns
                    ]
                else:
                    pivot.columns.name = None
                if 'a_to_b' not in pivot.columns:
                    pivot['a_to_b'] = 0
                if 'b_to_a' not in pivot.columns:
                    pivot['b_to_a'] = 0
                flow_data[scenario][str(year)] = [
                    {
                        'from': row['FROM'], 'to': row['TO'],
                        'from_lat': row['from_lat'],
                        'from_lon': row['from_lon'],
                        'to_lat': row['to_lat'],
                        'to_lon': row['to_lon'],
                        'value_a_to_b': round(float(row['a_to_b']), 4),
                        'value_b_to_a': round(float(row['b_to_a']), 4),
                    }
                    for _, row in pivot.iterrows()
                ]
            else:
                flow_data[scenario][str(year)] = []

            # --- Load-Capacity Ratio ---
            pbt = ydf[
                ydf[PRODUCTION_BY_TIMESLICE_COL].notna()
                & (ydf[PRODUCTION_BY_TIMESLICE_COL] != 0)
                & ydf['TIMESLICE'].notna()
            ].copy()
            cap_nonzero = ydf[ydf[CAPACITY_COL].notna() & (ydf[CAPACITY_COL] != 0)]

            if not pbt.empty and not cap_nonzero.empty and year_split:
                pbt['_ys'] = pbt['TIMESLICE'].map(year_split)
                pbt = pbt[pbt['_ys'].notna() & (pbt['_ys'] > 0)]
                pbt['_rate'] = pbt[PRODUCTION_BY_TIMESLICE_COL] / pbt['_ys']

                # Max rate per interconnection
                max_rate = (
                    pbt.groupby(['TECHNOLOGY', 'FROM', 'TO', 'from_lat', 'from_lon', 'to_lat', 'to_lon'])
                    ['_rate'].max().reset_index()
                )

                # Capacity per technology
                cap_by_tech = (
                    cap_nonzero.groupby(['TECHNOLOGY'])[CAPACITY_COL].sum().reset_index()
                )

                # Merge and compute ratio
                merged = max_rate.merge(cap_by_tech, on='TECHNOLOGY', how='inner')
                merged['_cta'] = merged['TECHNOLOGY'].map(cta_map).fillna(31.536)
                merged['_ratio'] = merged['_rate'] / (merged[CAPACITY_COL] * merged['_cta'])

                ratio_data[scenario][str(year)] = [
                    {
                        'from': row['FROM'], 'to': row['TO'],
                        'from_lat': row['from_lat'], 'from_lon': row['from_lon'],
                        'to_lat': row['to_lat'], 'to_lon': row['to_lon'],
                        'value': round(float(row['_ratio']), 4)
                    }
                    for _, row in merged.iterrows()
                ]
            else:
                ratio_data[scenario][str(year)] = []

    print(f"  Scenarios: {scenarios}")
    print(f"  Years: {years[0]}-{years[-1]} ({len(years)} years)")

    return capacity_data, flow_data, ratio_data


def build_node_list(centerpoints, capacity_data, flow_data, ratio_data):
    """Build list of all nodes involved in any interconnection."""
    nodes = set()
    for scenario_data in [capacity_data, flow_data, ratio_data]:
        for scenario in scenario_data.values():
            for year_links in scenario.values():
                for link in year_links:
                    nodes.add(link['from'])
                    nodes.add(link['to'])

    return [
        {'region': r, 'lat': centerpoints[r]['lat'], 'lon': centerpoints[r]['long']}
        for r in sorted(nodes) if r in centerpoints
    ]


def generate_html(capacity_data, flow_data, ratio_data, nodes, output_path):
    """Generate standalone HTML file with interactive transmission maps."""

    capacity_json = json.dumps(capacity_data)
    flow_json = json.dumps(flow_data)
    ratio_json = json.dumps(ratio_data)
    nodes_json = json.dumps(nodes)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OSTRAM - Transmission Maps</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; }}
  .header {{ background: linear-gradient(135deg, #1a237e, #0d47a1); color: white; padding: 20px 30px; }}
  .header h1 {{ font-size: 1.6em; font-weight: 600; }}
  .header p {{ font-size: 0.9em; opacity: 0.85; margin-top: 4px; }}
  .controls {{ display: flex; flex-wrap: wrap; gap: 16px; align-items: center;
               padding: 16px 30px; background: #fff; border-bottom: 1px solid #ddd; }}
  .control-group {{ display: flex; flex-direction: column; gap: 4px; }}
  .control-group label {{ font-size: 0.75em; font-weight: 600; text-transform: uppercase;
                          color: #666; letter-spacing: 0.5px; }}
  select, .toggle-group {{ font-size: 0.9em; padding: 6px 12px; border: 1px solid #ccc;
                           border-radius: 6px; background: #fff; }}
  .tabs {{ display: flex; gap: 0; }}
  .tab {{ padding: 8px 20px; border: 1px solid #ccc; background: #f5f5f5; cursor: pointer;
          font-size: 0.9em; font-weight: 500; transition: all 0.2s; }}
  .tab:first-child {{ border-radius: 6px 0 0 6px; }}
  .tab:last-child {{ border-radius: 0 6px 6px 0; }}
  .tab.active {{ background: #1a237e; color: white; border-color: #1a237e; }}
  .toggle-group {{ display: inline-flex; gap: 0; padding: 0; overflow: hidden; }}
  .toggle-btn {{ padding: 6px 14px; border: none; background: #f5f5f5; cursor: pointer;
                 font-size: 0.85em; font-weight: 500; transition: all 0.2s; }}
  .toggle-btn.active {{ background: #0d47a1; color: white; }}
  .toggle-btn:not(:last-child) {{ border-right: 1px solid #ddd; }}
  .download-btn {{ padding: 8px 18px; background: #2e7d32; color: white; border: none;
                   border-radius: 6px; cursor: pointer; font-size: 0.85em; font-weight: 500;
                   transition: background 0.2s; margin-left: auto; }}
  .download-btn:hover {{ background: #1b5e20; }}
  #map-container {{ padding: 20px 30px; }}
  #map {{ width: 100%; height: 75vh; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
  .info {{ padding: 8px 30px; font-size: 0.8em; color: #888; text-align: center; }}
</style>
</head>
<body>

<div class="header">
  <h1>OSTRAM Transmission Maps</h1>
  <p>Interactive visualization of interconnection capacity and energy flow</p>
</div>

<div class="controls">
  <div class="control-group">
    <label>Map Type</label>
    <div class="tabs" id="tabsContainer"></div>
  </div>
  <div class="control-group">
    <label>Scenario</label>
    <select id="scenarioSelect"></select>
  </div>
  <div class="control-group">
    <label>Year</label>
    <select id="yearSelect"></select>
  </div>
  <div class="control-group">
    <label>Unit</label>
    <div class="toggle-group" id="unitToggle"></div>
  </div>
  <button class="download-btn" onclick="downloadPNG()">&#11015; Download PNG</button>
</div>

<div id="map-container">
  <div id="map"></div>
</div>
<div class="info" id="infoBar"></div>

<script>
// ─── Embedded data ───────────────────────────────────────────────
const capacityData = {capacity_json};
const flowData = {flow_json};
const ratioData = {ratio_json};
const nodes = {nodes_json};

// ─── Unit definitions ────────────────────────────────────────────
const UNITS = {{
  capacity: [
    {{ label: 'GW', factor: 1 }},
    {{ label: 'TW', factor: 0.001 }}
  ],
  flow: [
    {{ label: 'PJ', factor: 1 }},
    {{ label: 'TWh', factor: 0.277778 }},
    {{ label: 'GWh', factor: 277.778 }}
  ],
  ratio: [
    {{ label: 'Ratio', factor: 1 }},
    {{ label: '%', factor: 100 }}
  ]
}};

// ─── State ───────────────────────────────────────────────────────
let currentTab = 'capacity';
let currentUnit = {{ capacity: 0, flow: 0, ratio: 0 }};
let currentScenario = '';
let currentYear = '';

// ─── Initialize controls ────────────────────────────────────────
function init() {{
  // Tabs
  const tabsEl = document.getElementById('tabsContainer');
  const tabLabels = {{ capacity: 'Transmission Capacity', flow: 'Transmission Flow', ratio: 'Load-Capacity Ratio' }};
  ['capacity', 'flow', 'ratio'].forEach(t => {{
    const btn = document.createElement('div');
    btn.className = 'tab' + (t === currentTab ? ' active' : '');
    btn.textContent = tabLabels[t];
    btn.onclick = () => {{ currentTab = t; updateAll(); }};
    tabsEl.appendChild(btn);
  }});

  // Scenarios
  const scenarioSel = document.getElementById('scenarioSelect');
  const scenarios = Object.keys(capacityData).sort();
  scenarios.forEach(s => {{
    const opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    scenarioSel.appendChild(opt);
  }});
  currentScenario = scenarios[0] || '';
  scenarioSel.onchange = () => {{ currentScenario = scenarioSel.value; updateMap(); }};

  // Years
  const yearSel = document.getElementById('yearSelect');
  const allYears = new Set();
  for (const sc of Object.values(capacityData))
    for (const y of Object.keys(sc)) allYears.add(y);
  for (const sc of Object.values(flowData))
    for (const y of Object.keys(sc)) allYears.add(y);
  for (const sc of Object.values(ratioData))
    for (const y of Object.keys(sc)) allYears.add(y);
  const years = [...allYears].sort((a, b) => Number(a) - Number(b));
  years.forEach(y => {{
    const opt = document.createElement('option');
    opt.value = y; opt.textContent = y;
    yearSel.appendChild(opt);
  }});
  currentYear = years[years.length - 1] || '';
  yearSel.value = currentYear;
  yearSel.onchange = () => {{ currentYear = yearSel.value; updateMap(); }};

  buildUnitToggle();
  updateMap();
}}

function buildUnitToggle() {{
  const container = document.getElementById('unitToggle');
  container.innerHTML = '';
  const units = UNITS[currentTab];
  const selIdx = currentUnit[currentTab];
  units.forEach((u, i) => {{
    const btn = document.createElement('button');
    btn.className = 'toggle-btn' + (i === selIdx ? ' active' : '');
    btn.textContent = u.label;
    btn.onclick = () => {{ currentUnit[currentTab] = i; updateAll(); }};
    container.appendChild(btn);
  }});
}}

function updateAll() {{
  // Update tab styling
  const tabKeys = ['capacity', 'flow', 'ratio'];
  document.querySelectorAll('.tab').forEach((el, i) => {{
    el.className = 'tab' + (tabKeys[i] === currentTab ? ' active' : '');
  }});
  buildUnitToggle();
  updateMap();
}}

// ─── Geometry helpers ────────────────────────────────────────────
function offsetLine(fromLat, fromLon, toLat, toLon, offsetDeg) {{
  const dx = toLon - fromLon;
  const dy = toLat - fromLat;
  const len = Math.sqrt(dx * dx + dy * dy) || 1;
  const nx = -dy / len * offsetDeg;
  const ny =  dx / len * offsetDeg;
  return {{
    fromLat: fromLat + ny, fromLon: fromLon + nx,
    toLat:   toLat + ny,   toLon:   toLon + nx
  }};
}}

function pointAlong(lat1, lon1, lat2, lon2, t) {{
  return {{ lat: lat1 + (lat2 - lat1) * t, lon: lon1 + (lon2 - lon1) * t }};
}}

function bearingDeg(lat1, lon1, lat2, lon2) {{
  // Returns clockwise degrees from north (for Plotly marker.angle)
  const dx = lon2 - lon1;
  const dy = lat2 - lat1;
  return (90 - Math.atan2(dy, dx) * 180 / Math.PI + 360) % 360;
}}

// ─── Map rendering ──────────────────────────────────────────────
function updateMap() {{
  const dataMap = {{ capacity: capacityData, flow: flowData, ratio: ratioData }};
  const data = dataMap[currentTab];
  const unitDef = UNITS[currentTab][currentUnit[currentTab]];
  const links = (data[currentScenario] || {{}})[currentYear] || [];
  const labelMap = {{ capacity: 'Capacity', flow: 'Flow', ratio: 'Load-Capacity Ratio' }};
  const baseLabel = labelMap[currentTab];

  const traces = [];

  if (currentTab === 'flow') {{
    // ── Directional flow rendering ──
    const allVals = [];
    links.forEach(l => {{
      allVals.push(l.value_a_to_b * unitDef.factor);
      allVals.push(l.value_b_to_a * unitDef.factor);
    }});
    const maxVal = Math.max(...allVals, 0.001);

    // Dynamic offset based on map extent
    const lons = nodes.map(n => n.lon);
    const lonSpan = Math.max(...lons) - Math.min(...lons);
    const OFFSET = Math.max(0.2, lonSpan * 0.006);

    links.forEach(link => {{
      // Direction A → B (offset to the left of the line)
      const ab = offsetLine(link.from_lat, link.from_lon, link.to_lat, link.to_lon, OFFSET);
      const valAB = link.value_a_to_b * unitDef.factor;
      const wAB = Math.max(0.5, (valAB / maxVal) * 6);
      const dashAB = valAB === 0 ? 'dot' : 'solid';

      traces.push({{
        type: 'scattergeo',
        lon: [ab.fromLon, ab.toLon], lat: [ab.fromLat, ab.toLat],
        mode: 'lines',
        line: {{ width: wAB, color: '#1565C0', dash: dashAB }},
        hoverinfo: 'text',
        text: `${{link.from}} → ${{link.to}}: ${{valAB.toFixed(2)}} ${{unitDef.label}}`,
        showlegend: false
      }});

      // Arrowhead A→B at 85% along offset line
      const tipAB = pointAlong(ab.fromLat, ab.fromLon, ab.toLat, ab.toLon, 0.85);
      const angleAB = bearingDeg(ab.fromLat, ab.fromLon, ab.toLat, ab.toLon);
      traces.push({{
        type: 'scattergeo',
        lon: [tipAB.lon], lat: [tipAB.lat],
        mode: 'markers',
        marker: {{
          symbol: 'triangle-up', size: Math.max(6, wAB * 2.5),
          color: '#1565C0', angle: angleAB,
          line: {{ width: 0 }}
        }},
        hoverinfo: 'text',
        text: [`${{link.from}} → ${{link.to}}: ${{valAB.toFixed(2)}} ${{unitDef.label}}`],
        showlegend: false
      }});

      // Label A→B at 50%
      const midAB = pointAlong(ab.fromLat, ab.fromLon, ab.toLat, ab.toLon, 0.5);
      traces.push({{
        type: 'scattergeo',
        lon: [midAB.lon], lat: [midAB.lat],
        mode: 'text',
        text: [`${{valAB.toFixed(1)}}`],
        textfont: {{ size: 9, color: '#0D47A1', family: 'Arial Black' }},
        hoverinfo: 'none', showlegend: false
      }});

      // Direction B → A (offset to the right of the line)
      const ba = offsetLine(link.from_lat, link.from_lon, link.to_lat, link.to_lon, -OFFSET);
      const valBA = link.value_b_to_a * unitDef.factor;
      const wBA = Math.max(0.5, (valBA / maxVal) * 6);
      const dashBA = valBA === 0 ? 'dot' : 'solid';

      // Draw B→A line reversed (from TO coords to FROM coords on the offset)
      traces.push({{
        type: 'scattergeo',
        lon: [ba.toLon, ba.fromLon], lat: [ba.toLat, ba.fromLat],
        mode: 'lines',
        line: {{ width: wBA, color: '#C62828', dash: dashBA }},
        hoverinfo: 'text',
        text: `${{link.to}} → ${{link.from}}: ${{valBA.toFixed(2)}} ${{unitDef.label}}`,
        showlegend: false
      }});

      // Arrowhead B→A at 85% along (from to_offset to from_offset)
      const tipBA = pointAlong(ba.toLat, ba.toLon, ba.fromLat, ba.fromLon, 0.85);
      const angleBA = bearingDeg(ba.toLat, ba.toLon, ba.fromLat, ba.fromLon);
      traces.push({{
        type: 'scattergeo',
        lon: [tipBA.lon], lat: [tipBA.lat],
        mode: 'markers',
        marker: {{
          symbol: 'triangle-up', size: Math.max(6, wBA * 2.5),
          color: '#C62828', angle: angleBA,
          line: {{ width: 0 }}
        }},
        hoverinfo: 'text',
        text: [`${{link.to}} → ${{link.from}}: ${{valBA.toFixed(2)}} ${{unitDef.label}}`],
        showlegend: false
      }});

      // Label B→A at 50%
      const midBA = pointAlong(ba.toLat, ba.toLon, ba.fromLat, ba.fromLon, 0.5);
      traces.push({{
        type: 'scattergeo',
        lon: [midBA.lon], lat: [midBA.lat],
        mode: 'text',
        text: [`${{valBA.toFixed(1)}}`],
        textfont: {{ size: 9, color: '#8E0000', family: 'Arial Black' }},
        hoverinfo: 'none', showlegend: false
      }});
    }});

    // Legend entries for flow directions
    traces.push({{
      type: 'scattergeo', lon: [null], lat: [null],
      mode: 'lines', line: {{ color: '#1565C0', width: 3 }},
      name: 'A → B', showlegend: true
    }});
    traces.push({{
      type: 'scattergeo', lon: [null], lat: [null],
      mode: 'lines', line: {{ color: '#C62828', width: 3 }},
      name: 'B → A', showlegend: true
    }});

  }} else {{
    // ── Non-directional rendering (capacity / ratio) ──
    const values = links.map(l => l.value * unitDef.factor);
    const maxVal = Math.max(...values, 0.001);

    links.forEach(link => {{
      const val = link.value * unitDef.factor;
      const width = Math.max(0.8, (val / maxVal) * 6);

      let lineColor = '#b71c1c';
      if (currentTab === 'ratio') {{
        const normalized = Math.min(Math.max(link.value, 0), 1);
        const r = Math.round(220 * normalized + 30);
        const g = Math.round(200 * (1 - normalized) + 30);
        lineColor = `rgb(${{r}}, ${{g}}, 30)`;
      }}

      traces.push({{
        type: 'scattergeo',
        lon: [link.from_lon, link.to_lon],
        lat: [link.from_lat, link.to_lat],
        mode: 'lines',
        line: {{ width: width, color: lineColor }},
        hoverinfo: 'text',
        text: `${{link.from}} → ${{link.to}}: ${{val.toFixed(2)}} ${{unitDef.label}}`,
        showlegend: false
      }});

      const midLat = (link.from_lat + link.to_lat) / 2;
      const midLon = (link.from_lon + link.to_lon) / 2;
      traces.push({{
        type: 'scattergeo',
        lon: [midLon], lat: [midLat],
        mode: 'text',
        text: [`${{val.toFixed(1)}}`],
        textfont: {{ size: 10, color: '#333', family: 'Arial Black' }},
        hoverinfo: 'none', showlegend: false
      }});
    }});
  }}

  // Draw nodes
  const nodeLons = nodes.map(n => n.lon);
  const nodeLats = nodes.map(n => n.lat);
  const nodeLabels = nodes.map(n => n.region);
  traces.push({{
    type: 'scattergeo',
    lon: nodeLons, lat: nodeLats,
    mode: 'markers+text',
    marker: {{ size: 8, color: '#0d47a1', line: {{ width: 1, color: '#fff' }} }},
    text: nodeLabels,
    textposition: 'top center',
    textfont: {{ size: 9, color: '#1a237e', family: 'Arial' }},
    hoverinfo: 'text',
    hovertext: nodeLabels,
    showlegend: false
  }});

  // Auto-compute map extent from nodes
  const pad = 3;
  const lonMin = Math.min(...nodeLons) - pad;
  const lonMax = Math.max(...nodeLons) + pad;
  const latMin = Math.min(...nodeLats) - pad;
  const latMax = Math.max(...nodeLats) + pad;

  const layout = {{
    title: {{
      text: `${{baseLabel}} — ${{currentScenario}} — ${{currentYear}} (${{unitDef.label}})`,
      font: {{ size: 16, family: 'Segoe UI', color: '#333' }}
    }},
    geo: {{
      scope: 'world',
      projection: {{ type: 'mercator' }},
      showland: true, landcolor: '#e8e8e8',
      showocean: true, oceancolor: '#d4e6f1',
      showcountries: true, countrycolor: '#bbb',
      showcoastlines: true, coastlinecolor: '#999',
      lonaxis: {{ range: [lonMin, lonMax] }},
      lataxis: {{ range: [latMin, latMax] }},
      resolution: 50
    }},
    margin: {{ l: 0, r: 0, t: 50, b: 0 }},
    height: window.innerHeight * 0.72,
    showlegend: currentTab === 'flow',
    legend: {{ x: 1, y: 1, bgcolor: 'rgba(255,255,255,0.8)', bordercolor: '#ccc', borderwidth: 1 }}
  }};

  Plotly.react('map', traces, layout, {{ responsive: true }});

  // Info bar
  const linkCount = currentTab === 'flow' ? links.length * 2 : links.length;
  const dirLabel = currentTab === 'flow' ? ' (bidirectional)' : '';
  document.getElementById('infoBar').textContent =
    `${{linkCount}} interconnections${{dirLabel}} | Scenario: ${{currentScenario}} | Year: ${{currentYear}} | Unit: ${{unitDef.label}}`;
}}

function downloadPNG() {{
  const unitDef = UNITS[currentTab][currentUnit[currentTab]];
  const dlLabels = {{ capacity: 'TransmissionCapacity', flow: 'TransmissionFlow', ratio: 'LoadCapacityRatio' }};
  const baseLabel = dlLabels[currentTab];
  Plotly.downloadImage('map', {{
    format: 'png', width: 1600, height: 1000,
    filename: `${{baseLabel}}_${{currentScenario}}_${{currentYear}}_${{unitDef.label}}`
  }});
}}

// ─── Start ──────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', init);
window.addEventListener('resize', () => {{
  Plotly.relayout('map', {{ height: window.innerHeight * 0.72 }});
}});
</script>
</body>
</html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"  HTML saved to {output_path}")


# ---------------------------------------------------------------------------
# Dispatch Chart Constants
# ---------------------------------------------------------------------------
DISPATCH_FUEL_PATTERN = re.compile(r'^ELC.{5}0[012]$')

TECH_STACK_ORDER = [
    'GEO', 'URN', 'BIO', 'WAS', 'COA', 'COG', 'CCS', 'OIL', 'PET',
    'NGS', 'HYD', 'LDS', 'SDS', 'SPV', 'CSP', 'WON', 'WOF', 'OTH', 'BCK', 'CRT',
]

TECH_COLORS = {
    'GEO': '#CDDC39', 'URN': '#9C27B0', 'BIO': '#2E7D32', 'WAS': '#795548',
    'COA': '#424242', 'COG': '#6D4C41', 'CCS': '#00897B', 'OIL': '#212121',
    'PET': '#B71C1C', 'NGS': '#689F38', 'HYD': '#4FC3F7', 'LDS': '#78909C',
    'SDS': '#BDBDBD', 'SPV': '#FDD835', 'CSP': '#FF9800', 'WON': '#1E88E5',
    'WOF': '#0D47A1', 'OTH': '#7B1FA2', 'BCK': '#F44336', 'CRT': '#E0E0E0',
}

TECH_LABELS = {
    'GEO': 'Geothermal', 'URN': 'Nuclear', 'BIO': 'Biomass', 'WAS': 'Waste',
    'COA': 'Coal', 'COG': 'Coal+Gas', 'CCS': 'CCS', 'OIL': 'Oil',
    'PET': 'Petroleum', 'NGS': 'Natural Gas', 'HYD': 'Hydro', 'LDS': 'Long Storage',
    'SDS': 'Short Storage', 'SPV': 'Solar PV', 'CSP': 'CSP', 'WON': 'Wind Onshore',
    'WOF': 'Wind Offshore', 'OTH': 'Other', 'BCK': 'Backstop', 'CRT': 'Curtailment',
}

TIMESLICE_ORDER = [
    'S1D1', 'S1D2', 'S1D3',
    'S2D1', 'S2D2', 'S2D3',
    'S3D1', 'S3D2', 'S3D3',
    'S4D1', 'S4D2', 'S4D3',
]


def load_dispatch_data(csv_path):
    """Load combined CSV and extract dispatch data for PWR* technologies.

    Returns (dispatch_df, year_split_dict).
    """
    print(f"  Loading dispatch data from {csv_path} ...")
    df = pd.read_csv(csv_path, low_memory=False)

    # Extract YearSplit lookup
    ys_rows = df[df[YEAR_SPLIT_COL].notna()][['TIMESLICE', YEAR_SPLIT_COL]].drop_duplicates()
    year_split = dict(zip(ys_rows['TIMESLICE'], ys_rows[YEAR_SPLIT_COL]))

    # Filter: PWR* technologies, ELC*00 or ELC*01 fuels, non-null production
    mask = (
        df['TECHNOLOGY'].astype(str).str.startswith('PWR')
        & df['FUEL'].astype(str).apply(lambda f: bool(DISPATCH_FUEL_PATTERN.match(f)))
        & df[PRODUCTION_BY_TIMESLICE_COL].notna()
        & (df[PRODUCTION_BY_TIMESLICE_COL] != 0)
        & df['TIMESLICE'].notna()
    )
    ddf = df[mask].copy()

    # Extract region from fuel: ELCBGDXX00 -> BGDXX
    ddf['REGION'] = ddf['FUEL'].str[3:-2]
    # Extract tech code: PWRSPVBGDXX01 -> SPV
    ddf['TECH_CODE'] = ddf['TECHNOLOGY'].str[3:6]
    # Extract fuel suffix: 00, 01, 02
    ddf['FUEL_SUFFIX'] = ddf['FUEL'].str[-2:]

    print(f"  Found {len(ddf):,} dispatch rows")
    print(f"  Regions: {sorted(ddf['REGION'].unique().tolist())}")
    print(f"  Tech codes: {sorted(ddf['TECH_CODE'].unique().tolist())}")

    return ddf, year_split


def prepare_dispatch_json(ddf, year_split):
    """Prepare dispatch data as nested dict for JSON embedding.

    Converts ProductionByTechnology to GW:
      rate_PJ_per_yr = ProductionByTechnology / YearSplit
      power_GW = rate_PJ_per_yr / 31.536

    Returns dict: { scenario: { year: { region: { timeslice: { tech: GW } } } } }
    """
    ddf = ddf.copy()
    ddf['YEAR'] = pd.to_numeric(ddf['YEAR'], errors='coerce')
    ddf = ddf.dropna(subset=['YEAR'])
    ddf['YEAR'] = ddf['YEAR'].astype(int)

    # Separate generation (ELC*00/01) from demand-side (ELC*02)
    gen_df = ddf[ddf['FUEL_SUFFIX'].isin(['00', '01'])].copy()
    dem_df = ddf[ddf['FUEL_SUFFIX'] == '02'].copy()

    # Aggregate generation: sum PJ by scenario, year, region, timeslice, tech_code
    agg = gen_df.groupby(['Scenario', 'YEAR', 'REGION', 'TIMESLICE', 'TECH_CODE'])[PRODUCTION_BY_TIMESLICE_COL].sum().reset_index()
    agg.rename(columns={PRODUCTION_BY_TIMESLICE_COL: 'PJ'}, inplace=True)

    # Compute curtailment per scenario/year/region/timeslice
    gen_total = gen_df.groupby(['Scenario', 'YEAR', 'REGION', 'TIMESLICE'])[PRODUCTION_BY_TIMESLICE_COL].sum().reset_index()
    gen_total.rename(columns={PRODUCTION_BY_TIMESLICE_COL: 'GEN_PJ'}, inplace=True)

    dem_total = dem_df.groupby(['Scenario', 'YEAR', 'REGION', 'TIMESLICE'])[PRODUCTION_BY_TIMESLICE_COL].sum().reset_index()
    dem_total.rename(columns={PRODUCTION_BY_TIMESLICE_COL: 'DEM_PJ'}, inplace=True)

    curt = gen_total.merge(dem_total, on=['Scenario', 'YEAR', 'REGION', 'TIMESLICE'], how='left')
    curt['DEM_PJ'] = curt['DEM_PJ'].fillna(0)
    curt['CRT_PJ'] = curt['GEN_PJ'] - curt['DEM_PJ']
    curt = curt[curt['CRT_PJ'] > 0]

    # Add curtailment rows to aggregation
    if not curt.empty:
        crt_rows = curt[['Scenario', 'YEAR', 'REGION', 'TIMESLICE', 'CRT_PJ']].copy()
        crt_rows['TECH_CODE'] = 'CRT'
        crt_rows.rename(columns={'CRT_PJ': 'PJ'}, inplace=True)
        agg = pd.concat([agg, crt_rows], ignore_index=True)

    dispatch_data = {}
    for _, row in agg.iterrows():
        sc = row['Scenario']
        yr = str(row['YEAR'])
        reg = row['REGION']
        ts = row['TIMESLICE']
        tc = row['TECH_CODE']
        val = round(float(row['PJ']), 6)

        dispatch_data.setdefault(sc, {}).setdefault(yr, {}).setdefault(reg, {}).setdefault(ts, {})[tc] = val

    return dispatch_data


def generate_dispatch_html(dispatch_data, output_path):
    """Generate standalone HTML with interactive stacked area dispatch chart."""

    data_json = json.dumps(dispatch_data)
    stack_order_json = json.dumps(TECH_STACK_ORDER)
    colors_json = json.dumps(TECH_COLORS)
    labels_json = json.dumps(TECH_LABELS)
    ts_order_json = json.dumps(TIMESLICE_ORDER)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OSTRAM - Dispatch Chart</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; }}
  .header {{ background: linear-gradient(135deg, #1a237e, #0d47a1); color: white; padding: 20px 30px; }}
  .header h1 {{ font-size: 1.6em; font-weight: 600; }}
  .header p {{ font-size: 0.9em; opacity: 0.85; margin-top: 4px; }}
  .controls {{ display: flex; flex-wrap: wrap; gap: 16px; align-items: center;
               padding: 16px 30px; background: #fff; border-bottom: 1px solid #ddd; }}
  .control-group {{ display: flex; flex-direction: column; gap: 4px; }}
  .control-group label {{ font-size: 0.75em; font-weight: 600; text-transform: uppercase;
                          color: #666; letter-spacing: 0.5px; }}
  select {{ font-size: 0.9em; padding: 6px 12px; border: 1px solid #ccc;
            border-radius: 6px; background: #fff; }}
  .toggle-group {{ display: inline-flex; gap: 0; padding: 0; overflow: hidden;
                   border: 1px solid #ccc; border-radius: 6px; }}
  .toggle-btn {{ padding: 6px 14px; border: none; background: #f5f5f5; cursor: pointer;
                 font-size: 0.85em; font-weight: 500; transition: all 0.2s; }}
  .toggle-btn.active {{ background: #0d47a1; color: white; }}
  .toggle-btn:not(:last-child) {{ border-right: 1px solid #ddd; }}
  .download-btn {{ padding: 8px 18px; background: #2e7d32; color: white; border: none;
                   border-radius: 6px; cursor: pointer; font-size: 0.85em; font-weight: 500;
                   transition: background 0.2s; margin-left: auto; }}
  .download-btn:hover {{ background: #1b5e20; }}
  #chart-container {{ padding: 20px 30px; }}
  #chart {{ width: 100%; height: 75vh; border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1); background: #fff; }}
  .info {{ padding: 8px 30px; font-size: 0.8em; color: #888; text-align: center; }}
</style>
</head>
<body>

<div class="header">
  <h1>OSTRAM Dispatch Chart</h1>
  <p>Generation mix by technology and timeslice (stacked area)</p>
</div>

<div class="controls">
  <div class="control-group">
    <label>Scenario</label>
    <select id="scenarioSelect"></select>
  </div>
  <div class="control-group">
    <label>Year</label>
    <select id="yearSelect"></select>
  </div>
  <div class="control-group">
    <label>Region</label>
    <select id="regionSelect"></select>
  </div>
  <div class="control-group">
    <label>Unit</label>
    <div class="toggle-group" id="unitToggle">
      <button class="toggle-btn active" onclick="setUnit(0)">PJ</button>
      <button class="toggle-btn" onclick="setUnit(1)">GWh</button>
      <button class="toggle-btn" onclick="setUnit(2)">MWh</button>
    </div>
  </div>
  <button class="download-btn" onclick="downloadPNG()">&#11015; Download PNG</button>
</div>

<div id="chart-container">
  <div id="chart"></div>
</div>
<div class="info" id="infoBar"></div>

<script>
// ─── Embedded data ───────────────────────────────────────────────
const dispatchData = {data_json};
const STACK_ORDER = {stack_order_json};
const TECH_COLORS = {colors_json};
const TECH_LABELS = {labels_json};
const TS_ORDER = {ts_order_json};

const UNITS = [
  {{ label: 'PJ', factor: 1 }},
  {{ label: 'GWh', factor: 277.778 }},
  {{ label: 'MWh', factor: 277778 }}
];

// Season labels for x-axis annotation
const SEASON_LABELS = {{
  'S1D1': 'Season 1', 'S2D1': 'Season 2', 'S3D1': 'Season 3', 'S4D1': 'Season 4'
}};

// ─── State ───────────────────────────────────────────────────────
let currentScenario = '';
let currentYear = '';
let currentRegion = '';
let currentUnitIdx = 0;

// ─── Initialize ─────────────────────────────────────────────────
function init() {{
  const scenarios = Object.keys(dispatchData).sort();
  const scenarioSel = document.getElementById('scenarioSelect');
  scenarios.forEach(s => {{
    const opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    scenarioSel.appendChild(opt);
  }});
  currentScenario = scenarios[0] || '';
  scenarioSel.onchange = () => {{ currentScenario = scenarioSel.value; populateYears(); populateRegions(); updateChart(); }};

  populateYears();
  populateRegions();
  updateChart();
}}

function populateYears() {{
  const yearSel = document.getElementById('yearSelect');
  yearSel.innerHTML = '';
  const years = Object.keys(dispatchData[currentScenario] || {{}}).sort((a, b) => Number(a) - Number(b));
  years.forEach(y => {{
    const opt = document.createElement('option');
    opt.value = y; opt.textContent = y;
    yearSel.appendChild(opt);
  }});
  if (years.includes(currentYear)) {{
    yearSel.value = currentYear;
  }} else {{
    currentYear = years[years.length - 1] || '';
    yearSel.value = currentYear;
  }}
  yearSel.onchange = () => {{ currentYear = yearSel.value; populateRegions(); updateChart(); }};
}}

function populateRegions() {{
  const regionSel = document.getElementById('regionSelect');
  regionSel.innerHTML = '';
  const yearData = (dispatchData[currentScenario] || {{}})[currentYear] || {{}};
  const regions = Object.keys(yearData).sort();
  regions.forEach(r => {{
    const opt = document.createElement('option');
    opt.value = r; opt.textContent = r;
    regionSel.appendChild(opt);
  }});
  if (regions.includes(currentRegion)) {{
    regionSel.value = currentRegion;
  }} else {{
    currentRegion = regions[0] || '';
    regionSel.value = currentRegion;
  }}
  regionSel.onchange = () => {{ currentRegion = regionSel.value; updateChart(); }};
}}

function setUnit(idx) {{
  currentUnitIdx = idx;
  document.querySelectorAll('.toggle-btn').forEach((btn, i) => {{
    btn.className = 'toggle-btn' + (i === idx ? ' active' : '');
  }});
  updateChart();
}}

// ─── Chart rendering ────────────────────────────────────────────
function updateChart() {{
  const yearData = (dispatchData[currentScenario] || {{}})[currentYear] || {{}};
  const regionData = yearData[currentRegion] || {{}};
  const unitDef = UNITS[currentUnitIdx];

  // Build traces in stack order
  const traces = [];
  const presentTechs = new Set();
  TS_ORDER.forEach(ts => {{
    const tsData = regionData[ts] || {{}};
    Object.keys(tsData).forEach(tc => presentTechs.add(tc));
  }});

  const orderedTechs = STACK_ORDER.filter(tc => presentTechs.has(tc));

  orderedTechs.forEach(tc => {{
    const yValues = TS_ORDER.map(ts => {{
      const val = (regionData[ts] || {{}})[tc] || 0;
      return val * unitDef.factor;
    }});
    traces.push({{
      x: TS_ORDER,
      y: yValues,
      name: TECH_LABELS[tc] || tc,
      type: 'scatter',
      mode: 'lines',
      fill: 'tonexty',
      stackgroup: 'one',
      line: {{ width: 0.5, color: TECH_COLORS[tc] || '#999' }},
      fillcolor: TECH_COLORS[tc] || '#999',
      hovertemplate: `%{{x}}<br>${{TECH_LABELS[tc] || tc}}: %{{y:.2f}} ${{unitDef.label}}<extra></extra>`
    }});
  }});

  // Season separator annotations
  const annotations = [];
  const shapes = [];
  Object.entries(SEASON_LABELS).forEach(([ts, label]) => {{
    const idx = TS_ORDER.indexOf(ts);
    if (idx >= 0) {{
      annotations.push({{
        x: ts, y: 1.06, xref: 'x', yref: 'paper',
        text: `<b>${{label}}</b>`, showarrow: false,
        font: {{ size: 12, color: '#555' }}
      }});
      if (idx > 0) {{
        shapes.push({{
          type: 'line', x0: idx - 0.5, x1: idx - 0.5,
          y0: 0, y1: 1, xref: 'x', yref: 'paper',
          line: {{ color: '#aaa', width: 1, dash: 'dot' }}
        }});
      }}
    }}
  }});

  const layout = {{
    title: {{
      text: `Generation Mix — ${{currentScenario}} — ${{currentYear}} — ${{currentRegion}} (${{unitDef.label}})`,
      font: {{ size: 16, family: 'Segoe UI', color: '#333' }}
    }},
    xaxis: {{
      title: 'Timeslice',
      tickangle: -45,
      type: 'category',
      categoryorder: 'array',
      categoryarray: TS_ORDER
    }},
    yaxis: {{
      title: unitDef.label,
      rangemode: 'tozero'
    }},
    annotations: annotations,
    shapes: shapes,
    margin: {{ l: 70, r: 30, t: 80, b: 80 }},
    height: window.innerHeight * 0.72,
    legend: {{
      orientation: 'h', y: -0.25, x: 0.5, xanchor: 'center',
      font: {{ size: 11 }},
      traceorder: 'normal'
    }},
    hovermode: 'x unified'
  }};

  Plotly.react('chart', traces, layout, {{ responsive: true }});

  document.getElementById('infoBar').textContent =
    `${{orderedTechs.length}} technologies | Region: ${{currentRegion}} | Scenario: ${{currentScenario}} | Year: ${{currentYear}}`;
}}

function downloadPNG() {{
  const unitDef = UNITS[currentUnitIdx];
  Plotly.downloadImage('chart', {{
    format: 'png', width: 1600, height: 900,
    filename: `DispatchChart_${{currentScenario}}_${{currentYear}}_${{currentRegion}}_${{unitDef.label}}`
  }});
}}

// ─── Start ──────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', init);
window.addEventListener('resize', () => {{
  Plotly.relayout('chart', {{ height: window.innerHeight * 0.72 }});
}});
</script>
</body>
</html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"  Dispatch HTML saved to {output_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    script_dir = Path(__file__).resolve().parent
    csv_path = script_dir / 'OSTRAM_Combined_Inputs_Outputs.csv'
    centerpoints_path = script_dir / 'Miscellaneous' / 'centerpoints.csv'
    output_dir = script_dir / 'Figures'
    output_path = output_dir / 'TransmissionMaps.html'
    dispatch_output_path = output_dir / 'DispatchChart.html'

    if not csv_path.exists():
        raise FileNotFoundError(f"Data file not found: {csv_path}")
    if not centerpoints_path.exists():
        raise FileNotFoundError(f"Centerpoints file not found: {centerpoints_path}")

    output_dir.mkdir(exist_ok=True)

    # --- Transmission Maps ---
    print("=" * 60)
    print("OSTRAM Transmission Maps Generator")
    print("=" * 60)

    df, year_split = load_data(str(csv_path))
    centerpoints = load_centerpoints(str(centerpoints_path))

    capacity_data, flow_data, ratio_data = prepare_json_data(df, centerpoints, year_split)
    node_list = build_node_list(centerpoints, capacity_data, flow_data, ratio_data)

    generate_html(capacity_data, flow_data, ratio_data, node_list, str(output_path))

    print(f"  Done! Open {output_path}")

    # --- Dispatch Chart ---
    print()
    print("=" * 60)
    print("OSTRAM Dispatch Chart Generator")
    print("=" * 60)

    ddf, ys = load_dispatch_data(str(csv_path))
    dispatch_data = prepare_dispatch_json(ddf, ys)
    generate_dispatch_html(dispatch_data, str(dispatch_output_path))

    print("=" * 60)
    print(f"Done! Open {dispatch_output_path}")
    print("=" * 60)


if __name__ == '__main__':
    main()
