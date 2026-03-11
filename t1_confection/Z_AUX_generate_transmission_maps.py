# -*- coding: utf-8 -*-
"""
Transmission Maps Generator for OSTRAM

Generates a standalone interactive HTML file with two map views:
  - Transmission Capacity (GW/TW)
  - Transmission Flow (PJ/TWh/GWh)

Features:
  - Year selector (all years discovered from data)
  - Scenario selector (all scenarios discovered from data)
  - Unit toggle (GW↔TW for capacity; PJ↔TWh↔GWh for flow)
  - Interactive Plotly Scattergeo maps with hover info
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
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
INTERCONNECTION_PATTERN = re.compile(r'^TRN[A-Z]{5}[A-Z]{5}$')

CAPACITY_COL = 'TotalCapacityAnnual'
FLOW_COL = 'ProductionByTechnologyAnnual'


def load_data(csv_path):
    """Load the combined CSV and filter to interconnection technologies."""
    print(f"  Loading data from {csv_path} ...")
    df = pd.read_csv(csv_path, low_memory=False)

    mask = df['TECHNOLOGY'].astype(str).apply(
        lambda t: bool(INTERCONNECTION_PATTERN.match(t))
    )
    df = df[mask].copy()
    print(f"  Found {len(df):,} rows with interconnection technologies")
    print(f"  Unique interconnections: {df['TECHNOLOGY'].nunique()}")
    return df


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


def prepare_json_data(df, centerpoints):
    """Prepare all data as nested dicts for JSON embedding in HTML.

    Returns two dicts (capacity_data, flow_data) structured as:
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
        return {}, {}

    # Add coordinates
    df['from_lat'] = df['FROM'].map(lambda r: centerpoints[r]['lat'])
    df['from_lon'] = df['FROM'].map(lambda r: centerpoints[r]['long'])
    df['to_lat'] = df['TO'].map(lambda r: centerpoints[r]['lat'])
    df['to_lon'] = df['TO'].map(lambda r: centerpoints[r]['long'])

    scenarios = sorted(df['Scenario'].dropna().unique().tolist())
    years = sorted(df['YEAR'].unique().tolist())

    capacity_data = {}
    flow_data = {}

    for scenario in scenarios:
        capacity_data[scenario] = {}
        flow_data[scenario] = {}
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

            # --- Flow ---
            flw = ydf[ydf[FLOW_COL].notna() & (ydf[FLOW_COL] != 0)]
            flw_agg = (
                flw.groupby(['FROM', 'TO', 'from_lat', 'from_lon', 'to_lat', 'to_lon'])
                [FLOW_COL].sum().reset_index()
            )
            flow_data[scenario][str(year)] = [
                {
                    'from': row['FROM'], 'to': row['TO'],
                    'from_lat': row['from_lat'], 'from_lon': row['from_lon'],
                    'to_lat': row['to_lat'], 'to_lon': row['to_lon'],
                    'value': round(float(row[FLOW_COL]), 4)
                }
                for _, row in flw_agg.iterrows()
            ]

    print(f"  Scenarios: {scenarios}")
    print(f"  Years: {years[0]}-{years[-1]} ({len(years)} years)")

    return capacity_data, flow_data


def build_node_list(centerpoints, capacity_data, flow_data):
    """Build list of all nodes involved in any interconnection."""
    nodes = set()
    for scenario_data in [capacity_data, flow_data]:
        for scenario in scenario_data.values():
            for year_links in scenario.values():
                for link in year_links:
                    nodes.add(link['from'])
                    nodes.add(link['to'])

    return [
        {'region': r, 'lat': centerpoints[r]['lat'], 'lon': centerpoints[r]['long']}
        for r in sorted(nodes) if r in centerpoints
    ]


def generate_html(capacity_data, flow_data, nodes, output_path):
    """Generate standalone HTML file with interactive transmission maps."""

    capacity_json = json.dumps(capacity_data)
    flow_json = json.dumps(flow_data)
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
  ]
}};

// ─── State ───────────────────────────────────────────────────────
let currentTab = 'capacity';
let currentUnit = {{ capacity: 0, flow: 0 }};
let currentScenario = '';
let currentYear = '';

// ─── Initialize controls ────────────────────────────────────────
function init() {{
  // Tabs
  const tabsEl = document.getElementById('tabsContainer');
  ['capacity', 'flow'].forEach(t => {{
    const btn = document.createElement('div');
    btn.className = 'tab' + (t === currentTab ? ' active' : '');
    btn.textContent = t === 'capacity' ? 'Transmission Capacity' : 'Transmission Flow';
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
  document.querySelectorAll('.tab').forEach((el, i) => {{
    el.className = 'tab' + ((i === 0 && currentTab === 'capacity') ||
                             (i === 1 && currentTab === 'flow') ? ' active' : '');
  }});
  buildUnitToggle();
  updateMap();
}}

// ─── Map rendering ──────────────────────────────────────────────
function updateMap() {{
  const data = currentTab === 'capacity' ? capacityData : flowData;
  const unitDef = UNITS[currentTab][currentUnit[currentTab]];
  const links = (data[currentScenario] || {{}})[currentYear] || [];
  const baseLabel = currentTab === 'capacity' ? 'Capacity' : 'Flow';

  // Compute max value for line width scaling
  const values = links.map(l => l.value * unitDef.factor);
  const maxVal = Math.max(...values, 0.001);

  const traces = [];

  // Draw lines
  links.forEach((link, idx) => {{
    const val = link.value * unitDef.factor;
    const width = Math.max(0.8, (val / maxVal) * 6);
    traces.push({{
      type: 'scattergeo',
      lon: [link.from_lon, link.to_lon],
      lat: [link.from_lat, link.to_lat],
      mode: 'lines',
      line: {{ width: width, color: '#b71c1c' }},
      hoverinfo: 'text',
      text: `${{link.from}} → ${{link.to}}: ${{val.toFixed(2)}} ${{unitDef.label}}`,
      showlegend: false
    }});

    // Midpoint label
    const midLat = (link.from_lat + link.to_lat) / 2;
    const midLon = (link.from_lon + link.to_lon) / 2;
    traces.push({{
      type: 'scattergeo',
      lon: [midLon], lat: [midLat],
      mode: 'text',
      text: [`${{val.toFixed(1)}}`],
      textfont: {{ size: 10, color: '#333', family: 'Arial Black' }},
      hoverinfo: 'none',
      showlegend: false
    }});
  }});

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
    height: window.innerHeight * 0.72
  }};

  Plotly.react('map', traces, layout, {{ responsive: true }});

  // Info bar
  document.getElementById('infoBar').textContent =
    `${{links.length}} interconnections | Scenario: ${{currentScenario}} | Year: ${{currentYear}} | Unit: ${{unitDef.label}}`;
}}

function downloadPNG() {{
  const unitDef = UNITS[currentTab][currentUnit[currentTab]];
  const baseLabel = currentTab === 'capacity' ? 'TransmissionCapacity' : 'TransmissionFlow';
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
# Main
# ---------------------------------------------------------------------------
def main():
    script_dir = Path(__file__).resolve().parent
    csv_path = script_dir / 'OSTRAM_Combined_Inputs_Outputs.csv'
    centerpoints_path = script_dir / 'Miscellaneous' / 'centerpoints.csv'
    output_dir = script_dir / 'Figures'
    output_path = output_dir / 'TransmissionMaps.html'

    if not csv_path.exists():
        raise FileNotFoundError(f"Data file not found: {csv_path}")
    if not centerpoints_path.exists():
        raise FileNotFoundError(f"Centerpoints file not found: {centerpoints_path}")

    output_dir.mkdir(exist_ok=True)

    print("=" * 60)
    print("OSTRAM Transmission Maps Generator")
    print("=" * 60)

    df = load_data(str(csv_path))
    centerpoints = load_centerpoints(str(centerpoints_path))

    capacity_data, flow_data = prepare_json_data(df, centerpoints)
    node_list = build_node_list(centerpoints, capacity_data, flow_data)

    generate_html(capacity_data, flow_data, node_list, str(output_path))

    print("=" * 60)
    print(f"Done! Open {output_path}")
    print("=" * 60)


if __name__ == '__main__':
    main()
