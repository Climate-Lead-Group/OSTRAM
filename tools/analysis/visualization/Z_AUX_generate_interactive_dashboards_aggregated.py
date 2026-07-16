"""
Interactive HTML Dashboard Generator for PWR Technologies - AGGREGATED VERSION

This script generates HTML dashboards with graphs AGGREGATED BY SCENARIO:

KEY DIFFERENCE from the standard version:
- Graphs show values SUMMED by scenario only
- If you select 5 countries and 6 fuels, all their values are summed
- The selection of fuels and countries automatically updates the included technologies

Interactive Filters:
- Scenario Selection
- Fuel Selection (FUEL) -> Filters technologies automatically
- Country Selection (COUNTRY) -> Filters technologies automatically
- Technology View (auto-selected, not editable)
- Quick buttons: Renewables / Non-Renewables / Select All / Reset
- Switch between line and bar charts

Included Charts (ALL AGGREGATED BY SCENARIO):
1. Renewability Shares (% based on ProductionByTechnology)
2. Total Sum by Scenario - Lower Limit (aggregated across all years)
3. Total Sum by Scenario - Production (aggregated across all years)
4. Temporal Evolution - Lower Limit (by year, summed by scenario)
5. Temporal Evolution - Production (by year, summed by scenario)

Features:
- Automatic filtering of valid PWR technologies (regex pattern)
- Auto-selection of technologies based on FUEL and COUNTRY
- Automatic aggregation by scenario
- Visual classification with colors (Green=Renewable, Red=Non-Renewable)
- Completely standalone (no server required)
- Real-time updates without page reload

Author: Climate Lead Group, Andrey Salazar-Vargas
Date: 2026-01-22
"""

import pandas as pd
import json
from datetime import datetime
import os
import glob
import re

# Fuel classification
RENEWABLE_FUELS = ['BIO', 'WAS', 'CSP', 'SPV', 'GEO', 'HYD', 'WAV', 'WON', 'WOF']
NON_RENEWABLE_FUELS = ['URN', 'NGS', 'COA', 'COG', 'OIL', 'PET', 'CCS', 'OTH']

# All valid fuels for PWR technologies
VALID_FUELS = RENEWABLE_FUELS + NON_RENEWABLE_FUELS

# Regex pattern for valid PWR technologies: PWR + FUEL(3 letters) + COUNTRY(3 letters) + XX
TECH_PATTERN = r'^PWR(' + '|'.join(VALID_FUELS) + r')[A-Z]{3}XX$'

# Colors
COLORS = {
    'background': '#f8f9fa',
    'card': '#ffffff',
    'primary': '#0066cc',
    'secondary': '#6c757d',
    'text': '#212529',
    'border': '#dee2e6',
    'renewable': '#2ecc71',
    'non_renewable': '#e74c3c'
}


def is_valid_pwr_technology(tech):
    """
    Validates whether a technology matches the expected PWR pattern

    Args:
        tech: Technology code (e.g.: 'PWRBIOARGXX')

    Returns:
        True if it matches the pattern, False otherwise
    """
    if pd.isna(tech):
        return False
    return bool(re.match(TECH_PATTERN, str(tech)))


def filter_pwr_technologies(df):
    """
    Filters only valid PWR technologies according to the specified pattern

    Args:
        df: DataFrame with TECHNOLOGY column

    Returns:
        DataFrame filtered with only valid PWR technologies
    """
    if 'TECHNOLOGY' not in df.columns:
        return df

    rows_before = len(df)

    # Apply valid technology filter
    df_filtered = df[df['TECHNOLOGY'].apply(is_valid_pwr_technology)].copy()

    rows_after = len(df_filtered)
    rows_removed = rows_before - rows_after

    if rows_removed > 0:
        print(f"   🔍 PWR technology filtering:")
        print(f"      - Rows before: {rows_before:,}")
        print(f"      - Rows after: {rows_after:,}")
        print(f"      - Rows removed: {rows_removed:,}")
        print(f"      - Valid unique technologies: {df_filtered['TECHNOLOGY'].nunique():,}")

    return df_filtered


def extract_fuel_country(df):
    """
    Extracts FUEL and COUNTRY from the TECHNOLOGY column

    Expected format: PWR[FUEL][COUNTRY]XX
    - FUEL: positions 3-5 (characters 3, 4, 5)
    - COUNTRY: positions 6-8 (characters 6, 7, 8)
    """
    df = df.copy()
    df['FUEL'] = df['TECHNOLOGY'].str[3:6]
    df['COUNTRY'] = df['TECHNOLOGY'].str[6:9]
    return df


def generate_interactive_dashboard(df, source_file):
    """Generates an interactive HTML dashboard"""
    print(f"\n📊 Processing: {source_file}")

    # Validate required columns
    required_cols = ['Scenario', 'YEAR', 'TECHNOLOGY',
                     'ProductionByTechnology',
                     'TotalTechnologyAnnualActivityLowerLimit']

    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        print(f"   ❌ Error: Missing required columns: {missing_cols}")
        return None

    # Clean and convert numeric columns
    print(f"   🔄 Cleaning data...")
    df = df.copy()

    numeric_cols = ['YEAR', 'ProductionByTechnology', 'TotalTechnologyAnnualActivityLowerLimit']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df.dropna(subset=['Scenario', 'YEAR', 'TECHNOLOGY'])
    df['ProductionByTechnology'] = df['ProductionByTechnology'].fillna(0)
    df['TotalTechnologyAnnualActivityLowerLimit'] = df['TotalTechnologyAnnualActivityLowerLimit'].fillna(0)

    if df.empty:
        print(f"   ❌ Error: No valid data")
        return None

    print(f"   ✅ Clean data: {len(df):,} valid rows")

    # Filter only valid PWR technologies
    df = filter_pwr_technologies(df)

    if df.empty:
        print(f"   ❌ Error: No valid PWR technologies after filtering")
        return None

    # Extract FUEL and COUNTRY
    df = extract_fuel_country(df)

    # Get unique lists
    scenarios = sorted(df['Scenario'].unique().tolist())
    fuels = sorted(df['FUEL'].unique().tolist())
    countries = sorted(df['COUNTRY'].unique().tolist())
    technologies = sorted(df['TECHNOLOGY'].unique().tolist())
    year_range = f"{df['YEAR'].min():.0f} - {df['YEAR'].max():.0f}"

    # Prepare data for JSON export
    print(f"   🔄 Preparing data for JavaScript...")
    df_export = df[['Scenario', 'YEAR', 'TECHNOLOGY', 'FUEL', 'COUNTRY',
                     'ProductionByTechnology', 'TotalTechnologyAnnualActivityLowerLimit']].copy()

    # Convert to JSON
    data_json = df_export.to_json(orient='records')

    # Generate filename (with Aggregated suffix)
    base_name = os.path.splitext(os.path.basename(source_file))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"Dashboard_Interactive_Aggregated_{base_name}_{timestamp}.html"

    # Create interactive HTML
    print(f"   🔄 Generating interactive HTML...")

    html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard PWR Interactivo - {base_name}</title>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: {COLORS['background']};
            padding: 20px;
            color: {COLORS['text']};
        }}

        .container {{ max-width: 1600px; margin: 0 auto; }}

        h1 {{
            color: {COLORS['primary']};
            text-align: center;
            margin-bottom: 10px;
            font-size: 2.5em;
        }}

        .subtitle {{
            text-align: center;
            color: {COLORS['secondary']};
            margin-bottom: 30px;
            font-size: 1.1em;
        }}

        .card {{
            background-color: {COLORS['card']};
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            border: 1px solid {COLORS['border']};
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}

        .filters-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }}

        .filter-group {{
            display: flex;
            flex-direction: column;
        }}

        .filter-group label {{
            font-weight: bold;
            margin-bottom: 5px;
            font-size: 0.95em;
            color: {COLORS['text']};
        }}

        select {{
            padding: 8px;
            border: 1px solid {COLORS['border']};
            border-radius: 5px;
            font-size: 0.95em;
            background-color: white;
        }}

        select[multiple] {{
            height: 120px;
        }}

        .controls {{
            display: flex;
            gap: 15px;
            align-items: center;
            flex-wrap: wrap;
            margin-top: 20px;
        }}

        .graph-type-group {{
            display: flex;
            gap: 15px;
            align-items: center;
        }}

        .graph-type-group label {{
            font-weight: bold;
            margin-right: 10px;
        }}

        .radio-option {{
            display: flex;
            align-items: center;
            gap: 5px;
        }}

        .btn {{
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 1em;
            font-weight: bold;
            transition: all 0.3s;
        }}

        .btn-primary {{
            background-color: {COLORS['primary']};
            color: white;
        }}

        .btn-primary:hover {{
            background-color: #0052a3;
        }}

        .btn-secondary {{
            background-color: {COLORS['secondary']};
            color: white;
        }}

        .btn-secondary:hover {{
            background-color: #545b62;
        }}

        .filter-info {{
            margin-top: 15px;
            padding: 10px;
            background-color: {COLORS['background']};
            border-radius: 5px;
            font-size: 0.9em;
            color: {COLORS['secondary']};
        }}

        .legend-box {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
            margin-top: 15px;
        }}

        .legend-item {{
            padding: 12px;
            border-radius: 8px;
            font-size: 0.9em;
        }}

        .renewable {{
            background-color: #d4edda;
            color: #155724;
            border-left: 4px solid {COLORS['renewable']};
        }}

        .non-renewable {{
            background-color: #f8d7da;
            color: #721c24;
            border-left: 4px solid {COLORS['non_renewable']};
        }}

        .graph-container {{
            min-height: 500px;
        }}

        .footer {{
            text-align: center;
            color: {COLORS['secondary']};
            font-size: 0.9em;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid {COLORS['border']};
        }}

        @media (max-width: 768px) {{
            .filters-grid {{ grid-template-columns: 1fr; }}
            .legend-box {{ grid-template-columns: 1fr; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Interactive PWR Dashboard - AGGREGATED BY SCENARIO</h1>
        <p class="subtitle">Analysis: {base_name}</p>
        <div style="text-align: center; margin-bottom: 20px; padding: 10px; background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 5px;">
            <p style="margin: 0; color: #856404; font-weight: bold;">
                ⚠️ AGGREGATED VERSION: Graphs show values summed by scenario only.
            </p>
            <p style="margin: 5px 0 0 0; font-size: 0.9em; color: #856404;">
                The Fuel and Country filters determine which technologies to include in the sum.
            </p>
        </div>

        <!-- Card: Information -->
        <div class="card">
            <h3 style="margin-bottom: 15px;">📋 Dataset Information</h3>
            <p><strong>File:</strong> {source_file}</p>
            <p><strong>Year range:</strong> {year_range}</p>
            <p><strong>Total rows:</strong> {len(df):,}</p>
            <p><strong>Unique PWR technologies:</strong> {df['TECHNOLOGY'].nunique():,}</p>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-top: 10px;">
                ℹ️ <strong>PWR technology pattern:</strong> PWR + [FUEL] + [COUNTRY] + XX<br>
                Where FUEL can be: {', '.join(VALID_FUELS)}
            </p>
            <div style="background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 5px; padding: 10px; margin-top: 15px;">
                <p style="margin: 0; color: #856404; font-size: 0.9em;">
                    <strong>⚠️ AGGREGATED VERSION:</strong> In this version, graphs show values <strong>summed by scenario</strong>.
                </p>
                <p style="margin: 5px 0 0 0; color: #856404; font-size: 0.85em;">
                    • FUEL and COUNTRY filters determine which technologies to include in the sum<br>
                    • Technologies are automatically selected based on filters<br>
                    • Ideal for comparing the total of each scenario, not individual technologies
                </p>
            </div>

            <div class="legend-box">
                <div class="legend-item renewable">
                    <strong>🌱 Renewable Fuels</strong><br>
                    BIO, WAS, CSP, SPV, GEO, HYD, WAV, WON, WOF
                </div>
                <div class="legend-item non-renewable">
                    <strong>⚫ Non-Renewable Fuels</strong><br>
                    URN, NGS, COA, COG, OIL, PET, CCS, OTH
                </div>
            </div>
        </div>

        <!-- Card: Filters -->
        <div class="card">
            <h3 style="margin-bottom: 15px;">🔍 Interactive Filters</h3>

            <div class="filters-grid">
                <div class="filter-group">
                    <label for="scenario-filter">Scenarios:</label>
                    <select id="scenario-filter" multiple>
                        <!-- Options loaded by JavaScript -->
                    </select>
                </div>

                <div class="filter-group">
                    <label for="fuel-filter">Fuels (FUEL):</label>
                    <select id="fuel-filter" multiple>
                        <!-- Options loaded by JavaScript -->
                    </select>
                </div>

                <div class="filter-group">
                    <label for="country-filter">Countries (COUNTRY):</label>
                    <select id="country-filter" multiple>
                        <!-- Options loaded by JavaScript -->
                    </select>
                </div>

                <div class="filter-group">
                    <label for="technology-filter">
                        Technologies (auto-selected):
                        <span style="font-size: 0.85em; font-weight: normal; color: {COLORS['secondary']};">
                            ℹ️ Automatically updated based on FUEL and COUNTRY
                        </span>
                    </label>
                    <select id="technology-filter" multiple disabled style="background-color: #f0f0f0; cursor: not-allowed;">
                        <!-- Options loaded and selected automatically by JavaScript -->
                    </select>
                </div>
            </div>

            <div class="controls">
                <div class="graph-type-group">
                    <label>Chart Type:</label>
                    <div class="radio-option">
                        <input type="radio" id="type-line" name="graph-type" value="line" checked>
                        <label for="type-line">📈 Lines</label>
                    </div>
                    <div class="radio-option">
                        <input type="radio" id="type-bar" name="graph-type" value="bar">
                        <label for="type-bar">📊 Bars</label>
                    </div>
                </div>

                <button class="btn btn-primary" onclick="updateAllGraphs()">🔄 Update Charts</button>
                <button class="btn btn-secondary" onclick="resetFilters()">↺ Reset Filters</button>
                <button class="btn btn-secondary" onclick="selectAll()">☑️ Select All</button>
                <button class="btn btn-secondary" onclick="selectRenewable()">🌱 Renewables Only</button>
                <button class="btn btn-secondary" onclick="selectNonRenewable()">⚫ Non-Renewables Only</button>
            </div>

            <div class="filter-info" id="filter-info"></div>
        </div>

        <!-- Charts -->
        <!-- Section: Renewability Shares -->
        <div class="card">
            <h3 style="margin-bottom: 15px;">🌱 Renewability Shares</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Percentage of renewable vs non-renewable generation - Respects all selected filters (Scenarios, Fuels, Countries)
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ When changing FUEL or COUNTRY, the chart updates automatically
            </p>
            <div id="renewability-graph" class="graph-container"></div>
        </div>

        <!-- Section: Total Sum by Technology Charts -->
        <div style="margin: 30px 0; padding: 15px; background: linear-gradient(to right, {COLORS['primary']}, {COLORS['primary']}33); border-radius: 10px;">
            <h2 style="color: white; text-align: center; margin: 0; font-size: 1.5em;">
                📊 TOTAL SUM BY TECHNOLOGY
            </h2>
            <p style="color: white; text-align: center; margin: 5px 0 0 0; font-size: 0.95em; opacity: 0.95;">
                Aggregated values across all years - Top 30 technologies
            </p>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 15px;">📊 Total Sum by Scenario - Lower Limit</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Total sum of TotalTechnologyAnnualActivityLowerLimit <strong>aggregated by scenario</strong> (sum of all selected years, technologies, fuels and countries)
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ Bars show the renewable (green) vs non-renewable (red) composition
            </p>
            <div id="total-lowerlimit-by-tech-graph" class="graph-container"></div>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 15px;">📊 Total Sum by Scenario - Production</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Total sum of ProductionByTechnology <strong>aggregated by scenario</strong> (sum of all selected years, technologies, fuels and countries)
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ Bars show the renewable (green) vs non-renewable (red) composition
            </p>
            <div id="total-production-by-tech-graph" class="graph-container"></div>
        </div>

        <!-- Section: Temporal Evolution Charts -->
        <div style="margin: 30px 0; padding: 15px; background: linear-gradient(to right, {COLORS['secondary']}, {COLORS['secondary']}33); border-radius: 10px;">
            <h2 style="color: white; text-align: center; margin: 0; font-size: 1.5em;">
                📈 TEMPORAL EVOLUTION
            </h2>
            <p style="color: white; text-align: center; margin: 5px 0 0 0; font-size: 0.95em; opacity: 0.95;">
                Values by year - Line or bar charts
            </p>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 15px;">📉 Total Technology Annual Activity Lower Limit (By Year - Aggregated by Scenario)</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Values aggregated by scenario - Sum of all selected technologies, fuels and countries
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ Tooltips show the renewable/non-renewable breakdown
            </p>
            <div id="lower-limit-graph" class="graph-container"></div>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 15px;">⚡ Production By Technology (By Year - Aggregated by Scenario)</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Values aggregated by scenario - Sum of all selected technologies, fuels and countries
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ Tooltips show the renewable/non-renewable breakdown
            </p>
            <div id="production-graph" class="graph-container"></div>
        </div>

        <div class="footer">
            <p><strong>Interactive Dashboard - Climate Lead Group | ReLAC-TX Project</strong></p>
            <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p>File: {output_file}</p>
        </div>
    </div>

    <script>
        // ============================================================================
        // DATA AND CONFIGURATION
        // ============================================================================
        const RAW_DATA = {data_json};
        const RENEWABLE_FUELS = {json.dumps(RENEWABLE_FUELS)};
        const NON_RENEWABLE_FUELS = {json.dumps(NON_RENEWABLE_FUELS)};

        const SCENARIOS = {json.dumps(scenarios)};
        const FUELS = {json.dumps(fuels)};
        const COUNTRIES = {json.dumps(countries)};
        const TECHNOLOGIES = {json.dumps(technologies)};

        // ============================================================================
        // INITIALIZATION
        // ============================================================================
        document.addEventListener('DOMContentLoaded', function() {{
            initializeFilters();
            updateAllGraphs();

            // Event listeners for automatic update
            document.querySelectorAll('input[name="graph-type"]').forEach(radio => {{
                radio.addEventListener('change', updateAllGraphs);
            }});

            // Event listeners to automatically update technologies when FUEL or COUNTRY change
            document.getElementById('fuel-filter').addEventListener('change', function() {{
                updateTechnologiesBasedOnFilters();
                updateAllGraphs();
            }});

            document.getElementById('country-filter').addEventListener('change', function() {{
                updateTechnologiesBasedOnFilters();
                updateAllGraphs();
            }});
        }});

        function initializeFilters() {{
            populateSelect('scenario-filter', SCENARIOS, SCENARIOS);
            populateSelect('fuel-filter', FUELS, FUELS);
            populateSelect('country-filter', COUNTRIES, COUNTRIES);

            // Technologies are automatically updated based on FUEL and COUNTRY
            updateTechnologiesBasedOnFilters();
        }}

        function populateSelect(id, options, selected = []) {{
            const select = document.getElementById(id);
            select.innerHTML = '';
            options.forEach(opt => {{
                const option = document.createElement('option');
                option.value = opt;
                option.textContent = opt;
                option.selected = selected.includes(opt);
                select.appendChild(option);
            }});
        }}

        function updateTechnologiesBasedOnFilters() {{
            // Get selected FUEL and COUNTRY
            const selectedFuels = getSelectedValues('fuel-filter');
            const selectedCountries = getSelectedValues('country-filter');

            // Filter technologies that match selected FUEL and COUNTRY
            const matchingTechs = TECHNOLOGIES.filter(tech => {{
                const fuel = tech.substring(3, 6);
                const country = tech.substring(6, 9);

                const fuelMatch = selectedFuels.length === 0 || selectedFuels.includes(fuel);
                const countryMatch = selectedCountries.length === 0 || selectedCountries.includes(country);

                return fuelMatch && countryMatch;
            }});

            // Update the technology selector
            populateSelect('technology-filter', matchingTechs, matchingTechs);

            return matchingTechs;
        }}

        function getSelectedValues(id) {{
            const select = document.getElementById(id);
            return Array.from(select.selectedOptions).map(opt => opt.value);
        }}

        function resetFilters() {{
            initializeFilters();
            updateAllGraphs();
        }}

        function selectAll() {{
            ['scenario-filter', 'fuel-filter', 'country-filter', 'technology-filter'].forEach(id => {{
                const select = document.getElementById(id);
                Array.from(select.options).forEach(opt => opt.selected = true);
            }});
            updateAllGraphs();
        }}

        function selectRenewable() {{
            const fuelSelect = document.getElementById('fuel-filter');
            Array.from(fuelSelect.options).forEach(opt => {{
                opt.selected = RENEWABLE_FUELS.includes(opt.value);
            }});
            updateAllGraphs();
        }}

        function selectNonRenewable() {{
            const fuelSelect = document.getElementById('fuel-filter');
            Array.from(fuelSelect.options).forEach(opt => {{
                opt.selected = NON_RENEWABLE_FUELS.includes(opt.value);
            }});
            updateAllGraphs();
        }}

        // ============================================================================
        // DATA FILTERING
        // ============================================================================
        function getFilteredData() {{
            const scenarios = getSelectedValues('scenario-filter');
            const fuels = getSelectedValues('fuel-filter');
            const countries = getSelectedValues('country-filter');
            const technologies = getSelectedValues('technology-filter');

            let filtered = RAW_DATA.filter(row => {{
                return (scenarios.length === 0 || scenarios.includes(row.Scenario)) &&
                       (fuels.length === 0 || fuels.includes(row.FUEL)) &&
                       (countries.length === 0 || countries.includes(row.COUNTRY)) &&
                       (technologies.length === 0 || technologies.includes(row.TECHNOLOGY));
            }});

            updateFilterInfo(scenarios.length, fuels.length, countries.length, technologies.length);
            return filtered;
        }}

        function updateFilterInfo(nScenarios, nFuels, nCountries, nTechs) {{
            const fuels = getSelectedValues('fuel-filter');
            let fuelTypeInfo = '';

            if (fuels.length > 0) {{
                const renewableCount = fuels.filter(f => RENEWABLE_FUELS.includes(f)).length;
                const nonRenewableCount = fuels.filter(f => NON_RENEWABLE_FUELS.includes(f)).length;
                fuelTypeInfo = ` (🌱 ${{renewableCount}} renewable, ⚫ ${{nonRenewableCount}} non-renewable)`;
            }}

            document.getElementById('filter-info').textContent =
                `📊 Filters applied: ${{nScenarios}} scenarios, ${{nFuels}} fuels${{fuelTypeInfo}}, ${{nCountries}} countries, ${{nTechs}} technologies`;
        }}

        // ============================================================================
        // RENEWABILITY SHARES CALCULATION
        // ============================================================================
        function calculateRenewabilityShares(data) {{
            // Use the same filters as getFilteredData() so it respects FUEL, COUNTRY and TECHNOLOGY
            const scenarios = getSelectedValues('scenario-filter');
            const fuels = getSelectedValues('fuel-filter');
            const countries = getSelectedValues('country-filter');
            const technologies = getSelectedValues('technology-filter');

            let filtered = RAW_DATA.filter(row => {{
                return (scenarios.length === 0 || scenarios.includes(row.Scenario)) &&
                       (fuels.length === 0 || fuels.includes(row.FUEL)) &&
                       (countries.length === 0 || countries.includes(row.COUNTRY)) &&
                       (technologies.length === 0 || technologies.includes(row.TECHNOLOGY));
            }});

            // Group and classify
            const grouped = {{}};
            filtered.forEach(row => {{
                let fuelType = 'Other';
                if (RENEWABLE_FUELS.includes(row.FUEL)) fuelType = 'Renewable';
                else if (NON_RENEWABLE_FUELS.includes(row.FUEL)) fuelType = 'Non-Renewable';

                const key = `${{row.Scenario}}|${{row.YEAR}}|${{fuelType}}`;
                if (!grouped[key]) {{
                    grouped[key] = {{
                        Scenario: row.Scenario,
                        YEAR: row.YEAR,
                        Type: fuelType,
                        Production: 0
                    }};
                }}
                grouped[key].Production += row.ProductionByTechnology || 0;
            }});

            // Calculate totals and shares
            const totals = {{}};
            Object.values(grouped).forEach(item => {{
                const key = `${{item.Scenario}}|${{item.YEAR}}`;
                totals[key] = (totals[key] || 0) + item.Production;
            }});

            const result = Object.values(grouped).map(item => {{
                const key = `${{item.Scenario}}|${{item.YEAR}}`;
                const total = totals[key];
                return {{
                    ...item,
                    Share: total > 0 ? (item.Production / total) * 100 : 0
                }};
            }}).filter(item => item.Type !== 'Other');

            return result;
        }}

        // ============================================================================
        // CHART CREATION
        // ============================================================================
        function updateAllGraphs() {{
            const graphType = document.querySelector('input[name="graph-type"]:checked').value;

            createRenewabilityGraph(graphType);
            createTotalByScenarioGraph('total-lowerlimit-by-tech-graph', 'TotalTechnologyAnnualActivityLowerLimit',
                                        'Total Sum by Scenario - Lower Limit');
            createTotalByScenarioGraph('total-production-by-tech-graph', 'ProductionByTechnology',
                                        'Total Sum by Scenario - Production');
            createMetricGraph('lower-limit-graph', 'TotalTechnologyAnnualActivityLowerLimit',
                            'Total Technology Annual Activity Lower Limit', graphType);
            createMetricGraph('production-graph', 'ProductionByTechnology',
                            'Production By Technology', graphType);
        }}

        function createRenewabilityGraph(graphType) {{
            const shares = calculateRenewabilityShares();
            if (shares.length === 0) {{
                const emptyLayout = {{
                    title: '🌱 Renewability Shares (% based on ProductionByTechnology) - Filters Applied',
                    annotations: [{{
                        text: 'No data to calculate shares with the selected filters',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Year' }},
                    yaxis: {{ title: 'Share (%)' }}
                }};
                Plotly.newPlot('renewability-graph', [], emptyLayout, {{responsive: true}});
                return;
            }}

            const traces = [];
            const scenarios = [...new Set(shares.map(d => d.Scenario))].sort();
            const colorMap = {{
                'Renewable': '#2ecc71',
                'Non-Renewable': '#e74c3c'
            }};

            scenarios.forEach(scenario => {{
                ['Renewable', 'Non-Renewable'].forEach(type => {{
                    const data = shares
                        .filter(d => d.Scenario === scenario && d.Type === type)
                        .sort((a, b) => a.YEAR - b.YEAR);

                    if (data.length === 0) return;

                    const icon = type === 'Renewable' ? '🌱' : '⚫';
                    const color = colorMap[type];

                    const trace = {{
                        x: data.map(d => d.YEAR),
                        y: data.map(d => d.Share),
                        name: `${{icon}} ${{scenario}} - ${{type}}`,
                        type: graphType === 'bar' ? 'bar' : 'scatter',
                        mode: graphType === 'line' ? 'lines+markers' : undefined,
                        marker: {{
                            color: color,
                            size: 10,
                            line: {{ width: 1, color: 'white' }}
                        }},
                        line: graphType === 'line' ? {{
                            color: color,
                            width: 3.5,
                            dash: type === 'Renewable' ? 'solid' : 'dash'
                        }} : undefined,
                        text: data.map(d => `${{d.Share.toFixed(1)}}%`),
                        textposition: graphType === 'bar' ? 'outside' : 'top center',
                        textfont: {{ size: 9 }},
                        hovertemplate: (
                            `<b>${{icon}} ${{scenario}} - ${{type}}</b><br>` +
                            `Year: %{{x}}<br>` +
                            `Share: %{{y:.2f}}%<br>` +
                            `Production: %{{customdata:.2f}} PJ<br>` +
                            `<extra></extra>`
                        ),
                        customdata: data.map(d => d.Production)
                    }};
                    traces.push(trace);
                }});
            }});

            const layout = {{
                title: {{
                    text: '🌱 Renewability Shares (% based on ProductionByTechnology) - Filters Applied',
                    font: {{ size: 18 }}
                }},
                xaxis: {{
                    title: 'Year',
                    dtick: 1,
                    gridcolor: '{COLORS['border']}'
                }},
                yaxis: {{
                    title: 'Share (%)',
                    range: [0, 105],
                    gridcolor: '{COLORS['border']}'
                }},
                hovermode: 'closest',
                plot_bgcolor: 'white',
                paper_bgcolor: 'white',
                barmode: graphType === 'bar' ? 'group' : undefined,
                margin: {{ l: 80, r: 250, t: 80, b: 80 }},
                height: 600,
                legend: {{
                    orientation: 'v',
                    yanchor: 'top',
                    y: 1,
                    xanchor: 'left',
                    x: 1.02,
                    bgcolor: 'rgba(255,255,255,0.9)',
                    bordercolor: '{COLORS['border']}',
                    borderwidth: 1
                }},
                shapes: [{{
                    type: 'line',
                    x0: 0,
                    x1: 1,
                    xref: 'paper',
                    y0: 50,
                    y1: 50,
                    line: {{
                        color: '{COLORS['secondary']}',
                        width: 2,
                        dash: 'dot'
                    }}
                }}],
                annotations: [{{
                    x: 0.02,
                    y: 50,
                    xref: 'paper',
                    yref: 'y',
                    text: '50% reference line',
                    showarrow: false,
                    font: {{ size: 10, color: '{COLORS['secondary']}' }},
                    xanchor: 'left'
                }}]
            }};

            Plotly.newPlot('renewability-graph', traces, layout, {{responsive: true}});
        }}

        function createTotalByScenarioGraph(divId, metric, title) {{
            const data = getFilteredData();
            if (data.length === 0) {{
                const emptyLayout = {{
                    title: title,
                    annotations: [{{
                        text: 'No data for the selected filters',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Total Value' }},
                    yaxis: {{ title: 'Scenario' }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Group ONLY by Scenario, summing EVERYTHING (years, technologies, fuels, countries)
            const grouped = {{}};
            data.forEach(row => {{
                const key = row.Scenario;
                if (!grouped[key]) {{
                    grouped[key] = {{
                        Scenario: row.Scenario,
                        Value: 0,
                        RenewableValue: 0,
                        NonRenewableValue: 0
                    }};
                }}
                const value = row[metric] || 0;
                grouped[key].Value += value;

                // Classify by fuel type
                if (RENEWABLE_FUELS.includes(row.FUEL)) {{
                    grouped[key].RenewableValue += value;
                }} else if (NON_RENEWABLE_FUELS.includes(row.FUEL)) {{
                    grouped[key].NonRenewableValue += value;
                }}
            }});

            const groupedArray = Object.values(grouped);

            // Filter significant values (greater than 0)
            const filteredArray = groupedArray.filter(d => d.Value > 0);

            if (filteredArray.length === 0) {{
                const emptyLayout = {{
                    title: title,
                    annotations: [{{
                        text: 'No values greater than 0 for this metric',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Total Value' }},
                    yaxis: {{ title: 'Scenario' }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Sort by descending value
            filteredArray.sort((a, b) => b.Value - a.Value);

            // Create traces: one bar per scenario with colors based on renewable/non-renewable composition
            const traces = [
                // Trace for renewables
                {{
                    x: filteredArray.map(d => d.RenewableValue),
                    y: filteredArray.map(d => d.Scenario),
                    name: '🌱 Renewable',
                    type: 'bar',
                    orientation: 'h',
                    marker: {{
                        color: '#2ecc71',
                        line: {{ color: 'white', width: 1 }}
                    }},
                    text: filteredArray.map(d => d.RenewableValue > 0 ? d.RenewableValue.toFixed(2) : ''),
                    textposition: 'inside',
                    textfont: {{ size: 10, color: 'white' }},
                    hovertemplate: filteredArray.map(d =>
                        `<b>🌱 Renewable</b><br>` +
                        `Scenario: ${{d.Scenario}}<br>` +
                        `Value: %{{x:.6f}}<br>` +
                        `% of total: ${{((d.RenewableValue / d.Value) * 100).toFixed(1)}}%<br>` +
                        `<extra></extra>`
                    )
                }},
                // Trace for non-renewables
                {{
                    x: filteredArray.map(d => d.NonRenewableValue),
                    y: filteredArray.map(d => d.Scenario),
                    name: '⚫ Non-Renewable',
                    type: 'bar',
                    orientation: 'h',
                    marker: {{
                        color: '#e74c3c',
                        line: {{ color: 'white', width: 1 }}
                    }},
                    text: filteredArray.map(d => d.NonRenewableValue > 0 ? d.NonRenewableValue.toFixed(2) : ''),
                    textposition: 'inside',
                    textfont: {{ size: 10, color: 'white' }},
                    hovertemplate: filteredArray.map(d =>
                        `<b>⚫ Non-Renewable</b><br>` +
                        `Scenario: ${{d.Scenario}}<br>` +
                        `Value: %{{x:.6f}}<br>` +
                        `% of total: ${{((d.NonRenewableValue / d.Value) * 100).toFixed(1)}}%<br>` +
                        `<extra></extra>`
                    )
                }}
            ];

            const layout = {{
                title: {{
                    text: title + ' (Aggregated)',
                    font: {{ size: 18 }}
                }},
                xaxis: {{
                    title: 'Total Value (sum of all years)',
                    gridcolor: '{COLORS['border']}'
                }},
                yaxis: {{
                    title: 'Scenario',
                    automargin: true,
                    tickfont: {{ size: 11 }}
                }},
                hovermode: 'closest',
                plot_bgcolor: 'white',
                paper_bgcolor: 'white',
                barmode: 'stack',
                margin: {{ l: 150, r: 80, t: 80, b: 80 }},
                height: Math.max(400, filteredArray.length * 80),
                legend: {{
                    orientation: 'v',
                    yanchor: 'top',
                    y: 1,
                    xanchor: 'left',
                    x: 1.02,
                    bgcolor: 'rgba(255,255,255,0.9)',
                    bordercolor: '{COLORS['border']}',
                    borderwidth: 1
                }},
                annotations: filteredArray.map((d, idx) => ({{
                    x: d.Value,
                    y: d.Scenario,
                    text: `Total: ${{d.Value.toFixed(2)}}`,
                    xanchor: 'left',
                    xshift: 5,
                    showarrow: false,
                    font: {{ size: 10, color: '{COLORS['text']}', weight: 'bold' }}
                }}))
            }};

            Plotly.newPlot(divId, traces, layout, {{responsive: true}});
        }}

        function createTotalByTechnologyGraph(divId, metric, title) {{
            const data = getFilteredData();
            if (data.length === 0) {{
                const emptyLayout = {{
                    title: title,
                    annotations: [{{
                        text: 'No data for the selected filters',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Total Value' }},
                    yaxis: {{ title: 'Technology' }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Group by Scenario and TECHNOLOGY, summing all years
            const grouped = {{}};
            data.forEach(row => {{
                const key = `${{row.Scenario}}|${{row.TECHNOLOGY}}`;
                if (!grouped[key]) {{
                    grouped[key] = {{
                        Scenario: row.Scenario,
                        TECHNOLOGY: row.TECHNOLOGY,
                        FUEL: row.FUEL,
                        COUNTRY: row.COUNTRY,
                        Value: 0
                    }};
                }}
                grouped[key].Value += row[metric] || 0;
            }});

            const groupedArray = Object.values(grouped);

            // Filter significant values (greater than 0)
            const filteredArray = groupedArray.filter(d => d.Value > 0);

            if (filteredArray.length === 0) {{
                const emptyLayout = {{
                    title: title,
                    annotations: [{{
                        text: 'No values greater than 0 for this metric',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Total Value' }},
                    yaxis: {{ title: 'Technology' }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Sort by descending value
            filteredArray.sort((a, b) => b.Value - a.Value);

            // Limit to top 30 technologies to avoid overloading the chart
            const topN = 30;
            const topTechs = filteredArray.slice(0, topN);

            // Function to get color based on fuel type
            const getColorForFuel = (fuel) => {{
                if (RENEWABLE_FUELS.includes(fuel)) {{
                    return '#2ecc71'; // Green for renewables
                }} else if (NON_RENEWABLE_FUELS.includes(fuel)) {{
                    return '#e74c3c'; // Red for non-renewables
                }} else {{
                    return '#95a5a6'; // Gray for others
                }}
            }};

            const scenarios = [...new Set(topTechs.map(d => d.Scenario))].sort();
            const traces = [];

            scenarios.forEach(scenario => {{
                const scenarioData = topTechs.filter(d => d.Scenario === scenario);

                if (scenarioData.length === 0) return;

                // Create one trace per scenario
                const trace = {{
                    x: scenarioData.map(d => d.Value),
                    y: scenarioData.map(d => `${{d.TECHNOLOGY}}`),
                    name: scenario,
                    type: 'bar',
                    orientation: 'h',
                    marker: {{
                        color: scenarioData.map(d => getColorForFuel(d.FUEL)),
                        line: {{
                            color: 'white',
                            width: 1
                        }}
                    }},
                    text: scenarioData.map(d => {{
                        const fuelType = RENEWABLE_FUELS.includes(d.FUEL) ? '🌱' :
                                       NON_RENEWABLE_FUELS.includes(d.FUEL) ? '⚫' : '❓';
                        return `${{d.Value.toFixed(2)}}`;
                    }}),
                    textposition: 'outside',
                    textfont: {{ size: 9 }},
                    hovertemplate: scenarioData.map(d => {{
                        const fuelType = RENEWABLE_FUELS.includes(d.FUEL) ? '🌱' :
                                       NON_RENEWABLE_FUELS.includes(d.FUEL) ? '⚫' : '❓';
                        return `<b>${{fuelType}} ${{d.TECHNOLOGY}}</b><br>` +
                               `Scenario: ${{scenario}}<br>` +
                               `FUEL: ${{d.FUEL}}<br>` +
                               `Country: ${{d.COUNTRY}}<br>` +
                               `Total Value: %{{x:.6f}}<br>` +
                               `<extra></extra>`;
                    }})
                }};

                traces.push(trace);
            }});

            const layout = {{
                title: {{
                    text: title + ` (Top ${{topN}})`,
                    font: {{ size: 18 }}
                }},
                xaxis: {{
                    title: 'Total Value (sum of all years)',
                    gridcolor: '{COLORS['border']}'
                }},
                yaxis: {{
                    title: 'Technology',
                    automargin: true,
                    tickfont: {{ size: 10 }}
                }},
                hovermode: 'closest',
                plot_bgcolor: 'white',
                paper_bgcolor: 'white',
                barmode: 'group',
                margin: {{ l: 150, r: 80, t: 80, b: 80 }},
                height: Math.max(600, topTechs.length * 25),
                legend: {{
                    orientation: 'v',
                    yanchor: 'top',
                    y: 1,
                    xanchor: 'left',
                    x: 1.02,
                    bgcolor: 'rgba(255,255,255,0.9)',
                    bordercolor: '{COLORS['border']}',
                    borderwidth: 1
                }}
            }};

            Plotly.newPlot(divId, traces, layout, {{responsive: true}});
        }}

        function createMetricGraph(divId, metric, title, graphType) {{
            const data = getFilteredData();
            if (data.length === 0) {{
                const emptyLayout = {{
                    title: title,
                    annotations: [{{
                        text: 'No data for the selected filters',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Year' }},
                    yaxis: {{ title: title }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Group ONLY by Scenario and YEAR (not by FUEL) - AGGREGATED VERSION
            const grouped = {{}};
            data.forEach(row => {{
                const key = `${{row.Scenario}}|${{row.YEAR}}`;
                if (!grouped[key]) {{
                    grouped[key] = {{
                        Scenario: row.Scenario,
                        YEAR: row.YEAR,
                        Value: 0,
                        RenewableValue: 0,
                        NonRenewableValue: 0
                    }};
                }}
                const value = row[metric] || 0;
                grouped[key].Value += value;

                // Classify by fuel type for additional information
                if (RENEWABLE_FUELS.includes(row.FUEL)) {{
                    grouped[key].RenewableValue += value;
                }} else if (NON_RENEWABLE_FUELS.includes(row.FUEL)) {{
                    grouped[key].NonRenewableValue += value;
                }}
            }});

            const groupedArray = Object.values(grouped);

            // Filter significant values (greater than 0)
            const filteredArray = groupedArray.filter(d => d.Value > 0);

            if (filteredArray.length === 0) {{
                const emptyLayout = {{
                    title: title,
                    annotations: [{{
                        text: 'No values greater than 0 for this metric',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Year' }},
                    yaxis: {{ title: title }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            const scenarios = [...new Set(filteredArray.map(d => d.Scenario))].sort();

            // Colors for each scenario (varied color palette)
            const scenarioColors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b',
                                   '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'];

            const traces = [];

            scenarios.forEach((scenario, idx) => {{
                const scenarioData = filteredArray
                    .filter(d => d.Scenario === scenario)
                    .sort((a, b) => a.YEAR - b.YEAR);

                if (scenarioData.length === 0) return;

                const color = scenarioColors[idx % scenarioColors.length];

                const trace = {{
                    x: scenarioData.map(d => d.YEAR),
                    y: scenarioData.map(d => d.Value),
                    name: scenario,
                    type: graphType === 'bar' ? 'bar' : 'scatter',
                    mode: graphType === 'line' ? 'lines+markers' : undefined,
                    marker: {{
                        color: color,
                        size: 10,
                        line: {{ width: 1, color: 'white' }}
                    }},
                    line: graphType === 'line' ? {{
                        width: 3,
                        color: color
                    }} : undefined,
                    text: scenarioData.map(d => d.Value.toFixed(2)),
                    textposition: graphType === 'bar' ? 'outside' : 'top center',
                    textfont: {{ size: 9 }},
                    hovertemplate: scenarioData.map(d =>
                        `<b>${{scenario}}</b><br>` +
                        `Year: %{{x}}<br>` +
                        `${{title}}: %{{y:.6f}}<br>` +
                        `🌱 Renewable: ${{d.RenewableValue.toFixed(6)}} (${{((d.RenewableValue / d.Value) * 100).toFixed(1)}}%)<br>` +
                        `⚫ Non-Renewable: ${{d.NonRenewableValue.toFixed(6)}} (${{((d.NonRenewableValue / d.Value) * 100).toFixed(1)}}%)<br>` +
                        `<extra></extra>`
                    ),
                    customdata: scenarioData.map(d => [d.RenewableValue, d.NonRenewableValue])
                }};
                traces.push(trace);
            }});

            const layout = {{
                title: {{
                    text: title + ' (Aggregated by Scenario)',
                    font: {{ size: 18 }}
                }},
                xaxis: {{
                    title: 'Year',
                    dtick: 1,
                    gridcolor: '{COLORS['border']}'
                }},
                yaxis: {{
                    title: title,
                    gridcolor: '{COLORS['border']}'
                }},
                hovermode: 'closest',
                plot_bgcolor: 'white',
                paper_bgcolor: 'white',
                barmode: graphType === 'bar' ? 'group' : undefined,
                margin: {{ l: 80, r: 250, t: 80, b: 80 }},
                height: 600,
                legend: {{
                    orientation: 'v',
                    yanchor: 'top',
                    y: 1,
                    xanchor: 'left',
                    x: 1.02,
                    bgcolor: 'rgba(255,255,255,0.9)',
                    bordercolor: '{COLORS['border']}',
                    borderwidth: 1
                }}
            }};

            Plotly.newPlot(divId, traces, layout, {{responsive: true}});
        }}
    </script>
</body>
</html>"""

    # Save file
    print(f"   💾 Saving dashboard: {output_file}")
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"   ✅ Interactive dashboard generated successfully!")
    return output_file


def find_csv_files():
    """Searches for CSV files in the current directory"""
    csv_files = glob.glob("*.csv")
    return sorted(csv_files)


def select_files_interactive():
    """Allows the user to select files interactively"""
    csv_files = find_csv_files()

    if not csv_files:
        print("\n⚠️  No CSV files found in the current directory.")
        return []

    print("\n📁 Available CSV files:")
    print("-" * 70)
    for idx, file in enumerate(csv_files, 1):
        size_mb = os.path.getsize(file) / (1024 * 1024)
        print(f"   {idx}. {file} ({size_mb:.2f} MB)")
    print("-" * 70)

    print("\n💡 Options:")
    print("   • Enter numbers separated by commas (e.g.: 1,3,5)")
    print("   • Enter 'all' to process all files")
    print("   • Enter 'q' to cancel")

    while True:
        selection = input("\n👉 Your selection: ").strip().lower()

        if selection == 'q':
            return []

        if selection == 'all':
            return csv_files

        try:
            indices = [int(x.strip()) for x in selection.split(',')]
            if all(1 <= idx <= len(csv_files) for idx in indices):
                selected_files = [csv_files[idx - 1] for idx in indices]
                return selected_files
            else:
                print(f"   ❌ Error: Numbers must be between 1 and {len(csv_files)}")
        except ValueError:
            print("   ❌ Error: Invalid format. Use numbers separated by commas.")


def main():
    """Main function"""
    print("=" * 70)
    print("  HTML DASHBOARD GENERATOR - PWR TECHNOLOGIES - AGGREGATED VERSION")
    print("=" * 70)
    print("\n⚠️  AGGREGATED VERSION: Graphs show values summed BY SCENARIO")
    print("\nThis script generates interactive HTML dashboards with:")
    print("  • ✅ Automatic filtering of valid PWR technologies")
    print(f"  • ✅ Pattern: {TECH_PATTERN}")
    print("  • ✅ Filters by Scenario, FUEL and Country")
    print("  • ✅ Auto-selection of technologies based on FUEL and COUNTRY")
    print("  • ✅ Quick buttons for Renewables/Non-Renewables")
    print("  • ✅ 5 charts AGGREGATED BY SCENARIO:")
    print("      - Renewability Shares")
    print("      - Total Sum by Scenario (Lower Limit)")
    print("      - Total Sum by Scenario (Production)")
    print("      - Temporal Evolution by Scenario (Lower Limit)")
    print("      - Temporal Evolution by Scenario (Production)")
    print("  • ✅ Renewable/non-renewable breakdown in charts and tooltips")
    print("  • ✅ Switch between line and bar charts")
    print("  • ✅ Completely standalone (no server required)")

    selected_files = select_files_interactive()

    if not selected_files:
        print("\n❌ No files selected. Exiting...")
        return

    print(f"\n🚀 Processing {len(selected_files)} file(s)...")
    print("=" * 70)

    generated_files = []
    errors = []

    for file in selected_files:
        try:
            print(f"\n📂 Loading file: {file}")
            df = pd.read_csv(file, low_memory=False)
            print(f"   ✅ File loaded: {len(df):,} rows")

            output_file = generate_interactive_dashboard(df, file)

            if output_file:
                generated_files.append(output_file)
            else:
                errors.append((file, "Could not generate the dashboard"))

        except FileNotFoundError:
            print(f"   ❌ Error: File not found: {file}")
            errors.append((file, "File not found"))

        except pd.errors.EmptyDataError:
            print(f"   ❌ Error: The file is empty: {file}")
            errors.append((file, "Empty file"))

        except Exception as e:
            print(f"   ❌ Unexpected error processing {file}")
            print(f"      Details: {type(e).__name__}: {str(e)}")
            errors.append((file, f"{type(e).__name__}: {str(e)}"))

    # Summary
    print("\n" + "=" * 70)
    print("  SUMMARY")
    print("=" * 70)
    print(f"\n✅ Interactive dashboards generated: {len(generated_files)}")

    if generated_files:
        print("\n📄 Generated files:")
        for file in generated_files:
            print(f"   • {file}")

    if errors:
        print(f"\n❌ Errors: {len(errors)}")
        for file, error in errors:
            print(f"   • {file}: {error}")

    print("\n💡 Features of the aggregated dashboards (AGGREGATED VERSION):")
    print("   • ⚠️  AGGREGATED BY SCENARIO: Charts sum all values by scenario")
    print("   • ✅ Automatic filtering of valid PWR technologies")
    print(f"   • ✅ Validation pattern: {TECH_PATTERN}")
    print("   • ✅ Auto-selection of technologies based on FUEL and COUNTRY filters")
    print("   • ✅ Interactive filters by Scenario, FUEL and Country")
    print("   • ✅ Quick buttons: Renewables / Non-Renewables")
    print("   • ✅ 5 charts AGGREGATED BY SCENARIO:")
    print("      - 🌱 Renewability Shares (with 50% reference line)")
    print("      - 📊 Total Sum by Scenario - Lower Limit (stacked bars renewable/non-renewable)")
    print("      - 📊 Total Sum by Scenario - Production (stacked bars renewable/non-renewable)")
    print("      - 📉 Temporal Evolution by Scenario - Lower Limit (one line per scenario)")
    print("      - ⚡ Temporal Evolution by Scenario - Production (one line per scenario)")
    print("   • ✅ Renewable/non-renewable breakdown in tooltips")
    print("   • ✅ Visual classification with colors")
    print("   • ✅ Dynamic switch between line and bar charts")
    print("   • ✅ Real-time updates without page reload")
    print("\n📌 Next steps:")
    print("   1. Open the HTML files in your browser")
    print("   2. Use the FUEL and COUNTRY filters to select what to include")
    print("   3. Technologies are automatically updated based on filters")
    print("   4. Charts show values SUMMED by scenario")
    print("   5. Switch between lines and bars as preferred")
    print("   6. Compare multiple scenarios in the same chart")
    print("\n💡 Key difference:")
    print("   • STANDARD VERSION: Shows each technology/fuel separately")
    print("   • AGGREGATED VERSION: Sums everything by scenario (this version)")
    print("\n" + "=" * 70)


if __name__ == '__main__':
    main()
