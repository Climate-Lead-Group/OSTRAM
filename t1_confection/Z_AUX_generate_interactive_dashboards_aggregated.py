"""
Generador de Dashboards HTML Interactivos para Tecnologías PWR - VERSIÓN AGREGADA

Este script genera dashboards HTML con gráficos AGREGADOS POR ESCENARIO:

DIFERENCIA CLAVE con la versión estándar:
- Los gráficos muestran valores SUMADOS por escenario únicamente
- Si seleccionas 5 países y 6 combustibles, todos sus valores se suman
- La selección de combustibles y países actualiza automáticamente las tecnologías incluidas

Filtros Interactivos:
- Selección de Escenarios
- Selección de Combustibles (FUEL) → Filtra tecnologías automáticamente
- Selección de Países (COUNTRY) → Filtra tecnologías automáticamente
- Vista de Tecnologías (auto-seleccionadas, no editable)
- Botones rápidos: Renovables / No Renovables / Seleccionar Todo / Resetear
- Cambio entre gráficos de líneas y barras

Gráficos Incluidos (TODOS AGREGADOS POR ESCENARIO):
1. Shares de Renovabilidad (% basado en ProductionByTechnology)
2. Suma Total por Escenario - Lower Limit (agregado de todos los años)
3. Suma Total por Escenario - Production (agregado de todos los años)
4. Evolución Temporal - Lower Limit (por año, sumado por escenario)
5. Evolución Temporal - Production (por año, sumado por escenario)

Características:
- Filtrado automático de tecnologías PWR válidas (patrón regex)
- Auto-selección de tecnologías basada en FUEL y COUNTRY
- Agregación automática por escenario
- Clasificación visual con colores (Verde=Renovable, Rojo=No Renovable)
- Completamente standalone (no requieren servidor)
- Actualización en tiempo real sin recargar la página

Autor: Climate Lead Group
Fecha: 2026-01-22
"""

import pandas as pd
import json
from datetime import datetime
import os
import glob
import re

# Clasificación de combustibles
RENEWABLE_FUELS = ['BIO', 'WAS', 'CSP', 'SPV', 'GEO', 'HYD', 'WAV', 'WON', 'WOF']
NON_RENEWABLE_FUELS = ['URN', 'NGS', 'COA', 'COG', 'OIL', 'PET', 'CCS', 'OTH']

# Todos los combustibles válidos para tecnologías PWR
VALID_FUELS = RENEWABLE_FUELS + NON_RENEWABLE_FUELS

# Patrón regex para tecnologías PWR válidas: PWR + FUEL(3 letras) + COUNTRY(3 letras) + XX
TECH_PATTERN = r'^PWR(' + '|'.join(VALID_FUELS) + r')[A-Z]{3}XX$'

# Colores
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
    Valida si una tecnología cumple con el patrón PWR esperado

    Args:
        tech: Código de tecnología (ej: 'PWRBIOARGXX')

    Returns:
        True si cumple el patrón, False en caso contrario
    """
    if pd.isna(tech):
        return False
    return bool(re.match(TECH_PATTERN, str(tech)))


def filter_pwr_technologies(df):
    """
    Filtra solo las tecnologías PWR válidas según el patrón especificado

    Args:
        df: DataFrame con columna TECHNOLOGY

    Returns:
        DataFrame filtrado con solo tecnologías PWR válidas
    """
    if 'TECHNOLOGY' not in df.columns:
        return df

    rows_before = len(df)

    # Aplicar filtro de tecnologías válidas
    df_filtered = df[df['TECHNOLOGY'].apply(is_valid_pwr_technology)].copy()

    rows_after = len(df_filtered)
    rows_removed = rows_before - rows_after

    if rows_removed > 0:
        print(f"   🔍 Filtrado de tecnologías PWR:")
        print(f"      - Filas antes: {rows_before:,}")
        print(f"      - Filas después: {rows_after:,}")
        print(f"      - Filas eliminadas: {rows_removed:,}")
        print(f"      - Tecnologías únicas válidas: {df_filtered['TECHNOLOGY'].nunique():,}")

    return df_filtered


def extract_fuel_country(df):
    """
    Extrae FUEL y COUNTRY desde la columna TECHNOLOGY

    Formato esperado: PWR[FUEL][COUNTRY]XX
    - FUEL: posiciones 3-5 (caracteres 3, 4, 5)
    - COUNTRY: posiciones 6-8 (caracteres 6, 7, 8)
    """
    df = df.copy()
    df['FUEL'] = df['TECHNOLOGY'].str[3:6]
    df['COUNTRY'] = df['TECHNOLOGY'].str[6:9]
    return df


def generate_interactive_dashboard(df, source_file):
    """Genera un dashboard HTML interactivo"""
    print(f"\n📊 Procesando: {source_file}")

    # Validar columnas requeridas
    required_cols = ['Scenario', 'YEAR', 'TECHNOLOGY',
                     'ProductionByTechnology',
                     'TotalTechnologyAnnualActivityLowerLimit']

    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        print(f"   ❌ Error: Faltan columnas requeridas: {missing_cols}")
        return None

    # Limpiar y convertir columnas numéricas
    print(f"   🔄 Limpiando datos...")
    df = df.copy()

    numeric_cols = ['YEAR', 'ProductionByTechnology', 'TotalTechnologyAnnualActivityLowerLimit']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df.dropna(subset=['Scenario', 'YEAR', 'TECHNOLOGY'])
    df['ProductionByTechnology'] = df['ProductionByTechnology'].fillna(0)
    df['TotalTechnologyAnnualActivityLowerLimit'] = df['TotalTechnologyAnnualActivityLowerLimit'].fillna(0)

    if df.empty:
        print(f"   ❌ Error: No hay datos válidos")
        return None

    print(f"   ✅ Datos limpios: {len(df):,} filas válidas")

    # Filtrar solo tecnologías PWR válidas
    df = filter_pwr_technologies(df)

    if df.empty:
        print(f"   ❌ Error: No hay tecnologías PWR válidas después del filtrado")
        return None

    # Extraer FUEL y COUNTRY
    df = extract_fuel_country(df)

    # Obtener listas únicas
    scenarios = sorted(df['Scenario'].unique().tolist())
    fuels = sorted(df['FUEL'].unique().tolist())
    countries = sorted(df['COUNTRY'].unique().tolist())
    technologies = sorted(df['TECHNOLOGY'].unique().tolist())
    year_range = f"{df['YEAR'].min():.0f} - {df['YEAR'].max():.0f}"

    # Preparar datos para exportar a JSON
    print(f"   🔄 Preparando datos para JavaScript...")
    df_export = df[['Scenario', 'YEAR', 'TECHNOLOGY', 'FUEL', 'COUNTRY',
                     'ProductionByTechnology', 'TotalTechnologyAnnualActivityLowerLimit']].copy()

    # Convertir a JSON
    data_json = df_export.to_json(orient='records')

    # Generar nombre de archivo (con sufijo Aggregated)
    base_name = os.path.splitext(os.path.basename(source_file))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"Dashboard_Interactive_Aggregated_{base_name}_{timestamp}.html"

    # Crear HTML interactivo
    print(f"   🔄 Generando HTML interactivo...")

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
        <h1>📊 Dashboard PWR Interactivo - AGREGADO POR ESCENARIO</h1>
        <p class="subtitle">Análisis: {base_name}</p>
        <div style="text-align: center; margin-bottom: 20px; padding: 10px; background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 5px;">
            <p style="margin: 0; color: #856404; font-weight: bold;">
                ⚠️ VERSIÓN AGREGADA: Los gráficos muestran valores sumados por escenario únicamente.
            </p>
            <p style="margin: 5px 0 0 0; font-size: 0.9em; color: #856404;">
                Los filtros de Combustibles y Países determinan qué tecnologías incluir en la suma.
            </p>
        </div>

        <!-- Card: Información -->
        <div class="card">
            <h3 style="margin-bottom: 15px;">📋 Información del Dataset</h3>
            <p><strong>Archivo:</strong> {source_file}</p>
            <p><strong>Rango de años:</strong> {year_range}</p>
            <p><strong>Total de filas:</strong> {len(df):,}</p>
            <p><strong>Tecnologías PWR únicas:</strong> {df['TECHNOLOGY'].nunique():,}</p>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-top: 10px;">
                ℹ️ <strong>Patrón de tecnologías PWR:</strong> PWR + [FUEL] + [COUNTRY] + XX<br>
                Donde FUEL puede ser: {', '.join(VALID_FUELS)}
            </p>
            <div style="background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 5px; padding: 10px; margin-top: 15px;">
                <p style="margin: 0; color: #856404; font-size: 0.9em;">
                    <strong>⚠️ VERSIÓN AGREGADA:</strong> En esta versión, los gráficos muestran valores <strong>sumados por escenario</strong>.
                </p>
                <p style="margin: 5px 0 0 0; color: #856404; font-size: 0.85em;">
                    • Los filtros de FUEL y COUNTRY determinan qué tecnologías incluir en la suma<br>
                    • Las tecnologías se seleccionan automáticamente según los filtros<br>
                    • Ideal para comparar el total de cada escenario, no tecnologías individuales
                </p>
            </div>

            <div class="legend-box">
                <div class="legend-item renewable">
                    <strong>🌱 Combustibles Renovables</strong><br>
                    BIO, WAS, CSP, SPV, GEO, HYD, WAV, WON, WOF
                </div>
                <div class="legend-item non-renewable">
                    <strong>⚫ Combustibles No Renovables</strong><br>
                    URN, NGS, COA, COG, OIL, PET, CCS, OTH
                </div>
            </div>
        </div>

        <!-- Card: Filtros -->
        <div class="card">
            <h3 style="margin-bottom: 15px;">🔍 Filtros Interactivos</h3>

            <div class="filters-grid">
                <div class="filter-group">
                    <label for="scenario-filter">Escenarios:</label>
                    <select id="scenario-filter" multiple>
                        <!-- Opciones cargadas por JavaScript -->
                    </select>
                </div>

                <div class="filter-group">
                    <label for="fuel-filter">Combustibles (FUEL):</label>
                    <select id="fuel-filter" multiple>
                        <!-- Opciones cargadas por JavaScript -->
                    </select>
                </div>

                <div class="filter-group">
                    <label for="country-filter">Países (COUNTRY):</label>
                    <select id="country-filter" multiple>
                        <!-- Opciones cargadas por JavaScript -->
                    </select>
                </div>

                <div class="filter-group">
                    <label for="technology-filter">
                        Tecnologías (auto-seleccionadas):
                        <span style="font-size: 0.85em; font-weight: normal; color: {COLORS['secondary']};">
                            ℹ️ Se actualizan automáticamente según FUEL y COUNTRY
                        </span>
                    </label>
                    <select id="technology-filter" multiple disabled style="background-color: #f0f0f0; cursor: not-allowed;">
                        <!-- Opciones cargadas y seleccionadas automáticamente por JavaScript -->
                    </select>
                </div>
            </div>

            <div class="controls">
                <div class="graph-type-group">
                    <label>Tipo de Gráfico:</label>
                    <div class="radio-option">
                        <input type="radio" id="type-line" name="graph-type" value="line" checked>
                        <label for="type-line">📈 Líneas</label>
                    </div>
                    <div class="radio-option">
                        <input type="radio" id="type-bar" name="graph-type" value="bar">
                        <label for="type-bar">📊 Barras</label>
                    </div>
                </div>

                <button class="btn btn-primary" onclick="updateAllGraphs()">🔄 Actualizar Gráficos</button>
                <button class="btn btn-secondary" onclick="resetFilters()">↺ Resetear Filtros</button>
                <button class="btn btn-secondary" onclick="selectAll()">☑️ Seleccionar Todo</button>
                <button class="btn btn-secondary" onclick="selectRenewable()">🌱 Solo Renovables</button>
                <button class="btn btn-secondary" onclick="selectNonRenewable()">⚫ Solo No Renovables</button>
            </div>

            <div class="filter-info" id="filter-info"></div>
        </div>

        <!-- Gráficos -->
        <!-- Sección: Shares de Renovabilidad -->
        <div class="card">
            <h3 style="margin-bottom: 15px;">🌱 Shares de Renovabilidad</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Porcentaje de generación renovable vs no renovable - Respeta todos los filtros seleccionados (Escenarios, Combustibles, Países)
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ Al cambiar FUEL o COUNTRY, el gráfico se actualiza automáticamente
            </p>
            <div id="renewability-graph" class="graph-container"></div>
        </div>

        <!-- Sección: Gráficos de Suma Total por Tecnología -->
        <div style="margin: 30px 0; padding: 15px; background: linear-gradient(to right, {COLORS['primary']}, {COLORS['primary']}33); border-radius: 10px;">
            <h2 style="color: white; text-align: center; margin: 0; font-size: 1.5em;">
                📊 SUMA TOTAL POR TECNOLOGÍA
            </h2>
            <p style="color: white; text-align: center; margin: 5px 0 0 0; font-size: 0.95em; opacity: 0.95;">
                Valores agregados de todos los años - Top 30 tecnologías
            </p>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 15px;">📊 Suma Total por Escenario - Lower Limit</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Suma total de TotalTechnologyAnnualActivityLowerLimit <strong>agregado por escenario</strong> (suma de todos los años, tecnologías, combustibles y países seleccionados)
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ Las barras muestran la composición renovable (verde) vs no renovable (rojo)
            </p>
            <div id="total-lowerlimit-by-tech-graph" class="graph-container"></div>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 15px;">📊 Suma Total por Escenario - Production</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Suma total de ProductionByTechnology <strong>agregado por escenario</strong> (suma de todos los años, tecnologías, combustibles y países seleccionados)
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ Las barras muestran la composición renovable (verde) vs no renovable (rojo)
            </p>
            <div id="total-production-by-tech-graph" class="graph-container"></div>
        </div>

        <!-- Sección: Gráficos de Evolución Temporal -->
        <div style="margin: 30px 0; padding: 15px; background: linear-gradient(to right, {COLORS['secondary']}, {COLORS['secondary']}33); border-radius: 10px;">
            <h2 style="color: white; text-align: center; margin: 0; font-size: 1.5em;">
                📈 EVOLUCIÓN TEMPORAL
            </h2>
            <p style="color: white; text-align: center; margin: 5px 0 0 0; font-size: 0.95em; opacity: 0.95;">
                Valores por año - Gráficos de líneas o barras
            </p>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 15px;">📉 Total Technology Annual Activity Lower Limit (Por Año - Agregado por Escenario)</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Valores agregados por escenario - Suma de todas las tecnologías, combustibles y países seleccionados
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ Los tooltips muestran el desglose renovable/no renovable
            </p>
            <div id="lower-limit-graph" class="graph-container"></div>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 15px;">⚡ Production By Technology (Por Año - Agregado por Escenario)</h3>
            <p style="font-size: 0.9em; color: {COLORS['secondary']}; margin-bottom: 10px;">
                Valores agregados por escenario - Suma de todas las tecnologías, combustibles y países seleccionados
            </p>
            <p style="font-size: 0.85em; color: {COLORS['primary']}; margin-bottom: 10px;">
                ℹ️ Los tooltips muestran el desglose renovable/no renovable
            </p>
            <div id="production-graph" class="graph-container"></div>
        </div>

        <div class="footer">
            <p><strong>Dashboard Interactivo - Climate Lead Group | Proyecto ReLAC-TX</strong></p>
            <p>Generado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p>Archivo: {output_file}</p>
        </div>
    </div>

    <script>
        // ============================================================================
        // DATOS Y CONFIGURACIÓN
        // ============================================================================
        const RAW_DATA = {data_json};
        const RENEWABLE_FUELS = {json.dumps(RENEWABLE_FUELS)};
        const NON_RENEWABLE_FUELS = {json.dumps(NON_RENEWABLE_FUELS)};

        const SCENARIOS = {json.dumps(scenarios)};
        const FUELS = {json.dumps(fuels)};
        const COUNTRIES = {json.dumps(countries)};
        const TECHNOLOGIES = {json.dumps(technologies)};

        // ============================================================================
        // INICIALIZACIÓN
        // ============================================================================
        document.addEventListener('DOMContentLoaded', function() {{
            initializeFilters();
            updateAllGraphs();

            // Event listeners para actualización automática
            document.querySelectorAll('input[name="graph-type"]').forEach(radio => {{
                radio.addEventListener('change', updateAllGraphs);
            }});

            // Event listeners para actualizar tecnologías automáticamente cuando cambien FUEL o COUNTRY
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

            // Las tecnologías se actualizan automáticamente basadas en FUEL y COUNTRY
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
            // Obtener FUEL y COUNTRY seleccionados
            const selectedFuels = getSelectedValues('fuel-filter');
            const selectedCountries = getSelectedValues('country-filter');

            // Filtrar tecnologías que coincidan con FUEL y COUNTRY seleccionados
            const matchingTechs = TECHNOLOGIES.filter(tech => {{
                const fuel = tech.substring(3, 6);
                const country = tech.substring(6, 9);

                const fuelMatch = selectedFuels.length === 0 || selectedFuels.includes(fuel);
                const countryMatch = selectedCountries.length === 0 || selectedCountries.includes(country);

                return fuelMatch && countryMatch;
            }});

            // Actualizar el selector de tecnologías
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
        // FILTRADO DE DATOS
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
                fuelTypeInfo = ` (🌱 ${{renewableCount}} renovables, ⚫ ${{nonRenewableCount}} no renovables)`;
            }}

            document.getElementById('filter-info').textContent =
                `📊 Filtros aplicados: ${{nScenarios}} escenarios, ${{nFuels}} combustibles${{fuelTypeInfo}}, ${{nCountries}} países, ${{nTechs}} tecnologías`;
        }}

        // ============================================================================
        // CÁLCULO DE SHARES DE RENOVABILIDAD
        // ============================================================================
        function calculateRenewabilityShares(data) {{
            // Usar los mismos filtros que getFilteredData() para que respete FUEL, COUNTRY y TECHNOLOGY
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

            // Agrupar y clasificar
            const grouped = {{}};
            filtered.forEach(row => {{
                let fuelType = 'Otro';
                if (RENEWABLE_FUELS.includes(row.FUEL)) fuelType = 'Renovable';
                else if (NON_RENEWABLE_FUELS.includes(row.FUEL)) fuelType = 'No Renovable';

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

            // Calcular totales y shares
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
            }}).filter(item => item.Type !== 'Otro');

            return result;
        }}

        // ============================================================================
        // CREACIÓN DE GRÁFICOS
        // ============================================================================
        function updateAllGraphs() {{
            const graphType = document.querySelector('input[name="graph-type"]:checked').value;

            createRenewabilityGraph(graphType);
            createTotalByScenarioGraph('total-lowerlimit-by-tech-graph', 'TotalTechnologyAnnualActivityLowerLimit',
                                        'Suma Total por Escenario - Lower Limit');
            createTotalByScenarioGraph('total-production-by-tech-graph', 'ProductionByTechnology',
                                        'Suma Total por Escenario - Production');
            createMetricGraph('lower-limit-graph', 'TotalTechnologyAnnualActivityLowerLimit',
                            'Total Technology Annual Activity Lower Limit', graphType);
            createMetricGraph('production-graph', 'ProductionByTechnology',
                            'Production By Technology', graphType);
        }}

        function createRenewabilityGraph(graphType) {{
            const shares = calculateRenewabilityShares();
            if (shares.length === 0) {{
                const emptyLayout = {{
                    title: '🌱 Shares de Renovabilidad (% basado en ProductionByTechnology) - Filtros Aplicados',
                    annotations: [{{
                        text: 'No hay datos para calcular shares con los filtros seleccionados',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Año' }},
                    yaxis: {{ title: 'Share (%)' }}
                }};
                Plotly.newPlot('renewability-graph', [], emptyLayout, {{responsive: true}});
                return;
            }}

            const traces = [];
            const scenarios = [...new Set(shares.map(d => d.Scenario))].sort();
            const colorMap = {{
                'Renovable': '#2ecc71',
                'No Renovable': '#e74c3c'
            }};

            scenarios.forEach(scenario => {{
                ['Renovable', 'No Renovable'].forEach(type => {{
                    const data = shares
                        .filter(d => d.Scenario === scenario && d.Type === type)
                        .sort((a, b) => a.YEAR - b.YEAR);

                    if (data.length === 0) return;

                    const icon = type === 'Renovable' ? '🌱' : '⚫';
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
                            dash: type === 'Renovable' ? 'solid' : 'dash'
                        }} : undefined,
                        text: data.map(d => `${{d.Share.toFixed(1)}}%`),
                        textposition: graphType === 'bar' ? 'outside' : 'top center',
                        textfont: {{ size: 9 }},
                        hovertemplate: (
                            `<b>${{icon}} ${{scenario}} - ${{type}}</b><br>` +
                            `Año: %{{x}}<br>` +
                            `Share: %{{y:.2f}}%<br>` +
                            `Producción: %{{customdata:.2f}} PJ<br>` +
                            `<extra></extra>`
                        ),
                        customdata: data.map(d => d.Production)
                    }};
                    traces.push(trace);
                }});
            }});

            const layout = {{
                title: {{
                    text: '🌱 Shares de Renovabilidad (% basado en ProductionByTechnology) - Filtros Aplicados',
                    font: {{ size: 18 }}
                }},
                xaxis: {{
                    title: 'Año',
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
                    text: '50% línea de referencia',
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
                        text: 'No hay datos para los filtros seleccionados',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Valor Total' }},
                    yaxis: {{ title: 'Escenario' }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Agrupar SOLO por Scenario, sumando TODO (años, tecnologías, combustibles, países)
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

                // Clasificar por tipo de combustible
                if (RENEWABLE_FUELS.includes(row.FUEL)) {{
                    grouped[key].RenewableValue += value;
                }} else if (NON_RENEWABLE_FUELS.includes(row.FUEL)) {{
                    grouped[key].NonRenewableValue += value;
                }}
            }});

            const groupedArray = Object.values(grouped);

            // Filtrar valores significativos (mayores a 0)
            const filteredArray = groupedArray.filter(d => d.Value > 0);

            if (filteredArray.length === 0) {{
                const emptyLayout = {{
                    title: title,
                    annotations: [{{
                        text: 'No hay valores mayores a 0 para esta métrica',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Valor Total' }},
                    yaxis: {{ title: 'Escenario' }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Ordenar por valor descendente
            filteredArray.sort((a, b) => b.Value - a.Value);

            // Crear trazas: una barra por escenario con colores según composición renovable/no renovable
            const traces = [
                // Traza para renovables
                {{
                    x: filteredArray.map(d => d.RenewableValue),
                    y: filteredArray.map(d => d.Scenario),
                    name: '🌱 Renovable',
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
                        `<b>🌱 Renovable</b><br>` +
                        `Escenario: ${{d.Scenario}}<br>` +
                        `Valor: %{{x:.6f}}<br>` +
                        `% del total: ${{((d.RenewableValue / d.Value) * 100).toFixed(1)}}%<br>` +
                        `<extra></extra>`
                    )
                }},
                // Traza para no renovables
                {{
                    x: filteredArray.map(d => d.NonRenewableValue),
                    y: filteredArray.map(d => d.Scenario),
                    name: '⚫ No Renovable',
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
                        `<b>⚫ No Renovable</b><br>` +
                        `Escenario: ${{d.Scenario}}<br>` +
                        `Valor: %{{x:.6f}}<br>` +
                        `% del total: ${{((d.NonRenewableValue / d.Value) * 100).toFixed(1)}}%<br>` +
                        `<extra></extra>`
                    )
                }}
            ];

            const layout = {{
                title: {{
                    text: title + ' (Agregado)',
                    font: {{ size: 18 }}
                }},
                xaxis: {{
                    title: 'Valor Total (suma de todos los años)',
                    gridcolor: '{COLORS['border']}'
                }},
                yaxis: {{
                    title: 'Escenario',
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
                        text: 'No hay datos para los filtros seleccionados',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Valor Total' }},
                    yaxis: {{ title: 'Tecnología' }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Agrupar por Scenario y TECHNOLOGY, sumando todos los años
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

            // Filtrar valores significativos (mayores a 0)
            const filteredArray = groupedArray.filter(d => d.Value > 0);

            if (filteredArray.length === 0) {{
                const emptyLayout = {{
                    title: title,
                    annotations: [{{
                        text: 'No hay valores mayores a 0 para esta métrica',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Valor Total' }},
                    yaxis: {{ title: 'Tecnología' }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Ordenar por valor descendente
            filteredArray.sort((a, b) => b.Value - a.Value);

            // Limitar a las top 30 tecnologías para no sobrecargar el gráfico
            const topN = 30;
            const topTechs = filteredArray.slice(0, topN);

            // Función para obtener color según tipo de combustible
            const getColorForFuel = (fuel) => {{
                if (RENEWABLE_FUELS.includes(fuel)) {{
                    return '#2ecc71'; // Verde para renovables
                }} else if (NON_RENEWABLE_FUELS.includes(fuel)) {{
                    return '#e74c3c'; // Rojo para no renovables
                }} else {{
                    return '#95a5a6'; // Gris para otros
                }}
            }};

            const scenarios = [...new Set(topTechs.map(d => d.Scenario))].sort();
            const traces = [];

            scenarios.forEach(scenario => {{
                const scenarioData = topTechs.filter(d => d.Scenario === scenario);

                if (scenarioData.length === 0) return;

                // Crear una traza por escenario
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
                               `Escenario: ${{scenario}}<br>` +
                               `FUEL: ${{d.FUEL}}<br>` +
                               `País: ${{d.COUNTRY}}<br>` +
                               `Valor Total: %{{x:.6f}}<br>` +
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
                    title: 'Valor Total (suma de todos los años)',
                    gridcolor: '{COLORS['border']}'
                }},
                yaxis: {{
                    title: 'Tecnología',
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
                        text: 'No hay datos para los filtros seleccionados',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Año' }},
                    yaxis: {{ title: title }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            // Agrupar SOLO por Scenario y YEAR (no por FUEL) - VERSIÓN AGREGADA
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

                // Clasificar por tipo de combustible para información adicional
                if (RENEWABLE_FUELS.includes(row.FUEL)) {{
                    grouped[key].RenewableValue += value;
                }} else if (NON_RENEWABLE_FUELS.includes(row.FUEL)) {{
                    grouped[key].NonRenewableValue += value;
                }}
            }});

            const groupedArray = Object.values(grouped);

            // Filtrar valores significativos (mayores a 0)
            const filteredArray = groupedArray.filter(d => d.Value > 0);

            if (filteredArray.length === 0) {{
                const emptyLayout = {{
                    title: title,
                    annotations: [{{
                        text: 'No hay valores mayores a 0 para esta métrica',
                        xref: 'paper',
                        yref: 'paper',
                        x: 0.5,
                        y: 0.5,
                        showarrow: false,
                        font: {{ size: 16, color: '{COLORS['secondary']}' }}
                    }}],
                    xaxis: {{ title: 'Año' }},
                    yaxis: {{ title: title }}
                }};
                Plotly.newPlot(divId, [], emptyLayout, {{responsive: true}});
                return;
            }}

            const scenarios = [...new Set(filteredArray.map(d => d.Scenario))].sort();

            // Colores para cada escenario (paleta de colores variados)
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
                        `Año: %{{x}}<br>` +
                        `${{title}}: %{{y:.6f}}<br>` +
                        `🌱 Renovable: ${{d.RenewableValue.toFixed(6)}} (${{((d.RenewableValue / d.Value) * 100).toFixed(1)}}%)<br>` +
                        `⚫ No Renovable: ${{d.NonRenewableValue.toFixed(6)}} (${{((d.NonRenewableValue / d.Value) * 100).toFixed(1)}}%)<br>` +
                        `<extra></extra>`
                    ),
                    customdata: scenarioData.map(d => [d.RenewableValue, d.NonRenewableValue])
                }};
                traces.push(trace);
            }});

            const layout = {{
                title: {{
                    text: title + ' (Agregado por Escenario)',
                    font: {{ size: 18 }}
                }},
                xaxis: {{
                    title: 'Año',
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

    # Guardar archivo
    print(f"   💾 Guardando dashboard: {output_file}")
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"   ✅ Dashboard interactivo generado exitosamente!")
    return output_file


def find_csv_files():
    """Busca archivos CSV en el directorio actual"""
    csv_files = glob.glob("*.csv")
    return sorted(csv_files)


def select_files_interactive():
    """Permite al usuario seleccionar archivos interactivamente"""
    csv_files = find_csv_files()

    if not csv_files:
        print("\n⚠️  No se encontraron archivos CSV en el directorio actual.")
        return []

    print("\n📁 Archivos CSV disponibles:")
    print("-" * 70)
    for idx, file in enumerate(csv_files, 1):
        size_mb = os.path.getsize(file) / (1024 * 1024)
        print(f"   {idx}. {file} ({size_mb:.2f} MB)")
    print("-" * 70)

    print("\n💡 Opciones:")
    print("   • Ingresa números separados por comas (ej: 1,3,5)")
    print("   • Ingresa 'all' para procesar todos los archivos")
    print("   • Ingresa 'q' para cancelar")

    while True:
        selection = input("\n👉 Tu selección: ").strip().lower()

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
                print(f"   ❌ Error: Los números deben estar entre 1 y {len(csv_files)}")
        except ValueError:
            print("   ❌ Error: Formato inválido. Usa números separados por comas.")


def main():
    """Función principal"""
    print("=" * 70)
    print("  GENERADOR DE DASHBOARDS HTML - TECNOLOGÍAS PWR - VERSIÓN AGREGADA")
    print("=" * 70)
    print("\n⚠️  VERSIÓN AGREGADA: Los gráficos muestran valores sumados POR ESCENARIO")
    print("\nEste script genera dashboards HTML interactivos con:")
    print("  • ✅ Filtrado automático de tecnologías PWR válidas")
    print(f"  • ✅ Patrón: {TECH_PATTERN}")
    print("  • ✅ Filtros por Escenario, FUEL y País")
    print("  • ✅ Auto-selección de tecnologías basada en FUEL y COUNTRY")
    print("  • ✅ Botones rápidos para Renovables/No Renovables")
    print("  • ✅ 5 gráficos AGREGADOS POR ESCENARIO:")
    print("      - Shares de Renovabilidad")
    print("      - Suma Total por Escenario (Lower Limit)")
    print("      - Suma Total por Escenario (Production)")
    print("      - Evolución Temporal por Escenario (Lower Limit)")
    print("      - Evolución Temporal por Escenario (Production)")
    print("  • ✅ Desglose renovable/no renovable en gráficos y tooltips")
    print("  • ✅ Cambio entre gráficos de líneas y barras")
    print("  • ✅ Completamente standalone (no requiere servidor)")

    selected_files = select_files_interactive()

    if not selected_files:
        print("\n❌ No se seleccionaron archivos. Saliendo...")
        return

    print(f"\n🚀 Procesando {len(selected_files)} archivo(s)...")
    print("=" * 70)

    generated_files = []
    errors = []

    for file in selected_files:
        try:
            print(f"\n📂 Cargando archivo: {file}")
            df = pd.read_csv(file, low_memory=False)
            print(f"   ✅ Archivo cargado: {len(df):,} filas")

            output_file = generate_interactive_dashboard(df, file)

            if output_file:
                generated_files.append(output_file)
            else:
                errors.append((file, "No se pudo generar el dashboard"))

        except FileNotFoundError:
            print(f"   ❌ Error: Archivo no encontrado: {file}")
            errors.append((file, "Archivo no encontrado"))

        except pd.errors.EmptyDataError:
            print(f"   ❌ Error: El archivo está vacío: {file}")
            errors.append((file, "Archivo vacío"))

        except Exception as e:
            print(f"   ❌ Error inesperado procesando {file}")
            print(f"      Detalles: {type(e).__name__}: {str(e)}")
            errors.append((file, f"{type(e).__name__}: {str(e)}"))

    # Resumen
    print("\n" + "=" * 70)
    print("  RESUMEN")
    print("=" * 70)
    print(f"\n✅ Dashboards interactivos generados: {len(generated_files)}")

    if generated_files:
        print("\n📄 Archivos generados:")
        for file in generated_files:
            print(f"   • {file}")

    if errors:
        print(f"\n❌ Errores: {len(errors)}")
        for file, error in errors:
            print(f"   • {file}: {error}")

    print("\n💡 Características de los dashboards agregados (VERSIÓN AGREGADA):")
    print("   • ⚠️  AGREGADO POR ESCENARIO: Los gráficos suman todos los valores por escenario")
    print("   • ✅ Filtrado automático de tecnologías PWR válidas")
    print(f"   • ✅ Patrón de validación: {TECH_PATTERN}")
    print("   • ✅ Auto-selección de tecnologías basada en filtros FUEL y COUNTRY")
    print("   • ✅ Filtros interactivos por Escenario, FUEL y País")
    print("   • ✅ Botones rápidos: Renovables / No Renovables")
    print("   • ✅ 5 gráficos AGREGADOS POR ESCENARIO:")
    print("      - 🌱 Shares de Renovabilidad (con línea de referencia al 50%)")
    print("      - 📊 Suma Total por Escenario - Lower Limit (barras apiladas renovable/no renovable)")
    print("      - 📊 Suma Total por Escenario - Production (barras apiladas renovable/no renovable)")
    print("      - 📉 Evolución Temporal por Escenario - Lower Limit (una línea por escenario)")
    print("      - ⚡ Evolución Temporal por Escenario - Production (una línea por escenario)")
    print("   • ✅ Desglose renovable/no renovable en tooltips")
    print("   • ✅ Clasificación visual con colores")
    print("   • ✅ Cambio dinámico entre gráficos de líneas y barras")
    print("   • ✅ Actualización en tiempo real sin recargar la página")
    print("\n📌 Próximos pasos:")
    print("   1. Abre los archivos HTML en tu navegador")
    print("   2. Usa los filtros de FUEL y COUNTRY para seleccionar qué incluir")
    print("   3. Las tecnologías se actualizan automáticamente según los filtros")
    print("   4. Los gráficos muestran valores SUMADOS por escenario")
    print("   5. Cambia entre líneas y barras según prefieras")
    print("   6. Compara múltiples escenarios en el mismo gráfico")
    print("\n💡 Diferencia clave:")
    print("   • VERSIÓN ESTÁNDAR: Muestra cada tecnología/combustible por separado")
    print("   • VERSIÓN AGREGADA: Suma todo por escenario (esta versión)")
    print("\n" + "=" * 70)


if __name__ == '__main__':
    main()
