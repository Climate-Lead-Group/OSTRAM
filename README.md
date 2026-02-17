# RELAC TX - Energy System Optimization Model

Sistema de modelado de optimización energética basado en OSeMOSYS para América Latina y el Caribe.

## Descripción

Este proyecto implementa un pipeline automatizado para la ejecución de modelos de optimización energética utilizando OSeMOSYS. El sistema soporta múltiples solvers (GLPK, CBC, CPLEX, Gurobi) y está diseñado para garantizar reproducibilidad completa de los resultados.

## Características Principales

- **Pipeline Automatizado**: Gestión completa del flujo de trabajo con DVC
- **Múltiples Solvers**: Soporte para GLPK, CBC, CPLEX y Gurobi
- **Reproducibilidad Garantizada**: Seeds configurables para resultados determinísticos
- **Medición de Rendimiento**: Timer integrado para monitorear tiempos de ejecución
- **Gestión Automática de Entorno**: Creación y actualización automática del entorno Conda

## Requisitos del Sistema

- Windows 10 o superior
- Git para Windows
- Miniconda o Anaconda
- Al menos un solver: GLPK, CBC, CPLEX o Gurobi

## Inicio Rápido

```bash
# Clonar el repositorio
git clone https://github.com/clg-admin/relac_tx.git
cd relac_tx

# Ejecutar el modelo (desde Anaconda Prompt)
python run.py
```

El script `run.py` gestiona automáticamente:
- Creación del entorno Conda
- Instalación de dependencias
- Ejecución del pipeline completo
- Generación de archivos de salida

## Documentación

Para instrucciones detalladas de instalación y configuración, consulta la guía completa:
- **Guía de Instalación y Ejecución**: `RELAC_TX_Guia_instalacion_ejecucion.md`

## Estructura de Archivos de Salida

Los resultados se generan en `t1_confection/` con los siguientes archivos:
- `RELAC_TX_Inputs.csv` / `RELAC_TX_Inputs_YYYY-MM-DD.csv`
- `RELAC_TX_Outputs.csv` / `RELAC_TX_Outputs_YYYY-MM-DD.csv`
- `RELAC_TX_Combined_Inputs_Outputs.csv` / `RELAC_TX_Combined_Inputs_Outputs_YYYY-MM-DD.csv`

Los archivos con fecha mantienen un histórico completo de ejecuciones.

## Configuración

El archivo principal de configuración es `t1_confection/MOMF_T1_AB.yaml`, donde puedes ajustar:
- Solver a utilizar (`solver: 'cplex'`)
- Número de threads para solvers comerciales
- Seeds para reproducibilidad
- Anualización de capital (`annualize_capital`)

## Matriz Tecnología-País

El sistema incluye una matriz configurable que permite especificar qué combinaciones de tecnología y país deben procesarse, además de unificar las tecnologías CCG y OCG en NGS.

### Uso

1. **Generar la matriz**:
   ```bash
   python t1_confection/A0_generate_tech_country_matrix.py
   ```
   Esto crea el archivo `Tech_Country_Matrix.xlsx` con las siguientes hojas:
   - **Matrix**: Matriz YES/NO para cada combinación tecnología-país
   - **NGS_Unification**: Configuración para unificar CCG + OCG → NGS
   - **Aggregation_Rules**: Reglas de agregación (avg/sum/disabled)
   - **Tech_Reference**: Descripción de tecnologías
   - **Country_Reference**: Descripción de países

2. **Configurar la matriz**:
   - En la hoja **Matrix**: Cambiar YES/NO para habilitar/deshabilitar combinaciones
   - En la hoja **NGS_Unification**: Cambiar a YES/NO para habilitar la unificación CCG+OCG→NGS

3. **Ejecutar el preprocesamiento**:
   ```bash
   python t1_confection/A1_Pre_processing_OG_csvs.py
   ```
   El script aplicará automáticamente:
   - Filtrado por matriz tecnología-país
   - Unificación NGS (si está habilitada)
   - Consolidación de regiones
   - Limpieza de tecnologías PWR

### Tecnologías en la Matriz

| Código | Descripción |
|--------|-------------|
| BCK | Backstop |
| BIO | Biomass |
| CCS | Carbon Capture Storage with Coal |
| COA | Coal |
| COG | Cogeneration |
| CSP | Concentrated Solar Power |
| GAS | Natural Gas |
| GEO | Geothermal |
| HYD | Hydroelectric |
| LDS | Long duration storage |
| NGS | Natural Gas (CCG + OCG unified) |
| OIL | Oil |
| OTH | Other |
| PET | Petroleum |
| SDS | Short duration storage |
| SPV | Solar Photovoltaic |
| URN | Nuclear |
| WAS | Waste |
| WAV | Wave |
| WOF | Offshore Wind |
| WON | Onshore Wind |

**Nota:** Los prefijos estructurales (ELC, MIN, PWR, RNW, TRN) no se incluyen en la matriz porque se combinan con los códigos anteriores para formar nombres de tecnología completos (ej: PWRBIOARGXX, MINCOAARGXX).

## Editor de Tecnologías Secundarias

El proyecto incluye un sistema para facilitar la edición de tecnologías secundarias (Secondary Techs) en los archivos de parametrización, con soporte para integración automática de datos OLADE.

### Uso del Editor

1. **Generar plantilla de edición**:
   ```bash
   python t1_confection/D1_generate_editor_template.py
   ```
   Esto crea el archivo `Secondary_Techs_Editor.xlsx` con dos hojas:
   - **Instructions**: Para edición manual con listas desplegables
   - **OLADE_Config**: Configuración de integración automática con datos OLADE

2. **Edición Manual** (Hoja "Instructions"):
   - Seleccionar: Escenario (BAU, NDC, NDC+ELC, NDC_NoRPO, o ALL)
   - Seleccionar: País, Tecnología (Tech.Name) y Parámetro
   - Ingresar los valores para los años deseados (2021-2050)
   - La columna "Tech" se completa automáticamente con VLOOKUP

3. **Integración OLADE** (Hoja "OLADE_Config"):

   Permite poblar automáticamente parámetros usando datos de OLADE.

   | Parámetro | Descripción |
   |-----------|-------------|
   | `ResidualCapacitiesFromOLADE` | YES/NO - Habilitar integración OLADE para capacidad instalada (ResidualCapacity) |
   | `PetroleumSplitMode` | OIL_only o Split_PET_OIL - Modo de división de petróleo |
   | `DemandFromOLADE` | YES/NO - Habilitar integración OLADE para demanda eléctrica |
   | `ActivityLowerLimitFromOLADE` | YES/NO - Habilitar integración OLADE para TotalTechnologyAnnualActivityLowerLimit |
   | `ActivityUpperLimitFromOLADE` | YES/NO - Habilitar integración OLADE para TotalTechnologyAnnualActivityUpperLimit |

   **PetroleumSplitMode**:
   - `OIL_only`: Asigna toda la capacidad de petróleo a OIL (Fuel oil)
   - `Split_PET_OIL`: Divide entre PET (Diésel) y OIL (Fuel oil + Búnker) usando proporciones del archivo `Shares_PET_OIL_Split.xlsx`

   **DemandFromOLADE**:
   - Cuando está habilitado, actualiza la demanda eléctrica en `A-O_Demand.xlsx` usando datos de generación de OLADE
   - Configura las tasas de crecimiento por país en la hoja `Demand_Growth`
   - Fórmula: `Demanda(año) = Demanda(2023) × (1 + tasa × (año - 2023))`

   **ActivityLowerLimit y ActivityUpperLimit**:
   - Cuando están habilitados, poblan automáticamente los límites de actividad en `A-O_Parametrization.xlsx`
   - Usa datos de generación eléctrica de OLADE combinados con shares de tecnologías de `Shares_Power_Generation_Technologies.xlsx`
   - Configura objetivos de renovabilidad opcionales en la hoja `Renewability_Targets`
   - Configura pesos personalizados de tecnologías en la hoja `Technology_Weights`
   - Fórmula: `ActivityLimit(tech,año) = Generación_Total(PJ) × (1 + tasa × (año - 2023)) × Share(tech,año)`
   - Incluye validación automática contra capacidades disponibles
   - Ver hoja `Documentation` en el editor para detalles completos de cálculo y validación

4. **Hojas Adicionales del Editor**:

   El archivo `Secondary_Techs_Editor.xlsx` también incluye:
   - **Renewability_Targets**: Define objetivos de % renovable por año para cada país/escenario (usado por Activity Limits)
   - **Technology_Weights**: Permite definir distribución personalizada de tecnologías renovables y no renovables
   - **Scenarios_Demand_Growth**: Configura tasas de crecimiento de demanda específicas por escenario y país
   - **Documentation**: Documentación técnica completa sobre el cálculo y validación de Activity Limits

5. **Aplicar cambios**:
   ```bash
   python t1_confection/D2_update_secondary_techs.py
   ```

### Características del Sistema

- **Listas desplegables**: Facilitan la selección de escenarios, países, tecnologías y parámetros
- **Mapeo Tech.Name → Tech**: Conversión automática de nombres descriptivos a códigos técnicos
- **Integración OLADE Capacidad**: Población automática de ResidualCapacity desde datos de capacidad instalada
- **Integración OLADE Demanda**: Población automática de demanda eléctrica desde datos de generación
- **Integración OLADE Activity Limits**: Población automática de TotalTechnologyAnnualActivityLowerLimit y UpperLimit
- **Conversión de unidades**: MW → GW (capacidad), GWh → PJ (demanda y activity)
- **Valores flat (capacidad)**: El mismo valor de capacidad se usa para todos los años
- **Crecimiento lineal (demanda y activity)**: Tasa de crecimiento configurable por país
- **Validación de Activity Limits**: Verifica automáticamente que los límites no excedan la capacidad disponible
- **Objetivos de renovabilidad**: Sistema de interpolación para alcanzar metas de % renovable
- **Respaldos automáticos**: Un backup por escenario antes de aplicar cambios
- **Projection.Mode**: Se actualiza automáticamente a "User defined" al modificar valores
- **Logs detallados**: Registro completo con identificación de país en cada operación

### Archivos Relacionados

| Archivo | Descripción |
|---------|-------------|
| `A0_generate_tech_country_matrix.py` | Genera la matriz tecnología-país |
| `D1_generate_editor_template.py` | Genera la plantilla Excel |
| `D2_update_secondary_techs.py` | Aplica los cambios a los escenarios |
| `Tech_Country_Matrix.xlsx` | Matriz tecnología-país (generada) |
| `Secondary_Techs_Editor.xlsx` | Plantilla de edición (generada) |
| `OLADE - Capacidad instalada por fuente - Anual.xlsx` | Datos fuente OLADE (capacidad instalada) |
| `OLADE - Generación eléctrica por fuente - Anual.xlsx` | Datos fuente OLADE (generación eléctrica) |
| `Shares_PET_OIL_Split.xlsx` | Proporciones para split de petróleo (Diesel, Fuel oil, Bunker) por escenario |
| `Shares_Power_Generation_Technologies.xlsx` | Proporciones de tecnologías de generación eléctrica por país/escenario/año |

### Mapeo de Países OLADE → Modelo

Algunos códigos de país difieren entre OLADE y el modelo:

| País | OLADE | Modelo |
|------|-------|--------|
| Barbados | BAR | JAM |
| Chile | CHI | CHL |
| Costa Rica | CRC | CRI |

## Herramientas de Gestión de Países

### Validador de Datos por País

Verifica que un país tiene todos los datos requeridos en los archivos CSV de entrada de OSeMOSYS.

```bash
python t1_confection/Z_validate_country_data.py                  # Validar todos los países RELAC
python t1_confection/Z_validate_country_data.py --country ARG    # Validar un país específico
python t1_confection/Z_validate_country_data.py --country NCC --report  # Generar reporte detallado
```

**Validaciones realizadas:**
- Presencia en sets (TECHNOLOGY, FUEL, EMISSION, STORAGE)
- Cantidad mínima de tecnologías por prefijo (PWR, MIN, RNW)
- Datos en todos los parámetros requeridos (costos, capacidad, factores, ratios, etc.)
- Patrones de fuels esperados por país

### Generador de Plantillas para Nuevo País

Crea un conjunto de archivos CSV con la estructura mínima necesaria para agregar un nuevo país, usando un país existente como referencia.

```bash
python t1_confection/Z_generate_country_template.py                              # Lee config desde YAML
python t1_confection/Z_generate_country_template.py --new NCC --ref ARG -i BOL PRY  # Override por CLI
```

**Configuración** (sección `template_generation` en `Config_country_codes.yaml`):

| Parámetro | Descripción |
|-----------|-------------|
| `new_country` | Código de 3 letras del nuevo país |
| `reference_country` | País existente del cual clonar datos |
| `region` | Código de región (default: XX) |
| `interconnections` | Lista de vecinos para interconexiones (vacío = sin interconexiones) |

**Características:**
- Genera CSVs en `templates/{CÓDIGO}/` sin modificar los archivos originales
- Manejo dinámico de interconexiones: soporta más, menos, igual o cero interconexiones respecto al país de referencia
- Transformación correcta de códigos de fuel y mode-of-operation para tecnologías TRN
- Incluye script `merge_into_inputs.py` en la carpeta generada para facilitar la integración

### Archivos Relacionados

| Archivo | Descripción |
|---------|-------------|
| `Z_validate_country_data.py` | Valida datos de un país en OG_csvs_inputs |
| `Z_generate_country_template.py` | Genera plantilla CSV para agregar un país |
| `Config_country_codes.yaml` | Configuración centralizada (incluye sección `template_generation`) |

## Licencia

Este proyecto está licenciado bajo la Licencia Apache 2.0 - consulta el archivo [LICENSE](LICENSE) para más detalles.

Copyright 2025 Climate Lead Group

Este proyecto está desarrollado por Climate Lead Group para análisis de sistemas energéticos en América Latina y el Caribe.
