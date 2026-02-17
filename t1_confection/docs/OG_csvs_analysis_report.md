# Informe de Análisis Estructural: OG_csvs_inputs

**Proyecto:** ReLAC-TX (Red de Energía Limpia de América Latina y el Caribe)
**Fecha:** 2026-02-16
**Directorio analizado:** `t1_confection/OG_csvs_inputs/`

---

## 1. Resumen Ejecutivo

### 1.1 Contexto: OSeMOSYS Global

**OSeMOSYS** (Open Source energy MOdelling SYStem) es un framework de código abierto para modelado de sistemas energéticos basado en **optimización lineal de largo plazo**. Minimiza el costo total descontado del sistema energético a lo largo de un horizonte de planificación, sujeto a restricciones técnicas, de demanda, de capacidad y de política.

**OSeMOSYS Global** es un dataset global pre-construido que utiliza el framework OSeMOSYS, proporcionando datos parametrizados para todos los países del mundo. El proyecto ReLAC extrae y adapta los datos para **19 países de América Latina y el Caribe**.

### 1.2 Hallazgos Clave

| Métrica | Valor |
|---------|-------|
| **Total de archivos CSV** | 64 (excluye backups) |
| **Archivos de conjuntos (Sets)** | 11 |
| **Archivos de parámetros** | 53 (29 con datos, 24 vacíos) |
| **Países representados** | 19 + INT (mercados internacionales) |
| **Horizonte temporal** | 2023–2050 (28 años) |
| **Resolución temporal** | 12 time slices (4 estaciones × 3 periodos diarios) |
| **Región del modelo** | GLOBAL (única) |
| **Total de tecnologías** | 1,007 |
| **Total de combustibles** | 415 |
| **Total de emisiones** | 171 (19 RELAC + 152 globales) |
| **Total de almacenamientos** | 50 |

### 1.3 Arquitectura del Modelo

El modelo utiliza una **única región "GLOBAL"** y codifica la diferenciación por país **dentro de los nombres de tecnologías y combustibles**. Esto permite modelar comercio internacional sin múltiples regiones en OSeMOSYS.

### 1.4 Nomenclatura

Las tecnologías siguen el patrón: `PREFIJO(3) + TIPO(3) + PAÍS(3) + REGIÓN(2) + [SUFIJO]`

Ejemplo: `PWRSPVARGXX01` = **P**o**W**e**R** + **S**olar **P**hoto**V**oltaic + **ARG**entina + región **XX** (todas) + **01** (nueva inversión)

> **Nota importante:** Los archivos de parámetros ya contienen datos en formato "consolidado" (sin sufijos 00/01), mientras que `TECHNOLOGY.csv` mantiene el formato original de OSeMOSYS Global con sufijos. Esto es intencional — el pipeline de preprocesamiento (A1) maneja la consolidación.

---

## 2. Inventario de Archivos

### 2.1 Archivos de Conjuntos (Sets)

Estos archivos definen los elementos del modelo. Tienen una sola columna `VALUE`.

| Archivo | Registros | Tamaño | Descripción |
|---------|-----------|--------|-------------|
| `REGION.csv` | 1 | 15 B | Región única: GLOBAL |
| `YEAR.csv` | 28 | 175 B | Años 2023–2050 |
| `TIMESLICE.csv` | 12 | 79 B | S1D1–S4D3 (4 estaciones × 3 periodos) |
| `SEASON.csv` | 4 | 19 B | Estaciones S1–S4 |
| `DAYTYPE.csv` | 1 | 10 B | Tipo de día único |
| `DAILYTIMEBRACKET.csv` | 3 | 16 B | Periodos diarios D1–D3 |
| `MODE_OF_OPERATION.csv` | 2 | 13 B | Modos de operación |
| `TECHNOLOGY.csv` | 1,007 | 14 KB | Todas las tecnologías del modelo |
| `FUEL.csv` | 415 | 4 KB | Todos los combustibles/commodities |
| `EMISSION.csv` | 171 | 1.4 KB | Tipos de emisiones (CO2 por país) |
| `STORAGE.csv` | 50 | 607 B | Almacenamientos (LDS + SDS por país) |

### 2.2 Archivos de Parámetros con Datos

#### Costos

| Archivo | Columnas | Filas | Tamaño | Parámetro OSeMOSYS | Unidades |
|---------|----------|-------|--------|---------------------|----------|
| `CapitalCost.csv` | REGION, TECHNOLOGY, YEAR, VALUE | 6,804 | 219 KB | Costo de inversión | M$/GW (≡ $/kW) |
| `FixedCost.csv` | REGION, TECHNOLOGY, YEAR, VALUE | 6,804 | 217 KB | Costo fijo O&M anual | M$/GW/año |
| `VariableCost.csv` | REGION, TECHNOLOGY, MODE_OF_OPERATION, YEAR, VALUE | 6,440 | 209 KB | Costo variable | M$/PJ |
| `CapitalCostStorage.csv` | REGION, STORAGE, YEAR, VALUE | 1,064 | 45 KB | Costo inversión almacenamiento | M$/GW |

#### Capacidad y Vida Útil

| Archivo | Columnas | Filas | Tamaño | Parámetro OSeMOSYS |
|---------|----------|-------|--------|---------------------|
| `ResidualCapacity.csv` | REGION, TECHNOLOGY, YEAR, VALUE | 3,528 | 115 KB | Capacidad existente (GW) |
| `TotalAnnualMaxCapacity.csv` | REGION, TECHNOLOGY, YEAR, VALUE | 2,072 | 72 KB | Máxima capacidad total anual |
| `TotalAnnualMaxCapacityInvestment.csv` | REGION, TECHNOLOGY, YEAR, VALUE | 1,820 | 53 KB | Máxima inversión anual |
| `OperationalLife.csv` | REGION, TECHNOLOGY, VALUE | 243 | 6 KB | Vida útil (años) |
| `CapacityToActivityUnit.csv` | REGION, TECHNOLOGY, VALUE | 262 | 7 KB | Factor conversión GW→PJ/año |

#### Rendimiento y Actividad

| Archivo | Columnas | Filas | Tamaño | Parámetro OSeMOSYS |
|---------|----------|-------|--------|---------------------|
| `CapacityFactor.csv` | REGION, TECHNOLOGY, TIMESLICE, YEAR, VALUE | 24,192 | 900 KB | Factor de capacidad por timeslice |
| `AvailabilityFactor.csv` | REGION, TECHNOLOGY, YEAR, VALUE | 4,480 | 130 KB | Disponibilidad anual |
| `InputActivityRatio.csv` | REGION, TECHNOLOGY, FUEL, MODE_OF_OPERATION, YEAR, VALUE | 9,800 | 400 KB | Ratio de consumo de combustible |
| `OutputActivityRatio.csv` | REGION, TECHNOLOGY, FUEL, MODE_OF_OPERATION, YEAR, VALUE | 18,508 | 753 KB | Ratio de producción |

#### Demanda

| Archivo | Columnas | Filas | Tamaño | Parámetro OSeMOSYS |
|---------|----------|-------|--------|---------------------|
| `SpecifiedAnnualDemand.csv` | REGION, FUEL, YEAR, VALUE | 504 | 15 KB | Demanda anual total (PJ) |
| `SpecifiedDemandProfile.csv` | REGION, FUEL, TIMESLICE, YEAR, VALUE | 6,384 | 259 KB | Perfil temporal de demanda |

#### Emisiones

| Archivo | Columnas | Filas | Tamaño | Parámetro OSeMOSYS |
|---------|----------|-------|--------|---------------------|
| `EmissionActivityRatio.csv` | REGION, TECHNOLOGY, EMISSION, MODE_OF_OPERATION, YEAR, VALUE | 3,836 | 157 KB | Emisiones por actividad (Mt/PJ) |

#### Almacenamiento

| Archivo | Columnas | Filas | Tamaño | Parámetro OSeMOSYS |
|---------|----------|-------|--------|---------------------|
| `ResidualStorageCapacity.csv` | REGION, STORAGE, YEAR, VALUE | 224 | 7 KB | Capacidad residual storage |
| `OperationalLifeStorage.csv` | REGION, STORAGE, VALUE | 38 | 934 B | Vida útil storage |
| `StorageLevelStart.csv` | REGION, STORAGE, VALUE | 38 | 820 B | Nivel inicial |
| `TechnologyToStorage.csv` | REGION, TECHNOLOGY, STORAGE, MODE_OF_OPERATION, VALUE | 76 | 3 KB | Techs que cargan storage |
| `TechnologyFromStorage.csv` | REGION, TECHNOLOGY, STORAGE, MODE_OF_OPERATION, VALUE | 76 | 3 KB | Techs que descargan storage |

#### Restricciones y Reservas

| Archivo | Columnas | Filas | Tamaño | Parámetro OSeMOSYS |
|---------|----------|-------|--------|---------------------|
| `TotalTechnologyAnnualActivityUpperLimit.csv` | REGION, TECHNOLOGY, YEAR, VALUE | 3,724 | 101 KB | Límite superior de actividad |
| `ReserveMarginTagTechnology.csv` | REGION, TECHNOLOGY, YEAR, VALUE | 6,076 | 176 KB | Tag de margen de reserva tech |
| `ReserveMarginTagFuel.csv` | REGION, FUEL, YEAR, VALUE | 532 | 15 KB | Tag de margen de reserva fuel |

#### Conversiones Temporales

| Archivo | Columnas | Filas | Tamaño | Parámetro OSeMOSYS |
|---------|----------|-------|--------|---------------------|
| `YearSplit.csv` | TIMESLICE, YEAR, VALUE | 336 | 10 KB | Fracción del año por timeslice |
| `DaySplit.csv` | DAILYTIMEBRACKET, YEAR, VALUE | 84 | 2 KB | División diaria |
| `Conversionls.csv` | TIMESLICE, SEASON, VALUE | 48 | 600 B | Mapeo timeslice→estación |
| `Conversionld.csv` | TIMESLICE, DAYTYPE, VALUE | 12 | 145 B | Mapeo timeslice→tipo de día |
| `Conversionlh.csv` | TIMESLICE, DAILYTIMEBRACKET, VALUE | 36 | 466 B | Mapeo timeslice→periodo diario |

### 2.3 Archivos de Parámetros Vacíos (24 archivos)

Estos archivos existen con headers pero sin datos. Representan parámetros que **usan valores por defecto de OSeMOSYS** o que se configuran en etapas posteriores del pipeline:

| Categoría | Archivos Vacíos |
|-----------|-----------------|
| **Costos/Económicos** | `DiscountRate`, `DiscountRateStorage`, `DepreciationMethod`, `EmissionsPenalty` |
| **Demanda** | `AccumulatedAnnualDemand` |
| **Emisiones** | `AnnualEmissionLimit`, `AnnualExogenousEmission`, `ModelPeriodEmissionLimit`, `ModelPeriodExogenousEmission` |
| **Capacidad mínima** | `TotalAnnualMinCapacity`, `TotalAnnualMinCapacityInvestment`, `CapacityOfOneTechnologyUnit` |
| **Actividad** | `TotalTechnologyAnnualActivityLowerLimit`, `TotalTechnologyModelPeriodActivityLowerLimit`, `TotalTechnologyModelPeriodActivityUpperLimit` |
| **Renovables** | `REMinProductionTarget`, `RETagTechnology`, `RETagFuel` |
| **Reserva** | `ReserveMargin` |
| **Almacenamiento** | `StorageMaxChargeRate`, `StorageMaxDischargeRate`, `MinStorageCharge` |
| **Comercio** | `TradeRoute` |
| **Temporal** | `DaysInDayType` |

---

## 3. Estructura de Datos por País

### 3.1 Los 19 Países del Modelo

| Código ISO-3 | País | Nombre OLADE |
|--------------|------|--------------|
| ARG | Argentina | Argentina |
| BOL | Bolivia | Bolivia |
| BRA | Brasil | Brasil |
| BRB | Barbados | Barbados |
| CHL | Chile | Chile |
| COL | Colombia | Colombia |
| CRI | Costa Rica | Costa Rica |
| DOM | República Dominicana | República Dominicana |
| ECU | Ecuador | Ecuador |
| GTM | Guatemala | Guatemala |
| HND | Honduras | Honduras |
| HTI | Haití | Haití |
| MEX | México | México |
| NIC | Nicaragua | Nicaragua |
| PAN | Panamá | Panamá |
| PER | Perú | Perú |
| PRY | Paraguay | Paraguay |
| SLV | El Salvador | El Salvador |
| URY | Uruguay | Uruguay |

Además existe `INT` (International Markets) para commodities fósiles importables.

### 3.2 Cómo se Identifica un País en Cada Archivo

Los países **NO se identifican mediante una columna dedicada**. En cambio, el código de 3 letras se **embebe en los nombres de tecnologías, combustibles, emisiones y almacenamientos**.

| Tipo de Elemento | Columna | Patrón de País | Ejemplo |
|------------------|---------|----------------|---------|
| Tecnología | `TECHNOLOGY` | Posición 4-6 (PWR) o 4-6 (MIN/RNW) | `PWR`**BIO**`ARG`XX, `MIN`**COA**`ARG` |
| Combustible | `FUEL` | Posición 4-6 (renovable) o 4-6 (fósil) | `ELC`**ARG**`XX02`, `COA`**ARG** |
| Emisión | `EMISSION` | Posición 4-6 | `CO2`**ARG** |
| Almacenamiento | `STORAGE` | Posición 4-6 | `LDS`**ARG**`XX01` |

### 3.3 Tecnologías por País

Cada país tiene entre **40–46 tecnologías** (excepto Brasil con **260** debido a sus 7 subregiones en los datos originales).

**Distribución por tipo de tecnología (fuera de BRA):**

| Prefijo | Tipo | Cantidad por País | Descripción |
|---------|------|-------------------|-------------|
| PWR | Power | ~25 | Plantas de generación (BIO, CCG, OCG, CCS, COA, COG, CSP, GEO, HYD, LDS, OIL, OTH, PET, SDS, SPV, URN, WAS, WAV, WOF, WON + BCK + TRN) |
| MIN | Mining | 7 | Extracción de commodities (COA, COG, GAS, OIL, OTH, PET, URN) |
| RNW | Renewable | 9 | Provisión de recursos renovables (BIO, CSP, GEO, HYD, SPV, WAS, WAV, WOF, WON) |
| TRN | Transmission | 0–6 | Interconexiones internacionales (varía por país) |

**Interconexiones por país:**

| País | Conexiones | Destinos |
|------|------------|----------|
| ARG | 5 | BOL, BRA(SO), CHL, PRY, URY |
| BOL | 5 | BRA(CW), BRA(WE), CHL, PER, PRY |
| BRA | 22 | Internas (CN↔CW↔NE↔NW↔SE↔SO↔WE) + ARG, BOL, COL, PER, PRY, URY |
| BRB | 0 | (isla sin interconexiones) |
| CHL | 3 | ARG, BOL, PER |
| COL | 4 | BRA(NW), ECU, PAN, PER |
| CRI | 2 | NIC, PAN |
| DOM | 1 | HTI |
| ECU | 2 | COL, PER |
| GTM | 3 | HND, MEX, SLV |
| HND | 3 | GTM, NIC, SLV |
| HTI | 1 | DOM |
| MEX | 1 | GTM |
| NIC | 2 | CRI, HND |
| PAN | 2 | COL, CRI |
| PER | 6 | BOL, BRA(NW), BRA(WE), CHL, COL, ECU |
| PRY | 4 | ARG, BOL, BRA(CW), BRA(SO) |
| SLV | 2 | GTM, HND |
| URY | 2 | ARG, BRA(SO) |

### 3.4 Combustibles por País

Cada país tiene exactamente **18 combustibles** (excepto BRA con 84 debido a subregiones):

| Tipo | Ejemplo (ARG) | Con Región | Descripción |
|------|---------------|------------|-------------|
| BIOARGXX | Sí (XX) | Biomasa |
| CSPARGXX | Sí (XX) | Recurso solar térmico |
| GEOARGXX | Sí (XX) | Recurso geotérmico |
| HYDARGXX | Sí (XX) | Recurso hídrico |
| SPVARGXX | Sí (XX) | Recurso solar fotovoltaico |
| WASARGXX | Sí (XX) | Residuos |
| WAVARGXX | Sí (XX) | Recurso undimotriz |
| WOFARGXX | Sí (XX) | Recurso eólico offshore |
| WONARGXX | Sí (XX) | Recurso eólico onshore |
| COAARG | No | Carbón (comerciable) |
| COGARG | No | Cogeneración |
| GASARG | No | Gas natural |
| OILARG | No | Petróleo |
| OTHARG | No | Otros |
| PETARG | No | Productos derivados |
| URNARG | No | Uranio |
| ELCARGXX01 | Sí (XX01) | Electricidad generada |
| ELCARGXX02 | Sí (XX02) | Electricidad demanda final |

Adicionalmente existen 7 combustibles internacionales: `COAINT`, `COGINT`, `GASINT`, `OILINT`, `OTHINT`, `PETINT`, `URNINT`.

---

## 4. Análisis de Contenido y Valores

### 4.1 Rangos de Valores Clave

| Parámetro | Mín | Máx | Media | Mediana | Unidades |
|-----------|-----|-----|-------|---------|----------|
| **CapitalCost** | 231.7 | 7,650.0 | 1,875.9 | 1,240.0 | $/kW |
| **FixedCost** | 0.8 | 290.0 | 50.8 | 22.5 | $/kW/año |
| **VariableCost** | 0.0 | 17.5 | 6.2 | 5.9 | M$/PJ |
| **ResidualCapacity** | 0.0 | 98.0 | 1.9 | 0.1 | GW |
| **CapacityFactor** | 0.0 | 0.8 | 0.27 | 0.24 | fracción |
| **InputActivityRatio** | 1.0 | 3.9 | 1.77 | 1.0 | PJ_in/PJ_out |
| **SpecifiedAnnualDemand** | 14.9 | 3,231.9 | 409.9 | 102.4 | PJ |
| **EmissionActivityRatio** | 0.09 | 0.28 | 0.19 | 0.22 | Mt CO2/PJ |
| **OperationalLife** | 30 | 100 | 41.1 | 30.0 | años |
| **CapacityToActivityUnit** | 31.536 | 31.536 | 31.536 | 31.536 | PJ/GW/año |

### 4.2 Valores por Tipo de Tecnología

**CapitalCost típicos ($/kW):**
- Solar PV (SPV): ~500–800
- Eólico onshore (WON): ~900–1,400
- Hidroeléctrica (HYD): ~1,500–2,500
- Gas natural ciclo combinado (CCG): ~800–1,200
- Nuclear (URN): ~5,000–7,650
- Backstop (BCK): ~6,000+

**OperationalLife típicos (años):**
- Solar PV/Eólica: 30
- Hidroeléctrica: 100
- Carbón/Gas: 30
- Nuclear: 40

**CapacityToActivityUnit:** Constante = 31.536 para todas las tecnologías (1 GW × 8,766 h × 0.0036 PJ/GWh).

### 4.3 Análisis Temporal

- **Horizonte:** 28 años (2023–2050)
- **Time Slices:** 12 periodos (4 estaciones × 3 periodos diarios)
  - Nomenclatura: S{1-4}D{1-3} (Ej: S1D1 = Estación 1, Periodo diurno 1)
  - YearSplit promedio: ~0.083 (cada timeslice ≈ 8.3% del año)
- **ResidualCapacity** decrece con el tiempo, llegando a 0 cuando las plantas alcanzan su fin de vida útil
- **CapitalCost** generalmente decrece para renovables (curva de aprendizaje)

---

## 5. Análisis de Completitud

### 5.1 Parámetros Obligatorios vs. Opcionales

#### Parámetros con datos para TODOS los 19 países (22 archivos — OBLIGATORIOS):

| # | Parámetro | Tipo |
|---|-----------|------|
| 1 | CapitalCost | Costo |
| 2 | FixedCost | Costo |
| 3 | VariableCost | Costo |
| 4 | CapitalCostStorage | Costo |
| 5 | CapacityFactor | Rendimiento |
| 6 | AvailabilityFactor | Rendimiento |
| 7 | InputActivityRatio | Rendimiento |
| 8 | OutputActivityRatio | Rendimiento |
| 9 | EmissionActivityRatio | Emisiones |
| 10 | ResidualCapacity | Capacidad |
| 11 | OperationalLife | Capacidad |
| 12 | CapacityToActivityUnit | Capacidad |
| 13 | TotalAnnualMaxCapacity | Restricciones |
| 14 | TotalAnnualMaxCapacityInvestment | Restricciones |
| 15 | TotalTechnologyAnnualActivityUpperLimit | Restricciones |
| 16 | ReserveMarginTagTechnology | Reservas |
| 17 | ReserveMarginTagFuel | Reservas |
| 18 | SpecifiedDemandProfile | Demanda |
| 19 | StorageLevelStart | Almacenamiento |
| 20 | OperationalLifeStorage | Almacenamiento |
| 21 | TechnologyToStorage | Almacenamiento |
| 22 | TechnologyFromStorage | Almacenamiento |

#### Parámetros con cobertura parcial:

| Parámetro | Países sin datos | Nota |
|-----------|------------------|------|
| **SpecifiedAnnualDemand** | HTI | **CRÍTICO**: Haití no tiene demanda definida |
| **ResidualStorageCapacity** | BRB, COL, CRI, DOM, ECU, GTM, HND, NIC, PAN, PER, PRY, SLV, URY | Solo ARG, BOL, BRA, HTI, MEX tienen almacenamiento residual |

#### Parámetros vacíos (24 archivos — OPCIONALES/DEFAULT):

Estos parámetros usan valores por defecto de OSeMOSYS (generalmente 0 o sin restricción). Se pueden agregar datos posteriormente para:
- Definir límites de emisiones (`AnnualEmissionLimit`)
- Establecer metas de renovables (`REMinProductionTarget`, `RETagTechnology`)
- Fijar margen de reserva (`ReserveMargin`)
- Definir tasa de descuento (`DiscountRate`)

### 5.2 Integridad Referencial

**Hallazgo importante:** Existe una discrepancia controlada entre `TECHNOLOGY.csv` y los archivos de parámetros:

| Aspecto | TECHNOLOGY.csv | Archivos de Parámetros |
|---------|----------------|------------------------|
| Sufijos | Con sufijos (00, 01) | Sin sufijos (consolidado) |
| Brasil | 7 subregiones | Consolidado a XX |
| Ejemplo | `PWRBIOARGXX01` | `PWRBIOARGXX` |

Esto es **intencional** — el pipeline de preprocesamiento (`A1_Pre_processing_OG_csvs.py`) maneja la consolidación de sufijos y regiones.

**Integridad de emisiones:** `EMISSION.csv` contiene 171 emisiones globales, pero solo 19 corresponden a países RELAC. Las 152 restantes son herencia del dataset global y se filtran durante el procesamiento.

---

## 6. Análisis Comparativo entre Países

### 6.1 Argentina vs Brasil vs México

#### Capacidad Instalada (ResidualCapacity, año base 2023)

| Tecnología | ARG (GW) | BRA (GW) | MEX (GW) |
|------------|----------|----------|----------|
| Hidroeléctrica | 10.00 | 98.05 | 12.44 |
| Carbón | 6.50 | — | 6.05 |
| Petróleo/Oil | — | 14.97 | 11.53 |
| Biomasa | — | 13.47 | — |
| Eólica onshore | — | 11.38 | 3.27 |
| Nuclear | — | — | 1.51 |
| **Total** | **31.58** | **156.69** | **36.83** |

#### Demanda Eléctrica (año base 2023)

| País | Demanda (PJ) | Demanda (TWh) | Crecimiento 2023→2050 |
|------|-------------|---------------|----------------------|
| ARG | 555.9 | 154.4 | → 914.2 PJ (+64%) |
| BRA | 2,217.9 | 616.1 | → 2,918.2 PJ (+32%) |
| MEX | 1,212.6 | 336.8 | → 3,231.9 PJ (+167%) |

#### Volumen de Datos

| Parámetro | ARG (filas) | BRA (filas) | MEX (filas) |
|-----------|-------------|-------------|-------------|
| CapitalCost | 504 | 532 | 420 |
| CapacityFactor | 1,344 | 1,344 | 1,344 |
| ResidualCapacity | 364 | 364 | 280 |
| Technologies | 45 | 260 | 41 |

### 6.2 País Más Completo como Referencia

**Argentina (ARG)** es el país más completo como referencia para agregar nuevos países:
- Tiene 45 tecnologías (representativo sin complejidad de subregiones como BRA)
- Tiene datos en todos los parámetros obligatorios
- Tiene almacenamiento residual definido
- Tiene 5 interconexiones internacionales
- Tiene demanda definida con perfil temporal completo

### 6.3 País Más Incompleto

**Haití (HTI)** es el país con menos datos:
- **Falta SpecifiedAnnualDemand** (demanda no definida)
- Solo 41 tecnologías
- Una sola interconexión (DOM)
- Menor volumen de datos en parámetros de costo y capacidad

---

## 7. Requisitos de Datos para Agregar un Nuevo País

### 7.1 Checklist Completo

Para agregar un nuevo país con código ISO-3 `NCC` (Nuevo Código de País):

#### A. Archivos de Conjuntos (Sets) a Modificar

| # | Archivo | Acción | Registros a agregar |
|---|---------|--------|---------------------|
| 1 | `TECHNOLOGY.csv` | Agregar | ~40 tecnologías PWR, MIN, RNW + TRN |
| 2 | `FUEL.csv` | Agregar | 18 combustibles (9 renovables + 7 fósiles + 2 ELC) |
| 3 | `EMISSION.csv` | Agregar | 1 entrada: `CO2NCC` |
| 4 | `STORAGE.csv` | Agregar | 2 entradas: `LDSNCCXX01`, `SDSNCCXX01` |

#### B. Archivos de Parámetros Obligatorios

| # | Archivo | Registros estimados | Datos necesarios |
|---|---------|---------------------|------------------|
| 1 | `CapitalCost.csv` | ~18 techs × 28 años = 504 | Costos de inversión por tecnología y año |
| 2 | `FixedCost.csv` | ~504 | Costos fijos O&M |
| 3 | `VariableCost.csv` | ~350 | Costos variables (principalmente MIN) |
| 4 | `ResidualCapacity.csv` | ~12 techs × 28 años = 336 | Capacidad existente con retiro progresivo |
| 5 | `CapacityFactor.csv` | ~4 techs × 12 TS × 28 años = 1,344 | Factor de capacidad por timeslice |
| 6 | `AvailabilityFactor.csv` | ~16 techs × 28 años = 448 | Disponibilidad anual |
| 7 | `InputActivityRatio.csv` | ~29 techs × 28 años = 812 | Ratios de entrada (eficiencia) |
| 8 | `OutputActivityRatio.csv` | ~45 techs × 28 años = 1,260 | Ratios de salida |
| 9 | `EmissionActivityRatio.csv` | ~9 techs × 28 años = 252 | Emisiones CO2 por actividad |
| 10 | `SpecifiedAnnualDemand.csv` | 28 | Demanda eléctrica anual (PJ) |
| 11 | `SpecifiedDemandProfile.csv` | 12 TS × 28 años = 336 | Perfil temporal de demanda |
| 12 | `OperationalLife.csv` | ~18 | Vida útil por tecnología |
| 13 | `CapacityToActivityUnit.csv` | ~19 | Factor de conversión (31.536) |
| 14 | `TotalAnnualMaxCapacity.csv` | ~4 techs × 28 = 112 | Límites de capacidad |
| 15 | `TotalAnnualMaxCapacityInvestment.csv` | ~3-5 techs × 28 = 84-140 | Límites de inversión |
| 16 | `TotalTechnologyAnnualActivityUpperLimit.csv` | 7 techs × 28 = 196 | Límites de actividad |
| 17 | `ReserveMarginTagTechnology.csv` | ~13 techs × 28 = 364 | Tags para margen de reserva |
| 18 | `ReserveMarginTagFuel.csv` | 28 | Tag de fuel para reserva |
| 19 | `CapitalCostStorage.csv` | 2 × 28 = 56 | Costos de almacenamiento |
| 20 | `OperationalLifeStorage.csv` | 2 | Vida útil de almacenamiento |
| 21 | `StorageLevelStart.csv` | 2 | Nivel inicial storage |
| 22 | `TechnologyToStorage.csv` | 4 | Links tech→storage |
| 23 | `TechnologyFromStorage.csv` | 4 | Links storage→tech |

#### C. Archivos de Configuración

| # | Archivo | Acción |
|---|---------|--------|
| 1 | `Config_country_codes.yaml` | Agregar código, nombre inglés, nombre OLADE |
| 2 | `Tech_Country_Matrix.xlsx` | Agregar columna con YES/NO por tecnología |

#### D. Archivos Temporales (no requieren cambios por país)

Los siguientes archivos son **globales** y no necesitan modificación:
- `YearSplit.csv`, `DaySplit.csv`, `Conversionls.csv`, `Conversionld.csv`, `Conversionlh.csv`
- `REGION.csv`, `YEAR.csv`, `TIMESLICE.csv`, `SEASON.csv`, `DAYTYPE.csv`, `DAILYTIMEBRACKET.csv`
- `MODE_OF_OPERATION.csv`

### 7.2 Perfil Mínimo de Datos (MVP)

Para un país funcional mínimo se necesitan al menos:

1. **1 emisión:** CO2NCC
2. **18 combustibles:** 9 renovables + 7 fósiles + 2 electricidad
3. **~40 tecnologías:** 7 MIN + 9 RNW + ~22 PWR + 1 PWRTRN + 1 PWRBCK
4. **2 almacenamientos:** LDS + SDS
5. **Demanda anual** con perfil temporal (28 + 336 filas)
6. **Costos** para todas las tecnologías (CapitalCost, FixedCost, VariableCost)
7. **Capacidad residual** para tecnologías existentes
8. **Factor de capacidad** para renovables variables (SPV, WON, WOF, HYD)
9. **Ratios de actividad** (Input/Output) para todas las tecnologías
10. **Emisiones** por actividad para tecnologías fósiles

### 7.3 Fuentes de Datos Recomendadas

| Dato | Fuente Primaria | Fuente Alternativa |
|------|----------------|-------------------|
| Capacidad instalada | OLADE, IRENA | Estadísticas nacionales |
| Demanda eléctrica | IEA, OLADE | Operador del sistema eléctrico |
| Costos de inversión | IRENA, NREL ATB | World Energy Outlook (IEA) |
| Factores de capacidad | IRENA, Global Wind Atlas, Global Solar Atlas | Datos nacionales |
| Emisiones | IPCC Guidelines, EPA | Inventarios nacionales GEI |
| Interconexiones | OLADE, CIER | Operadores regionales |

---

## 8. Guía Paso a Paso para Agregar un Nuevo País

### Paso 1: Definir el País en la Configuración

Editar `Config_country_codes.yaml`:

```yaml
countries:
  # ... países existentes ...
  NCC:
    english_name: "New Country"
    olade_name: "Nuevo País"
```

### Paso 2: Generar los Identificadores

Para un país con código `NCC`:

**Tecnologías a crear:**

```
# Generación (PWR) — 22 tecnologías
PWRBIOARGXX → PWRBIONCCXX    # Biomasa
PWRCCGNCCXX00                 # Gas ciclo combinado existente
PWRCCGNCCXX01                 # Gas ciclo combinado nuevo
PWRCCSNCCXX01                 # CCS
PWRCOANCCXX01                 # Carbón
PWRCOGNCCXX01                 # Cogeneración
PWRCSPNCCXX01                 # Solar térmica
PWRGEONCCXX01                 # Geotérmica
PWRHYDNCCXX01                 # Hidroeléctrica
PWRLDSNCCXX01                 # Almacenamiento largo
PWROCGNCCXX00                 # Gas ciclo abierto existente
PWROCGNCCXX01                 # Gas ciclo abierto nuevo
PWROILNCCXX01                 # Petróleo
PWROTHNCCXX01                 # Otros
PWRPETNCCXX01                 # Diesel
PWRSDSNCCXX01                 # Almacenamiento corto
PWRSPVNCCXX01                 # Solar PV
PWRURNNCCXX01                 # Nuclear
PWRWASNCCXX01                 # Residuos
PWRWAVNCCXX01                 # Olas
PWRWOFNCCXX01                 # Eólica offshore
PWRWONNCCXX01                 # Eólica onshore
PWRBCKNCCXX                   # Backstop
PWRTRNNCCXX                   # Transmisión intranacional

# Extracción (MIN) — 7 tecnologías
MINCOANCC, MINCOGNCC, MINGASNCC, MINOILNCC, MINOTHNCC, MINPETNCC, MINURNNCC

# Renovables (RNW) — 9 tecnologías
RNWBIONCCXX, RNWCSPNCCXX, RNWGEONCCXX, RNWHYDNCCXX, RNWSPVNCCXX
RNWWASNCCXX, RNWWAVNCCXX, RNWWOFNCCXX, RNWWONNCCXX

# Transmisión internacional (TRN) — según conexiones
TRNNCCXXOTRXX    # NCC → Otro país (por cada conexión)
```

**Combustibles a crear:**

```
# Renovables (con región)
BIONCCXX, CSPNCCXX, GEONCCXX, HYDNCCXX, SPVNCCXX
WASNCCXX, WAVNCCXX, WOFNCCXX, WONNCCXX

# Fósiles (sin región)
COANCC, COGNCC, GASNCC, OILNCC, OTHNCC, PETNCC, URNNCC

# Electricidad
ELCNCCXX01    # Generada
ELCNCCXX02    # Demanda final
```

### Paso 3: Poblar los Parámetros

Para cada parámetro, crear filas con el formato correspondiente. A continuación un ejemplo usando ARG como plantilla:

**CapitalCost.csv** (ejemplo de 3 tecnologías):
```csv
REGION,TECHNOLOGY,YEAR,VALUE
GLOBAL,PWRBIONCCXX,2023,2150.0
GLOBAL,PWRBIONCCXX,2024,2120.0
...
GLOBAL,PWRSPVNCCXX,2023,750.0
GLOBAL,PWRSPVNCCXX,2024,720.0
...
GLOBAL,PWRHYDNCCXX,2023,2100.0
...
```

**SpecifiedAnnualDemand.csv:**
```csv
REGION,FUEL,YEAR,VALUE
GLOBAL,ELCNCCXX02,2023,100.0
GLOBAL,ELCNCCXX02,2024,103.0
...
```

**SpecifiedDemandProfile.csv** (debe sumar 1.0 por año):
```csv
REGION,FUEL,TIMESLICE,YEAR,VALUE
GLOBAL,ELCNCCXX02,S1D1,2023,0.08
GLOBAL,ELCNCCXX02,S1D2,2023,0.10
...
```

### Paso 4: Registrar en Tech_Country_Matrix.xlsx

Agregar una columna `NCC` en la hoja "Matrix" con `YES` o `NO` para cada una de las 21 tecnologías agregadas.

### Paso 5: Ejecutar el Pipeline

```bash
python A0_generate_tech_country_matrix.py  # Regenerar si es necesario
python A1_Pre_processing_OG_csvs.py       # Preprocesar CSVs
python A2_AddTx.py                        # Agregar transmisión
python D1_generate_editor_template.py     # Generar editor
python D2_update_secondary_techs.py       # Aplicar ajustes OLADE
```

### Paso 6: Validar

Ejecutar el script de validación (ver sección siguiente) para verificar completitud e integridad referencial.

---

## 9. Cadena de Referencia Energética

El flujo de energía en el modelo sigue esta cadena:

```
Recurso             Extracción/Provisión    Generación         Transmisión         Demanda
─────────           ────────────────────    ──────────         ───────────         ───────
RNWSPVNCCXX ──────────────────────────────> PWRSPVNCCXX ──┐
RNWHYDNCCXX ──────────────────────────────> PWRHYDNCCXX ──┤
RNWWONNCCXX ──────────────────────────────> PWRWONNCCXX ──┤
                                                           ├──> PWRTRNNCCXX ──> ELCNCCXX02
MINCOANCC ──> COANCC ─────────────────────> PWRCOANCCXX ──┤      (demanda)
MINGASNCC ──> GASNCC ─────────────────────> PWRCCGNCCXX ──┤
MINOILNCC ──> OILNCC ─────────────────────> PWROILNCCXX ──┤
MINURNNCC ──> URNNCC ─────────────────────> PWRURNNCCXX ──┤
                                                           │
COAINT ──────────────> COANCC ─────────────> (importación) ┤
GASINT ──────────────> GASNCC ─────────────> (importación) ┘

TRNNCCXXOTRXX ──────────────────────────────────────────> (exportación a otro país)
TRNOTRXXNCCXX <──────────────────────────────────────── (importación desde otro país)
```

**Flujo de electricidad:**
1. `ELC{NCC}XX01` = electricidad generada por plantas PWR
2. `PWRTRN{NCC}XX` convierte `ELCNCCXX01` → `ELCNCCXX02`
3. `ELC{NCC}XX02` = electricidad en punto de demanda final

---

## 10. Notas y Automatizaciones

### 10.1 Notas sobre la Estructura de Datos

- **Formato de TECHNOLOGY.csv:** El archivo `TECHNOLOGY.csv` mantiene el formato original de OSeMOSYS Global con sufijos `00` (capacidad existente) y `01` (nueva inversión). Este es el formato correcto para estos archivos de entrada. La consolidación a formato sin sufijos se realiza posteriormente en el pipeline de preprocesamiento (`A1_Pre_processing_OG_csvs.py`).

- **Demanda de Haití (HTI):** La ausencia de `ELCHTIXX02` en `SpecifiedAnnualDemand.csv` es esperada en esta etapa. Scripts posteriores del pipeline se encargan de agregar la demanda.

- **Emisiones globales en EMISSION.csv:** Las 152 emisiones de países no-RELAC son herencia del dataset global de OSeMOSYS Global y no requieren limpieza — se filtran durante el procesamiento.

### 10.2 Automatización: Generador de País

Se incluye el script `Z_generate_country_template.py` que automatiza la creación de todas las filas necesarias para un nuevo país. El script:

1. Toma como entrada un **país de referencia** existente (ej: ARG)
2. Clona todos los registros de ese país en todos los archivos CSV
3. Reemplaza únicamente el **código de país** (ej: ARG → NCC)
4. Genera los archivos CSV listos para revisar y mergear

```bash
# Generar template para un nuevo país NCC basado en ARG
python Z_generate_country_template.py --new NCC --ref ARG

# Luego: revisar y ajustar valores, y ejecutar el merge
python templates/NCC/merge_into_inputs.py
```

Los valores numéricos (costos, capacidades, factores) se copian del país de referencia y deben ser ajustados manualmente con los datos reales del nuevo país.

---

## Apéndice A: Estructura Completa de Archivos

```
t1_confection/OG_csvs_inputs/
├── SETS (11 archivos)
│   ├── REGION.csv              (1 entry: GLOBAL)
│   ├── YEAR.csv                (28 entries: 2023-2050)
│   ├── TIMESLICE.csv           (12 entries: S1D1-S4D3)
│   ├── SEASON.csv              (4 entries)
│   ├── DAYTYPE.csv             (1 entry)
│   ├── DAILYTIMEBRACKET.csv    (3 entries)
│   ├── MODE_OF_OPERATION.csv   (2 entries)
│   ├── TECHNOLOGY.csv          (1,007 entries)
│   ├── FUEL.csv                (415 entries)
│   ├── EMISSION.csv            (171 entries)
│   └── STORAGE.csv             (50 entries)
│
├── PARAMETERS - Con datos (29 archivos)
│   ├── Costos: CapitalCost, FixedCost, VariableCost, CapitalCostStorage
│   ├── Capacidad: ResidualCapacity, OperationalLife, CapacityToActivityUnit
│   │              TotalAnnualMaxCapacity, TotalAnnualMaxCapacityInvestment
│   ├── Rendimiento: CapacityFactor, AvailabilityFactor
│   │                InputActivityRatio, OutputActivityRatio
│   ├── Demanda: SpecifiedAnnualDemand, SpecifiedDemandProfile
│   ├── Emisiones: EmissionActivityRatio
│   ├── Almacenamiento: ResidualStorageCapacity, OperationalLifeStorage
│   │                   StorageLevelStart, TechnologyToStorage, TechnologyFromStorage
│   ├── Restricciones: TotalTechnologyAnnualActivityUpperLimit
│   │                  ReserveMarginTagTechnology, ReserveMarginTagFuel
│   └── Temporal: YearSplit, DaySplit, Conversionls, Conversionld, Conversionlh
│
└── PARAMETERS - Vacíos (24 archivos)
    ├── DiscountRate, DiscountRateStorage, DepreciationMethod
    ├── AccumulatedAnnualDemand
    ├── AnnualEmissionLimit, AnnualExogenousEmission, EmissionsPenalty
    ├── ModelPeriodEmissionLimit, ModelPeriodExogenousEmission
    ├── TotalAnnualMinCapacity, TotalAnnualMinCapacityInvestment
    ├── CapacityOfOneTechnologyUnit
    ├── TotalTechnologyAnnualActivityLowerLimit
    ├── TotalTechnologyModelPeriodActivityLowerLimit
    ├── TotalTechnologyModelPeriodActivityUpperLimit
    ├── REMinProductionTarget, RETagTechnology, RETagFuel
    ├── ReserveMargin
    ├── StorageMaxChargeRate, StorageMaxDischargeRate, MinStorageCharge
    ├── TradeRoute
    └── DaysInDayType
```

## Apéndice B: Nomenclatura Completa de Tecnologías

### Prefijos de Tecnología

| Prefijo | Significado | Descripción |
|---------|------------|-------------|
| `PWR` | Power | Planta de generación eléctrica |
| `MIN` | Mining | Extracción/provisión de commodity fósil |
| `RNW` | Renewable | Provisión de recurso renovable local |
| `TRN` | Transmission | Interconexión internacional |

### Tipos de Tecnología (21 + Backstop + Transmisión)

| Código | Tipo | Renovable | Descripción |
|--------|------|-----------|-------------|
| BIO | Biomass | Sí | Biomasa |
| BCK | Backstop | — | Tecnología ficticia de respaldo |
| CCG | Combined Cycle Gas | No | Ciclo combinado gas natural |
| CCS | Carbon Capture & Storage | No | Captura y almacenamiento CO2 |
| COA | Coal | No | Carbón mineral |
| COG | Cogeneration | No | Cogeneración |
| CSP | Concentrated Solar Power | Sí | Solar térmica de concentración |
| GEO | Geothermal | Sí | Geotérmica |
| HYD | Hydroelectric | Sí | Hidroeléctrica |
| LDS | Long Duration Storage | — | Almacenamiento larga duración |
| NGS | Natural Gas (unified) | No | Gas natural (versión RELAC) |
| OCG | Open Cycle Gas | No | Ciclo abierto gas natural |
| OIL | Oil / Fuel Oil | No | Petróleo pesado |
| OTH | Other | No | Otra tecnología |
| PET | Petroleum / Diesel | No | Diesel y derivados |
| SDS | Short Duration Storage | — | Almacenamiento corta duración |
| SPV | Solar Photovoltaic | Sí | Solar fotovoltaica |
| URN | Uranium / Nuclear | No | Energía nuclear |
| WAS | Waste | Sí | Residuos sólidos |
| WAV | Wave | Sí | Energía undimotriz |
| WOF | Wind Offshore | Sí | Eólica marina |
| WON | Wind Onshore | Sí | Eólica terrestre |
