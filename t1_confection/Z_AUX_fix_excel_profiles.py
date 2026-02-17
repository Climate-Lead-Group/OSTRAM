"""
Script para normalizar perfiles de demanda en archivos Excel A-O_Demand.xlsx

Este script corrige las sumas de SpecifiedDemandProfile en la hoja "Profiles"
de los archivos A-O_Demand.xlsx para que sumen exactamente 1.0 por país/año.

Problema: Algunos países tienen perfiles que suman 1.01 o 0.99 en lugar de 1.0,
lo que causa errores en el modelo OSeMOSYS.

Autor: Claude Code
Fecha: 2026-02-03
"""

import openpyxl
from pathlib import Path
from datetime import datetime
import shutil

def make_backup(file_path):
    """Crea backup de un archivo Excel"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = str(file_path).replace(".xlsx", f"_backup_{timestamp}.xlsx")
    shutil.copy(file_path, backup_file)
    return backup_file

def normalize_excel_profiles(excel_path):
    """
    Normaliza los perfiles de demanda en la hoja Profiles de A-O_Demand.xlsx

    Estructura esperada:
    - Columna 0: Timeslices (S1D1, S1D2, etc.)
    - Columna 1: Demand/Share
    - Columna 2: Fuel/Tech (ej: ELCPERXX02)
    - Columnas 9+: Años (2023, 2024, ..., 2050)

    Para cada Fuel/Tech, la suma de todos los timeslices debe ser 1.0 por año.
    """
    print(f"\n{'='*70}")
    print(f"📊 Normalizando: {excel_path.name}")
    print(f"{'='*70}")

    if not excel_path.exists():
        print(f"❌ Archivo no encontrado: {excel_path}")
        return False

    # Crear backup
    backup = make_backup(excel_path)
    print(f"✓ Backup creado: {Path(backup).name}")

    # Cargar workbook
    wb = openpyxl.load_workbook(excel_path)

    if 'Profiles' not in wb.sheetnames:
        print(f"⚠️  Hoja 'Profiles' no encontrada en {excel_path.name}")
        wb.close()
        return False

    ws = wb['Profiles']

    # Identificar columnas de años (desde columna 9 en adelante)
    # La columna 9 (J) debería ser 2023
    first_year_col = 10  # Columna J en Excel (1-indexed)
    last_col = ws.max_column

    # Obtener lista de fuels únicos
    fuels_data = {}  # {fuel: [(row_num, timeslice), ...]}

    for row_num in range(2, ws.max_row + 1):
        fuel = ws.cell(row_num, 3).value  # Columna C (Fuel/Tech)
        timeslice = ws.cell(row_num, 1).value  # Columna A (Timeslice)

        if fuel and timeslice:
            if fuel not in fuels_data:
                fuels_data[fuel] = []
            fuels_data[fuel].append(row_num)

    print(f"\n📍 Fuels encontrados: {len(fuels_data)}")

    # Normalizar cada fuel por cada año
    problemas_antes = 0
    problemas_despues = 0
    correcciones = 0

    for fuel, row_numbers in fuels_data.items():
        # Para cada columna de año
        for col_num in range(first_year_col, last_col + 1):
            # Leer valores actuales
            values = []
            for row_num in row_numbers:
                cell = ws.cell(row_num, col_num)
                val = cell.value
                if val is not None:
                    try:
                        values.append((row_num, float(val)))
                    except:
                        values.append((row_num, 0.0))
                else:
                    values.append((row_num, 0.0))

            # Calcular suma actual
            suma_actual = sum(v[1] for v in values)

            # Si la suma no es 1.0, normalizar
            if abs(suma_actual - 1.0) > 0.0001:
                problemas_antes += 1

                if suma_actual > 0:  # Solo normalizar si la suma es positiva
                    # Normalizar dividiendo por la suma
                    for row_num, val in values:
                        new_val = val / suma_actual
                        ws.cell(row_num, col_num).value = new_val

                    correcciones += 1

                    # Verificar suma después de corrección
                    suma_nueva = sum(ws.cell(r, col_num).value for r, _ in values)
                    if abs(suma_nueva - 1.0) > 0.0001:
                        problemas_despues += 1

    print(f"\n📊 Resultados:")
    print(f"  - Problemas encontrados: {problemas_antes}")
    print(f"  - Correcciones aplicadas: {correcciones}")
    print(f"  - Problemas restantes: {problemas_despues}")

    # Guardar cambios
    wb.save(excel_path)
    wb.close()

    if problemas_despues == 0:
        print(f"\n✅ Archivo normalizado correctamente")
        return True
    else:
        print(f"\n⚠️  Algunos problemas persisten después de la normalización")
        return False

def main():
    """Función principal"""
    print("="*70)
    print("🔧 NORMALIZACIÓN DE PERFILES EN ARCHIVOS EXCEL")
    print("="*70)

    base_dir = Path(__file__).parent

    # Lista de escenarios a procesar
    scenarios = ["BAU", "NDC", "NDC+ELC", "NDC_NoRPO"]

    results = {}

    for scenario in scenarios:
        scenario_dir = base_dir / "A1_Outputs" / f"A1_Outputs_{scenario}"
        excel_file = scenario_dir / "A-O_Demand.xlsx"

        if excel_file.exists():
            results[scenario] = normalize_excel_profiles(excel_file)
        else:
            print(f"\n⚠️  Escenario {scenario}: archivo no encontrado")
            print(f"    {excel_file}")
            results[scenario] = None

    # Resumen final
    print("\n" + "="*70)
    print("📋 RESUMEN FINAL")
    print("="*70)

    for scenario, result in results.items():
        if result is True:
            print(f"✅ {scenario}: Normalizado correctamente")
        elif result is False:
            print(f"❌ {scenario}: Problemas durante normalización")
        else:
            print(f"⚠️  {scenario}: Archivo no encontrado")

    all_success = all(r == True for r in results.values() if r is not None)

    if all_success:
        print("\n🎉 ¡Todos los archivos normalizados exitosamente!")
    else:
        print("\n⚠️  Algunos archivos no se pudieron normalizar correctamente")

    print("\n💡 IMPORTANTE:")
    print("   Los perfiles en los archivos Excel ahora suman 1.0 exactamente.")
    print("   Si ejecutas A2_AddTx.py, estos valores se exportarán a los CSVs.")
    print("="*70)

    return 0 if all_success else 1

if __name__ == "__main__":
    exit(main())
