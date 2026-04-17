# -*- coding: utf-8 -*-
"""
Created on 2025

@author: Climate Lead Group, Andrey Salazar-Vargas
"""

import os
import pandas as pd
import yaml
import subprocess
import sys
import platform  
import shutil
import time
from datetime import date, datetime
import multiprocessing as mp
import math
from typing import List, Any
from pathlib import Path
import numpy as np

########################################################################################
def ensure_env_tool_paths():
    """Expose the active Python environment's executable folders to subprocesses."""
    env_root = Path(sys.executable).resolve().parent
    candidate_dirs = [
        env_root / "Scripts",
        env_root / "Library" / "bin",
        env_root / "bin",
    ]
    current_path = os.environ.get("PATH", "")
    path_entries = current_path.split(os.pathsep) if current_path else []
    for candidate in candidate_dirs:
        candidate_str = str(candidate)
        if candidate.exists() and candidate_str not in path_entries:
            os.environ["PATH"] = candidate_str + os.pathsep + os.environ.get("PATH", "")
            path_entries.insert(0, candidate_str)


def get_env_executable(executable_name):
    """Return the full path to an executable inside the active environment when available."""
    ensure_env_tool_paths()
    env_root = Path(sys.executable).resolve().parent
    suffix = ".exe" if platform.system() == "Windows" else ""
    candidate_dirs = [
        env_root / "Scripts",
        env_root / "Library" / "bin",
        env_root / "bin",
    ]
    for candidate_dir in candidate_dirs:
        candidate = candidate_dir / f"{executable_name}{suffix}"
        if candidate.exists():
            return str(candidate)
    return executable_name


def sort_csv_files_in_folder(folder_path):
    if not os.path.isdir(folder_path):
        print(f"La ruta es inválida: {folder_path}")
        return
    print('################################################################')
    print('Ordenar archivos csv.')
    for filename in sorted(os.listdir(folder_path)):
        if filename.endswith(".csv"):
            file_path = os.path.join(folder_path, filename)
            print(f"Procesando: {filename}")
            try:
                # Leer el CSV preservando el encabezado
                df = pd.read_csv(file_path)

                # Ordenar usando todas las columnas
                df_sorted = df.sort_values(by=list(df.columns))

                # Sobrescribir el archivo original
                df_sorted.to_csv(file_path, index=False)
            except Exception as e:
                print(f"Error procesando {filename}: {e}")

    print("✅ Todos los archivos fueron ordenados.")
    print('################################################################\n')

def process_scenario_folder(base_input_path, template_path, base_output_path, scenario_name):
    """
    Procesa una carpeta de escenario: lee sus archivos CSV, los alinea con la estructura de plantilla,
    mapea 'Value' a 'VALUE', excluye columnas específicas y guarda los resultados en la salida.
    También asegura que VALUE sea int() para ciertos archivos de plantilla.
    """

    # Paso 1: Definir ruta de entrada del escenario
    scenario_input_path = os.path.join(base_input_path, scenario_name)

    # Paso 2: Omitir si no es directorio o es 'Default'
    if not os.path.isdir(scenario_input_path) or scenario_name == 'Default':
        return

    # Paso 3: Leer y limpiar CSVs del escenario
    scenario_files = {}
    for f in sorted(os.listdir(scenario_input_path)):
        if f.endswith('.csv'):
            df = pd.read_csv(os.path.join(scenario_input_path, f))

            # Eliminar columnas no deseadas
            df = df.drop(columns=[col for col in ['PARAMETERT', 'Scenario'] if col in df.columns])
            df = df.dropna(axis=1, how='all')

            # Renombrar 'Value' a 'VALUE'
            if 'Value' in df.columns:
                df = df.rename(columns={'Value': 'VALUE'})

            scenario_files[f] = df

    # Paso 4: Leer archivos de plantilla
    template_files = {
        f: pd.read_csv(os.path.join(template_path, f))
        for f in sorted(os.listdir(template_path))
        if f.endswith('.csv')
    }

    # Paso 5: Crear ruta de salida
    scenario_output_path = os.path.join(base_output_path, scenario_name)
    os.makedirs(scenario_output_path, exist_ok=True)
    
    # Paso 6: Llenar plantillas con datos del escenario
    for template_name, template_df in template_files.items():
        output_file_path = os.path.join(scenario_output_path, template_name)
        
        if template_name in scenario_files:
            input_df = scenario_files[template_name]
            common_columns = [col for col in template_df.columns if col in input_df.columns]
            filled_df = template_df.copy()
            filled_df[common_columns] = input_df[common_columns]

            # Paso 7: Convertir VALUE a int si es necesario
            if template_name in [
                'DAYTYPE.csv', 'DAILYTIMEBRACKET.csv', 'SEASON.csv',
                'MODE_OF_OPERATION.csv', 'YEAR.csv', 'EMISSION.csv',
                'FUEL.csv', 'REGION.csv', 'STORAGE.csv', 'TECHNOLOGY.csv',
                'TIMESLICE.csv', 'Conversionls.csv'
            ]:
                if 'VALUE' in filled_df.columns:
                    # Eliminar filas con NaN o cadena vacía (incluyendo solo espacios)
                    filled_df = filled_df[filled_df['VALUE'].notna() & (filled_df['VALUE'].astype(str).str.strip() != '')]
            
                    # Convertir a int si es necesario
                    if template_name in [
                        'DAYTYPE.csv', 'DAILYTIMEBRACKET.csv', 'SEASON.csv',
                        'MODE_OF_OPERATION.csv', 'YEAR.csv'
                    ]:
                        filled_df['VALUE'] = filled_df['VALUE'].astype(int)

            filled_df.to_csv(output_file_path, index=False)
        else:
            template_df.to_csv(output_file_path, index=False)
            
    folder_to_sort = os.path.join(base_output_path,scenario_name)
    sort_csv_files_in_folder(folder_to_sort)

    print(f"✅ Escenario '{scenario_name}': plantillas completadas y guardadas exitosamente.\n")
    print('#------------------------------------------------------------------------------#')

def run_otoole_conversion(base_output_path, scenario_name, params):
    """
    Ejecuta el comando corregido 'otoole convert csv datafile' para un escenario dado.

    Parámetros:
        base_output_path (str): Ruta donde se almacenan los archivos CSV del escenario.
        scenario_name (str): El nombre del escenario.
        params (dict): Diccionario cargado desde el archivo YAML con rutas requeridas.
    """
    # Paso 1: Definir rutas
    input_folder = os.path.join(base_output_path, scenario_name)
    scenario_exec_dir = os.path.join(HERE, params['executables'], scenario_name + '_0')
    output_file = os.path.join(scenario_exec_dir, f"{scenario_name}_0.txt")
    config_file = os.path.join(HERE, params['Miscellaneous'], params['otoole_config'])

    # Paso 2: Asegurar que la carpeta ejecutable del escenario exista
    os.makedirs(scenario_exec_dir, exist_ok=True)

    # Paso 3: Construir el comando
    otoole_exe = get_env_executable('otoole')
    command = [
        otoole_exe, 'convert', 'csv', 'datafile',
        input_folder,
        output_file,
        config_file
    ]

    print(f"Ejecutando comando: {' '.join(command)}")

    # Paso 4: Ejecutar el comando
    result = subprocess.run(command, capture_output=True, text=True)

    # Paso 5: Manejar salida
    if result.returncode != 0:
        print(f"❌ Error al convertir escenario '{scenario_name}':\n{result.stderr}")
        print('#------------------------------------------------------------------------------#')
        return False
    else:
        print(f"✅ Escenario '{scenario_name}' convertido exitosamente.\n{result.stdout}")
        print('#------------------------------------------------------------------------------#')
        return True

def run_preprocessing_script(params, scenario_name):
    """
    Ejecuta el script de preprocesamiento Python especificado en el archivo YAML de parámetros para un escenario dado.

    Parámetros:
        params (dict): Parámetros cargados desde el archivo YAML.
        scenario_name (str): El nombre del escenario a preprocesar.
    """
    # Paso 1: Definir rutas
    script_path = os.path.join(params['Miscellaneous'], params['preprocess_data'])
    input_file = os.path.join(params['executables'], scenario_name + '_0', f"{scenario_name}_0.txt")
    output_file = os.path.join(params['executables'], scenario_name + '_0', f"{params['preprocess_data_name']}{scenario_name}_0.txt")

    # Paso 2: Construir comando
    command = [sys.executable, script_path, input_file, output_file]

    print(f"Ejecutando script de preprocesamiento para escenario '{scenario_name}_0':")
    print(' '.join(command))

    # Paso 3: Ejecutar el script
    result = subprocess.run(command, capture_output=True, text=True)

    # Paso 4: Resultado de salida
    if result.returncode != 0:
        print(f"❌ Error durante preprocesamiento del escenario '{scenario_name}':\n{result.stderr}")
        print('#------------------------------------------------------------------------------#')
    else:
        print(f"✅ Preprocesamiento completado para escenario '{scenario_name}':\n{result.stdout}")
        print('#------------------------------------------------------------------------------#')

def check_enviro_variables(solver_command):
    ensure_env_tool_paths()
    # Determinar el comando según el sistema operativo
    command = 'where' if platform.system() == 'Windows' else 'which'

    # Ejecutar el comando apropiado
    where_solver = subprocess.run([command, solver_command], capture_output=True, text=True)
    paths = where_solver.stdout.splitlines()
    
    if paths:  # Asegurar que al menos una ruta fue encontrada
        path_solver = paths[0]
        
        # Verificar si la ruta ya está en la variable de entorno PATH
        if path_solver not in os.environ["PATH"]:
            # Si no está en PATH, agregarla
            os.environ["PATH"] += os.pathsep + path_solver
            print("Ruta agregada:", path_solver)
    else:
        print(f"No se encontró '{solver_command}' en el sistema.")
    #

def get_config_main_path(here, base_folder):
    # Navegar a la raíz del repositorio (padre de t1_confection)
    repo_root = Path(here).parent
    return str(repo_root / base_folder)

def main_executer(params, scenario_name, HERE):
    
    folder_scenario = os.path.join(HERE, params['executables'], scenario_name + '_0')                             
    
    # Construir rutas para el archivo de datos y el archivo de salida, adaptando a diferencias del sistema de archivos
    data_file = os.path.join(folder_scenario, params['preprocess_data_name'] + scenario_name + '_0')
    output_file = os.path.join(folder_scenario, params['preprocess_data_name'] + scenario_name + '_0' + params['output_files'])
    this_case = scenario_name + '_0.txt'

    # Determinar el solver según los parámetros
    solver = params['solver']
    commands = []

    if solver == 'glpk':
        if params['execute_model']:
            # Usando opciones más nuevas de GLPK
                                       
            check_enviro_variables('glpsol')
            
            # Componer el comando para resolver el modelo con las nuevas opciones
            str_solve = f'glpsol -m {params["osemosys_model"]} -d {data_file}.txt --wglp {output_file}.glp --write {output_file}.sol'
            commands.append(str_solve)
        
    else:
        if params['create_matrix']:
            # Para modelos LP
            str_solve = f'glpsol -m {params["osemosys_model"]} -d {data_file}.txt --wlp {output_file}.lp --check'
            commands.append(str_solve)
        
        if solver == 'cbc':
            # Usando solver CBC
            if params['execute_model']:
                if os.path.exists(output_file + '.sol'):
                    os.remove(output_file + '.sol')

                check_enviro_variables('cbc')

                # Obtener semilla aleatoria para reproducibilidad
                cbc_random_seed = params.get('cbc_random_seed', 12345)

                # Componer el comando para solver CBC con semillas aleatorias para comportamiento determinístico
                str_solve = f'cbc {output_file}.lp randomSeed {cbc_random_seed} randomCbcSeed {cbc_random_seed} -seconds {params["iteration_time"]} solve -solu {output_file}.sol'
                commands.append(str_solve)
            
        elif solver == 'cplex':
            # Usando solver CPLEX
            if params['execute_model']:
                if os.path.exists(output_file + '.sol'):
                    os.remove(output_file + '.sol')

                # Número de hilos que usa cplex
                cplex_threads = params['cplex_threads']

                # Obtener semilla aleatoria para reproducibilidad
                cplex_random_seed = params.get('cplex_random_seed', 12345)

                check_enviro_variables('cplex')

                # Componer el comando para solver CPLEX con semilla aleatoria para comportamiento determinístico
                str_solve = f'cplex -c "read {output_file}.lp" "set threads {cplex_threads}" "set randomseed {cplex_random_seed}" "set parallel 1" "optimize" "write {output_file}.sol"'
                commands.append(str_solve)

        elif solver == 'gurobi':
            # Usando solver Gurobi
            if params['execute_model']:
                if os.path.exists(output_file + '.sol'):
                    os.remove(output_file + '.sol')

                # Número de hilos que usa gurobi
                gurobi_threads = params['gurobi_threads']

                # Obtener semilla aleatoria para reproducibilidad
                gurobi_seed = params.get('gurobi_seed', 12345)

                check_enviro_variables('gurobi_cl')

                # Componer el comando para solver Gurobi con semilla para comportamiento determinístico
                str_solve = f'gurobi_cl Threads={gurobi_threads} Seed={gurobi_seed} ResultFile={output_file}.sol {output_file}.lp'
                commands.append(str_solve)

    if params['execute_model'] or params['create_matrix']:
        for cmd in commands:
            subprocess.run(cmd, shell=True, check=True)
        
    print(f'✅ Escenario {scenario_name}_0 resuelto exitosamente.')
    print('\n#------------------------------------------------------------------------------#')

    # Rutas para convertir salidas
    file_path_conv_format = os.path.join(HERE, params['Miscellaneous'], params['conv_format'])
    # file_path_template = os.path.join(params['Miscellaneous'], params['templates'])
    file_path_template = os.path.join(HERE, params['A2_output_otoole'], scenario_name)
    file_path_outputs = os.path.join(folder_scenario, params['outputs'])

    # Convertir salidas de .sol a formato csv
    if solver == 'glpk' and params['glpk_option'] == 'new':
        str_outputs = f'"{get_env_executable("otoole")}" results {solver} csv "{output_file}.sol" "{file_path_outputs}" datafile "{data_file}.txt" "{file_path_conv_format}" --glpk_model "{output_file}.glp"'
        if params['execute_model']:
            subprocess.run(str_outputs, shell=True, check=True)

    elif solver in ['cbc', 'cplex', 'gurobi']:

        str_outputs = f'"{get_env_executable("otoole")}" results {solver} csv "{output_file}.sol" "{file_path_outputs}" csv "{file_path_template}" "{file_path_conv_format}" 2> "{output_file}.log"'
        if params['execute_model']:
            subprocess.run(str_outputs, shell=True, check=True)

    # Módulo para concatenar salidas csv de otoole
    if solver in ['glpk', 'cbc', 'cplex', 'gurobi']:
        file_conca_csvs = get_config_main_path(HERE, params['concatenate_folder'])
        script_concate_csv = os.path.join(file_conca_csvs, params['concat_csvs'])
        str_otoole_concate_csv = f'"{sys.executable}" -u "{script_concate_csv}" "{file_path_outputs}" "{output_file}"'  # last int is the ID tier
        if params['concat_otoole_csv']:
            subprocess.run(str_otoole_concate_csv, shell=True, check=True)
        print(f'✅ Salidas concatenadas a {scenario_name}_0_Output.csv exitosamente.')
        print('\n#------------------------------------------------------------------------------#')

def delete_files(file, data_file, solver):
    # Eliminar archivos
    if file and os.path.exists(file):
        shutil.os.remove(file)
    if data_file and os.path.exists(data_file):
        shutil.os.remove(data_file)
    
    # Verificar si el archivo .sol existe y está vacío
    log_file = file.replace('.sol', '.log')
    if os.path.exists(log_file) and os.path.getsize(log_file) == 0:
        if os.path.exists(log_file):
            os.remove(log_file)
    
    if solver == 'glpk':
        glp_file = file.replace('sol', 'glp')
        if os.path.exists(glp_file):
            shutil.os.remove(glp_file)
    else:
        lp_file = file.replace('sol', 'lp')
        if os.path.exists(lp_file):
            shutil.os.remove(lp_file)
    
    # Eliminar archivos log cuando el solver es 'cplex' y del_files es True
    if solver == 'cplex':
        for filename in ['cplex.log', 'clone1.log', 'clone2.log']:
            if os.path.exists(filename):
                os.remove(filename)

    # Eliminar archivos log cuando el solver es 'gurobi' y del_files es True
    if solver == 'gurobi':
        if os.path.exists('gurobi.log'):
            os.remove('gurobi.log')

def read_csv_files(input_dir):
    """Lee todos los archivos CSV en el directorio dado y retorna un diccionario de DataFrames."""
    data_dict = {}
    for filename in sorted(os.listdir(input_dir)):
        if filename.endswith(".csv"):
            file_path = os.path.join(input_dir, filename)
            df = pd.read_csv(file_path)
            key = os.path.splitext(filename)[0]
            data_dict[key] = df
    return data_dict

def generate_combined_input_file(input_folder, output_folder, scenario_name):
    """
    Lee CSVs desde input_folder, filtra claves de metadatos, renombra columnas VALUE por clave,
    concatena todos los DataFrames no vacíos, ordena columnas y guarda el resultado en un archivo CSV.
    """
    keys_sets_delete = ['REGION', 'YEAR', 'TECHNOLOGY', 'FUEL', 'EMISSION', 'MODE_OF_OPERATION',
                        'TIMESLICE', 'STORAGE', 'SEASON', 'DAYTYPE', 'DAILYTIMEBRACKET']

    inputs_dataframes = []
    print(input_folder)
    print(sorted(os.listdir(input_folder)))
    for filename in sorted(os.listdir(input_folder)):
        if not filename.endswith(".csv"):
            continue
        key = filename.replace(".csv", "")
        if key in keys_sets_delete:
            continue
        path = os.path.join(input_folder, filename)
        df = pd.read_csv(path)
        if df.empty or 'VALUE' not in df.columns:
            continue
        df = df.rename(columns={'VALUE': key})
        inputs_dataframes.append(df)

    if not inputs_dataframes:
        print("[Advertencia] No se encontraron dataframes válidos para concatenar.")
        return None, None

    # Concatenar todos los dataframes no vacíos
    inputs_data = pd.concat(inputs_dataframes, ignore_index=True, sort=True)  # Ordenar para orden de columnas determinístico

    # Reordenar columnas
    present_keys = [col for col in keys_sets_delete if col in inputs_data.columns]
    other_columns = sorted([col for col in inputs_data.columns if col not in present_keys])
    inputs_data = inputs_data[present_keys + other_columns]

    # Guardar a CSV
    os.makedirs(output_folder, exist_ok=True)
    output_path = os.path.join(output_folder, f"{scenario_name}_Input.csv")
    inputs_data.to_csv(output_path, index=False)

    print(f'✅ Inputs concatenados a {scenario_name}_Input.csv exitosamente.')
    print('\n#------------------------------------------------------------------------------#')

    return output_path, inputs_data.head()


def export_root_datafile(here, params, scenario_name, export_name='OSTRAM_data.txt'):
    """
    Copy the preprocessed main-scenario datafile to the repository root so the
    user has a single easy-to-find model datafile next to `t1_confection/`.
    """
    repo_root = Path(here).parent
    source_path = (
        Path(here)
        / params['executables']
        / f"{scenario_name}_0"
        / f"{params['preprocess_data_name']}{scenario_name}_0.txt"
    )
    target_path = repo_root / export_name

    if not source_path.exists():
        print(f"[WARN] Root datafile export skipped because source was not found: {source_path}")
        return None

    shutil.copy2(source_path, target_path)
    print(f"✅ Datafile exported to repository root: {target_path}")
    print('#------------------------------------------------------------------------------#')
    return target_path



def concatenate_all_scenarios(HERE, params):
    """
    Itera sobre todas las carpetas de escenarios en `base_input_path` (excluyendo 'Default'),
    lee archivos *_Input.csv y *_Output.csv, agrega columnas de metadatos de escenario, los concatena
    en archivos CSV únicos para inputs, outputs y combinado, y devuelve sus rutas.

    Args:
        params (dict):
          - executables (str): Ruta al directorio base que contiene las carpetas de escenarios.
          - prefix_final_files (str): Carpeta/ruta donde guardar los resultados.
          - inputs_file (str): Nombre base para el CSV de inputs.
          - outputs_file (str): Nombre base para el CSV de outputs.
          - combined_file (str, opcional): Nombre base para el CSV combinado inputs+outputs.
    Returns:
        tuple: (input_csv_path, output_csv_path, combined_csv_path)
    """
    # Columnas de metadatos que movemos al frente
    keys_sets_delete = [
        'REGION','YEAR','TECHNOLOGY','FUEL','EMISSION','MODE_OF_OPERATION',
        'TIMESLICE','STORAGE','SEASON','DAYTYPE','DAILYTIMEBRACKET'
    ]

    combined_inputs = []
    combined_outputs = []
    combined_inputs_outputs = []
    base_input_path = params['executables']

    for scenario_future_name in sorted(os.listdir(base_input_path)):
        if scenario_future_name.lower() in ['default', '__pycache__', 'local_dataset_creator_0.py']:
            continue

        scenario_path = os.path.join(HERE, base_input_path, scenario_future_name)
        parts = scenario_future_name.rsplit("_", 1)
        scenario = parts[0]
        future = parts[1]

        input_file = os.path.join(scenario_path, f"{scenario_future_name}_Input.csv")
        output_file = os.path.join(scenario_path, f"Pre_processed_{scenario_future_name}_Output.csv")

        if os.path.exists(input_file):
            df_in = pd.read_csv(input_file, low_memory=False)
            df_in.insert(0, "Future", future)
            df_in.insert(1, "Scenario", scenario)
            combined_inputs.append(df_in)
            combined_inputs_outputs.append(df_in)

        if os.path.exists(output_file):
            df_out = pd.read_csv(output_file, low_memory=False)
            df_out.insert(0, "Future", future)
            df_out.insert(1, "Scenario", scenario)
            combined_outputs.append(df_out)
            combined_inputs_outputs.append(df_out)

    # Concatenar inputs y outputs por separado
    df_inputs_all = pd.concat(combined_inputs, ignore_index=True) if combined_inputs else pd.DataFrame()
    df_outputs_all = pd.concat(combined_outputs, ignore_index=True) if combined_outputs else pd.DataFrame()
    # df_inputs_outputs_all = pd.concat(combined_inputs_outputs, ignore_index=True) if combined_inputs_outputs else pd.DataFrame()
    # df_list = []
    # df_list.append(combined_inputs)
    # df_list.append(combined_outputs)
    df_inputs_outputs_all = pd.concat([df_inputs_all,df_outputs_all], ignore_index=True, sort=True)  # Sort for deterministic column order
    

    # Función para reordenar columnas: metadatos primero, luego alfabético
    def reorder_columns(df):
        front = ['Future','Scenario'] + [c for c in keys_sets_delete if c in df.columns]
        rest = sorted([c for c in df.columns if c not in front])
        return df[front + rest]

    today = date.today().isoformat()  # 'YYYY-MM-DD'

    # 1) Guardar inputs
    if not df_inputs_all.empty:
        df_inputs_all = reorder_columns(df_inputs_all)
        # Ordenar filas para salida determinística
        sort_cols = [c for c in ['Future', 'Scenario', 'REGION', 'TECHNOLOGY', 'YEAR'] if c in df_inputs_all.columns]
        if sort_cols:
            df_inputs_all = df_inputs_all.sort_values(by=sort_cols).reset_index(drop=True)
        path_in = os.path.join(HERE,params['prefix_final_files'] + params['inputs_file'])
        df_inputs_all.to_csv(path_in, index=False)
        dated = path_in.replace('.csv', f'_{today}.csv')
        df_inputs_all.to_csv(dated, index=False)
    else:
        path_in = None

    # 2) Guardar outputs
    if not df_outputs_all.empty:
        df_outputs_all = reorder_columns(df_outputs_all)
        # Ordenar filas para salida determinística
        sort_cols = [c for c in ['Future', 'Scenario', 'REGION', 'TECHNOLOGY', 'YEAR'] if c in df_outputs_all.columns]
        if sort_cols:
            df_outputs_all = df_outputs_all.sort_values(by=sort_cols).reset_index(drop=True)
        path_out = os.path.join(HERE,params['prefix_final_files'] + params['outputs_file'])
        df_outputs_all.to_csv(path_out, index=False)
        dated = path_out.replace('.csv', f'_{today}.csv')
        df_outputs_all.to_csv(dated, index=False)
    else:
        path_out = None

    # 3) Nuevamente, combinar ambos DataFrames en uno solo y guardarlo
    combined_name = params.get('combined_file', 'Combined_Inputs_Outputs.csv')
    if not df_inputs_outputs_all.empty and not df_outputs_all.empty:
        # df_combined = pd.concat([df_inputs_all, df_outputs_all],
        #                         ignore_index=True, sort=False)
        df_combined = reorder_columns(df_inputs_outputs_all)
        # Ordenar filas para salida determinística
        sort_cols = [c for c in ['Future', 'Scenario', 'REGION', 'TECHNOLOGY', 'YEAR'] if c in df_combined.columns]
        if sort_cols:
            df_combined = df_combined.sort_values(by=sort_cols).reset_index(drop=True)
        
        
        #########################################################################################
        # Calcular AccumulatedTotalAnnualMinCapacityInvestment
        # Debe agrupar por (Future, Scenario, TECHNOLOGY) y acumular dentro de cada grupo
        if "TotalAnnualMinCapacityInvestment" in df_combined.columns:
            df = df_combined.copy()

            # Inicializar la columna acumulada con NaN
            df['AccumulatedTotalAnnualMinCapacityInvestment'] = np.nan

            # Definir columnas de agrupación (excluir YEAR ya que acumulamos sobre años)
            group_cols = ['Future', 'Scenario', 'TECHNOLOGY']
            group_cols = [c for c in group_cols if c in df.columns]

            if group_cols:
                # Ordenar por columnas de grupo + YEAR para asegurar orden correcto para cumsum
                sort_cols = group_cols + ['YEAR']
                df = df.sort_values(by=sort_cols).reset_index(drop=True)

                # Calcular suma acumulativa dentro de cada grupo
                # Solo para filas que tengan un valor en TotalAnnualMinCapacityInvestment
                mask = df['TotalAnnualMinCapacityInvestment'].notna()
                df.loc[mask, 'AccumulatedTotalAnnualMinCapacityInvestment'] = (
                    df.loc[mask]
                    .groupby(group_cols, sort=False)['TotalAnnualMinCapacityInvestment']
                    .cumsum()
                )
            else:
                # Respaldo: si no hay columnas de grupo, hacer un cumsum simple
                mask = df['TotalAnnualMinCapacityInvestment'].notna()
                df.loc[mask, 'AccumulatedTotalAnnualMinCapacityInvestment'] = (
                    df.loc[mask, 'TotalAnnualMinCapacityInvestment'].cumsum()
                )

            df_combined = df
        #########################################################################################
        
        
        path_comb = os.path.join(HERE,params['prefix_final_files'] + combined_name)
        df_combined.to_csv(path_comb, index=False)
        # Nota: la copia con fecha con datos anualizados se creará después de la anualización (si está habilitada)
    else:
        path_comb = None

    return path_in, path_out, path_comb






def chunk_scenarios(
    scenarios: List[Any],
    max_x_per_iter: int,
) -> List[List[Any]]:
    """
    Divide la lista de entrada ``scenarios`` en bloques de tamaño ``max_x_per_iter``.

    Parámetros
    ----------
    scenarios : List[Any]
        La lista que contiene todos los valores de escenarios.
    max_x_per_iter : int
        Número máximo de elementos permitidos en cada bloque.

    Retorna
    -------
    List[List[Any]]
        Una lista donde cada elemento es una sub-lista de ``scenarios`` con longitud
        de hasta ``max_x_per_iter``.
    """
    if max_x_per_iter <= 0:
        raise ValueError("max_x_per_iter must be a positive integer")

    # Construir los bloques usando slicing en una comprensión
    scenarios_list_max_per_iter: List[List[Any]] = [
        scenarios[i : i + max_x_per_iter]  # noqa: E203 (spacing around :)
        for i in range(0, len(scenarios), max_x_per_iter)
    ]
    return scenarios_list_max_per_iter

########################################################################################
if __name__ == "__main__":
    # Iniciar temporizador
    start1 = time.time()
    
    # Carpeta donde vive este script: .../OSTRAM/t1_confection
    global HERE
    def get_here() -> Path:
        # 1) Script normal (ejecución directa)
        if '__file__' in globals():
            return Path(__file__).resolve().parent
        # 2) Algunos IDEs exponen __main__.__file__
        main = sys.modules.get('__main__')
        if hasattr(main, '__file__'):
            return Path(main.__file__).resolve().parent
        # 3) Ejecución en consola/interactiva: directorio de trabajo actual
        return Path.cwd().resolve()
    
    HERE = get_here()
    
    
    # (Opcional) Cambiar CWD a la carpeta del script
    if Path.cwd() != HERE:
        os.chdir(HERE)
        print(f"[INFO] Working dir -> {HERE}")
        
    # Cargar parámetros desde YAML
    with open('Config_MOMF_T1_AB.yaml', 'r') as f:
        params = yaml.safe_load(f)
        
    # Cargar parámetros desde YAML
    with open('Config_MOMF_T1_A.yaml', 'r') as f:
        params_A2 = yaml.safe_load(f)
    
    # Definir rutas base de origen y destino
    base_input_path = os.path.join(HERE, params['A2_output'])
    template_path = os.path.join(HERE, params['Miscellaneous'], params['templates'])
    base_output_path = os.path.join(HERE, params['A2_output_otoole'])

    scenarios=sorted(os.listdir(base_input_path))
    try:
        scenarios.remove('Default')
    except ValueError:
        pass
    
    if params['only_main_scenario']:
        scenarios = []
        scenarios.append(params_A2['xtra_scen']['Main_Scenario'])

    main_scenario_name = params_A2['xtra_scen']['Main_Scenario']
    
    ###############################################################################################
    # Escribir modelo txt
    for scenario_name in scenarios:
         
        if params['A2_otoole_outputs']:
            process_scenario_folder(
                base_input_path=base_input_path,
                template_path=template_path,
                base_output_path=base_output_path,
                scenario_name=scenario_name
            )
        if params['write_txt_model']:
            conversion_ok = run_otoole_conversion(
                base_output_path=base_output_path,
                scenario_name=scenario_name,
                params=params
            )

            if conversion_ok:
                run_preprocessing_script(params, scenario_name)
            else:
                print(f"❌ Se omite el preprocesamiento para '{scenario_name}' porque falló la conversión con otoole.")
                print('#------------------------------------------------------------------------------#')
                continue


        input_folder = os.path.join(HERE, base_output_path, scenario_name)
        output_folder = os.path.join(HERE, params['executables'], scenario_name + '_0')
        
        # Listar archivos disponibles para vista previa (solo para verificar configuración)
        os.makedirs(input_folder, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)
        
        # Concatenar inputs
        generate_combined_input_file(input_folder, output_folder, scenario_name + '_0')

        #
    ###############################################################################################

    if params['write_txt_model'] and main_scenario_name in scenarios:
        export_root_datafile(HERE, params, main_scenario_name)
        
        
        
    ###############################################################################################
    # Ejecutar modelo txt
    if params['execute_model'] or params['create_matrix']:
        if params['parallel']:
            print('Iniciada paralelización de ejecución del modelo')
            max_x_per_iter = params['max_x_per_iter'] # FLAG: This is an input
            scenarios_list_max_per_iter = chunk_scenarios(scenarios, max_x_per_iter)
            #
            for scens_list in scenarios_list_max_per_iter:
                processes = []
                for scenario_name in scens_list:
                    p = mp.Process(target=main_executer, args=(params, scenario_name, HERE) )
                    processes.append(p)
                    p.start()
                #
                for process in processes:
                    process.join()
            
        # Esto es para la versión lineal
        else:
            print('Iniciadas ejecuciones lineales')
            for scenario_num in scenarios:
                main_executer(params, scenario_num, HERE)
    
    ###############################################################################################
    # Eliminar archivos
    for scenario_name in scenarios:
        # Eliminar carpeta Outputs con archivos csvs de otoole
        if params['del_files']:
            # Eliminar carpeta Outputs con archivos csvs de otoole
            folder_scenario = os.path.join(HERE, params['executables'], scenario_name + '_0') 
            outputs_otoole_csvs = os.path.join(HERE, folder_scenario, params['outputs'])
            data_file = os.path.join(HERE, folder_scenario, scenario_name + '_0' + '.txt')
            sol_file = os.path.join(HERE, folder_scenario, params['preprocess_data_name'] + scenario_name + '_0' + params['output_files'] + '.sol')
            if os.path.exists(outputs_otoole_csvs):
                shutil.rmtree(outputs_otoole_csvs)
        
            # Eliminar archivos glp, lp, txt y sol
            if params['solver'] in ['glpk', 'cbc', 'cplex']:
                delete_files(sol_file, data_file, params['solver'])
            
            print(f'✅ Archivos intermedios del escenario {scenario_name}_0 eliminados exitosamente.')
            print('\n#------------------------------------------------------------------------------#')

    ###############################################################################################

    end_1 = time.time()   
    time_elapsed_1 = -start1 + end_1
    print( str( time_elapsed_1 ) + ' segundos /', str( time_elapsed_1/60 ) + ' minutos' )

    start2 = time.time()
    
    ###############################################################################################
    # Concatenar inputs y outputs
    if params['concat_scenarios_csv']:
        input_output_path, output_output_path, combined_output_path = concatenate_all_scenarios(HERE,params)
        print(f'✅ Inputs y outputs concatenados para todos los escenarios exitosamente.')
        print(f'Los archivos son: ({input_output_path}), ({output_output_path}) y ({combined_output_path})')
    ###############################################################################################

    ###############################################################################################
    # Anualizar inversión de capital
    if params.get('annualize_capital', False):
        try:
            print('\n')
            print('#'*80)
            print('# ANUALIZACIÓN DE INVERSIÓN DE CAPITAL')
            print('#'*80)

            # Importar la función de anualización
            from Z_AUX_capital_annualization_script import annualize_capital_investment

            # Definir la ruta al archivo combinado
            combined_file_path = os.path.join(HERE, params['prefix_final_files'] + 'Combined_Inputs_Outputs.csv')

            # Verificar si el archivo existe
            if os.path.exists(combined_file_path):
                print(f'Iniciando anualización para: {combined_file_path}')

                # Llamar a la función de anualización
                annualize_capital_investment(
                    input_file_path=combined_file_path,
                    verbose=True
                )

                print(f'✅ Anualización de inversión de capital completada exitosamente.')

                # Crear copia con fecha con datos anualizados
                today = date.today().isoformat()  # 'YYYY-MM-DD'
                dated_combined = combined_file_path.replace('.csv', f'_{today}.csv')
                shutil.copy2(combined_file_path, dated_combined)
                print(f'✅ Archivo anualizado copiado a: {dated_combined}')
                print('#'*80)
            else:
                print(f'⚠️  ADVERTENCIA: Archivo combinado no encontrado en {combined_file_path}')
                print('Omitiendo anualización de inversión de capital.')
                print('#'*80)

        except Exception as e:
            print(f'❌ ERROR durante anualización de inversión de capital: {e}')
            print('Continuando sin anualización...')
            import traceback
            traceback.print_exc()
            print('#'*80)
    else:
        # Si la anualización está deshabilitada, aún crear copia con fecha del archivo combinado
        combined_file_path = os.path.join(HERE, params['prefix_final_files'] + 'Combined_Inputs_Outputs.csv')
        if os.path.exists(combined_file_path):
            today = date.today().isoformat()  # 'YYYY-MM-DD'
            dated_combined = combined_file_path.replace('.csv', f'_{today}.csv')
            shutil.copy2(combined_file_path, dated_combined)
            print(f'✅ Archivo combinado copiado a: {dated_combined}')
    ###############################################################################################




    # # 1. Carga los dataframes desde los CSV
    # df_inputs_all = pd.read_csv('REALC_TX_Inputs.csv', low_memory=False)
    # df_outputs_all = pd.read_csv('REALC_TX_Outputs.csv', low_memory=False)
    
    # # 2. Concatenarlos verticalmente (uno debajo del otro)
    # df_combined = pd.concat([df_inputs_all, df_outputs_all], ignore_index=True, sort=False)
    
    # # 3. (Opcional) Reordenar columnas si se desea,
    # #    por ejemplo, poniendo 'Scenario' y 'Future' al frente
    # cols_front = ['Scenario', 'Future']
    # other_cols = [c for c in df_combined.columns if c not in cols_front]
    # df_combined = df_combined[cols_front + other_cols]
    
    # # 4. Guarda el CSV combinado
    # today = date.today().isoformat()  # e.g. '2025-07-14'
    # combined_filename = f"{params['prefix_final_files']}Combined_Inputs_Outputs_{today}.csv"
    # df_combined.to_csv(combined_filename, index=False)
    
    # print(f"Archivo combinado guardado en: {combined_filename}")
    ###############################################################################################
    
    
    end_2 = time.time()   
    time_elapsed_2 = -start2 + end_2
    print( str( time_elapsed_2 ) + ' segundos /', str( time_elapsed_2/60 ) + ' minutos' )
    print('\n#------------------------------------------------------------------------------#')
    
    time_elapsed_3 = -start1 + end_2
    print( str( time_elapsed_3 ) + ' seconds /', str( time_elapsed_3/60 ) + ' minutes' )
    print('*: For all effects, we have finished the work of this script.')

            

