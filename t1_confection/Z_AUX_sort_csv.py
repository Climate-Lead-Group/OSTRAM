# -*- coding: utf-8 -*-
"""
Created on Tue May 27 11:05:00 2025

@author: ClimateLeadGroup
"""

import os
import pandas as pd

def sort_csv_files_in_folder(folder_path):
    if not os.path.isdir(folder_path):
        print(f"La ruta proporcionada no es válida: {folder_path}")
        return

    for filename in os.listdir(folder_path):
        if filename.endswith(".csv"):
            file_path = os.path.join(folder_path, filename)
            print(f"Procesando: {filename}")
            try:
                # Leer el CSV preservando la cabecera
                df = pd.read_csv(file_path)

                # Ordenar usando todas las columnas
                df_sorted = df.sort_values(by=list(df.columns))

                # Sobrescribir el archivo original
                df_sorted.to_csv(file_path, index=False)
            except Exception as e:
                print(f"Error procesando {filename}: {e}")

    print("Todos los archivos han sido procesados.")

# Ejemplo de uso
if __name__ == "__main__":
    folder = input("Ingresa la ruta de la carpeta con los archivos CSV: ").strip()
    sort_csv_files_in_folder(folder)
