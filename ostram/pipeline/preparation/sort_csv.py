# -*- coding: utf-8 -*-
"""Sort every CSV in a folder by all columns and overwrite it in place.

Standalone utility; the pipeline uses its own copy of the routine in
base_inputs.py.

Inputs:
    A folder path typed at the prompt
Outputs:
    The same CSV files, rewritten sorted (no backup)
Usage:
    python -B -m ostram.pipeline.preparation.sort_csv
Authors:
    2025-05-27 Climate Lead Group, Andrey Salazar-Vargas
"""

import os
import pandas as pd

def sort_csv_files_in_folder(folder_path):
    if not os.path.isdir(folder_path):
        print(f"The provided path is not valid: {folder_path}")
        return

    for filename in os.listdir(folder_path):
        if filename.endswith(".csv"):
            file_path = os.path.join(folder_path, filename)
            print(f"Processing: {filename}")
            try:
                # Read the CSV preserving the header
                df = pd.read_csv(file_path)

                # Sort using all columns
                df_sorted = df.sort_values(by=list(df.columns))

                # Overwrite the original file
                df_sorted.to_csv(file_path, index=False)
            except Exception as e:
                print(f"Error processing {filename}: {e}")

    print("All files have been processed.")

# Usage example
if __name__ == "__main__":
    folder = input("Enter the path to the folder with CSV files: ").strip()
    sort_csv_files_in_folder(folder)
