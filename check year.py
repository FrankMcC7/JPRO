Python script to identify mislabelled folders beyond 2025

import pandas as pd import os import re

Configurations

EXCEL_FILE = 'folder_paths.xlsx'  # Path to your Excel file SHEET_NAME = 0                 # Sheet index or name COLUMN_NAME = 'path'           # Column containing folder addresses

def load_paths(excel_file, sheet_name, column_name): """Load folder paths from Excel into a list.""" df = pd.read_excel(excel_file, sheet_name=sheet_name) return df[column_name].dropna().tolist()

def find_invalid_year_folders(base_paths, max_year=2025): """Return list of tuples (base_path, invalid_folder) for folders named as year > max_year.""" pattern = re.compile(r'^\d{4}$') invalid = [] for base in base_paths: if not os.path.isdir(base): print(f"Warning: '{base}' is not a valid directory.") continue for name in os.listdir(base): full = os.path.join(base, name) if os.path.isdir(full) and pattern.match(name): if int(name) > max_year: invalid.append((base, name)) return invalid

def main(): paths = load_paths(EXCEL_FILE, SHEET_NAME, COLUMN_NAME) invalid_folders = find_invalid_year_folders(paths) if invalid_folders: print('Invalid year-labelled folders found:') for base, folder in invalid_folders: print(f"- {base}: {folder}") else: print('No mislabelled folders found.')

if name == 'main': main()

