# Python script to identify mislabelled folders and generate a hyperlink report

import pandas as pd
import os
import re

# Configuration
EXCEL_FILE = 'folder_paths.xlsx'  # Path to input Excel file
SHEET_NAME = 0                    # Sheet index or name
COLUMN_NAME = 'path'              # Column with folder addresses
OUTPUT_FILE = 'invalid_folders_report.xlsx'  # Path for output report
MAX_YEAR = 2025


def load_paths(excel_file, sheet_name, column_name):
    """Load folder paths from Excel into a list."""
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    return df[column_name].dropna().tolist()


def find_invalid_year_folders(base_paths):
    """Return set of base paths containing folders named as years > MAX_YEAR."""
    pattern = re.compile(r'^\d{4}$')
    invalid = set()
    for base in base_paths:
        if not os.path.isdir(base):
            continue
        for name in os.listdir(base):
            if not pattern.match(name):
                continue
            year = int(name)
            if year > MAX_YEAR:
                invalid.add(base)
                break
    return sorted(invalid)


def write_report(invalid_paths, output_file):
    """Write an Excel file with hyperlinks to each invalid base path."""
    df = pd.DataFrame({'Path': invalid_paths})
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='InvalidFolders')
        workbook  = writer.book
        worksheet = writer.sheets['InvalidFolders']
        for idx, path in enumerate(invalid_paths, start=1):
            # Create a clickable hyperlink
            worksheet.write_url(idx, 0, f'file://{path}', string=path)
    print(f'Report generated: {output_file}')


def main():
    paths = load_paths(EXCEL_FILE, SHEET_NAME, COLUMN_NAME)
    invalid_paths = find_invalid_year_folders(paths)
    if invalid_paths:
        write_report(invalid_paths, OUTPUT_FILE)
    else:
        print('No mislabelled folders found.')


if __name__ == '__main__':
    main()


I’ve updated the script to generate an Excel report (invalid_folders_report.xlsx) that lists each parent address containing mislabelled folders. Each path is now a clickable hyperlink, so you can directly jump to the folder locations for review.

