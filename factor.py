import pandas as pd

# Hard-coded paths (change these to your actual file locations)
input_csv_path = r"C:\path\to\your\input.csv"
output_excel_path = r"C:\path\to\your\output.xlsx"

# Read the CSV data
df = pd.read_csv(input_csv_path)

# Pivot the data: one row per RFAD_Fund_CoperID, one column per factor
pivot_df = (
    df.pivot_table(
        index='RFAD_Fund_CoperID',
        columns='IRR_Scorecard_factor',
        values='IRR_Scorecard_factor_value',
        aggfunc='first'
    )
    .reset_index()
)

# Optionally, flatten the column index if needed
pivot_df.columns.name = None

# Write the transformed data to a new Excel file
pivot_df.to_excel(output_excel_path, index=False)

print(f"Transformed data written to {output_excel_path}")
