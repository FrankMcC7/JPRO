import pandas as pd

# 1. Hard-coded file locations
nav_tracker_path = r'C:\path\to\NAV Tracker.xlsm'   # macro-enabled file
ate_path         = r'C:\path\to\ATE File.xlsx'
output_path      = r'C:\path\to\RF Supplement.xlsx'

# 2. Read 'Portfolio' sheet from NAV Tracker and select table columns
nav_df = pd.read_excel(
    nav_tracker_path,
    sheet_name='Portfolio',
    engine='openpyxl'
)[['Fund GCI', 'ECA India Analyst', 'Fund Manager GCI',
   'Trigger/Non-Trigger', 'NAV Source']]

# 3. Create RF Supplement DataFrame and add empty 'NPS Trigger'
rf_df = nav_df.copy()
rf_df['NPS Trigger'] = ''

# 5. Read ATE file, skip first row, convert to DataFrame
ate_df = pd.read_excel(
    ate_path,
    skiprows=1,  # delete first row
    engine='openpyxl'
)

# 6. Determine which Fund GCI have 'NAV per share' in Trigger Type
mask = ate_df['Trigger Type'].astype(str).str.contains('NAV per share', case=False, na=False)
gcis_with_nav = set(ate_df.loc[mask, 'Fund GCI'])

# 7. Populate 'NPS Trigger' with Yes/No
rf_df['NPS Trigger'] = rf_df['Fund GCI'].apply(
    lambda x: 'Yes' if x in gcis_with_nav else 'No'
)

# 3 & final: Write to new Excel file
rf_df.to_excel(output_path, index=False)

print(f"RF Supplement created at: {output_path}")
