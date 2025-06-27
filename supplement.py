import pandas as pd
import warnings

# Suppress pandas UserWarnings (e.g. date outside bounds)
warnings.filterwarnings('ignore', category=UserWarning)

# 1. Hard-coded file locations
nav_tracker_path = r'C:\path\to\NAV Tracker.xlsm'   # Excel macro-enabled
ate_path         = r'C:\path\to\ATE File.csv'       # Now CSV
output_path      = r'C:\path\to\RF Supplement.csv'  # Output as CSV

# 2. Read 'Portfolio' sheet from NAV Tracker
nav_df = pd.read_excel(
    nav_tracker_path,
    sheet_name='Portfolio',
    engine='openpyxl'
)[[
    'Fund GCI',
    'ECA India Analyst',
    'Fund Manager GCI',
    'Trigger/Non-Trigger',
    'NAV Source'
]]

# 3. Build RF Supplement and add empty 'NPS Trigger'
rf_df = nav_df.copy()
rf_df['NPS Trigger'] = ''

# 5. Read ATE CSV, skip its first row
ate_df = pd.read_csv(
    ate_path,
    skiprows=1
)

# 6. Find which Fund GCI entries have any 'NAV per share' in Trigger Type
mask = ate_df['Trigger Type'].astype(str).str.contains('NAV per share', case=False, na=False)
gcis_with_nav = set(ate_df.loc[mask, 'Fund GCI'])

# 7. Populate 'NPS Trigger' with Yes/No
rf_df['NPS Trigger'] = rf_df['Fund GCI'].apply(
    lambda x: 'Yes' if x in gcis_with_nav else 'No'
)

# 3 & Final: Write out to CSV
rf_df.to_csv(output_path, index=False)

print(f"RF Supplement created at: {output_path}")
