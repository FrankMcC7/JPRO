import pandas as pd
import warnings

# Suppress pandas UserWarnings (e.g. date outside bounds)
warnings.filterwarnings('ignore', category=UserWarning)

# 1. Hard-coded file locations
nav_tracker_path = r'C:\path\to\NAV Tracker.xlsm'   # Excel macro-enabled
ate_path         = r'C:\path\to\ATE File.csv'       # CSV input
output_path      = r'C:\path\to\RF Supplement.csv'  # CSV output

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
ate_df = pd.read_csv(ate_path, skiprows=1)

# 6. Normalize headers and fields
ate_df.columns       = ate_df.columns.str.strip().str.lower()
ate_df['trigger type'] = ate_df['trigger type'].astype(str).str.strip().str.lower()
ate_df['fund gci']     = ate_df['fund gci'].astype(str).str.strip()

# Identify all Fund GCI with any 'nav per share' trigger
mask = ate_df['trigger type'].str.contains('nav per share', na=False)
gcis_with_nav = set(ate_df.loc[mask, 'fund gci'])

# 7. Populate 'NPS Trigger' with Yes/No
rf_df['NPS Trigger'] = (
    rf_df['Fund GCI']
      .astype(str)
      .str.strip()
      .apply(lambda x: 'Yes' if x in gcis_with_nav else 'No')
)

# 8. Write out to CSV and notify
rf_df.to_csv(output_path, index=False)
print(f"RF Supplement created at: {output_path}")
