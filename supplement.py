import pandas as pd
import warnings

# Suppress pandas UserWarnings (e.g. date‐overflow)
warnings.filterwarnings('ignore', category=UserWarning)

# 1. Hard‐coded file locations
nav_tracker_path = r'C:\path\to\NAV Tracker.xlsm'
ate_path         = r'C:\path\to\ATE File.csv'
output_path      = r'C:\path\to\RF Supplement.csv'

# 2. Read only the five needed columns as strings
nav_df = pd.read_excel(
    nav_tracker_path,
    sheet_name='Portfolio',
    usecols=[
        'Fund GCI',
        'ECA India Analyst',
        'Fund Manager GCI',
        'Trigger/Non-Trigger',
        'NAV Source'
    ],
    dtype=str,
    engine='openpyxl'
)  :contentReference[oaicite:0]{index=0}

# 3. Prepare RF Supplement
rf_df = nav_df.copy()
rf_df['NPS Trigger'] = ''

# 4. Load ATE CSV (skip header row) as strings
ate_df = pd.read_csv(ate_path, skiprows=1, dtype=str)

# 5. Normalize columns & values
ate_df.columns        = ate_df.columns.str.strip().str.lower()
ate_df['trigger type'] = ate_df['trigger type'].str.strip().str.lower()
ate_df['fund gci']      = ate_df['fund gci'].str.strip().str.lower()

# 6. Find all GCI with any 'nav per share' entry
mask = ate_df['trigger type'].str.contains('nav per share', na=False)
gcis_with_nav = set(ate_df.loc[mask, 'fund gci'])

# 7. Flag Yes/No in RF Supplement
rf_df['NPS Trigger'] = (
    rf_df['Fund GCI']
      .astype(str)
      .str.strip()
      .str.lower()
      .apply(lambda x: 'Yes' if x in gcis_with_nav else 'No')
)

# 8. Write out and notify
rf_df.to_csv(output_path, index=False)
print(f"RF Supplement created at: {output_path}")
