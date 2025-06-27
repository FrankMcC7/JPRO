import warnings
warnings.filterwarnings('ignore', category=UserWarning)

# --- Monkey-patch openpyxl’s date parser to avoid overflow errors ---
import openpyxl.utils.datetime as dtutil
_orig = dtutil.from_excel
def _safe_from_excel(val, datemode):
    try:
        return _orig(val, datemode)
    except OverflowError:
        return None
dtutil.from_excel = _safe_from_excel

import pandas as pd

# 1. File paths
nav_tracker_path = r'C:\path\to\NAV Tracker.xlsm'
ate_path         = r'C:\path\to\ATE File.csv'
output_path      = r'C:\path\to\RF Supplement.csv'

# 2. Read only needed columns as strings
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
)

# 3. Build RF Supplement
rf_df = nav_df.copy()
rf_df['NPS Trigger'] = ''

# 4. Load ATE CSV (skip first row) as strings
ate_df = pd.read_csv(ate_path, skiprows=1, dtype=str)

# 5. Normalize headers & values
ate_df.columns         = ate_df.columns.str.strip().str.lower()
ate_df['trigger type'] = ate_df['trigger type'].str.strip().str.lower()
ate_df['fund gci']     = ate_df['fund gci'].str.strip().str.lower()

# 6. Identify GCIs with any “nav per share”
mask = ate_df['trigger type'].str.contains('nav per share', na=False)
gcis_with_nav = set(ate_df.loc[mask, 'fund gci'])

# 7. Flag Yes/No
rf_df['NPS Trigger'] = (
    rf_df['Fund GCI']
       .astype(str)
       .str.strip()
       .str.lower()
       .apply(lambda x: 'Yes' if x in gcis_with_nav else 'No')
)

# 8. Output and notify
rf_df.to_csv(output_path, index=False)
print(f"RF Supplement created at: {output_path}")
