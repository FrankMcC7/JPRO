"""
RF-Supplement builder
--------------------------------
• Reads five columns from the NAV Tracker (xlsm).  
• Reads the ATE file (csv) after dropping its first row.  
• Flags GCIs that have at least one “NAV per share” trigger.  
• Writes the result to RF Supplement.csv.  

Author: <your-name> | Date: <today>
"""
import warnings
warnings.filterwarnings('ignore', category=UserWarning)          # mute date warnings

# ------------------------------------------------------------------
# 1) Patch OpenPyXL’s date converter everywhere it is referenced
# ------------------------------------------------------------------
import openpyxl.utils.datetime as dtutil
import openpyxl.worksheet._reader as wsreader          # same function imported a 2nd time

_original_from_excel = dtutil.from_excel

def _safe_from_excel(value, datemode):
    """Return None instead of raising OverflowError on bad serial dates."""
    try:
        return _original_from_excel(value, datemode)
    except OverflowError:                               # out-of-bounds serial
        return None

dtutil.from_excel   = _safe_from_excel                  # patch primary reference
wsreader.from_excel = _safe_from_excel                  # patch secondary import

# ------------------------------------------------------------------
# 2) Main processing
# ------------------------------------------------------------------
import pandas as pd

# --- Hard-coded paths ------------------------------------------------
NAV_TRACKER   = r"C:\path\to\NAV Tracker.xlsm"          # macro-enabled workbook
ATE_FILE      = r"C:\path\to\ATE File.csv"              # ATE as csv
OUTPUT_CSV    = r"C:\path\to\RF Supplement.csv"         # destination file

# --- 2.1  Read required columns from NAV Tracker --------------------
nav_cols = [
    "Fund GCI",
    "ECA India Analyst",
    "Fund Manager GCI",
    "Trigger/Non-Trigger",
    "NAV Source",
]
nav_df = pd.read_excel(
    NAV_TRACKER,
    sheet_name="Portfolio",
    usecols=nav_cols,
    dtype=str,                                         # everything as text
    engine="openpyxl"
)

# --- 2.2  Build RF Supplement skeleton ------------------------------
rf_df = nav_df.copy()
rf_df["NPS Trigger"] = ""                               # placeholder

# --- 2.3  Read & normalise ATE CSV ----------------------------------
ate_df = pd.read_csv(ATE_FILE, skiprows=1, dtype=str)   # drop the first row
ate_df.columns          = ate_df.columns.str.strip().str.lower()

ate_df["trigger type"]  = ate_df["trigger type"].str.strip().str.lower()
ate_df["fund gci"]      = ate_df["fund gci"].str.strip().str.lower()

# --- 2.4  Determine which GCIs have any ‘nav per share’ trigger -----
has_nav_mask   = ate_df["trigger type"].str.contains("nav per share", na=False)
gcis_with_nav  = set(ate_df.loc[has_nav_mask, "fund gci"])

# --- 2.5  Populate NPS Trigger flag ---------------------------------
rf_df["NPS Trigger"] = (
    rf_df["Fund GCI"]
      .astype(str)
      .str.strip()
      .str.lower()
      .apply(lambda gci: "Yes" if gci in gcis_with_nav else "No")
)

# --- 2.6  Save & notify ---------------------------------------------
rf_df.to_csv(OUTPUT_CSV, index=False)
print(f"✅  RF Supplement created at {OUTPUT_CSV}")
