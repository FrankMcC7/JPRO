import warnings, pandas as pd
warnings.filterwarnings('ignore', category=UserWarning)          # mute date noise

# ---------- 1 ▸ patch openpyxl’s date converter everywhere ----------
import openpyxl.utils.datetime as dtutil
_orig_from_excel = dtutil.from_excel                        # keep original

def safe_from_excel(*args, **kwargs):
    """Return None when Excel serial dates overflow Python’s range."""
    try:
        return _orig_from_excel(*args, **kwargs)
    except OverflowError:
        return None

dtutil.from_excel = safe_from_excel                         # primary patch
import openpyxl.worksheet._reader as wsreader
wsreader.from_excel = safe_from_excel                       # secondary import

# ---------- 2 ▸ file locations (edit as needed) ---------------------
NAV_TRACKER = r"C:\path\to\NAV Tracker.xlsm"
ATE_FILE    = r"C:\path\to\ATE File.csv"
OUTPUT_CSV  = r"C:\path\to\RF Supplement.csv"

# ---------- 3 ▸ load NAV Tracker (five columns, as text) ------------
nav_cols = ["Fund GCI", "ECA India Analyst", "Fund Manager GCI",
            "Trigger/Non-Trigger", "NAV Source"]
nav_df = pd.read_excel(NAV_TRACKER, sheet_name="Portfolio",
                       usecols=nav_cols, dtype=str, engine="openpyxl")

rf_df = nav_df.copy()
rf_df["NPS Trigger"] = ""

# ---------- 4 ▸ load & clean ATE CSV -------------------------------
ate_df = (pd.read_csv(ATE_FILE, skiprows=1, dtype=str)
            .rename(columns=str.lower)
            .apply(lambda col: col.str.strip().str.lower()))

gcis_with_nav = set(ate_df.loc[
    ate_df["trigger type"].str.contains("nav per share", na=False),
    "fund gci"
])

# ---------- 5 ▸ flag GCIs ------------------------------------------
rf_df["NPS Trigger"] = (rf_df["Fund GCI"].str.strip().str.lower()
                        .apply(lambda g: "Yes" if g in gcis_with_nav else "No"))

rf_df.to_csv(OUTPUT_CSV, index=False)
print(f"✅  RF Supplement created at {OUTPUT_CSV}")
