import sys, warnings, pandas as pd
warnings.filterwarnings('ignore', category=UserWarning)          # silence noisy cast msgs

# ---------- 1 ▸ robust patch for ALL from_excel aliases -------------
import openpyxl.utils.datetime as dtutil
_orig = dtutil.from_excel

def _safe_from_excel(*args, **kwargs):
    """Return None instead of raising OverflowError on bad Excel serials."""
    try:
        return _orig(*args, **kwargs)
    except OverflowError:
        return None

# replace every existing alias that points at the original function
for name, mod in sys.modules.items():
    if name.startswith("openpyxl") and hasattr(mod, "from_excel"):
        if getattr(mod, "from_excel") is _orig:
            setattr(mod, "from_excel", _safe_from_excel)

# ---------- 2 ▸ paths ------------------------------------------------
NAV_TRACKER = r"C:\path\to\NAV Tracker.xlsm"
ATE_FILE    = r"C:\path\to\ATE File.csv"
OUTPUT_CSV  = r"C:\path\to\RF Supplement.csv"

# ---------- 3 ▸ load NAV Tracker ------------------------------------
nav_cols = ["Fund GCI", "ECA India Analyst", "Fund Manager GCI",
            "Trigger/Non-Trigger", "NAV Source"]
nav_df = pd.read_excel(NAV_TRACKER, sheet_name="Portfolio",
                       usecols=nav_cols, dtype=str, engine="openpyxl")

rf_df = nav_df.assign(**{"NPS Trigger": ""})

# ---------- 4 ▸ load & normalise ATE CSV -----------------------------
ate_df = (pd.read_csv(ATE_FILE, skiprows=1, dtype=str)
            .rename(columns=str.lower)
            .apply(lambda c: c.str.strip().str.lower()))

# ---------- 5 ▸ flag GCIs with any “nav per share” trigger ----------
gcis_with_nav = set(ate_df.loc[
    ate_df["trigger type"].str.contains("nav per share", na=False),
    "fund gci"
])
rf_df["NPS Trigger"] = (rf_df["Fund GCI"].str.strip().str.lower()
                        .map(lambda g: "Yes" if g in gcis_with_nav else "No"))

# ---------- 6 ▸ save -------------------------------------------------
rf_df.to_csv(OUTPUT_CSV, index=False)
print(f"✅ RF Supplement created → {OUTPUT_CSV}")
