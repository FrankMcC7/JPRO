# -*- coding: utf-8 -*-
"""
Approved Funds pipeline (fixed header handling):
- CSV row 1 is junk; row 2 holds headers.
- Delete row 1, promote row 2 to headers, then process.
- Save as Excel Table, count totals & Alternative=Yes, email via Outlook.
"""

import os
import sys
import datetime as dt
import pandas as pd

# pip install pandas openpyxl pywin32
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import win32com.client as win32

# =========================
# ===== USER SETTINGS =====
# =========================
CSV_PATH = r"D:\path\to\your\approved_funds.csv"       # <-- hard-code CSV path
OUTPUT_DIR = r"D:\path\to\your"                         # <-- hard-code output folder
SENDER_EMAIL = "your.alias@company.com"                 # <-- Outlook account to send from
TO_LIST = ["recipient1@company.com", "recipient2@company.com"]
CC_LIST = ["cc1@company.com"]
SIGNATURE_NAME = "formal"
CSV_ENCODING = "utf-8-sig"

TABLE_COLUMNS = [
    "Investment Manager",
    "IM CoPER",
    "Fund Name",
    "Fund GCI",
    "PDO",
    "PMA",
    "Alternative Fund?"
]

# =========================
# ====== UTILITIES ========
# =========================

def get_last_friday(today=None):
    if today is None:
        today = dt.date.today()
    return today - dt.timedelta(days=(today.weekday() - 4) % 7)

def monday_to_friday_window(last_friday_date):
    return last_friday_date - dt.timedelta(days=4), last_friday_date

def ddmmyyyy(d):
    return d.strftime("%d%m%Y")

def ordinal_day(d):
    n = d.day
    if 11 <= n % 100 <= 13:
        suf = "th"
    else:
        suf = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suf}"

def normalize_colname(s: str) -> str:
    # Lowercase, strip, drop non-alphanumerics to match variants like "Alternative Fund ?" / "Alternative_Fund?"
    import re
    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

def find_column(df: pd.DataFrame, target_name: str) -> str:
    target_norm = normalize_colname(target_name)
    mapping = {c: normalize_colname(c) for c in df.columns}
    for orig, norm in mapping.items():
        if norm == target_norm:
            return orig
    # Soft fallback: look for something starting with 'alternativefund'
    for orig, norm in mapping.items():
        if norm.startswith("alternativefund"):
            return orig
    raise KeyError(f"Column '{target_name}' not found after header promotion. Columns seen: {list(df.columns)}")

def load_signature_html(signature_name):
    sig_dir = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Signatures")
    if not os.path.isdir(sig_dir):
        return ""
    sig_htm = os.path.join(sig_dir, f"{signature_name}.htm")
    if not os.path.isfile(sig_htm):
        return ""
    try:
        with open(sig_htm, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        return ""

# =========================
# ====== CORE STEPS =======
# =========================

def read_and_prepare_dataframe(csv_path):
    """
    Read CSV where:
      - Row 1: junk
      - Row 2: REAL HEADERS
      - Row 3+: data
    Steps:
      1) read with no header
      2) drop first row
      3) set new header from next row
      4) remaining rows = data
    """
    raw = pd.read_csv(csv_path, header=None, encoding=CSV_ENCODING)
    if raw.shape[0] < 2:
        raise ValueError("CSV doesn't have enough rows to promote headers after dropping the first row.")

    # Delete the first (junk) row
    tmp = raw.iloc[1:].reset_index(drop=True)

    # Promote the new first row to header
    new_header = tmp.iloc[0].astype(str).str.strip().tolist()
    df = tmp.iloc[1:].copy()
    df.columns = new_header

    # Trim column names and cell strings
    df.columns = [str(c).strip() for c in df.columns]
    df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

    # Normalize Alternative Fund? values to Yes/No where possible (case-insensitive)
    alt_col = find_column(df, "Alternative Fund?")
    df[alt_col] = df[alt_col].astype(str).str.strip()
    df[alt_col] = df[alt_col].apply(lambda v: "Yes" if str(v).lower() == "yes" else ("No" if str(v).lower() == "no" else str(v)))

    return df, alt_col

def save_as_xlsx_table(df, out_path, table_name="ApprovedFundsTable"):
    df.to_excel(out_path, index=False, sheet_name="Approved Funds")
    wb = load_workbook(out_path)
    ws = wb.active

    ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    tbl = Table(displayName=table_name, ref=ref)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False)
    tbl.tableStyleInfo = style
    ws.add_table(tbl)

    wb.save(out_path)
    return out_path

def dataframe_to_html_table(df, columns):
    cols = [c for c in columns if c in df.columns]
    html = df[cols].to_html(index=False, border=0, justify="left", classes="alt-funds-table")
    style = """
    <style>
      table.alt-funds-table { border-collapse: collapse; font-family: Segoe UI, Arial, sans-serif; font-size: 11pt; }
      .alt-funds-table th, .alt-funds-table td { border: 1px solid #d0d0d0; padding: 6px 8px; }
      .alt-funds-table th { background: #f2f2f2; text-align: left; }
    </style>
    """
    return style + html

def compose_email_html(monday, friday, total_rows, alt_yes_count, table_html, signature_html):
    mon_str = f"{ordinal_day(monday)} {monday.strftime('%B')}"
    fri_str = f"{ordinal_day(friday)} {friday.strftime('%B')}"
    total_str = f"<b>{total_rows}</b>"
    alt_str = f"<b>{alt_yes_count}</b>"

    body = f"""
    <div style="font-family: Segoe UI, Arial, sans-serif; font-size: 11pt; color:#222;">
      <p>Hi All,</p>

      <p>Please find attached approved funds from <b>{mon_str}</b> to <b>{fri_str}</b> in EMEA region.
      Total {total_str} Funds got approved out of which {alt_str} Funds are marked as Alternative fund.</p>

      <p><b>ACTION FOR CPO:</b> Please review approved funds to check their Alternative fund (Column BF) status
      from the attached list or below table for quick reference-</p>

      {table_html}

      <p>Regards,</p>
      {signature_html}
    </div>
    """
    return body

def create_and_send_email(sender_email, to_list, cc_list, subject, html_body, attachment_path):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Attempt to send using specified account (and on behalf)
    try:
        for acct in outlook.Session.Accounts:
            if str(acct.SmtpAddress).lower() == sender_email.lower():
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acct))  # SendUsingAccount
                break
        mail.SentOnBehalfOfName = sender_email
    except Exception:
        pass

    mail.Subject = subject
    mail.HTMLBody = html_body
    if to_list: mail.To = "; ".join(to_list)
    if cc_list: mail.CC = "; ".join(cc_list)
    if attachment_path and os.path.isfile(attachment_path):
        mail.Attachments.Add(Source=attachment_path)
    mail.Send()

# =========================
# ========= MAIN ==========
# =========================

def main():
    if not os.path.isfile(CSV_PATH):
        print(f"ERROR: CSV not found: {CSV_PATH}")
        sys.exit(1)
    if not os.path.isdir(OUTPUT_DIR):
        print(f"ERROR: Output directory not found: {OUTPUT_DIR}")
        sys.exit(1)

    # Read, delete first row, promote new header
    df, alt_col = read_and_prepare_dataframe(CSV_PATH)

    # Counts
    total_rows = len(df)
    alt_yes_count = (df[alt_col].astype(str).str.strip().str.lower() == "yes").sum()

    # Dates
    last_fri = get_last_friday()
    mon, fri = monday_to_friday_window(last_fri)

    # Save Excel with table
    out_filename = f"Approved Funds_{ddmmyyyy(last_fri)}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    out_path = save_as_xlsx_table(df, out_path)

    # Filtered table for Alternative = Yes, and only requested columns (keep order if present)
    alt_yes_df = df[df[alt_col].astype(str).str.strip().str.lower() == "yes"].copy()
    # Ensure the “Alternative Fund?” column name in the email matches your display list
    # If your CSV had variant name, add a pretty alias column:
    if alt_col != "Alternative Fund?":
        alt_yes_df.rename(columns={alt_col: "Alternative Fund?"}, inplace=True)

    table_html = dataframe_to_html_table(alt_yes_df, TABLE_COLUMNS)

    # Body + signature
    signature_html = load_signature_html(SIGNATURE_NAME)
    body_html = compose_email_html(mon, fri, total_rows, alt_yes_count, table_html, signature_html)

    # Subject (per your instruction)
    subject = "ACTION required by CPO: Weekly Alternative Fund Status Review"

    # Send
    create_and_send_email(SENDER_EMAIL, TO_LIST, CC_LIST, subject, body_html, out_path)

    print("Email sent and task completed successfully.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR:", str(e))
        sys.exit(1)