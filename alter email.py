# -*- coding: utf-8 -*-
"""
Approved Funds pipeline (empty-table safe + no signature)

- Skips junk first row, uses second row as headers.
- Saves to Excel Table, counts totals and Alternative Fund? == Yes.
- Sends Outlook email with robust HTML table (even when 0 rows).
- Subject: "ACTION required by CPO: Weekly Alternative Fund Status Review"
- Ends with "Regards," and a hard-coded USER_FULL_NAME (no Outlook signature).
"""

import os
import sys
import csv
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
CSV_PATH = r"D:\path\to\your\approved_funds.csv"        # <-- Hard-code the CSV file path
OUTPUT_DIR = r"D:\path\to\your"                          # <-- Hard-code the output folder
SENDER_EMAIL = "your.alias@company.com"                  # <-- Outlook account to send from
TO_LIST = ["recipient1@company.com", "recipient2@company.com"]  # <-- TO recipients
CC_LIST = ["cc1@company.com"]                            # <-- CC recipients
CSV_ENCODING = "utf-8-sig"                               # <-- Change if your CSV encoding differs
USER_FULL_NAME = "Your Name"                             # <-- <--- HARD-CODE YOUR NAME HERE

# Columns to include in the embedded HTML table (Alternative Fund? == Yes only)
TABLE_COLUMNS = [
    "Investment Manager",
    "IM CoPER",
    "Fund Name",
    "Fund GCI",
    "PDO",
    "PMA",
    "Alternative Fund?"
]

EMAIL_SUBJECT = "ACTION required by CPO: Weekly Alternative Fund Status Review"

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
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"

def normalize_colname(s: str) -> str:
    import re
    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

def find_column(df: pd.DataFrame, target_name: str) -> str:
    target_norm = normalize_colname(target_name)
    mapping = {c: normalize_colname(c) for c in df.columns}
    for orig, norm in mapping.items():
        if norm == target_norm:
            return orig
    for orig, norm in mapping.items():
        if norm.startswith("alternativefund"):
            return orig
    raise KeyError(f"Column '{target_name}' not found. Columns seen: {list(df.columns)}")

# =========================
# ====== CORE STEPS =======
# =========================

def read_and_prepare_dataframe(csv_path):
    """
    Row 1: junk/title
    Row 2: real headers
    """
    try:
        df = pd.read_csv(
            csv_path,
            skiprows=1, header=0,
            sep=None, engine="python",
            dtype=str, keep_default_na=False,
            quoting=csv.QUOTE_MINIMAL,
            encoding=CSV_ENCODING
        )
    except Exception as e1:
        df = None
        for sep in [",", ";", "|", "\t"]:
            try:
                df = pd.read_csv(
                    csv_path,
                    skiprows=1, header=0,
                    sep=sep, engine="python",
                    dtype=str, keep_default_na=False,
                    quoting=csv.QUOTE_MINIMAL,
                    encoding=CSV_ENCODING
                )
                break
            except Exception:
                df = None
        if df is None:
            raise RuntimeError(f"Failed to parse CSV. Original error: {e1}")

    # Clean headers, trim strings
    df.columns = [str(c).strip() for c in df.columns]
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(lambda x: x.strip() if isinstance(x, str) else x)

    alt_col = find_column(df, "Alternative Fund?")
    df[alt_col] = df[alt_col].astype(str).str.strip()
    df[alt_col] = df[alt_col].apply(
        lambda v: "Yes" if v.lower() == "yes" else ("No" if v.lower() == "no" else v)
    )
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

def build_html_table(data_rows, columns):
    """
    Build a robust HTML table for Outlook (even when 0 rows).
    data_rows: list of dicts (already filtered to Alternative == Yes)
    columns: ordered list of column names to display
    """
    # Basic inline styles for Outlook reliability
    th_style = "background:#f2f2f2;border:1px solid #d0d0d0;padding:6px 8px;text-align:left;"
    td_style = "border:1px solid #d0d0d0;padding:6px 8px;text-align:left;"
    table_style = "border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:11pt;"

    # Header row
    thead = "<thead><tr>" + "".join([f'<th style="{th_style}">{col}</th>' for col in columns]) + "</tr></thead>"

    # Body rows
    if data_rows:
        trs = []
        for row in data_rows:
            tds = []
            for col in columns:
                val = row.get(col, "")
                tds.append(f'<td style="{td_style}">{val}</td>')
            trs.append("<tr>" + "".join(tds) + "</tr>")
        tbody = "<tbody>" + "".join(trs) + "</tbody>"
    else:
        # Zero rows: show a single full-width row
        tbody = f'<tbody><tr><td colspan="{len(columns)}" style="{td_style}">No records found this week.</td></tr></tbody>'

    return f'<table style="{table_style}">{thead}{tbody}</table>'

def compose_email_html(monday, friday, total_rows, alt_yes_count, table_html, user_full_name):
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

      <p>Regards,<br>{user_full_name}</p>
    </div>
    """
    return body

def create_and_send_email(sender_email, to_list, cc_list, subject, html_body, attachment_path):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Try to send using specified account
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

    df, alt_col = read_and_prepare_dataframe(CSV_PATH)

    total_rows = len(df)
    alt_yes_mask = df[alt_col].astype(str).str.strip().str.lower() == "yes"
    alt_yes_count = int(alt_yes_mask.sum())

    last_fri = get_last_friday()
    mon, fri = monday_to_friday_window(last_fri)

    out_filename = f"Approved Funds_{ddmmyyyy(last_fri)}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    out_path = save_as_xlsx_table(df, out_path)

    # Build data rows (dicts) for the HTML table in the original display column names
    display_df = df.copy()
    if alt_col != "Alternative Fund?":
        display_df = display_df.rename(columns={alt_col: "Alternative Fund?"})

    filtered = display_df.loc[alt_yes_mask, :]
    data_rows = []
    for _, r in filtered.iterrows():
        row = {col: ("" if col not in filtered.columns else ("" if pd.isna(r.get(col)) else str(r.get(col)))) for col in TABLE_COLUMNS}
        data_rows.append(row)

    table_html = build_html_table(data_rows, TABLE_COLUMNS)
    body_html = compose_email_html(mon, fri, total_rows, alt_yes_count, table_html, USER_FULL_NAME)

    create_and_send_email(SENDER_EMAIL, TO_LIST, CC_LIST, EMAIL_SUBJECT, body_html, out_path)

    print("Email sent and task completed successfully.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR:", str(e))
        sys.exit(1)