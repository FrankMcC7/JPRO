# -*- coding: utf-8 -*-
"""
Approved Funds pipeline:
1) Read CSV from a hard-coded path
2) Delete the first data row
3) Save to .xlsx as a proper Excel Table
4) Count total rows (approved funds) & count 'Yes' in 'Alternative Fund?'
5) Compute last Friday and the Monday–Friday window for the email text
6) Build an Outlook email (from a hard-coded account), embed a filtered HTML table (Alternative Fund? == Yes),
   attach the .xlsx file, include 'formal' signature, and send.
7) Print 'Task completed' on success.
"""

import os
import sys
import datetime as dt
import pandas as pd

# ---- Optional deps (install via pip if needed) ----
# pip install pandas openpyxl pywin32
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import win32com.client as win32  # Requires Outlook for sending email

# =========================
# ===== USER SETTINGS =====
# =========================
CSV_PATH = r"D:\path\to\your\approved_funds.csv"     # <-- Hard-code your CSV file path here
OUTPUT_DIR = r"D:\path\to\your"                       # <-- Hard-code your output folder (can be os.path.dirname(CSV_PATH))
SENDER_EMAIL = "your.alias@company.com"               # <-- The Outlook account to send from
TO_LIST = ["recipient1@company.com", "recipient2@company.com"]  # <-- Hard-code TO recipients
CC_LIST = ["cc1@company.com"]                         # <-- Hard-code CC recipients (or leave empty [])
SIGNATURE_NAME = "formal"                             # <-- Outlook signature name (as seen in Outlook > Signatures)

# If your CSV uses a specific encoding, set it here (e.g., 'utf-8-sig', 'cp1252', etc.)
CSV_ENCODING = "utf-8-sig"

# Columns to include in the email’s embedded table (in this order)
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
# ====== CORE LOGIC =======
# =========================

def get_last_friday(today=None):
    """Return date of the most recent Friday (including today if Friday)."""
    if today is None:
        today = dt.date.today()
    weekday = today.weekday()  # Mon=0 ... Sun=6
    # Friday is 4
    days_since_friday = (weekday - 4) % 7
    return today - dt.timedelta(days=days_since_friday)

def monday_to_friday_window(last_friday_date):
    """Given a Friday date, return (monday_date, friday_date)."""
    monday = last_friday_date - dt.timedelta(days=4)
    return monday, last_friday_date

def ddmmyyyy(d):
    return d.strftime("%d%m%Y")

def ordinal_day(d):
    """Return '1st', '2nd', '3rd', ... for a date object."""
    n = d.day
    if 11 <= n % 100 <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"

def load_signature_html(signature_name):
    """
    Load the specified Outlook signature's HTML.
    Returns HTML string or empty string if not found.
    """
    # Typical signature folder
    sig_dir = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Signatures")
    if not os.path.isdir(sig_dir):
        return ""
    # Signatures are saved as <name>.htm (plus related files)
    sig_htm = os.path.join(sig_dir, f"{signature_name}.htm")
    if not os.path.isfile(sig_htm):
        return ""
    try:
        with open(sig_htm, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        return ""

def read_and_prepare_dataframe(csv_path):
    """Read CSV, drop the first data row, standardize 'Alternative Fund?' values, return DataFrame."""
    df = pd.read_csv(csv_path, encoding=CSV_ENCODING)
    if df.shape[0] == 0:
        raise ValueError("CSV has no rows.")
    # Delete first data row (keep header)
    if df.shape[0] == 1:
        # If there's only one row, dropping it leaves empty; still proceed per instruction
        df = df.iloc[0:0]
    else:
        df = df.iloc[1:].copy()

    # Trim column names & values
    df.columns = [str(c).strip() for c in df.columns]
    if "Alternative Fund?" not in df.columns:
        raise KeyError("Column 'Alternative Fund?' not found in the CSV.")

    # Normalize 'Alternative Fund?' values to Yes/No (case-insensitive), keep original otherwise
    df["Alternative Fund?"] = df["Alternative Fund?"].astype(str).str.strip()
    df["Alternative Fund?"] = df["Alternative Fund?"].str.replace(r"^\s*$", "No", regex=True)
    df["Alternative Fund?"] = df["Alternative Fund?"].apply(lambda x: "Yes" if str(x).strip().lower() == "yes" else ("No" if str(x).strip().lower() == "no" else str(x).strip()))

    return df

def save_as_xlsx_table(df, out_path, table_name="ApprovedFundsTable"):
    """
    Save DataFrame to .xlsx with an Excel Table and a neutral built-in style.
    Returns the final saved path.
    """
    # Write with pandas first
    df.to_excel(out_path, index=False, sheet_name="Approved Funds")
    wb = load_workbook(out_path)
    ws = wb.active

    # Determine ref range (A1:??)
    max_row = ws.max_row
    max_col = ws.max_column
    from openpyxl.utils import get_column_letter
    ref = f"A1:{get_column_letter(max_col)}{max_row}"

    # Add Table
    tbl = Table(displayName=table_name, ref=ref)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False)
    tbl.tableStyleInfo = style
    ws.add_table(tbl)

    wb.save(out_path)
    return out_path

def dataframe_to_html_table(df, columns):
    """Return a clean HTML table string for the email body."""
    # Keep only requested columns that exist
    cols = [c for c in columns if c in df.columns]
    html = df[cols].to_html(index=False, border=0, justify="left", classes="alt-funds-table")
    # Light styling for readability in Outlook
    style = """
    <style>
      table.alt-funds-table { border-collapse: collapse; font-family: Segoe UI, Arial, sans-serif; font-size: 11pt; }
      .alt-funds-table th, .alt-funds-table td { border: 1px solid #d0d0d0; padding: 6px 8px; }
      .alt-funds-table th { background: #f2f2f2; text-align: left; }
    </style>
    """
    return style + html

def compose_email_html(monday, friday, total_rows, alt_yes_count, table_html, signature_html):
    """
    Build the HTML email body per spec.
    Dates must be bold and shown as '22nd September to 26th September', derived from the Monday–Friday window.
    Also bold the fund counts where indicated.
    """
    mon_str = f"{ordinal_day(monday)} {monday.strftime('%B')}"
    fri_str = f"{ordinal_day(friday)} {friday.strftime('%B')}"
    total_str = f"<b>{total_rows}</b>"
    alt_str = f"<b>{alt_yes_count}</b>"

    # Body per instruction
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

def create_and_send_email(
    sender_email, to_list, cc_list, subject, html_body, attachment_path
):
    """Create and send Outlook email with given HTML body and attachment."""
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem

    # Try to send using the specified account
    session = outlook.Session
    try:
        accounts = session.Accounts
        for acct in accounts:
            if str(acct.SmtpAddress).lower() == sender_email.lower():
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acct))  # SendUsingAccount
                break
        # Also set SentOnBehalfOfName in case of delegates/shared mailbox
        mail.SentOnBehalfOfName = sender_email
    except Exception:
        # Best effort; if it fails, Outlook default account will be used
        pass

    mail.Subject = subject
    mail.HTMLBody = html_body

    if to_list:
        mail.To = "; ".join(to_list)
    if cc_list:
        mail.CC = "; ".join(cc_list)

    if attachment_path and os.path.isfile(attachment_path):
        mail.Attachments.Add(Source=attachment_path)

    mail.Send()

def main():
    # Sanity checks
    if not os.path.isfile(CSV_PATH):
        print(f"ERROR: CSV file not found at {CSV_PATH}")
        sys.exit(1)
    if not os.path.isdir(OUTPUT_DIR):
        print(f"ERROR: Output directory not found at {OUTPUT_DIR}")
        sys.exit(1)

    # 1–2) Read CSV and delete the first row
    df = read_and_prepare_dataframe(CSV_PATH)

    # 3–4) Convert to Excel Table & count totals
    total_rows = len(df)  # each row represents one approved fund
    alt_yes_count = (df["Alternative Fund?"].str.strip().str.lower() == "yes").sum()

    # 5–6) Build filename with last Friday DDMMYYYY and save
    last_fri = get_last_friday()
    mon, fri = monday_to_friday_window(last_fri)

    out_filename = f"Approved Funds_{ddmmyyyy(last_fri)}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    out_path = save_as_xlsx_table(df, out_path)

    # 7B) Prepare filtered table (Alternative Fund? == Yes) with specific columns
    alt_yes_df = df[df["Alternative Fund?"].str.strip().str.lower() == "yes"].copy()
    table_html = dataframe_to_html_table(alt_yes_df, TABLE_COLUMNS)

    # 8) Compose HTML body with signature
    signature_html = load_signature_html(SIGNATURE_NAME)
    # If signature missing, keep just 'Regards,' (already in body)
    body_html = compose_email_html(mon, fri, total_rows, alt_yes_count, table_html, signature_html)

    # Subject (formal & informative)
    subject = f"Approved Funds – {mon.strftime('%d %b')} to {fri.strftime('%d %b %Y')} – EMEA"

    # 10–11) Create, attach, send, and print success
    create_and_send_email(
        SENDER_EMAIL,
        TO_LIST,
        CC_LIST,
        subject,
        body_html,
        out_path
    )

    print("Email sent and task completed successfully.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR:", str(e))
        sys.exit(1)