# -*- coding: utf-8 -*-
"""
Recent File Updates Alerter — DOCX+PDF pairing, 10-day window, oldest→newest, hyperlink UI
Sends Outlook email with a readable, hyperlinked table (no file attachments).
Environment: Windows + Outlook desktop; Python 3.9+; pywin32
"""

import sys
import csv
import math
import socket
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Optional

# ============================= CONFIGURATION =================================
# Hard-coded shared folder (non-recursive)
SHARED_DIR = r"\\SERVER\Shared\TeamDocs"       # <-- update to your network path

# Threshold and display
DAYS_THRESHOLD = 10                            # inclusion window (≤ N days)
TIMEZONE_LABEL = "IST"                         # display label only

# Email routing and fixed subject
TO_RECIPIENTS = ["ops-team@example.com"]       # one or more
CC_RECIPIENTS = []                             # optional
SUBJECT_TITLE = "ABC DEF"                      # fixed subject header; date appended
ATTACH_CSV = True                              # attach CSV summary for audit
CSV_FILENAME = "recent_doc_updates.csv"

# UI look-and-feel
TITLE = "Recent Word/PDF Updates"
PRIMARY_COLOR = "#0f766e"                      # teal-700
MUTED_TEXT = "#6b7280"
ROW_STRIPE = "#f9fafb"
BORDER = "#e5e7eb"

# =============================== UTILITIES ===================================

def human_ts(dt: datetime) -> str:
    return dt.strftime(f"%d-%b-%Y %H:%M {TIMEZONE_LABEL}")

def human_size(num_bytes: int) -> str:
    if num_bytes is None or num_bytes == 0:
        return "0 B"
    units = ["B", "KB", "MB", "GB", "TB"]
    i = int(math.floor(math.log(num_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(num_bytes / p, 1)
    return f"{s} {units[i]}"

def age_str(dt: datetime) -> str:
    delta = datetime.now() - dt
    days = delta.days
    hours = delta.seconds // 3600
    mins = (delta.seconds % 3600) // 60
    if days > 0:
        return f"{days}d {hours}h"
    if hours > 0:
        return f"{hours}h {mins}m"
    return f"{mins}m"

def parse_readable_name_from_stem(stem: str, title_case: bool = False) -> str:
    """
    From 'abc_asset_limi_090925_1234' -> 'abc asset' (first two underscore tokens).
    """
    parts = [p for p in stem.split("_") if p]
    if len(parts) >= 2:
        name = f"{parts[0]} {parts[1]}"
    elif parts:
        name = parts[0]
    else:
        name = stem
    return name.title() if title_case else name

def path_to_uri(p: Path) -> str:
    # Convert Windows path to file:// URI that Outlook can click
    return f"file:///{str(p.resolve()).replace('\\', '/')}"

# ============================== CORE LOGIC ===================================

def scan_folder_pairs(folder: Path) -> Dict[str, Dict]:
    """
    Build a dict keyed by exact basename (no extension).
    Value holds metadata for docx/pdf siblings (same stem).
    """
    pairs: Dict[str, Dict] = {}
    for p in folder.iterdir():
        if not p.is_file():
            continue
        ext = p.suffix.lower()
        if ext not in (".docx", ".pdf"):
            continue
        stem = p.stem
        st = p.stat()
        mtime = datetime.fromtimestamp(st.st_mtime)
        entry = pairs.setdefault(stem, {"stem": stem, "docx": None, "pdf": None})
        rec = {"path": str(p.resolve()), "uri": path_to_uri(p), "mtime": mtime, "size": st.st_size}
        if ext == ".docx":
            entry["docx"] = rec
        else:
            entry["pdf"] = rec
    return pairs

def group_rows_from_pairs(pairs: Dict[str, Dict], days: int) -> List[Dict]:
    """
    Return one row per stem. Include if ANY sibling (docx/pdf) is within the window.
    Sort by OLDEST modified to highlight items 'sitting' the longest (oldest→newest).
    """
    cutoff = datetime.now() - timedelta(days=days)
    rows: List[Dict] = []
    for stem, entry in pairs.items():
        mtimes = []
        sizes = 0
        if entry["docx"]:
            mtimes.append(entry["docx"]["mtime"])
            sizes += entry["docx"]["size"]
        if entry["pdf"]:
            mtimes.append(entry["pdf"]["mtime"])
            sizes += entry["pdf"]["size"]
        if not mtimes:
            continue

        newest = max(mtimes)
        oldest = min(mtimes)

        if newest >= cutoff:  # include if any sibling is recent
            rows.append({
                "stem": stem,
                "readable": parse_readable_name_from_stem(stem, title_case=False),
                "docx": entry["docx"],
                "pdf": entry["pdf"],
                "mtime_oldest": oldest,
                "mtime_newest": newest,
                "size_total": sizes
            })

    rows.sort(key=lambda r: r["mtime_oldest"])  # oldest → newest
    return rows

def count_scanned(folder: Path) -> int:
    """KPI: count .docx/.pdf files in folder (non-recursive)."""
    return sum(1 for p in folder.iterdir() if p.is_file() and p.suffix.lower() in (".docx", ".pdf"))

# ============================== EMAIL RENDER =================================

def build_header_html(folder: Path, rows: List[Dict], scanned_count: int) -> str:
    host = socket.gethostname()
    now = datetime.now()
    count = len(rows)
    total_size = sum(r["size_total"] for r in rows) if rows else 0
    newest = rows[-1]["mtime_newest"] if count else None  # sorted by oldest
    oldest = rows[0]["mtime_oldest"] if count else None

    kpis = f"""
    <div style="display:flex;gap:10px;flex-wrap:wrap;margin:10px 0 0;">
      <span style="background:{PRIMARY_COLOR};color:white;padding:6px 10px;border-radius:999px;font-size:12px;">Scanned: {scanned_count}</span>
      <span style="background:{PRIMARY_COLOR};color:white;padding:6px 10px;border-radius:999px;font-size:12px;">Matches: {count}</span>
      <span style="background:{PRIMARY_COLOR};color:white;padding:6px 10px;border-radius:999px;font-size:12px;">Total Size: {human_size(total_size)}</span>
      <span style="background:{PRIMARY_COLOR};color:white;padding:6px 10px;border-radius:999px;font-size:12px;">Window: ≤{DAYS_THRESHOLD} days</span>
      <span style="background:{PRIMARY_COLOR};color:white;padding:6px 10px;border-radius:999px;font-size:12px;">Oldest: {human_ts(oldest) if oldest else '-'}</span>
      <span style="background:{PRIMARY_COLOR};color:white;padding:6px 10px;border-radius:999px;font-size:12px;">Newest: {human_ts(newest) if newest else '-'}</span>
    </div>
    """
    return f"""
    <div style="font-family:Segoe UI, Arial, sans-serif;">
      <h2 style="margin:0;color:{PRIMARY_COLOR};">{TITLE}</h2>
      <div style="color:{MUTED_TEXT};font-size:12px;margin-top:2px;">
        Folder: <code>{folder}</code> &nbsp;|&nbsp; Host: {host} &nbsp;|&nbsp; Scan: {human_ts(now)}
      </div>
      {kpis}
    </div>
    """

def build_table_html(rows: List[Dict]) -> str:
    if not rows:
        return f"<p style='font-family:Segoe UI, Arial, sans-serif;'>No DOCX/PDF updates in the last {DAYS_THRESHOLD} day(s).</p>"

    header = f"""
    <table border="1" cellspacing="0" cellpadding="6"
           style="border-collapse:collapse;font-family:Segoe UI, Arial, sans-serif;font-size:12px;border:1px solid {BORDER};margin-top:12px;">
      <thead style="background:{ROW_STRIPE};">
        <tr>
          <th align="left">Name</th>
          <th align="left">Files</th>
          <th align="left">Oldest Modified</th>
          <th align="left">Newest Modified</th>
          <th align="left">Age (Oldest)</th>
          <th align="right">Total Size</th>
          <th align="left">Paths</th>
        </tr>
      </thead>
      <tbody>
    """
    body = []
    for i, r in enumerate(rows):
        stripe = ROW_STRIPE if i % 2 else "#ffffff"
        # Prefer DOCX hyperlink; fallback to PDF if DOCX absent
        primary = r["docx"] or r["pdf"]
        name_link = f"<a href='{primary['uri']}'>{r['readable']}</a>"

        files_cell = []
        if r["docx"]:
            files_cell.append(f"<a href='{r['docx']['uri']}'>DOCX</a>")
        else:
            files_cell.append("<span style='color:#9ca3af;'>DOCX - n/a</span>")
        if r["pdf"]:
            files_cell.append(f"<a href='{r['pdf']['uri']}'>PDF</a>")
        else:
            files_cell.append("<span style='color:#9ca3af;'>PDF - n/a</span>")

        paths_cell = []
        if r["docx"]:
            paths_cell.append(r["docx"]["path"])
        if r["pdf"]:
            paths_cell.append(r["pdf"]["path"])

        body.append(
            f"<tr style='background:{stripe};'>"
            f"<td>{name_link}</td>"
            f"<td>{' | '.join(files_cell)}</td>"
            f"<td>{human_ts(r['mtime_oldest'])}</td>"
            f"<td>{human_ts(r['mtime_newest'])}</td>"
            f"<td>{age_str(r['mtime_oldest'])}</td>"
            f"<td align='right'>{human_size(r['size_total'])}</td>"
            f"<td style='max-width:520px;word-break:break-all;'>{'<br>'.join(paths_cell)}</td>"
            f"</tr>"
        )
    footer = "</tbody></table>"
    return header + "\n".join(body) + footer

def write_csv(rows: List[Dict], target_path: Path) -> Path:
    with target_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([
            "Readable Name", "DOCX Path", "DOCX Modified",
            "PDF Path", "PDF Modified",
            "Oldest Modified", "Newest Modified", "Age (Oldest)", "Total Size"
        ])
        if rows:
            for r in rows:
                w.writerow([
                    r["readable"],
                    r["docx"]["path"] if r["docx"] else "",
                    human_ts(r["docx"]["mtime"]) if r["docx"] else "",
                    r["pdf"]["path"] if r["pdf"] else "",
                    human_ts(r["pdf"]["mtime"]) if r["pdf"] else "",
                    human_ts(r["mtime_oldest"]),
                    human_ts(r["mtime_newest"]),
                    age_str(r["mtime_oldest"]),
                    human_size(r["size_total"])
                ])
        else:
            w.writerow([f"No DOCX/PDF updates in last {DAYS_THRESHOLD} day(s)."])
    return target_path

# ============================== EMAIL SENDER =================================

def send_outlook_email(html_body: str, csv_path: Optional[Path]) -> None:
    try:
        import win32com.client as win32
    except ImportError:
        raise RuntimeError("pywin32 is required. Install with: pip install pywin32")

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Fixed subject: "ABC DEF - {Today's Date}"
    mail.Subject = f"{SUBJECT_TITLE} - {datetime.now().strftime('%d-%b-%Y')}"

    mail.HTMLBody = (
        "<div style='font-family:Segoe UI, Arial, sans-serif;font-size:12px;'>"
        f"{html_body}"
        f"<p style='color:{MUTED_TEXT};margin-top:10px;'>"
        f"This is an automated notification. Use the hyperlinks to open files."
        f"</p></div>"
    )

    if TO_RECIPIENTS:
        mail.To = "; ".join(TO_RECIPIENTS)
    if CC_RECIPIENTS:
        mail.CC = "; ".join(CC_RECIPIENTS)

    if csv_path and csv_path.exists() and ATTACH_CSV:
        mail.Attachments.Add(str(csv_path))

    mail.Send()

# ================================== MAIN =====================================

def main():
    try:
        folder = Path(SHARED_DIR)
        if not folder.exists() or not folder.is_dir():
            raise FileNotFoundError(f"Shared folder not found or not a directory: {folder}")

        scanned_count = count_scanned(folder)
        pairs = scan_folder_pairs(folder)
        rows = group_rows_from_pairs(pairs, DAYS_THRESHOLD)

        header_html = build_header_html(folder, rows, scanned_count)
        table_html = build_table_html(rows)
        html = header_html + table_html

        csv_path = (Path.cwd() / CSV_FILENAME) if ATTACH_CSV else None
        if csv_path:
            write_csv(rows, csv_path)

        send_outlook_email(html, csv_path)

        print(f"Email sent. Scanned: {scanned_count} | Groups: {len(rows)} "
              f"| Window: ≤{DAYS_THRESHOLD}d | Ordered: oldest→newest")

    except Exception as e:
        print("ERROR:", str(e), file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()