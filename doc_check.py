# -*- coding: utf-8 -*-
"""
Recent DOCX/PDF Alert — step-by-step tracer with robust logging
- Pairs DOCX+PDF by exact basename; unmatched files become independent rows
- 10-day window; include if any sibling is within window
- Oldest → Newest ordering to surface oldest items first
- Email: hyperlinked readable Name (DOCX else PDF); optional CSV; fixed subject "ABC DEF - {Today’s Date}"
Env: Windows, Python 3.9+, Outlook desktop, pywin32
"""

import os
import sys
import csv
import math
import socket
import traceback
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Optional

# ============================= CONFIGURATION =================================
SHARED_DIR = r"\\SERVER\Shared\TeamDocs"     # <<< non-recursive folder to scan
DAYS_THRESHOLD = 10                          # include files modified in ≤ N days
TIMEZONE_LABEL = "IST"                       # display label only

TO_RECIPIENTS = ["ops-team@example.com"]     # recipients (semicolon joined by Outlook)
CC_RECIPIENTS = []                           # optional
SUBJECT_TITLE = "ABC DEF"                    # email subject header (date appended)

ATTACH_CSV = True
CSV_FILENAME = "recent_doc_updates.csv"

# Diagnostics / Ops
LOG_FILE = "recent_doc_updates.log"
LOG_LEVEL = logging.INFO                     # set to logging.DEBUG for deeper telemetry
LOG_MAX_BYTES = 1_000_000                    # ~1 MB per file
LOG_BACKUP_COUNT = 3
SMOKE_TEST = False                           # True => send a tiny test email regardless of matches
SAVE_MSG_COPY = True                         # save a .msg draft pre-send for evidence

# UI look-and-feel
TITLE = "Recent Word/PDF Updates"
PRIMARY_COLOR = "#0f766e"                    # teal-700
MUTED_TEXT = "#6b7280"
ROW_STRIPE = "#f9fafb"
BORDER = "#e5e7eb"

# ============================= LOGGING & TRACER ==============================
def setup_logger() -> logging.Logger:
    logger = logging.getLogger("recent_updates")
    logger.setLevel(LOG_LEVEL)
    if logger.handlers:
        return logger
    log_path = Path.cwd() / LOG_FILE
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    fh = RotatingFileHandler(log_path, maxBytes=LOG_MAX_BYTES, backupCount=LOG_BACKUP_COUNT, encoding="utf-8")
    fh.setFormatter(fmt); fh.setLevel(LOG_LEVEL); logger.addHandler(fh)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt); ch.setLevel(LOG_LEVEL); logger.addHandler(ch)
    logger.info("Logger initialized. CWD=%s | LOG=%s | LEVEL=%s", Path.cwd(), log_path, logging.getLevelName(LOG_LEVEL))
    return logger

log = setup_logger()

class StepRunner:
    """Simple step sequencer to print & log [n/N] Step ... OK/FAIL with timing."""
    def __init__(self, total: int):
        self.total = total
        self.idx = 0

    def run(self, title: str, fn, *args, **kwargs):
        self.idx += 1
        tag = f"[{self.idx}/{self.total}] {title}"
        start = datetime.now()
        print(f"{tag} ..."); log.info("%s ...", title)
        try:
            result = fn(*args, **kwargs)
            sec = (datetime.now() - start).total_seconds()
            print(f"{tag} — OK ({sec:.2f}s)"); log.info("%s — OK (%.2fs)", title, sec)
            return result
        except Exception as e:
            sec = (datetime.now() - start).total_seconds()
            print(f"{tag} — FAIL ({sec:.2f}s): {e}"); log.exception("%s — FAIL: %r", title, e)
            raise

# =============================== UTILITIES ===================================
def human_ts(dt: datetime) -> str:
    return dt.strftime(f"%d-%b-%Y %H:%M {TIMEZONE_LABEL}")

def human_size(num_bytes: int) -> str:
    if not num_bytes:
        return "0 B"
    units = ["B", "KB", "MB", "GB", "TB"]
    i = int(math.floor(math.log(num_bytes, 1024)))
    p = math.pow(1024, i)
    return f"{round(num_bytes / p, 1)} {units[i]}"

def age_str(dt: datetime) -> str:
    d = datetime.now() - dt
    days = d.days; hours = d.seconds // 3600; mins = (d.seconds % 3600) // 60
    if days > 0: return f"{days}d {hours}h"
    if hours > 0: return f"{hours}h {mins}m"
    return f"{mins}m"

def parse_readable_name_from_stem(stem: str, title_case: bool=False) -> str:
    """From 'abc_asset_limi_090925_1234' -> 'abc asset' (first two underscore tokens)."""
    parts = [p for p in stem.split("_") if p]
    name = f"{parts[0]} {parts[1]}" if len(parts) >= 2 else (parts[0] if parts else stem)
    return name.title() if title_case else name

def path_to_uri(p: Path) -> str:
    return f"file:///{str(p.resolve()).replace('\\', '/')}"

# ============================== CORE LOGIC ===================================
def env_snapshot():
    """Collect lightweight runtime metadata for diagnostics."""
    snap = {
        "python": sys.version.split()[0],
        "executable": sys.executable,
        "cwd": str(Path.cwd()),
        "user": os.environ.get("USERNAME") or os.environ.get("USER") or "n/a",
        "host": socket.gethostname(),
    }
    log.info("ENV: %s", snap); return snap

def outlook_healthcheck():
    """Verify pywin32 & Outlook profile; enumerate accounts for visibility."""
    try:
        import win32com.client as win32
    except Exception as e:
        log.exception("pywin32 import failed: %r", e)
        raise

    try:
        app = win32.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
        accounts = []
        try:
            for i in range(ns.Accounts.Count):
                accounts.append(str(ns.Accounts.Item(i+1).SmtpAddress))
        except Exception:
            # Some builds restrict SmtpAddress; fallback to Name
            try:
                accounts = [str(ns.Accounts.Item(i+1).DisplayName) for i in range(ns.Accounts.Count)]
            except Exception:
                accounts = ["<unable to enumerate accounts>"]
        log.info("Outlook OK. Accounts: %s", accounts)
        return {"accounts": accounts}
    except Exception as e:
        log.exception("Outlook COM healthcheck failed: %r", e)
        raise

def ensure_folder(folder: Path) -> Path:
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"Shared folder not found or not a directory: {folder}")
    # Simple read probe
    _ = list(folder.iterdir())
    log.info("Folder accessible: %s", folder)
    return folder

def scan_folder_pairs(folder: Path) -> Dict[str, Dict]:
    """Build dict keyed by stem, each holding docx/pdf siblings."""
    pairs: Dict[str, Dict] = {}
    seen = 0
    for p in folder.iterdir():
        try:
            if not p.is_file(): continue
            ext = p.suffix.lower()
            if ext not in (".docx", ".pdf"): continue
            seen += 1
            stem = p.stem
            st = p.stat(); mtime = datetime.fromtimestamp(st.st_mtime)
            entry = pairs.setdefault(stem, {"stem": stem, "docx": None, "pdf": None})
            rec = {"path": str(p.resolve()), "uri": path_to_uri(p), "mtime": mtime, "size": st.st_size}
            if ext == ".docx": entry["docx"] = rec
            else:              entry["pdf"]  = rec
        except Exception as e:
            log.exception("Error scanning: %s | %r", p, e)
    log.info("Scan complete. Eligible files seen: %s | Pair groups: %s", seen, len(pairs))
    return pairs

def group_rows_from_pairs(pairs: Dict[str, Dict], days: int) -> List[Dict]:
    """One row per stem; include if any sibling is within window. Sort by oldest mtime."""
    cutoff = datetime.now() - timedelta(days=days)
    rows: List[Dict] = []
    for stem, entry in pairs.items():
        mtimes, sizes = [], 0
        if entry["docx"]: mtimes.append(entry["docx"]["mtime"]); sizes += entry["docx"]["size"]
        if entry["pdf"]:  mtimes.append(entry["pdf"]["mtime"]);  sizes += entry["pdf"]["size"]
        if not mtimes: continue
        newest, oldest = max(mtimes), min(mtimes)
        if newest >= cutoff:
            rows.append({
                "stem": stem,
                "readable": parse_readable_name_from_stem(stem, False),
                "docx": entry["docx"], "pdf": entry["pdf"],
                "mtime_oldest": oldest, "mtime_newest": newest,
                "size_total": sizes
            })
    rows.sort(key=lambda r: r["mtime_oldest"])
    log.info("Grouping complete. Rows within window (≤%sd): %s", days, len(rows))
    return rows

def count_scanned(folder: Path) -> int:
    try:
        return sum(1 for p in folder.iterdir() if p.is_file() and p.suffix.lower() in (".docx", ".pdf"))
    except Exception as e:
        log.exception("Count scan failed: %r", e); return 0

# ============================== EMAIL RENDER =================================
def build_header_html(folder: Path, rows: List[Dict], scanned_count: int) -> str:
    host = socket.gethostname(); now = datetime.now(); count = len(rows)
    total_size = sum(r["size_total"] for r in rows) if rows else 0
    newest = rows[-1]["mtime_newest"] if count else None
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
        primary = r["docx"] or r["pdf"]
        name_link = f"<a href='{primary['uri']}'>{r['readable']}</a>"
        files_cell = []
        files_cell.append(f"<a href='{r['docx']['uri']}'>DOCX</a>" if r["docx"] else "<span style='color:#9ca3af;'>DOCX - n/a</span>")
        files_cell.append(f"<a href='{r['pdf']['uri']}'>PDF</a>" if r["pdf"] else "<span style='color:#9ca3af;'>PDF - n/a</span>")
        paths_cell = []
        if r["docx"]: paths_cell.append(r["docx"]["path"])
        if r["pdf"]:  paths_cell.append(r["pdf"]["path"])
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
    return header + "\n".join(body) + "</tbody></table>"

def write_csv(rows: List[Dict], target_path: Path) -> Path:
    try:
        with target_path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Readable Name","DOCX Path","DOCX Modified","PDF Path","PDF Modified","Oldest Modified","Newest Modified","Age (Oldest)","Total Size"])
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
        log.info("CSV written: %s", target_path)
    except Exception as e:
        log.exception("CSV write failed: %r", e)
    return target_path

# ============================== EMAIL (COM) ==================================
def compose_email(html_body: str, csv_path: Optional[Path]):
    try:
        import win32com.client as win32
    except Exception as e:
        log.exception("pywin32 import failed: %r", e); raise

    try:
        outlook = win32.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")  # touch MAPI to assert profile
        mail = outlook.CreateItem(0)
    except Exception as e:
        log.exception("Outlook COM init failed: %r", e); raise

    mail.Subject = f"{SUBJECT_TITLE} - {datetime.now():%d-%b-%Y}"
    mail.HTMLBody = (
        "<div style='font-family:Segoe UI, Arial, sans-serif;font-size:12px;'>"
        f"{html_body}"
        f"<p style='color:{MUTED_TEXT};margin-top:10px;'>Automated notification. Hyperlinks require network access.</p>"
        "</div>"
    )
    if TO_RECIPIENTS: mail.To = "; ".join(TO_RECIPIENTS)
    if CC_RECIPIENTS: mail.CC = "; ".join(CC_RECIPIENTS)

    if csv_path and csv_path.exists() and ATTACH_CSV:
        try:
            mail.Attachments.Add(str(csv_path))
            log.info("CSV attached to email.")
        except Exception as e:
            log.warning("Attachment add failed (non-fatal): %r", e)

    if SAVE_MSG_COPY:
        try:
            msg_path = (Path.cwd() / f"draft_{datetime.now():%Y%m%d_%H%M%S}.msg").resolve()
            mail.SaveAs(str(msg_path))
            log.info("Saved draft MSG: %s", msg_path)
        except Exception as e:
            log.warning("SaveAs(.msg) failed (non-fatal): %r", e)

    return mail

def send_mail(mail) -> None:
    try:
        mail.Send()
        log.info("Outlook .Send() invoked successfully.")
    except Exception as e:
        log.exception("Outlook send failed: %r", e)
        raise

# ================================== MAIN =====================================
def main() -> int:
    # Define total steps (adjust if you toggle SMOKE_TEST)
    total_steps = 10 if not SMOKE_TEST else 6
    s = StepRunner(total_steps)

    try:
        s.run("Environment snapshot", env_snapshot)
        s.run("Outlook COM healthcheck", outlook_healthcheck)
        folder = s.run("Validate folder access", ensure_folder, Path(SHARED_DIR))

        if SMOKE_TEST:
            # Minimal payload to validate COM send end-to-end
            html = "<p>Smoke test payload.</p>"
            mail = s.run("Compose smoke-test email", compose_email, html, None)
            s.run("Send smoke-test email", send_mail, mail)
            s.run("End (smoke)", lambda: None)
            return 0

        scanned_count = s.run("Count eligible files (.docx/.pdf)", count_scanned, folder)
        pairs = s.run("Scan folder and pair DOCX/PDF", scan_folder_pairs, folder)
        rows = s.run(f"Group rows within ≤{DAYS_THRESHOLD} days (oldest→newest)", group_rows_from_pairs, pairs, DAYS_THRESHOLD)

        header_html = s.run("Build email header (KPIs)", build_header_html, folder, rows, scanned_count)
        table_html  = s.run("Build email table (hyperlinks)", build_table_html, rows)
        html = header_html + table_html

        csv_path = (Path.cwd() / CSV_FILENAME) if ATTACH_CSV else None
        if csv_path:
            s.run("Write CSV audit file", write_csv, rows, csv_path)

        mail = s.run("Compose email (subject, body, recipients)", compose_email, html, csv_path)
        s.run("Send email via Outlook", send_mail, mail)
        s.run("End", lambda: None)
        return 0

    except Exception:
        # The failing step, stack trace, and message are already logged.
        return 1

if __name__ == "__main__":
    sys.exit(main())
