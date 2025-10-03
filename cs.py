pip install pyautogui opencv-python pillow pywin32 pygetwindow

import os
import sys
import time
from datetime import datetime
import pyautogui as pag
import pygetwindow as gw
import pyperclip
import win32com.client as win32

# ================== USER SETTINGS ==================
ASSETS_DIR = r"C:\dhinka\cs_assets"           # <- put your 4 PNGs here
SAVE_DIR   = r"D:\Exports"                     # <- must exist
BASE_NAME  = "CreditStudio_CoperExport"        # filename prefix
CLIENT_IDS_DEFAULT = "ABC123,XYZ456"           # used if no CLI arg

# Image filenames (inside ASSETS_DIR)
IMG_TASKBAR_ICON          = "taskbar_icon.png"
IMG_CLIENT_LABEL          = "client_label.png"
IMG_SEARCH_BUTTON         = "search_button.png"
IMG_EXPORT_EXCEL_BUTTON   = "export_excel_button.png"

# Matching parameters
FIND_RETRIES = 25
FIND_INTERVAL_SEC = 0.4
CONFIDENCE = 0.88        # 0.85–0.92 is typical when screenshots are clean

# Safety/UX
pag.PAUSE = 0.15         # small delay after each action
pag.FAILSAFE = True      # move mouse to top-left to abort
# ====================================================

def die(msg: str, code: int = 1):
    print(f"\n[ERROR] {msg}")
    sys.exit(code)

def ensure_path():
    if not os.path.isdir(ASSETS_DIR):
        die(f"Assets folder not found: {ASSETS_DIR}")
    if not os.path.isdir(SAVE_DIR):
        die(f"Save folder not found: {SAVE_DIR}")

def image_path(name: str) -> str:
    return os.path.join(ASSETS_DIR, name)

def find_on_screen(img_name: str, confidence=CONFIDENCE, retries=FIND_RETRIES):
    path = image_path(img_name)
    if not os.path.isfile(path):
        die(f"Missing asset image: {path}")
    for _ in range(retries):
        box = pag.locateOnScreen(path, confidence=confidence, grayscale=True)
        if box:
            return box
        time.sleep(FIND_INTERVAL_SEC)
    return None

def click_center(img_name: str, click_times=1):
    box = find_on_screen(img_name)
    if not box:
        die(f"Could not find '{img_name}' on screen. "
            f"Check scaling=100%, app visible, and the screenshot accuracy.")
    center = pag.center(box)
    pag.moveTo(center.x, center.y, duration=0.1)
    for _ in range(click_times):
        pag.click()
        time.sleep(0.15)

def activate_credit_studio_by_taskbar():
    """Click the taskbar icon (preferred, robust for OpenFin)."""
    print("[i] Activating Credit Studio via taskbar icon…")
    click_center(IMG_TASKBAR_ICON)
    time.sleep(0.6)

    # Try maximize the active window just in case
    try:
        win = gw.getActiveWindow()
        if win and not win.isMaximized:
            win.maximize()
    except Exception:
        pass

def paste_text(text: str):
    pyperclip.copy(text)
    pag.hotkey('ctrl', 'v')

def wait_for_excel_and_save(save_dir: str, base_name: str, timeout_sec: int = 60) -> str:
    """Attach to Excel and save the active workbook as xlsx."""
    print("[i] Waiting for Excel workbook to open…")
    t0 = time.time()
    excel = None
    while time.time() - t0 < timeout_sec:
        try:
            excel = win32.GetObject(Class="Excel.Application")
            if excel and excel.Workbooks.Count >= 1:
                break
        except Exception:
            pass
        time.sleep(1.0)

    if not excel:
        die("Excel did not open the exported workbook within the timeout.")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    full_path = os.path.join(save_dir, f"{base_name}_{ts}.xlsx")

    wb = excel.ActiveWorkbook
    if wb is None:
        die("Excel ActiveWorkbook is None. Is the export actually opening a workbook?")
    print(f"[i] Saving workbook to: {full_path}")
    # 51 = xlOpenXMLWorkbook (.xlsx)
    wb.SaveAs(full_path, FileFormat=51)
    print("[✓] Excel save complete.")
    return full_path

def main():
    ensure_path()

    # Take COPER IDs from CLI if provided
    client_ids = CLIENT_IDS_DEFAULT
    if len(sys.argv) >= 2 and sys.argv[1].strip():
        client_ids = sys.argv[1].strip()
    print(f"[i] Using Client IDs: {client_ids}")

    # Bring Credit Studio to front
    activate_credit_studio_by_taskbar()

    # 1) Focus the Client Name/ID input
    # Strategy: click near the label, then TAB or click into the nearest input.
    print("[i] Locating 'Client Name/ID' label…")
    label_box = find_on_screen(IMG_CLIENT_LABEL)
    if not label_box:
        die("Could not find the Client Name/ID label on screen.")
    # Heuristic: textbox is usually to the right of label; click at label.right + offset
    x = label_box.left + label_box.width + 120
    y = label_box.top + label_box.height // 2
    pag.moveTo(x, y, duration=0.1)
    pag.click()
    time.sleep(0.15)

    # 2) Paste the IDs
    paste_text(client_ids)
    time.sleep(0.2)

    # 3) Click Search
    print("[i] Clicking Search…")
    click_center(IMG_SEARCH_BUTTON)
    time.sleep(1.2)  # wait for results grid to fill

    # 4) Click Export to Excel
    print("[i] Clicking Export to Excel…")
    click_center(IMG_EXPORT_EXCEL_BUTTON)
    time.sleep(2.0)  # give Excel time to spin up

    # 5) Save via Excel COM
    saved = wait_for_excel_and_save(SAVE_DIR, BASE_NAME)
    print(f"[✓] DONE. File saved at: {saved}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[!] Aborted by user (Ctrl+C).")
    except Exception as e:
        print("\n[ERROR]")
        print(str(e))
        raise
