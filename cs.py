import time
import sys
import os
from datetime import datetime
import traceback

from pywinauto.application import Application
from pywinauto import timings
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.controls.uia_controls import EditWrapper, ButtonWrapper, MenuWrapper
from pywinauto.keyboard import send_keys

import win32com.client as win32
import pyperclip


# ==== USER SETTINGS (EDIT THESE) ====
# Comma-delimited list; you can also pass via CLI: python credit_studio_export.py "ID1,ID2,ID3"
CLIENT_IDS = "ABC123,XYZ456"
SAVE_DIR   = r"D:\Exports"            # Ensure this directory exists
BASE_NAME  = "CreditStudio_CoperExport"  # Prefix for the saved file
WINDOW_TITLE_RE = ".*Credit Studio.*" # Regex to match the app window title
MENU_STUDIO_TOOLS = "Studio Tools"
MENU_COUNTERPARTY_QUICK_UPDATE = "Counterparty Quick Update"

# Control captions (may vary slightly in your build; adjust if needed)
CLIENT_FIELD_LABEL_SUBSTR = "Client Name/ID"
SEARCH_BUTTON_TEXT_SUBSTR = "Search"
EXPORT_BUTTON_TEXT_SUBSTR = "Export to Excel"

# Timeouts
ATTACH_TIMEOUT_SEC = 20
DIALOG_TIMEOUT_SEC = 20
EXCEL_ATTACH_TIMEOUT_SEC = 40


def ensure_dir(path: str):
    if not os.path.isdir(path):
        raise FileNotFoundError(f"Save directory does not exist: {path}")


def attach_credit_studio():
    print("[i] Attaching to Credit Studio window...")
    try:
        app = Application(backend="uia").connect(title_re=WINDOW_TITLE_RE, timeout=ATTACH_TIMEOUT_SEC)
        # If multiple windows match, pick the top-level active one
        win = app.top_window()
        win.set_focus()
        print("[✓] Attached.")
        return app, win
    except Exception as e:
        raise RuntimeError(f"Could not attach to Credit Studio window matching /{WINDOW_TITLE_RE}/. "
                           f"Make sure it is open and visible. Error: {e}")


def open_counterparty_quick_update(win):
    """
    Tries menu route first; if not a classic menu, uses Alt key navigation fallback.
    """
    print("[i] Opening 'Counterparty Quick Update' from 'Studio Tools'...")
    try:
        # Try to find a classic menu bar
        menubars = win.descendants(control_type="MenuBar")
        if menubars:
            mb = MenuWrapper(menubars[0])
            mb.select(f"{MENU_STUDIO_TOOLS}->{MENU_COUNTERPARTY_QUICK_UPDATE}")
            print("[✓] Opened via menu bar.")
            return
    except Exception:
        pass

    # Try clickable menu items directly
    try:
        studio_tools = win.child_window(title=MENU_STUDIO_TOOLS, control_type="MenuItem")
        studio_tools.click_input()
        time.sleep(0.5)
        cqu = win.child_window(title=MENU_COUNTERPARTY_QUICK_UPDATE, control_type="MenuItem")
        cqu.click_input()
        print("[✓] Opened via direct menu items.")
        return
    except Exception:
        pass

    # Fallback: keyboard navigation (adjust if your app uses different accelerators)
    print("[!] Falling back to keyboard navigation. You may see the UI flash.")
    win.set_focus()
    send_keys("%")  # Activate menu
    time.sleep(0.3)
    # Try to type the menu text letters; adapt if your app uses accelerators
    send_keys(MENU_STUDIO_TOOLS.replace(" ", "")[:1])  # First letter heuristic
    time.sleep(0.3)
    send_keys(MENU_COUNTERPARTY_QUICK_UPDATE.replace(" ", "")[:1])
    time.sleep(1.0)
    print("[~] If the dialog did not open, you may need to adjust menu navigation.")


def get_child_by_partial_text(container, substr, control_type=None):
    """
    Find a child control whose name/title contains `substr` (case-insensitive).
    Optionally filter by control_type.
    """
    substr_low = substr.lower()
    for ctrl in container.descendants():
        try:
            name = (ctrl.window_text() or "").lower()
            ctype = getattr(ctrl, "friendly_class_name", lambda: "")().lower()
            if substr_low in name and (control_type is None or ctrl.element_info.control_type == control_type):
                return ctrl
        except Exception:
            continue
    return None


def find_counterparty_window(app):
    """
    Heuristic: after opening the module, a new dialog/window appears.
    We look for a new top-level window containing the expected field/button labels.
    """
    print("[i] Locating the 'Counterparty Quick Update' window...")
    deadline = time.time() + DIALOG_TIMEOUT_SEC
    while time.time() < deadline:
        for w in app.windows():
            title = w.window_text() or ""
            if not title.strip():
                continue
            try:
                # Check for presence of the expected field or buttons
                if get_child_by_partial_text(w, CLIENT_FIELD_LABEL_SUBSTR) or \
                   get_child_by_partial_text(w, SEARCH_BUTTON_TEXT_SUBSTR) or \
                   get_child_by_partial_text(w, EXPORT_BUTTON_TEXT_SUBSTR):
                    print(f"[✓] Found: '{title}'")
                    return w
            except Exception:
                continue
        time.sleep(0.5)
    raise ElementNotFoundError("Could not find the Counterparty Quick Update dialog window.")


def fill_client_ids_and_search(cqu_win, client_ids_text):
    print("[i] Filling Client Name/ID and clicking Search...")
    # Find the edit box near the label or by name
    edit_ctrl = None

    # Try to get the Edit directly by label
    labeled = get_child_by_partial_text(cqu_win, CLIENT_FIELD_LABEL_SUBSTR, control_type="Text")
    if labeled:
        # Often the Edit is a sibling or a next descendant
        # Try immediate edit descendants first
        for desc in cqu_win.descendants(control_type="Edit"):
            # Heuristic: pick the first visible, enabled edit
            try:
                ew = EditWrapper(desc)
                if ew.is_enabled() and ew.is_visible():
                    edit_ctrl = ew
                    break
            except Exception:
                continue

    if not edit_ctrl:
        # Fall back: first visible edit on the window
        for desc in cqu_win.descendants(control_type="Edit"):
            try:
                ew = EditWrapper(desc)
                if ew.is_enabled() and ew.is_visible():
                    edit_ctrl = ew
                    break
            except Exception:
                continue

    if not edit_ctrl:
        raise ElementNotFoundError("Client Name/ID edit box not found. Inspect control names and adjust selectors.")

    # Set text via clipboard (more reliable on some custom controls)
    pyperclip.copy(client_ids_text)
    try:
        edit_ctrl.set_focus()
        edit_ctrl.select()  # highlight existing contents
    except Exception:
        pass
    send_keys("^v")  # paste
    time.sleep(0.3)

    # Click Search
    search_btn_el = get_child_by_partial_text(cqu_win, SEARCH_BUTTON_TEXT_SUBSTR, control_type="Button")
    if not search_btn_el:
        # try any button that has 'Search' in automation id/name
        for desc in cqu_win.descendants(control_type="Button"):
            try:
                bw = ButtonWrapper(desc)
                if "search" in (bw.window_text() or "").lower():
                    search_btn_el = desc
                    break
            except Exception:
                continue

    if not search_btn_el:
        raise ElementNotFoundError("Search button not found. Adjust SEARCH_BUTTON_TEXT_SUBSTR or use Inspect.exe.")
    ButtonWrapper(search_btn_el).click_input()
    print("[✓] Search clicked.")
    time.sleep(1.0)  # give results time to populate


def export_to_excel(cqu_win):
    print("[i] Clicking 'Export to Excel'...")
    export_el = get_child_by_partial_text(cqu_win, EXPORT_BUTTON_TEXT_SUBSTR, control_type="Button")
    if not export_el:
        # Try any button containing 'Export'
        for desc in cqu_win.descendants(control_type="Button"):
            try:
                bw = ButtonWrapper(desc)
                label = (bw.window_text() or "").lower()
                if "export" in label and "excel" in label:
                    export_el = desc
                    break
            except Exception:
                continue

    if not export_el:
        raise ElementNotFoundError("Export to Excel button not found. Adjust EXPORT_BUTTON_TEXT_SUBSTR or selectors.")
    ButtonWrapper(export_el).click_input()
    print("[✓] Export clicked. Waiting for Excel to open...")
    time.sleep(2.0)


def find_excel_instance(timeout_sec=EXCEL_ATTACH_TIMEOUT_SEC):
    print("[i] Waiting for Excel instance/workbook...")
    deadline = time.time() + timeout_sec
    excel = None
    while time.time() < deadline:
        try:
            excel = win32.GetObject(Class="Excel.Application")
            if excel is not None:
                # Ensure at least one workbook opened
                if excel.Workbooks.Count >= 1:
                    return excel
        except Exception:
            pass
        time.sleep(1.0)
    raise TimeoutError("Excel did not open the export within expected time.")


def save_active_workbook(excel, save_dir, base_name):
    ensure_dir(save_dir)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    full_path = os.path.join(save_dir, f"{base_name}_{ts}.xlsx")

    wb = excel.ActiveWorkbook
    if wb is None:
        raise RuntimeError("No ActiveWorkbook in Excel. Check if export actually opened a workbook.")

    print(f"[i] Saving workbook to: {full_path}")
    # 51 = xlOpenXMLWorkbook (.xlsx), 56 would be .xls
    wb.SaveAs(full_path, FileFormat=51)
    print("[✓] Saved.")
    return full_path


def main():
    # Allow client IDs via CLI arg override
    client_ids_text = CLIENT_IDS
    if len(sys.argv) >= 2 and sys.argv[1].strip():
        client_ids_text = sys.argv[1].strip()

    print(f"[i] Using Client IDs: {client_ids_text}")

    app, main_win = attach_credit_studio()
    open_counterparty_quick_update(main_win)

    # Find the dialog that contains our fields
    cqu_win = find_counterparty_window(app)
    cqu_win.set_focus()

    fill_client_ids_and_search(cqu_win, client_ids_text)
    export_to_excel(cqu_win)

    excel = find_excel_instance()
    saved_path = save_active_workbook(excel, SAVE_DIR, BASE_NAME)

    print(f"[✓] DONE. File saved at: {saved_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n[ERROR]")
        print(str(e))
        print("\n[TRACEBACK]")
        traceback.print_exc()
        sys.exit(1)
