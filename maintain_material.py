"""
maintain_material.py
Maintain Material Module — RPA PT Mayora Indah Tbk

Launched from launcher.py — receives back_callback to return to menu.

Flow:
  1. User browses material codes Excel file (Kode Barang column)
  2. User browses List_Plant.xlsx (Plant + Prch.Grp)
  3. User picks plants from checklist
  4. Click Run → SAP MM01 automation starts
     Outer loop : each selected plant
     Inner loop : each material code
       Step 1 — MM01 Initial Screen (COM wait)
       Step 2 — Select Views popup  (COM wait)
       Step 3 — Org Levels popup    (COM wait)
       Step 4 — Purchasing tab      (COM wait)
       Step 5 — Accounting 1 tab    (COM wait)
"""

import os
import time
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import queue
import pyautogui
import win32gui
import win32con
import openpyxl
from datetime import datetime

# ─────────────────────────────────────────────
# ENABLE DPI AWARENESS - MUST BE FIRST
# ─────────────────────────────────────────────

try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

try:
    from theme_manager import get_theme_manager
    theme = get_theme_manager()
except ImportError:
    theme = None

# ─────────────────────────────────────────────
# CONSTANTS - CRISP FONTS
# ─────────────────────────────────────────────

# Use Segoe UI for crisp, clear rendering on Windows
FONT       = "Segoe UI"
FONT_DISP  = "Segoe UI"
FONT_TITLE = ("Segoe UI", 12, "bold")
FONT_SUB   = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)


# ─────────────────────────────────────────────
# SAP CONSTANTS
# ─────────────────────────────────────────────

INDUSTRY      = "F"
MATERIAL_TYPE = "FERT"
SLOC          = "WH01"
SALES_ORG     = "CS00"
MOVING_PRICE  = "1"
DC_GT         = "41"
DC_MT         = "21"
COPY_PLANT_GT = "B100"
COPY_PLANT_MT = "B242"

SAP_CLASSES = ["SAP_FRONTEND_SESSION", "SAPFrontend", "SAPGUI"]
SAP_KEYWORDS = [
    "SAP Easy Access", "SAP R/3", "SAP NetWeaver",
    "Create Material", "MM01",
]
SKIP_TITLES = [
    "SAP Logon", "Firefox", "Chrome", "Edge",
    "Visual Studio", "Code", "Notepad", "Claude",
]

pyautogui.FAILSAFE = True
pyautogui.PAUSE    = 0.25


# ─────────────────────────────────────────────
# DATA LOADERS
# ─────────────────────────────────────────────

def load_plant_data(filepath: str) -> dict:
    """
    Read List_Plant.xlsx.
    Return: { plant_code: { name, prch_grp, type, dc, copy_plant, copy_dc } }
    """
    wb     = openpyxl.load_workbook(filepath, data_only=True)
    result = {}

    for sheet_name in wb.sheetnames:
        ws           = wb[sheet_name]
        header_found = False

        for row in ws.iter_rows(values_only=True):
            if not row or not row[0]:
                continue
            cell0 = str(row[0]).strip()

            if cell0 == "Plant":
                header_found = True
                continue
            if not header_found:
                continue

            plant_code = cell0
            plant_name = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            prch_grp   = ""
            if len(row) > 2 and row[2] is not None:
                try:
                    prch_grp = str(int(float(str(row[2])))).strip()
                except Exception:
                    prch_grp = str(row[2]).strip()

            is_mt = plant_name.upper().strip().endswith(" MT")
            result[plant_code] = {
                "name":       plant_name,
                "prch_grp":   prch_grp,
                "type":       "MT" if is_mt else "GT",
                "dc":         DC_MT if is_mt else DC_GT,
                "copy_plant": COPY_PLANT_MT if is_mt else COPY_PLANT_GT,
                "copy_dc":    DC_MT if is_mt else DC_GT,
            }

        if result:
            break

    return result


def load_material_codes(filepath: str) -> list:
    """
    Find 'Kode Barang' column in Excel and return all values.
    Return: ['410944', '410974', ...]
    """
    wb    = openpyxl.load_workbook(filepath, data_only=True)
    codes = []

    for sheet_name in wb.sheetnames:
        ws         = wb[sheet_name]
        kode_col   = None
        header_row = None

        for ri, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if not row:
                continue
            for ci, cell in enumerate(row):
                if cell and str(cell).strip() == "Kode Barang":
                    kode_col, header_row = ci, ri
                    break
            if kode_col is not None:
                break

        if kode_col is None:
            continue

        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if not row or kode_col >= len(row):
                continue
            val = row[kode_col]
            if val is None:
                continue
            try:
                code = str(int(float(str(val)))).strip()
            except Exception:
                code = str(val).strip()
            if code and code not in codes:
                codes.append(code)

        if codes:
            break

    return codes


# ─────────────────────────────────────────────
# SAP HELPERS
# ─────────────────────────────────────────────

def _get_session():
    from sap_com_helper import get_sap_session
    return get_sap_session()


def _focus_sap():
    priority, fallback = [], []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        cls   = win32gui.GetClassName(hwnd)
        if not title or any(s in title for s in SKIP_TITLES):
            return
        if any(c in cls for c in SAP_CLASSES):
            priority.append(hwnd)
        elif any(kw in title for kw in SAP_KEYWORDS):
            fallback.append(hwnd)

    win32gui.EnumWindows(cb, None)
    hwnd = priority[0] if priority else (fallback[0] if fallback else None)
    if not hwnd:
        raise Exception("SAP window not found. Make sure SAP is open.")
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.5)


def _tab(n: int = 1, delay: float = 0.18):
    for _ in range(n):
        pyautogui.press("tab")
        time.sleep(delay)


def _type(value: str, interval: float = 0.07):
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.08)
    pyautogui.typewrite(str(value), interval=interval)
    time.sleep(0.08)


def _com_enter(session, window: str = "wnd[0]", wait: float = 0.3):
    """Send Enter via SAP COM — more reliable than pyautogui.press('enter')."""
    session.findById(window).sendVKey(0)
    time.sleep(wait)


# ─────────────────────────────────────────────
# COM SCREEN DETECTION
# ─────────────────────────────────────────────

def _wait_com(session, condition_fn, timeout: float = 60.0,
              poll: float = 0.3, label: str = "") -> bool:
    """
    Poll until condition_fn(session) is True or timeout.
    Robot waits here until SAP actually shows the expected screen.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if condition_fn(session):
                return True
        except Exception:
            pass
        time.sleep(poll)
    return False


def _has_popup(session) -> bool:
    try:
        session.findById("wnd[1]")
        return True
    except Exception:
        return False


def _popup_title(session) -> str:
    try:
        return session.findById("wnd[1]").Text or ""
    except Exception:
        return ""


def _no_popup(session) -> bool:
    try:
        session.findById("wnd[1]")
        return False
    except Exception:
        return True


def _on_mm01(session) -> bool:
    try:
        return (session.Info.Transaction.upper() == "MM01"
                and _no_popup(session))
    except Exception:
        return False


# ─────────────────────────────────────────────
# MM01 — ONE MATERIAL × ONE PLANT
# ─────────────────────────────────────────────

def _mm01_one(session, material: str, plant_info: dict,
              plant_code: str, log_fn=None) -> bool:
    """
    Run MM01 for one material × one plant.
    COM handles screen detection. pyautogui handles field filling.
    Returns True on success, False on error.
    """
    def _l(msg, level="INFO"):
        if log_fn:
            log_fn(msg, level)

    prch_grp   = plant_info["prch_grp"]
    dc         = plant_info["dc"]
    copy_plant = plant_info["copy_plant"]
    copy_dc    = plant_info["copy_dc"]
    ptype      = plant_info["type"]

    _l(f"  [{ptype}] {plant_code} | {material} | "
       f"PrchGrp={prch_grp} | DC={dc} | CopyFrom={copy_plant}")

    try:
        # Navigate to MM01 via COM
        from sap_com_helper import sap_tcode
        sap_tcode("MM01", session=session, wait=0.5)

        # COM: wait for MM01 initial screen
        if not _wait_com(session, _on_mm01, timeout=60,
                         label="MM01 Initial Screen"):
            raise Exception("MM01 initial screen did not appear")

        _focus_sap()
        pyautogui.hotkey("ctrl", "Home")
        time.sleep(0.3)

        # ── STEP 1: Initial Screen ─────────────────────────
        # Field 1: Material (Kode Barang)
        _type(material)
        time.sleep(0.15)

        # Tab → Industry Sector = F (Food)
        _tab(1)
        _type(INDUSTRY)
        time.sleep(0.15)

        # Tab → Material Type = FERT
        _tab(1)
        _type(MATERIAL_TYPE)
        time.sleep(0.15)

        # Tab Tab → Copy from Material (same Kode Barang)
        _tab(2)
        _type(material)
        time.sleep(0.15)

        # Enter → SAP shows Select Views popup
        _com_enter(session, wait=0.5)
        if not _wait_com(session, _has_popup, timeout=60,
                         label="Select Views popup"):
            raise Exception("Select Views popup did not appear")

        # COM: confirm it's the Select Views popup
        _wait_com(session,
                  lambda s: "view" in _popup_title(s).lower(),
                  timeout=10, label="Confirm Select Views")

        # ── STEP 2: Select Views popup ─────────────────────
        # Views defaulted (Purchasing + Accounting 1) — just Enter
        _com_enter(session, window="wnd[1]", wait=0.3)

        # COM: wait for Org Levels popup
        if not _wait_com(
            session,
            lambda s: _has_popup(s) and (
                "organ" in _popup_title(s).lower() or
                "org"   in _popup_title(s).lower()
            ),
            timeout=60, label="Org Levels popup"
        ):
            raise Exception("Organizational Levels popup did not appear")

        # ── STEP 3: Organizational Levels popup ────────────
        _focus_sap()
        pyautogui.hotkey("ctrl", "Home")
        time.sleep(0.25)

        # LEFT side — Destination (cursor starts on Plant)
        _type(plant_code);  time.sleep(0.15)   # Plant = current plant
        _tab(1);            _type(SLOC);        time.sleep(0.15)   # SLoc = WH01
        _tab(1);            time.sleep(0.12)    # Valuation area (skip — auto)
        _tab(1);            _type(SALES_ORG);   time.sleep(0.15)   # Sales Org = CS00
        _tab(1);            _type(dc);          time.sleep(0.15)   # DC = 41 (GT) / 21 (MT)

        # RIGHT side — Copy From (only Plant and DC; SLoc/Sales Org use defaults)
        _tab(1);            _type(copy_plant);  time.sleep(0.15)   # Copy Plant = B100 (GT) / B242 (MT)
        _tab(2);            _type(copy_dc);     time.sleep(0.15)   # Copy DC = 41 (GT) / 21 (MT)

        # Enter → popup closes → Purchasing tab
        _com_enter(session, window="wnd[1]", wait=0.3)
        if not _wait_com(session, _no_popup, timeout=60,
                         label="Purchasing tab (popup closed)"):
            raise Exception("Org Levels popup did not close")

        # ── STEP 4: Purchasing tab ─────────────────────────
        # COM: wait for Purchasing tab to be active (popup closed)
        if not _wait_com(session, _no_popup, timeout=60,
                         label="Purchasing tab ready"):
            raise Exception("Did not reach Purchasing tab after Org Levels")

        _focus_sap()
        _tab(3)           # Tab ×3 → Purchasing Group field
        _type(prch_grp);  time.sleep(0.15)

        # Enter → SAP moves to Accounting 1 tab
        # (only Purchasing + Accounting 1 were selected in Step 2, so Enter goes straight there)
        _com_enter(session, wait=0.5)

        # COM: wait for Accounting 1 tab to be ready (no popup, still MM01)
        _wait_com(session,
                  lambda s: _no_popup(s) and _on_mm01(s),
                  timeout=30, label="Accounting 1 tab ready")

        # ── STEP 5: Accounting 1 tab ───────────────────────
        _focus_sap()
        _tab(8)               # Tab ×8 → Moving price field
        _type(MOVING_PRICE);  time.sleep(0.15)

        # Enter until SAP saves and returns to MM01 initial screen
        _com_enter(session, wait=0.4)   # Enter 1 — confirm value
        _com_enter(session, wait=0.4)   # Enter 2 — trigger save
        _com_enter(session, wait=0.6)   # Enter 3 — confirm save / dismiss any message

        # COM: confirm saved
        if not _wait_com(session, _on_mm01, timeout=60,
                         label="Saved — back to MM01"):
            _l(f"  ⚠ Could not confirm save for {material} — check SAP", "WARN")
        else:
            _l(f"  ✓ {material} → {plant_code}", "OK")

        return True

    except Exception as e:
        _l(f"  ✗ {material} → {plant_code} : {e}", "ERROR")
        try:
            for _ in range(5):
                pyautogui.press("escape")
                time.sleep(0.3)
            from sap_com_helper import sap_tcode
            sap_tcode("MM01", session=session, wait=1.0)
        except Exception:
            pass
        return False


# ─────────────────────────────────────────────
# MAINTAIN MATERIAL GUI CLASS
# ─────────────────────────────────────────────

class MaintainMaterialGui:
    """
    Maintain Material module — launched from launcher.py.
    parent         : tk.Frame provided by launcher
    back_callback  : callable to return to main menu
    theme_manager  : ThemeManager instance for dark/light mode (optional)
    """

    def __init__(self, parent: tk.Frame, back_callback=None, theme_manager=None):
        self.parent        = parent
        self.back_callback = back_callback
        self.theme         = theme_manager or get_theme_manager() if theme else None
        self._log_queue    = queue.Queue()
        self._running      = False
        self._stop_flag    = threading.Event()
        self._plant_vars   = {}      # { plant_code: BooleanVar }
        self._plant_data   = {}      # loaded plant metadata

        # Configure for sharp rendering
        self.parent.tk.call('tk', 'scaling', 2.0)

        self._refresh_colors()
        self._build_ui()
        self._poll_log()

    def _refresh_colors(self):
        """Update all colors from current theme"""
        if self.theme:
            self.BG        = self.theme.get_color("bg")
            self.BG_CARD   = self.theme.get_color("bg_card")
            self.BG_INPUT  = self.theme.get_color("input_bg")
            self.TEXT      = self.theme.get_color("text_dark")
            self.TEXT2     = self.theme.get_color("text_mid")
            self.TEXT3     = self.theme.get_color("text_light")
            self.BORDER    = self.theme.get_color("border")
            self.SUCCESS   = self.theme.get_color("success")
            self.WARNING   = self.theme.get_color("warning")
            self.DANGER    = self.theme.get_color("danger")
            self.BG_DARK   = self.theme.get_color("bg_active")
        else:
            # Fallback to light theme if theme_manager not available
            self.BG        = "#F5F5F5"
            self.BG_CARD   = "#FFFFFF"
            self.BG_INPUT  = "#FFFFFF"
            self.TEXT      = "#1A1A1A"
            self.TEXT2     = "#555555"
            self.TEXT3     = "#999999"
            self.BORDER    = "#E0E0E0"
            self.SUCCESS   = "#2E7D32"
            self.WARNING   = "#E65100"
            self.DANGER    = "#C62828"
            self.BG_DARK   = "#1A1A1A"

    # ── BUILD UI ─────────────────────────────────────────────

    def _build_ui(self):
        self.parent.configure(bg=self.BG)

        main = tk.Frame(self.parent, bg=self.BG)
        main.pack(fill="both", expand=True)

        # ── LEFT PANEL (Scrollable) ───────────────────────────
        left_outer = tk.Frame(main, bg=self.BG, width=340)
        left_outer.pack(side="left", fill="both", padx=(32, 0), pady=24)
        left_outer.pack_propagate(False)

        # Create canvas for scrollable content
        left_canvas = tk.Canvas(left_outer, bg=self.BG,
                                highlightthickness=0, width=320)
        left_vsb = tk.Scrollbar(left_outer, orient="vertical",
                                command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_vsb.set)

        left_vsb.pack(side="right", fill="y")
        left_canvas.pack(side="left", fill="both", expand=True)

        left = tk.Frame(left_canvas, bg=self.BG)
        _left_win = left_canvas.create_window((0, 0), window=left, anchor="nw")

        def _on_left_cfg(e):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))

        def _on_canvas_cfg(e):
            left_canvas.itemconfig(_left_win, width=e.width)

        left.bind("<Configure>", _on_left_cfg)
        left_canvas.bind("<Configure>", _on_canvas_cfg)

        def _scroll_left(e):
            left_canvas.yview_scroll(int(-1*(e.delta/120)), "units")

        def _on_mousewheel_global(event):
            """Global mouse wheel handler that checks if mouse is over left panel."""
            # Get mouse position
            try:
                x = event.x_root
                y = event.y_root
                # Get left canvas bounding box
                x1 = left_outer.winfo_rootx()
                y1 = left_outer.winfo_rooty()
                x2 = x1 + left_outer.winfo_width()
                y2 = y1 + left_outer.winfo_height()

                # If mouse is over left panel, scroll
                if x1 <= x <= x2 and y1 <= y <= y2:
                    _scroll_left(event)
            except:
                pass

        # Bind scroll directly to the canvas - this will work for canvas area
        left_canvas.bind("<MouseWheel>", _scroll_left)
        # Bind global mouse wheel to check if over left panel
        self.parent.bind_all("<MouseWheel>", _on_mousewheel_global)

        tk.Label(left, text="Maintain Material",
                 font=(FONT_DISP, 15, "bold"), fg=self.TEXT, bg=self.BG,
                 ).pack(anchor="w", pady=(0, 2))
        tk.Label(left, text="SAP MM01 — Create Material per Plant",
                 font=(FONT, 8), fg=self.TEXT3, bg=self.BG,
                 ).pack(anchor="w", pady=(0, 16))

        # ── SECTION: Files ────────────────────────────────────
        self._section_label(left, "Files")
        f_card = self._card(left)

        # Material codes file
        self._field_row(f_card, "Material Codes File")
        mf_row = tk.Frame(f_card, bg=self.BG_CARD)
        mf_row.pack(fill="x", pady=(2, 6))
        self.mat_file_var = tk.StringVar()
        tk.Entry(mf_row, textvariable=self.mat_file_var,
                 font=(FONT, 8), bg=self.BG_INPUT, fg=self.TEXT,
                 insertbackground=self.TEXT, relief="flat", bd=0,
                 highlightthickness=1, highlightbackground=self.BORDER,
                 highlightcolor=self.BG_DARK, width=24,
                 ).pack(side="left", ipady=4)
        tk.Button(mf_row, text="📁", font=(FONT, 9),
                  bg=self.BG_CARD, fg=self.TEXT2, relief="flat", bd=0,
                  cursor="hand2", padx=6,
                  command=self._browse_material_file,
                  ).pack(side="left", padx=(4, 0))
        self._mat_file_lbl = tk.Label(
            f_card, text="  No file selected",
            font=(FONT, 7), fg=self.TEXT3, bg=self.BG_CARD, anchor="w")
        self._mat_file_lbl.pack(anchor="w", pady=(0, 4))

        # Plant list file
        self._field_row(f_card, "Plant List  (List_Plant.xlsx)")
        pl_row = tk.Frame(f_card, bg=self.BG_CARD)
        pl_row.pack(fill="x", pady=(2, 6))
        self.plant_file_var = tk.StringVar()
        tk.Entry(pl_row, textvariable=self.plant_file_var,
                 font=(FONT, 8), bg=self.BG_INPUT, fg=self.TEXT,
                 insertbackground=self.TEXT, relief="flat", bd=0,
                 highlightthickness=1, highlightbackground=self.BORDER,
                 highlightcolor=self.BG_DARK, width=24,
                 ).pack(side="left", ipady=4)
        tk.Button(pl_row, text="📁", font=(FONT, 9),
                  bg=self.BG_CARD, fg=self.TEXT2, relief="flat", bd=0,
                  cursor="hand2", padx=6,
                  command=self._browse_plant_file,
                  ).pack(side="left", padx=(4, 0))
        self._plant_file_lbl = tk.Label(
            f_card, text="  No file selected",
            font=(FONT, 7), fg=self.TEXT3, bg=self.BG_CARD, anchor="w")
        self._plant_file_lbl.pack(anchor="w", pady=(0, 4))

        # ── SECTION: Plants ───────────────────────────────────
        self._section_label(left, "Select Plants")
        p_card = self._card(left)

        # Header row: count + All / Clear
        ph = tk.Frame(p_card, bg=self.BG_CARD)
        ph.pack(fill="x", pady=(0, 4))
        self._plant_count_lbl = tk.Label(
            ph, text="Load plant list file first",
            font=(FONT, 7), fg=self.TEXT3, bg=self.BG_CARD)
        self._plant_count_lbl.pack(side="left")
        tk.Button(ph, text="All", font=(FONT, 7),
                  fg=self.TEXT2, bg=self.BG_CARD, relief="flat", bd=0,
                  cursor="hand2", command=self._plant_select_all,
                  ).pack(side="right", padx=(4, 0))
        tk.Button(ph, text="Clear", font=(FONT, 7),
                  fg=self.TEXT3, bg=self.BG_CARD, relief="flat", bd=0,
                  cursor="hand2", command=self._plant_clear_all,
                  ).pack(side="right")

        tk.Frame(p_card, bg=self.BORDER, height=1).pack(fill="x", pady=(0, 4))

        # Scrollable checklist
        grid_canvas = tk.Canvas(p_card, bg=self.BG_CARD,
                                highlightthickness=0, height=160)
        grid_vsb = tk.Scrollbar(p_card, orient="vertical",
                                command=grid_canvas.yview)
        grid_canvas.configure(yscrollcommand=grid_vsb.set)
        grid_vsb.pack(side="right", fill="y")
        grid_canvas.pack(fill="both", expand=True)

        self._plant_grid = tk.Frame(grid_canvas, bg=self.BG_CARD)
        _win = grid_canvas.create_window(
            (0, 0), window=self._plant_grid, anchor="nw")

        def _on_cfg(e):
            grid_canvas.configure(scrollregion=grid_canvas.bbox("all"))

        def _on_canvas_cfg(e):
            grid_canvas.itemconfig(_win, width=e.width)

        self._plant_grid.bind("<Configure>", _on_cfg)
        grid_canvas.bind("<Configure>", _on_canvas_cfg)

        def _scroll(e):
            grid_canvas.yview_scroll(int(-1*(e.delta/120)), "units")

        grid_canvas.bind("<MouseWheel>", _scroll)
        self._plant_grid.bind("<MouseWheel>", _scroll)

        # ── SECTION: Info ─────────────────────────────────────
        self._section_label(left, "Info")
        i_card = self._card(left)
        for dot_c, msg in [
            (self.SUCCESS, "GT → DC=41, Copy from B100"),
            (self.WARNING, "MT → DC=21, Copy from B242"),
            (self.TEXT2,   "Outer loop = plants"),
            (self.TEXT2,   "Inner loop = material codes"),
            (self.WARNING, "SAP must be open before Run"),
        ]:
            rf = tk.Frame(i_card, bg=self.BG_CARD)
            rf.pack(fill="x", pady=1)
            tk.Label(rf, text="●", fg=dot_c, bg=self.BG_CARD,
                     font=(FONT, 8)).pack(side="left", padx=(0, 6))
            tk.Label(rf, text=msg, fg=self.TEXT2, bg=self.BG_CARD,
                     font=(FONT, 8)).pack(side="left")

        # ── RUN / STOP buttons ────────────────────────────────
        btn_frame = tk.Frame(left, bg=self.BG)
        btn_frame.pack(fill="x", pady=(16, 0))

        self.stop_btn = tk.Button(
            btn_frame,
            text="■  Stop",
            font=(FONT_DISP, 10, "bold"),
            fg=self.BG_CARD, bg="#8B0000",
            activebackground="#C62828",
            activeforeground=self.BG_CARD,
            relief="flat", bd=0,
            padx=14, pady=10,
            cursor="hand2", state="disabled",
            command=self._on_stop,
        )
        self.stop_btn.pack(side="right", padx=(8, 0))
        self.stop_btn.bind("<Enter>",
            lambda e: self.stop_btn.configure(bg="#C62828")
            if self.stop_btn["state"] != "disabled" else None)
        self.stop_btn.bind("<Leave>",
            lambda e: self.stop_btn.configure(bg="#8B0000")
            if self.stop_btn["state"] != "disabled" else None)

        self.run_btn = tk.Button(
            btn_frame,
            text="▶   Run MM01",
            font=(FONT_DISP, 11, "bold"),
            fg=self.BG_CARD, bg=self.BG_DARK,
            activebackground="#333333",
            activeforeground=self.BG_CARD,
            relief="flat", bd=0,
            padx=20, pady=10,
            cursor="hand2",
            command=self._on_run,
        )
        self.run_btn.pack(side="left", fill="x", expand=True)
        self.run_btn.bind("<Enter>",
            lambda e: self.run_btn.configure(bg="#333333")
            if self.run_btn["state"] != "disabled" else None)
        self.run_btn.bind("<Leave>",
            lambda e: self.run_btn.configure(bg=self.BG_DARK)
            if self.run_btn["state"] != "disabled" else None)

        # ── RIGHT PANEL: log ──────────────────────────────────
        right = tk.Frame(main, bg=self.BG)
        right.pack(side="left", fill="both", expand=True,
                   padx=24, pady=24)

        log_hdr = tk.Frame(right, bg=self.BG)
        log_hdr.pack(fill="x", pady=(0, 6))
        tk.Label(log_hdr, text="ACTIVITY LOG",
                 font=(FONT, 8, "bold"), fg=self.TEXT3, bg=self.BG,
                 ).pack(side="left")
        tk.Button(log_hdr, text="Clear", font=(FONT, 7),
                  fg=self.TEXT3, bg=self.BG, relief="flat", bd=0,
                  cursor="hand2", command=self._clear_log,
                  ).pack(side="right")

        log_outer = tk.Frame(right, bg=self.BORDER, bd=0)
        log_outer.pack(fill="both", expand=True)
        log_inner = tk.Frame(log_outer, bg=self.BG_CARD)
        log_inner.pack(fill="both", padx=1, pady=1)

        self.log_box = scrolledtext.ScrolledText(
            log_inner,
            bg=self.BG_CARD, fg=self.TEXT2,
            font=(FONT, 9),
            relief="flat", bd=0,
            state="disabled", wrap="word",
            padx=12, pady=10,
        )
        self.log_box.pack(fill="both", expand=True)
        self.log_box.tag_config("OK",    foreground=self.SUCCESS)
        self.log_box.tag_config("ERROR", foreground=self.DANGER)
        self.log_box.tag_config("WARN",  foreground=self.WARNING)
        self.log_box.tag_config("INFO",  foreground=self.TEXT2)

        self._write_log("Maintain Material ready.", "INFO")
        self._write_log("1. Browse material codes file", "INFO")
        self._write_log("2. Browse plant list file", "INFO")
        self._write_log("3. Select plants → Run MM01", "INFO")

    # ── UI HELPERS ───────────────────────────────────────────

    def _section_label(self, parent, text):
        f = tk.Frame(parent, bg=self.BG)
        f.pack(fill="x", pady=(12, 4))
        tk.Frame(f, bg=self.BG_DARK, width=3, height=14).pack(
            side="left", padx=(0, 6))
        tk.Label(f, text=text.upper(),
                 font=(FONT, 8, "bold"), fg=self.TEXT, bg=self.BG,
                 ).pack(side="left")

    def _card(self, parent):
        outer = tk.Frame(parent, bg=self.BORDER)
        outer.pack(fill="x", pady=(0, 4))
        inner = tk.Frame(outer, bg=self.BG_CARD, padx=14, pady=10)
        inner.pack(fill="x", padx=1, pady=1)
        return inner

    def _field_row(self, parent, label, width=None):
        tk.Label(parent, text=label,
                 font=(FONT, 9), fg=self.TEXT2, bg=self.BG_CARD, anchor="w",
                 ).pack(anchor="w", pady=(6, 0))

    def _entry(self, parent, var, width=28, placeholder=""):
        e = tk.Entry(
            parent, textvariable=var, width=width,
            bg=self.BG_INPUT, fg=self.TEXT, insertbackground=self.TEXT,
            font=(FONT, 10), relief="flat", bd=0,
            highlightthickness=1,
            highlightbackground=self.BORDER, highlightcolor=self.BG_DARK,
        )
        e.pack(anchor="w", ipady=5, pady=(2, 4))
        if placeholder and not var.get():
            e.insert(0, placeholder)
            e.config(fg=self.TEXT3)

            def _fi(event, entry=e, ph=placeholder):
                if entry.get() == ph:
                    entry.delete(0, "end")
                    entry.config(fg=self.TEXT)

            def _fo(event, entry=e, ph=placeholder):
                if not entry.get():
                    entry.insert(0, ph)
                    entry.config(fg=self.TEXT3)

            e.bind("<FocusIn>",  _fi)
            e.bind("<FocusOut>", _fo)
        return e

    # ── FILE BROWSERS ─────────────────────────────────────────

    def _browse_material_file(self):
        path = filedialog.askopenfilename(
            title="Select Material Codes Excel File",
            filetypes=[("Excel File", "*.xlsx *.xls"),
                       ("All Files", "*.*")],
        )
        if not path:
            return
        self.mat_file_var.set(path)
        try:
            codes = load_material_codes(path)
            self._mat_file_lbl.config(
                text=f"  ✔ {len(codes)} codes — {os.path.basename(path)}",
                fg=self.SUCCESS)
            self._write_log(
                f"Material file: {os.path.basename(path)} "
                f"→ {len(codes)} codes", "OK")
        except Exception as e:
            self._mat_file_lbl.config(text=f"  ⚠ {e}", fg=self.WARNING)
            self._write_log(f"Material file error: {e}", "ERROR")

    def _browse_plant_file(self):
        path = filedialog.askopenfilename(
            title="Select Plant List Excel File",
            filetypes=[("Excel File", "*.xlsx *.xls"),
                       ("All Files", "*.*")],
        )
        if not path:
            return
        self.plant_file_var.set(path)
        try:
            self._plant_data = load_plant_data(path)
            self._plant_file_lbl.config(
                text=f"  ✔ {len(self._plant_data)} plants — "
                     f"{os.path.basename(path)}",
                fg=self.SUCCESS)
            self._write_log(
                f"Plant list: {os.path.basename(path)} "
                f"→ {len(self._plant_data)} plants", "OK")
            self._rebuild_plant_checklist()
        except Exception as e:
            self._plant_file_lbl.config(text=f"  ⚠ {e}", fg=self.WARNING)
            self._write_log(f"Plant file error: {e}", "ERROR")

    # ── PLANT CHECKLIST ───────────────────────────────────────

    def _rebuild_plant_checklist(self):
        """Rebuild plant checkboxes from loaded plant data."""
        for w in self._plant_grid.winfo_children():
            w.destroy()

        old_states     = {p: v.get() for p, v in self._plant_vars.items()}
        self._plant_vars = {}

        plants = sorted(self._plant_data.keys())
        for i, plant in enumerate(plants):
            info  = self._plant_data[plant]
            is_mt = info["type"] == "MT"
            label = f"{plant}  [{info['type']}]"
            color = self.WARNING if is_mt else self.TEXT2

            var = tk.BooleanVar(value=old_states.get(plant, True))
            self._plant_vars[plant] = var

            cb = tk.Checkbutton(
                self._plant_grid, text=label,
                variable=var,
                font=(FONT, 8),
                fg=color, bg=self.BG_CARD,
                selectcolor=self.BG_CARD,
                activebackground=self.BG_CARD,
                activeforeground=color,
                relief="flat", bd=0,
            )
            cb.grid(row=i // 2, column=i % 2, sticky="w",
                    padx=4, pady=1)

        n   = len(plants)
        gt  = sum(1 for p in self._plant_data.values() if p["type"] == "GT")
        mt  = sum(1 for p in self._plant_data.values() if p["type"] == "MT")
        self._plant_count_lbl.config(
            text=f"{n} plants  (GT:{gt}  MT:{mt})",
            fg=self.TEXT2)

    def _plant_select_all(self):
        for v in self._plant_vars.values():
            v.set(True)

    def _plant_clear_all(self):
        for v in self._plant_vars.values():
            v.set(False)

    def _get_selected_plants(self) -> list:
        return [p for p, v in self._plant_vars.items() if v.get()]

    # ── CLEAR LOG ─────────────────────────────────────────────

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ── RUN / STOP ────────────────────────────────────────────

    def _on_run(self):
        if self._running:
            return

        mat_file   = self.mat_file_var.get().strip()
        plant_file = self.plant_file_var.get().strip()
        plants     = self._get_selected_plants()

        if not mat_file or not os.path.exists(mat_file):
            messagebox.showwarning("Missing Input",
                                   "Please select a material codes file.")
            return
        if not plant_file or not os.path.exists(plant_file):
            messagebox.showwarning("Missing Input",
                                   "Please select a plant list file.")
            return
        if not plants:
            messagebox.showwarning("Missing Input",
                                   "Please select at least one plant.")
            return

        self._stop_flag.clear()
        self._running = True
        self.run_btn.configure(state="disabled",
                               text="⏳  Running...", bg="#555555")
        self.stop_btn.configure(state="normal")

        self._write_log("━" * 45, "INFO")
        self._write_log("Maintain Material started", "OK")
        self._write_log(f"Plants   : {', '.join(plants)}", "INFO")
        self._write_log(
            f"Material : {os.path.basename(mat_file)}", "INFO")
        self._write_log("━" * 45, "INFO")

        def _run():
            self._execute_mm01(mat_file, plant_file, plants)

        threading.Thread(target=_run, daemon=True).start()

    def _on_stop(self):
        self._stop_flag.set()
        self._write_log("─" * 45, "WARN")
        self._write_log("Stop requested — finishing current step...", "WARN")
        self._write_log("─" * 45, "WARN")
        self.stop_btn.configure(state="disabled")

    def _reset_buttons(self):
        """Re-enable Run button (called on main thread after run ends)."""
        self._running = False
        self.run_btn.configure(
            state="normal", text="▶   Run MM01", bg=self.BG_DARK)
        self.stop_btn.configure(state="disabled")

    # ── SAP EXECUTION ─────────────────────────────────────────

    def _execute_mm01(self, mat_file: str, plant_file: str,
                      plants: list):
        """
        Background thread: run MM01 for all plants × all materials.
        Outer loop = plants  (finish all materials before next plant)
        Inner loop = material codes
        """
        def _l(msg, level="INFO"):
            self._log(msg, level)

        try:
            # Load data
            _l("Loading material codes...")
            materials = load_material_codes(mat_file)
            if not materials:
                _l("No material codes found in file!", "ERROR")
                return
            _l(f"{len(materials)} material codes loaded", "OK")

            _l("Loading plant data...")
            plant_data = load_plant_data(plant_file)
            _l(f"{len(plant_data)} plants loaded", "OK")

            # Connect to SAP via COM
            _l("Connecting to SAP via COM scripting...")
            session = _get_session()
            _l(f"SAP connected — User={session.Info.User} | "
               f"System={session.Info.SystemName}", "OK")

            _l(f"Total operations: "
               f"{len(plants)} plants × {len(materials)} materials "
               f"= {len(plants)*len(materials)}", "INFO")
            _l("SAP is now under robot control — "
               "do not touch keyboard/mouse!", "WARN")

            total_ok   = 0
            total_fail = 0

            # ── OUTER LOOP: plants ─────────────────────────
            for plant_code in plants:
                if self._stop_flag.is_set():
                    _l("Stopped by user.", "WARN")
                    break

                if plant_code not in plant_data:
                    _l(f"Plant {plant_code} not in plant list — skip",
                       "WARN")
                    continue

                info  = plant_data[plant_code]
                ptype = info["type"]

                _l("─" * 40, "INFO")
                _l(f"Plant {plant_code} — {info['name']} [{ptype}]",
                   "INFO")
                _l(f"DC={info['dc']} | PrchGrp={info['prch_grp']} | "
                   f"CopyFrom={info['copy_plant']}", "INFO")

                ok_count   = 0
                fail_count = 0
                failed     = []

                # ── INNER LOOP: materials ──────────────────
                for idx, material in enumerate(materials, start=1):
                    if self._stop_flag.is_set():
                        _l("Stopped by user.", "WARN")
                        break

                    _l(f"[{idx}/{len(materials)}] "
                       f"{material} → {plant_code}...")

                    ok = _mm01_one(
                        session    = session,
                        material   = material,
                        plant_info = info,
                        plant_code = plant_code,
                        log_fn     = _l,
                    )

                    if ok:
                        ok_count  += 1
                        total_ok  += 1
                    else:
                        fail_count += 1
                        total_fail += 1
                        failed.append(material)

                    time.sleep(0.3)

                _l(f"Plant {plant_code} done — "
                   f"✓ {ok_count} | ✗ {fail_count}",
                   "OK" if fail_count == 0 else "WARN")
                if failed:
                    _l(f"  Failed: {', '.join(failed)}", "WARN")

            # Summary
            _l("━" * 45, "INFO")
            _l("SUMMARY", "INFO")
            _l(f"Total: ✓ {total_ok} created | ✗ {total_fail} failed",
               "OK" if total_fail == 0 else "WARN")
            _l("━" * 45, "INFO")

        except Exception as e:
            _l(f"FATAL: {e}", "ERROR")

        finally:
            self.parent.after(0, self._reset_buttons)

    # ── LOGGING ──────────────────────────────────────────────

    def _log(self, msg: str, level: str = "INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_queue.put((level, f"{ts}  {msg}"))

    def _write_log(self, msg: str, level: str = "INFO"):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n", level)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _poll_log(self):
        try:
            while True:
                level, msg = self._log_queue.get_nowait()
                self._write_log(msg, level)
        except queue.Empty:
            pass
        self.parent.after(80, self._poll_log)


# ─────────────────────────────────────────────
# STANDALONE ENTRY POINT (for testing)
# ─────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Maintain Material — MM01")
    root.geometry("1100x680")
    root.configure(bg="#F5F5F5")
    frame = tk.Frame(root, bg="#F5F5F5")
    frame.pack(fill="both", expand=True)
    MaintainMaterialGui(frame, back_callback=root.destroy)
    root.mainloop()