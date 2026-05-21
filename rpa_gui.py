"""
rpa_gui.py
GUI untuk RPA Stock Reconciliation — UI Modern
Jalankan: python rpa_gui.py
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import queue
import os
from datetime import datetime

# ─────────────────────────────────────────────
# QUEUE & EVENTS
# ─────────────────────────────────────────────
log_queue   = queue.Queue()
stop_event  = threading.Event()
login_event = threading.Event()
sap_event   = threading.Event()  # sinyal dari GUI ke robot: SAP sudah terbuka

# ─────────────────────────────────────────────
# COLOR PALETTE — GitHub-inspired dark
# ─────────────────────────────────────────────
C = {
    "bg":        "#0D1117",
    "surface":   "#161B22",
    "surface2":  "#21262D",
    "border":    "#30363D",
    "accent":    "#2F81F7",
    "accent2":   "#1F6FEB",
    "success":   "#3FB950",
    "warning":   "#D29922",
    "danger":    "#F85149",
    "text":      "#E6EDF3",
    "text2":     "#8B949E",
    "text3":     "#484F58",
    "run_bg":    "#238636",
    "run_hov":   "#2EA043",
    "stop_bg":   "#6E3535",
    "stop_hov":  "#DA3633",
}


def send_log(msg: str, level: str = "INFO"):
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_queue.put((level, f"{timestamp}  {msg}"))


# ─────────────────────────────────────────────
# ROBOT LOGIC
# ─────────────────────────────────────────────

def wait_for_login(send_log):
    """
    Tampilkan dialog di GUI thread — minta user konfirmasi sudah login.
    Dipanggil dari background thread via queue.
    """
    log_queue.put(("WAIT_LOGIN", ""))


def run_robot(plants, posting_date, email_to, email_cc, mode):
    try:
        # Cek apakah Chrome + CDP sudah aktif
        import urllib.request
        cdp_ok = False
        for host in ["127.0.0.1", "localhost"]:
            try:
                urllib.request.urlopen(f"http://{host}:9222/json/version", timeout=2)
                cdp_ok = True
                break
            except Exception:
                continue

        if not cdp_ok:
            # Chrome belum terbuka dengan CDP — buka dulu
            send_log("Membuka Chrome ke halaman login portal...", "INFO")
            import subprocess, os
            chrome_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            ]
            chrome = None
            for cp in chrome_paths:
                if os.path.exists(cp):
                    chrome = cp
                    break
            if not chrome:
                send_log("Chrome tidak ditemukan!", "ERROR")
                log_queue.put(("DONE", ""))
                return

            # Tutup Chrome yang ada
            subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"],
                           capture_output=True, timeout=5)
            import time
            time.sleep(2)

            # Buka Chrome ke halaman login
            subprocess.Popen([
                chrome,
                "--remote-debugging-port=9222",
                "--remote-debugging-address=0.0.0.0",
                "--user-data-dir=C:\\ChromeRPA",
                "--no-first-run",
                "--no-default-browser-check",
                "https://portal.mayora.co.id/v2login",
            ])

            # Tunggu Chrome siap
            for _ in range(40):
                time.sleep(0.5)
                try:
                    urllib.request.urlopen("http://127.0.0.1:9222/json/version", timeout=1)
                    break
                except Exception:
                    continue

            send_log("Chrome terbuka — silakan LOGIN di browser!", "WARN")
            send_log("Setelah login, klik OK di dialog konfirmasi.", "WARN")

            # Minta user konfirmasi login via GUI dialog
            log_queue.put(("WAIT_LOGIN", ""))

            # Tunggu konfirmasi dari GUI (event)
            login_event.wait()
            login_event.clear()
            send_log("Login dikonfirmasi — melanjutkan RPA...", "OK")

        from test_compare import run_full_pipeline
        send_log("━" * 50, "INFO")
        send_log(f"RPA dimulai  ·  mode={mode}  ·  {posting_date}", "OK")
        send_log(f"Email → {email_to}" + (f"  CC: {email_cc}" if email_cc else ""), "INFO")
        send_log("━" * 50, "INFO")
        send_log("FASE 1: Scan portal untuk plant dengan selisih...", "INFO")

        # Inject hook ke send_log untuk detect kapan SAP dikontrol
        original_send_log = send_log
        def send_log_with_sap_warning(msg, level="INFO"):
            if "SAP akan dikontrol" in msg:
                log_queue.put(("SAP_WARNING", ""))
            elif level == "SAP_WAIT":
                log_queue.put(("SAP_WAIT", ""))
                sap_event.wait()
                sap_event.clear()
                # Kirim SAP_WAIT_DONE ke test_compare via special call
                return
            original_send_log(msg, level)

        items = run_full_pipeline(
            plants       = None if mode == "auto_scan" else plants,
            posting_date = posting_date,
            email_to     = email_to,
            email_cc     = email_cc,
            send_log     = send_log_with_sap_warning,
        )

        send_log("", "INFO")
        send_log("━" * 50, "INFO")
        send_log("SUMMARY", "INFO")
        if items:
            for plant, itms in sorted(items.items()):
                t917 = sum(1 for i in itms if i.mvt_type == "917")
                t918 = sum(1 for i in itms if i.mvt_type == "918")
                send_log(f"Plant {plant}  ·  {len(itms)} item  ·  917:{t917}  918:{t918}", "OK")
        else:
            send_log("Tidak ada selisih ditemukan.", "INFO")

        send_log("RPA selesai!", "OK")
        log_queue.put(("DONE", ""))
    except Exception as e:
        send_log(f"FATAL: {e}", "ERROR")
        log_queue.put(("DONE", ""))


# ─────────────────────────────────────────────
# HOVER BUTTON
# ─────────────────────────────────────────────

class HBtn(tk.Button):
    def __init__(self, master, n, h, **kw):
        self._n, self._h = n, h
        super().__init__(master, bg=n, activebackground=h, **kw)
        self.bind("<Enter>", lambda e: self.config(bg=self._h) if str(self["state"]) != "disabled" else None)
        self.bind("<Leave>", lambda e: self.config(bg=self._n) if str(self["state"]) != "disabled" else None)

    def recolor(self, n, h):
        self._n, self._h = n, h
        self.config(bg=n, activebackground=h)


# ─────────────────────────────────────────────
# MAIN GUI CLASS
# ─────────────────────────────────────────────

class RpaGui:
    def __init__(self, root):
        self.root   = root
        self.root.title("RPA Stock Reconciliation")
        self.root.geometry("1020x760")
        self.root.minsize(860, 640)
        self.root.configure(bg=C["bg"])
        self._running = False
        self._build()
        self._load_email_defaults()
        self._poll()

    # ── UI BUILD ─────────────────────────────────────────────

    def _build(self):
        from config import Config

        # TOP BAR
        top = tk.Frame(self.root, bg=C["surface"], height=52)
        top.pack(fill="x")
        top.pack_propagate(False)

        tk.Label(top, text="◈  RPA Stock Reconciliation",
                 font=("Calibri", 13, "bold"),
                 fg=C["text"], bg=C["surface"]
                 ).pack(side="left", padx=20, pady=14)

        pill = tk.Frame(top, bg=C["surface2"],
                        highlightbackground=C["border"], highlightthickness=1)
        pill.pack(side="right", padx=16, pady=14)
        self._dot = tk.Label(pill, text="●", fg=C["success"],
                             bg=C["surface2"], font=("Calibri", 9))
        self._dot.pack(side="left", padx=(10, 3), pady=3)
        self._stat = tk.Label(pill, text="Ready",
                              fg=C["text2"], bg=C["surface2"],
                              font=("Calibri", 9))
        self._stat.pack(side="left", padx=(0, 10), pady=3)

        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x")

        # BODY
        body = tk.Frame(self.root, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=18, pady=14)

        # LEFT — config (scrollable)
        left_outer = tk.Frame(body, bg=C["bg"], width=330)
        left_outer.pack(side="left", fill="y", padx=(0, 14))
        left_outer.pack_propagate(False)

        self._left_canvas = tk.Canvas(
            left_outer, bg=C["bg"], highlightthickness=0, width=314
        )
        left_sb = tk.Scrollbar(
            left_outer, orient="vertical",
            command=self._left_canvas.yview
        )
        self._left_canvas.configure(yscrollcommand=left_sb.set)

        left_sb.pack(side="right", fill="y")
        self._left_canvas.pack(side="left", fill="both", expand=True)

        self._left_frame = tk.Frame(self._left_canvas, bg=C["bg"])
        self._left_win   = self._left_canvas.create_window(
            (0, 0), window=self._left_frame, anchor="nw"
        )

        self._left_frame.bind("<Configure>", self._on_left_configure)
        self._left_canvas.bind("<Configure>", self._on_canvas_configure)

        # Scroll mouse wheel saat hover di panel kiri
        self._left_canvas.bind("<Enter>",
            lambda e: self._left_canvas.bind_all(
                "<MouseWheel>", self._on_mousewheel))
        self._left_canvas.bind("<Leave>",
            lambda e: self._left_canvas.unbind_all("<MouseWheel>"))

        self._build_left(self._left_frame, Config)

        # RIGHT — log
        right = tk.Frame(body, bg=C["bg"])
        right.pack(side="left", fill="both", expand=True)
        self._build_right(right)

        # BOTTOM
        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x")
        self._build_bottom()

    def _on_left_configure(self, event):
        self._left_canvas.configure(
            scrollregion=self._left_canvas.bbox("all")
        )

    def _on_canvas_configure(self, event):
        self._left_canvas.itemconfig(self._left_win, width=event.width)

    def _on_mousewheel(self, event):
        self._left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_portal_change(self, event=None):
        """Update URL portal aktif + T-code SAP yang sesuai saat dropdown berubah."""
        from config import Config

        label = self._portal_var.get()
        url   = self._portal_options.get(label, "")

        # Update URL di test_compare
        try:
            from test_compare import set_portal_url
            set_portal_url(url)
        except Exception:
            pass

        # Update T-code aktif berdasarkan portal yang dipilih
        tmap = Config.PORTAL_TCODE_MAP.get(label, {})
        Config.ACTIVE_TCODE_SAPSTK = tmap.get("sapstk", "ZPGD_SAPSTK")
        Config.ACTIVE_TCODE_U2C    = tmap.get("u2c",    "ZPGD_U2C")

        # Update label URL
        self._portal_url_lbl.config(text=f"  → {url}")

        # Update label T-code (jika widget sudah dibuat)
        if hasattr(self, "_tcode_info_lbl"):
            self._tcode_info_lbl.config(
                text=self._tcode_info_text(label)
            )

    def _tcode_info_text(self, portal_label: str = None) -> str:
        """Return string T-code SAP aktif untuk ditampilkan di label."""
        from config import Config
        if portal_label is None:
            portal_label = self._portal_var.get() if hasattr(self, "_portal_var") else "PGDMTX"
        tmap = Config.PORTAL_TCODE_MAP.get(portal_label, {})
        sapstk = tmap.get("sapstk", Config.ACTIVE_TCODE_SAPSTK)
        u2c    = tmap.get("u2c",    Config.ACTIVE_TCODE_U2C)
        return f"  SAP: /{sapstk}  +  /{u2c}"

    def _load_plants_from_excel(self, filepath: str = None) -> list:
        """
        Baca daftar plant dari file Excel mapping (sheet Plant_CostCenter, kolom B).
        Fallback ke Config.PLANTS hanya jika file benar-benar tidak bisa dibaca,
        dan tampilkan pesan error yang jelas di log.
        """
        from config import Config as Cfg
        if filepath is None:
            filepath = self.plant_map_var.get() if hasattr(self, 'plant_map_var') else Cfg.PLANT_MAPPING_FILE

        # Simpan path ke Config supaya konsisten di seluruh session
        Cfg.PLANT_MAPPING_FILE = filepath

        try:
            import openpyxl, os
            if not os.path.exists(filepath):
                raise FileNotFoundError(f"File tidak ditemukan: {filepath}")

            wb = openpyxl.load_workbook(filepath, data_only=True)

            if "Plant_CostCenter" not in wb.sheetnames:
                raise ValueError(
                    f"Sheet 'Plant_CostCenter' tidak ada di file.\n"
                    f"Sheet yang tersedia: {wb.sheetnames}"
                )

            ws = wb["Plant_CostCenter"]
            plants = []
            for row in ws.iter_rows(min_row=5, values_only=True):
                if row[1]:  # kolom B = Plant Code
                    code = str(row[1]).strip()
                    if code and code not in plants:
                        plants.append(code)

            if not plants:
                raise ValueError("Tidak ada data plant di sheet Plant_CostCenter (mulai baris 5, kolom B)")

            return sorted(plants)

        except Exception as e:
            # Tampilkan ke log GUI jika sudah ada, juga print ke console
            err_msg = f"[Plant Map] Gagal baca '{filepath}': {e}"
            print(err_msg)
            # Kirim ke log GUI setelah widget siap (pakai after agar tidak crash di __init__)
            try:
                self.root.after(500, lambda: self._write_log(
                    f"⚠  {err_msg}\n   → Pakai hardcode plants sebagai fallback. "
                    f"Pilih file yang benar via tombol 📁 Plant Map.",
                    "WARN"
                ))
            except Exception:
                pass
            return list(Cfg.PLANTS)

    def _rebuild_plant_checklist(self, grid_frame, plants: list):
        """Hapus dan rebuild checkbox plant di grid."""
        for w in grid_frame.winfo_children():
            w.destroy()
        # Simpan state lama
        old_states = {p: v.get() for p, v in self._plant_vars.items()}
        # Rebuild BooleanVar
        self._plant_vars = {}
        for plant in plants:
            var = tk.BooleanVar(value=old_states.get(plant, True))
            self._plant_vars[plant] = var
            cb = tk.Checkbutton(
                grid_frame, text=plant,
                variable=var,
                font=("Calibri", 9),
                fg=C["text"], bg=C["surface2"],
                selectcolor=C["surface"],
                activebackground=C["surface2"],
                activeforeground=C["text"],
                relief="flat", bd=0,
                command=self._update_plant_summary,
            )
            i = plants.index(plant)
            cb.grid(row=i//2, column=i%2, sticky="w", padx=4, pady=1)
        self._update_plant_summary()

    def _build_left(self, p, Config):

        def section(title):
            f = tk.Frame(p, bg=C["bg"])
            f.pack(fill="x", pady=(12, 3))
            tk.Label(f, text=title, fg=C["text3"], bg=C["bg"],
                     font=("Calibri", 7, "bold")).pack(side="left")

        def card():
            o = tk.Frame(p, bg=C["border"])
            o.pack(fill="x", pady=(0, 2))
            i = tk.Frame(o, bg=C["surface"])
            i.pack(fill="both", padx=1, pady=1)
            return i

        def row_entry(parent, label, var, width=22, show=""):
            f = tk.Frame(parent, bg=C["surface"])
            f.pack(fill="x", pady=1)
            tk.Label(f, text=label, fg=C["text2"], bg=C["surface"],
                     font=("Calibri", 9), width=13, anchor="w"
                     ).pack(side="left", padx=(12, 4), pady=7)
            e = tk.Entry(f, textvariable=var, show=show, width=width,
                         bg=C["surface2"], fg=C["text"],
                         insertbackground=C["accent"],
                         font=("Calibri", 9), relief="flat", bd=5,
                         highlightthickness=1,
                         highlightbackground=C["border"],
                         highlightcolor=C["accent"])
            e.pack(side="left", padx=(0, 10))
            return e

        # ── CONFIGURATION ──────────────────────────────────
        section("CONFIGURATION")
        c1 = card()

        self.date_var      = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))
        self.plant_map_var = tk.StringVar(value=Config.PLANT_MAPPING_FILE)

        # Baca daftar plant dari file Excel mapping.
        # Kalau file tidak ada / belum diset → fallback Config.PLANTS dulu,
        # pesan error muncul di Activity Log setelah GUI siap.
        _initial_plants = self._load_plants_from_excel(Config.PLANT_MAPPING_FILE)

        # Cek apakah hasil benar-benar dari Excel atau dari fallback hardcode
        _from_excel = sorted(_initial_plants) != sorted(list(Config.PLANTS)) or \
                      len(_initial_plants) != len(Config.PLANTS)
        self._plant_map_loaded_ok = _from_excel  # flag untuk status label

        self._plant_vars = {p: tk.BooleanVar(value=True) for p in _initial_plants}

        row_entry(c1, "Posting Date", self.date_var, width=13)

        # ── Portal EOD URL ────────────────────────────────
        from test_compare import PORTAL_EOD_URLS
        self._portal_options = PORTAL_EOD_URLS   # { label: url }
        self._portal_var = tk.StringVar(
            value=list(self._portal_options.keys())[0]
        )

        purl_row = tk.Frame(c1, bg=C["surface"])
        purl_row.pack(fill="x", pady=1)
        tk.Label(purl_row, text="Portal EOD", fg=C["text2"], bg=C["surface"],
                 font=("Calibri", 9), width=13, anchor="w"
                 ).pack(side="left", padx=(12, 4), pady=7)

        self._portal_combo = ttk.Combobox(
            purl_row,
            textvariable=self._portal_var,
            values=list(self._portal_options.keys()),
            state="readonly",
            font=("Calibri", 9),
            width=18,
        )
        self._portal_combo.pack(side="left", padx=(0, 6))
        self._portal_combo.bind("<<ComboboxSelected>>", self._on_portal_change)

        # Tampilkan URL yang dipilih (pendek)
        self._portal_url_lbl = tk.Label(
            c1,
            text=f"  → {list(self._portal_options.values())[0]}",
            fg=C["text3"], bg=C["surface"],
            font=("Calibri", 7), anchor="w",
            wraplength=260
        )
        self._portal_url_lbl.pack(anchor="w", padx=12, pady=(0, 1))

        # Label T-code SAP yang aktif — ikut berubah saat portal diganti
        _default_portal = list(self._portal_options.keys())[0]
        self._tcode_info_lbl = tk.Label(
            c1,
            text=self._tcode_info_text(_default_portal),
            fg=C["accent"], bg=C["surface"],
            font=("Calibri", 7, "bold"), anchor="w",
        )
        self._tcode_info_lbl.pack(anchor="w", padx=12, pady=(0, 6))

        # ── Email To — textarea multi-baris ───────────────
        for lbl_txt, attr in [("Email To", "_email_to_txt"),
                               ("Email CC", "_email_cc_txt")]:
            rf = tk.Frame(c1, bg=C["surface"])
            rf.pack(fill="x", padx=12, pady=(3, 0))
            tk.Label(rf, text=lbl_txt, fg=C["text2"], bg=C["surface"],
                     font=("Calibri", 9), width=13, anchor="nw"
                     ).pack(side="left", anchor="n", pady=4)
            txt = tk.Text(rf, height=3, wrap="word",
                          bg=C["surface2"], fg=C["text"],
                          insertbackground=C["accent"],
                          font=("Calibri", 9), relief="flat", bd=4,
                          highlightthickness=1,
                          highlightbackground=C["border"],
                          highlightcolor=C["accent"],
                          selectbackground=C["accent"],
                          selectforeground=C["bg"],
                          width=24)
            txt.pack(side="left", fill="x", expand=True, pady=2)
            setattr(self, attr, txt)

        tk.Label(c1, text="  Satu email per baris atau pisah koma",
                 fg=C["text3"], bg=C["surface"],
                 font=("Calibri", 8)).pack(anchor="w", padx=12, pady=(2, 4))

        # ── Plants Checklist Dropdown ─────────────────────
        prow = tk.Frame(c1, bg=C["surface"])
        prow.pack(fill="x", padx=12, pady=(4, 0))
        tk.Label(prow, text="Plants", fg=C["text2"], bg=C["surface"],
                 font=("Calibri", 9), width=13, anchor="w").pack(side="left")

        self._plant_summary = tk.StringVar()
        self._update_plant_summary()

        plant_btn = tk.Button(
            prow, textvariable=self._plant_summary,
            bg=C["surface2"], fg=C["text"],
            font=("Calibri", 9), relief="flat", bd=4,
            highlightthickness=1,
            highlightbackground=C["border"],
            highlightcolor=C["accent"],
            anchor="w", cursor="hand2", width=22,
            command=lambda: self._toggle_plant_popup(plant_btn)
        )
        plant_btn.pack(side="left", padx=(4, 0))
        tk.Label(prow, text="▾", fg=C["text3"], bg=C["surface"],
                 font=("Calibri", 9)).pack(side="left", padx=(2, 0))

        # Frame popup checklist (hidden by default)
        self._plant_popup = tk.Frame(
            c1, bg=C["surface2"],
            highlightbackground=C["border"], highlightthickness=1
        )
        self._plant_popup_visible = False

        # Header popup: label + Select All + Clear
        ph = tk.Frame(self._plant_popup, bg=C["surface2"])
        ph.pack(fill="x", padx=8, pady=(6, 3))
        self._plant_count_lbl = tk.Label(ph,
                 text=f"{len(self._plant_vars)} plant tersedia",
                 fg=C["text3"], bg=C["surface2"],
                 font=("Calibri", 8))
        self._plant_count_lbl.pack(side="left")
        tk.Button(ph, text="All", font=("Calibri", 8),
                  fg=C["accent"], bg=C["surface2"], relief="flat", bd=0,
                  cursor="hand2",
                  command=self._plant_select_all).pack(side="right", padx=(4,0))
        tk.Button(ph, text="Clear", font=("Calibri", 8),
                  fg=C["text3"], bg=C["surface2"], relief="flat", bd=0,
                  cursor="hand2",
                  command=self._plant_clear_all).pack(side="right")

        tk.Frame(self._plant_popup, bg=C["border"], height=1).pack(fill="x")

        # Grid 2 kolom untuk checkboxes
        self._plant_grid = tk.Frame(self._plant_popup, bg=C["surface2"])
        self._plant_grid.pack(fill="x", padx=8, pady=6)
        for i, plant in enumerate(list(self._plant_vars.keys())):
            cb = tk.Checkbutton(
                self._plant_grid, text=plant,
                variable=self._plant_vars[plant],
                font=("Calibri", 9),
                fg=C["text"], bg=C["surface2"],
                selectcolor=C["surface"],
                activebackground=C["surface2"],
                activeforeground=C["text"],
                relief="flat", bd=0,
                command=self._update_plant_summary,
            )
            cb.grid(row=i//2, column=i%2, sticky="w", padx=4, pady=1)

        tk.Label(c1, text="", bg=C["surface"]).pack(pady=3)

        # ── Plant Mapping File ────────────────────────────
        pf = tk.Frame(c1, bg=C["surface"])
        pf.pack(fill="x", padx=12, pady=(0, 4))
        tk.Label(pf, text="Plant Map", fg=C["text2"], bg=C["surface"],
                 font=("Calibri", 9), width=10, anchor="w").pack(side="left")
        tk.Entry(pf, textvariable=self.plant_map_var,
                 bg=C["surface2"], fg=C["text"],
                 insertbackground=C["accent"],
                 font=("Calibri", 9), relief="flat", bd=4,
                 highlightthickness=1,
                 highlightbackground=C["border"],
                 highlightcolor=C["accent"],
                 width=20).pack(side="left", padx=(4, 6))
        HBtn(pf, C["surface2"], C["border"],
             text="📁", font=("Calibri", 10),
             fg=C["text2"], relief="flat", bd=0,
             padx=5, pady=2, cursor="hand2",
             command=self._browse_plant_map
             ).pack(side="left")

        # Label status: apakah berhasil baca dari Excel atau masih hardcode
        _status_text  = "  ✔ Plant dibaca dari Excel" if self._plant_map_loaded_ok \
                        else "  ⚠ File tidak ditemukan — pilih via 📁"
        _status_color = C["success"] if self._plant_map_loaded_ok else C["warning"]
        self._plant_map_status_lbl = tk.Label(
            c1, text=_status_text,
            fg=_status_color, bg=C["surface"],
            font=("Calibri", 8), anchor="w",
        )
        self._plant_map_status_lbl.pack(anchor="w", pady=(0, 8), padx=12)

        # ── U2C FILE PATH ───────────────────────────────────
        section("U2C FILE PATH")
        c_u2c = card()

        # Load saved path dari config
        try:
            from u2c_upload import get_u2c_filepath
            _default_u2c = get_u2c_filepath()
        except Exception:
            _default_u2c = r"C:\Users\User\Documents\PGD\EOD\U2C.txt"

        self.u2c_path_var = tk.StringVar(value=_default_u2c)

        # Row: label + entry + tombol Browse
        uf = tk.Frame(c_u2c, bg=C["surface"])
        uf.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(uf, text="File Path", fg=C["text2"], bg=C["surface"],
                 font=("Calibri", 9), width=9, anchor="w").pack(side="left")
        tk.Entry(uf, textvariable=self.u2c_path_var,
                 bg=C["surface2"], fg=C["text"],
                 insertbackground=C["accent"],
                 font=("Calibri", 8), relief="flat", bd=4,
                 highlightthickness=1,
                 highlightbackground=C["border"],
                 highlightcolor=C["accent"],
                 width=24).pack(side="left", padx=(4, 6))
        HBtn(uf, C["surface2"], C["border"],
             text="📁", font=("Calibri", 10),
             fg=C["text2"], relief="flat", bd=0,
             padx=6, pady=2, cursor="hand2",
             command=self._browse_u2c_path
             ).pack(side="left")

        # Baris tombol Simpan
        sf = tk.Frame(c_u2c, bg=C["surface"])
        sf.pack(fill="x", padx=12, pady=(0, 10))
        HBtn(sf, C["accent2"], C["accent"],
             text="💾  Simpan Path U2C",
             font=("Calibri", 9, "bold"), fg="white",
             activeforeground="white", relief="flat", bd=0,
             padx=12, pady=5, cursor="hand2",
             command=self._save_u2c_path
             ).pack(side="left")
        self._u2c_save_lbl = tk.Label(sf, text="", fg=C["success"],
                                       bg=C["surface"], font=("Calibri", 8))
        self._u2c_save_lbl.pack(side="left", padx=8)

        # ── RUN MODE ───────────────────────────────────────
        section("RUN MODE")
        c2 = card()
        self.mode_var = tk.StringVar(value="list_plants")

        for val, label, sub, color in [
            ("list_plants", "Proses dari daftar Plants",
             "Hanya plant yang tercantum di atas", C["text"]),
            ("auto_scan",   "Auto-scan Not Completed",
             "Scan semua Not Completed di ListEod", C["warning"]),
        ]:
            rf = tk.Frame(c2, bg=C["surface"])
            rf.pack(fill="x", padx=12, pady=(8, 0))
            tk.Radiobutton(
                rf, text=label, variable=self.mode_var, value=val,
                fg=color, bg=C["surface"], selectcolor=C["surface2"],
                activebackground=C["surface"], activeforeground=color,
                font=("Calibri", 9), cursor="hand2"
            ).pack(anchor="w")
            tk.Label(rf, text=sub, fg=C["text3"], bg=C["surface"],
                     font=("Calibri", 8)).pack(anchor="w", padx=22)

        tk.Frame(c2, bg=C["surface"], height=8).pack()

        # ── INFO ───────────────────────────────────────────
        section("INFO")
        c3 = card()

        for dot_c, msg in [
            (C["accent"],   "SMTP direct — tanpa Thunderbird"),
            (C["success"],  "Chrome dibuka otomatis via CDP"),
            (C["success"],  "Login portal otomatis"),
            (C["warning"],  "Pastikan SAP sudah terbuka"),
        ]:
            rf = tk.Frame(c3, bg=C["surface"])
            rf.pack(fill="x", padx=12, pady=2)
            tk.Label(rf, text="●", fg=dot_c, bg=C["surface"],
                     font=("Calibri", 8)).pack(side="left", padx=(0, 6))
            tk.Label(rf, text=msg, fg=C["text2"], bg=C["surface"],
                     font=("Calibri", 8)).pack(side="left")

        tk.Frame(c3, bg=C["surface"], height=4).pack()

        HBtn(c3, C["surface2"], C["border"],
             text="⚙  Konfigurasi Email SMTP",
             font=("Calibri", 9), fg=C["text2"],
             relief="flat", bd=0, pady=7, cursor="hand2",
             command=self._open_email_config
             ).pack(fill="x", padx=12, pady=(4, 12))

    def _build_right(self, p):
        # ── ACTIVITY LOG (atas, flex) ─────────────────────
        top = tk.Frame(p, bg=C["bg"])
        top.pack(fill="both", expand=True)

        hdr = tk.Frame(top, bg=C["bg"])
        hdr.pack(fill="x", pady=(0, 6))
        tk.Label(hdr, text="ACTIVITY LOG", fg=C["text3"], bg=C["bg"],
                 font=("Calibri", 8, "bold")).pack(side="left")
        HBtn(hdr, C["bg"], C["surface2"],
             text="Clear", font=("Calibri", 8), fg=C["text3"],
             relief="flat", bd=0, padx=8, pady=2, cursor="hand2",
             command=self._clear_log
             ).pack(side="right")

        outer = tk.Frame(top, bg=C["border"])
        outer.pack(fill="both", expand=True)
        inner = tk.Frame(outer, bg=C["surface"])
        inner.pack(fill="both", padx=1, pady=1)

        self.log = scrolledtext.ScrolledText(
            inner, bg=C["surface"], fg=C["text2"],
            font=("Calibri", 9), relief="flat", bd=0,
            state="disabled", wrap="word",
            padx=14, pady=12,
        )
        self.log.pack(fill="both", expand=True)
        self.log.tag_config("INFO",  foreground=C["text2"])
        self.log.tag_config("OK",    foreground=C["success"])
        self.log.tag_config("WARN",  foreground=C["warning"])
        self.log.tag_config("ERROR", foreground=C["danger"])

        # Progress bar
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("RPA.Horizontal.TProgressbar",
                        troughcolor=C["surface2"],
                        background=C["accent"],
                        bordercolor=C["surface2"],
                        lightcolor=C["accent"],
                        darkcolor=C["accent2"],
                        thickness=3)
        self.prog = ttk.Progressbar(top, mode="indeterminate",
                                    style="RPA.Horizontal.TProgressbar")
        self.prog.pack(fill="x", pady=(6, 0))

        # ── SEPARATOR ─────────────────────────────────────
        tk.Frame(p, bg=C["border"], height=1).pack(fill="x", pady=(10, 0))

        # ── LIMIT ALERT (bawah, fixed height) ────────────
        bot = tk.Frame(p, bg=C["bg"])
        bot.pack(fill="x", pady=(8, 0))

        lhdr = tk.Frame(bot, bg=C["bg"])
        lhdr.pack(fill="x", pady=(0, 6))

        # Indikator dot — merah kalau ada alert, abu kalau bersih
        self._alert_dot = tk.Label(lhdr, text="●", fg=C["text3"], bg=C["bg"],
                                    font=("Calibri", 10))
        self._alert_dot.pack(side="left", padx=(0, 5))

        tk.Label(lhdr, text="LIMIT ALERT", fg=C["text3"], bg=C["bg"],
                 font=("Calibri", 8, "bold")).pack(side="left")

        self._alert_count = tk.Label(lhdr, text="", fg=C["warning"], bg=C["bg"],
                                      font=("Calibri", 8))
        self._alert_count.pack(side="left", padx=(8, 0))

        HBtn(lhdr, C["bg"], C["surface2"],
             text="Clear", font=("Calibri", 8), fg=C["text3"],
             relief="flat", bd=0, padx=8, pady=2, cursor="hand2",
             command=self._clear_alert_log
             ).pack(side="right")

        alert_outer = tk.Frame(bot, bg=C["border"])
        alert_outer.pack(fill="x")
        alert_inner = tk.Frame(alert_outer, bg=C["surface"])
        alert_inner.pack(fill="both", padx=1, pady=1)

        self.alert_log = scrolledtext.ScrolledText(
            alert_inner, bg=C["surface"], fg=C["text2"],
            font=("Calibri", 9), relief="flat", bd=0,
            state="disabled", wrap="word",
            padx=14, pady=10, height=7,
        )
        self.alert_log.pack(fill="x")
        self.alert_log.tag_config("HEAD",  foreground=C["warning"],   font=("Calibri", 9, "bold"))
        self.alert_log.tag_config("ITEM",  foreground=C["text2"])
        self.alert_log.tag_config("OVER",  foreground=C["danger"])
        self.alert_log.tag_config("PLANT", foreground=C["accent"],    font=("Calibri", 8, "bold"))

        # Counter internal
        self._alert_total = 0

    def _build_bottom(self):
        bar = tk.Frame(self.root, bg=C["surface"], height=60)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        self._lastrun = tk.Label(bar, text="Belum pernah dijalankan",
                                 fg=C["text3"], bg=C["surface"],
                                 font=("Calibri", 8))
        self._lastrun.pack(side="left", padx=20)

        btn_area = tk.Frame(bar, bg=C["surface"])
        btn_area.pack(side="right", padx=16, pady=10)

        self.stop_btn = HBtn(
            btn_area, C["surface2"], C["stop_hov"],
            text="■  Stop",
            font=("Calibri", 10, "bold"), fg=C["text2"],
            activeforeground="white", relief="flat", bd=0,
            padx=18, pady=8, cursor="hand2", state="disabled",
            command=self._on_stop
        )
        self.stop_btn.pack(side="right", padx=(8, 0))

        self.run_btn = HBtn(
            btn_area, C["run_bg"], C["run_hov"],
            text="▶   Run RPA",
            font=("Calibri", 11, "bold"), fg="white",
            activeforeground="white", relief="flat", bd=0,
            padx=28, pady=8, cursor="hand2",
            command=self._on_run
        )
        self.run_btn.pack(side="right")

    # ── HELPERS ──────────────────────────────────────────────

    def _load_email_defaults(self):
        try:
            from email_config_ui import load_credentials
            cred = load_credentials()

            # Isi email_to tags
            self._email_to_tags.clear()
            for e in cred.get("email_to", "").split(","):
                e = e.strip()
                if e:
                    self._email_to_tags.append(e)
            if hasattr(self, "_refresh__email_to_tags"):
                self._refresh__email_to_tags()

            # Isi email_cc tags
            self._email_cc_tags.clear()
            for e in cred.get("email_cc", "").split(","):
                e = e.strip()
                if e:
                    self._email_cc_tags.append(e)
            if hasattr(self, "_refresh__email_cc_tags"):
                self._refresh__email_cc_tags()
        except Exception:
            pass

    def _open_email_config(self):
        try:
            import subprocess, sys
            subprocess.Popen([sys.executable, "email_config_ui.py"])
            self.root.after(3000, self._load_email_defaults)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal buka konfigurasi:\n{e}")

    def _set_running(self, on: bool):
        self._running = on
        if on:
            self.run_btn.config(state="disabled", text="⏳  Running...")
            self.run_btn.recolor(C["surface2"], C["surface2"])
            self.stop_btn.config(state="normal")
            self.stop_btn.recolor(C["stop_bg"], C["stop_hov"])
            self._dot.config(fg=C["warning"])
            self._stat.config(text="Running", fg=C["warning"])
            self.prog.start(12)
        else:
            self.run_btn.config(state="normal", text="▶   Run RPA")
            self.run_btn.recolor(C["run_bg"], C["run_hov"])
            self.stop_btn.config(state="disabled")
            self.stop_btn.recolor(C["surface2"], C["stop_hov"])
            self._dot.config(fg=C["success"])
            self._stat.config(text="Ready", fg=C["text2"])
            self.prog.stop()
            self._lastrun.config(
                text=f"Last run: {datetime.now().strftime('%d %b %Y  %H:%M')}")

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _clear_alert_log(self):
        self.alert_log.config(state="normal")
        self.alert_log.delete("1.0", "end")
        self.alert_log.config(state="disabled")
        self._alert_total = 0
        self._alert_dot.config(fg=C["text3"])
        self._alert_count.config(text="")

    def _append_alert(self, plant: str, items_skip: list, limits: dict):
        """
        Tulis item yang lewat limit ke panel Limit Alert.
        Dipanggil dari _poll saat level == LIMIT_ALERT.
        """
        if not items_skip:
            return

        self._alert_total += len(items_skip)
        self._alert_dot.config(fg=C["danger"])
        self._alert_count.config(
            text=f"{self._alert_total} item lewat limit"
        )

        self.alert_log.config(state="normal")

        # Header plant
        ts = datetime.now().strftime("%H:%M:%S")
        self.alert_log.insert("end",
            f"[{ts}]  Plant {plant}  \u2014  {len(items_skip)} item lewat batas\n",
            "HEAD")

        for item in items_skip:
            lim   = limits.get(item["material"], {})
            lim_p = lim.get("limit_plus",  "N/A")
            lim_m = lim.get("limit_minus", "N/A")
            diff  = item["diff"]
            mat   = item["material"]
            sloc  = item["sloc"]

            if diff > 0:
                batas  = f"limit+ {lim_p}"
                arah   = f"{diff:+.3f} > {lim_p}"
            else:
                batas  = f"limit- {lim_m}"
                arah   = f"{diff:+.3f} < {lim_m}"

            self.alert_log.insert("end",
                f"  \u26a0  {mat}  SLoc={sloc}  diff={diff:+.6f}  [{batas}]  {arah}\n",
                "OVER")

        self.alert_log.insert("end", "\n", "ITEM")
        self.alert_log.see("end")
        self.alert_log.config(state="disabled")

    def _write_log(self, msg: str, level: str = "INFO"):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n", level)
        self.log.see("end")
        self.log.config(state="disabled")

    # ── HANDLERS ─────────────────────────────────────────────

    def _toggle_plant_popup(self, anchor_btn=None):
        """Tampilkan/sembunyikan checklist plant."""
        if self._plant_popup_visible:
            self._plant_popup.pack_forget()
            self._plant_popup_visible = False
        else:
            # Tampilkan di bawah tombol plants (setelah label Plant Map)
            self._plant_popup.pack(fill="x", padx=12, pady=(0, 4))
            self._plant_popup_visible = True

    def _update_plant_summary(self):
        """Update teks ringkasan di tombol Plants."""
        selected = [p for p, v in self._plant_vars.items() if v.get()]
        total    = len(self._plant_vars)
        if len(selected) == 0:
            self._plant_summary.set("— tidak ada —")
        elif len(selected) == total:
            self._plant_summary.set(f"Semua ({total} plant)")
        elif len(selected) <= 3:
            self._plant_summary.set(", ".join(selected))
        else:
            self._plant_summary.set(f"{', '.join(selected[:3])} +{len(selected)-3} lagi")

    def _plant_select_all(self):
        for v in self._plant_vars.values():
            v.set(True)
        self._update_plant_summary()

    def _plant_clear_all(self):
        for v in self._plant_vars.values():
            v.set(False)
        self._update_plant_summary()

    def get_selected_plants(self) -> list:
        """Ambil list plant yang dicentang."""
        return [p for p, v in self._plant_vars.items() if v.get()]

    def _show_tooltip(self, widget, text: str):
        """Tampilkan tooltip dengan email lengkap saat hover tag."""
        self._hide_tooltip()
        x = widget.winfo_rootx() + 10
        y = widget.winfo_rooty() + widget.winfo_height() + 4
        self._tooltip = tk.Toplevel(self.root)
        self._tooltip.wm_overrideredirect(True)
        self._tooltip.wm_geometry(f"+{x}+{y}")
        tk.Label(self._tooltip, text=text,
                 bg="#1E293B", fg="#E2E8F0",
                 font=("Calibri", 9), padx=6, pady=3,
                 relief="flat").pack()

    def _hide_tooltip(self):
        if hasattr(self, "_tooltip") and self._tooltip:
            try:
                self._tooltip.destroy()
            except Exception:
                pass
            self._tooltip = None

    def _build_tag_field(self, parent, label: str,
                         tag_list: list, attr_name: str) -> tk.Frame:
        """
        Buat widget tag input (chip) untuk satu field email.
        Ketik email → Enter atau koma → jadi tag biru.
        Klik × di tag → hapus.
        """
        outer = tk.Frame(parent, bg=C["surface"])
        outer.pack(fill="x", padx=12, pady=(3, 0))

        tk.Label(outer, text=label, fg=C["text2"], bg=C["surface"],
                 font=("Calibri", 9), width=13, anchor="w").pack(side="left",
                                                                    anchor="n",
                                                                    pady=4)

        # Box tempat tag + input — scrollable, max height 56px (~3 baris tag)
        box = tk.Frame(outer, bg=C["surface2"],
                       highlightbackground=C["border"],
                       highlightthickness=1,
                       relief="flat")
        box.pack(side="left", fill="x", expand=True, pady=2)

        # Canvas + scrollbar untuk wrap tags
        canvas = tk.Canvas(box, bg=C["surface2"], highlightthickness=0,
                           height=28)
        vsb    = tk.Scrollbar(box, orient="vertical", command=canvas.yview,
                              width=8)
        canvas.configure(yscrollcommand=vsb.set)

        inner = tk.Frame(canvas, bg=C["surface2"])
        inner_win = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_configure(event, c=canvas, iw=inner_win):
            c.configure(scrollregion=c.bbox("all"))
            # Auto-expand height max 56px (3 baris), min 28px
            h = min(max(inner.winfo_reqheight() + 4, 28), 56)
            c.configure(height=h)
            if inner.winfo_reqheight() > 56:
                vsb.pack(side="right", fill="y")
            else:
                vsb.pack_forget()

        def _on_canvas_configure(event, c=canvas, iw=inner_win):
            c.itemconfig(iw, width=event.width)

        inner.bind("<Configure>", _on_inner_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.pack(side="left", fill="x", expand=True, padx=4, pady=3)

        # Entry input
        entry_var = tk.StringVar()
        entry = tk.Entry(inner, textvariable=entry_var,
                         bg=C["surface2"], fg=C["text"],
                         insertbackground=C["accent"],
                         font=("Calibri", 9), relief="flat", bd=0,
                         width=16)
        entry.pack(side="left", padx=2, pady=1)

        def commit(event=None):
            raw = entry_var.get()
            # Split by koma — bisa paste beberapa email sekaligus
            emails = [e.strip().rstrip(",") for e in raw.replace(",", "\n").splitlines()]
            for email in emails:
                if email and email not in tag_list:
                    tag_list.append(email)
            entry_var.set("")
            _refresh()

        def _refresh():
            # Hapus semua tag lama
            for w in inner.winfo_children():
                if getattr(w, "_is_tag", False):
                    w.destroy()
            # Rebuild tags — wrap otomatis karena inner expand horizontal
            for email in list(tag_list):
                tag_frame = tk.Frame(inner, bg=C["accent2"])
                tag_frame._is_tag = True
                tag_frame.pack(side="left", padx=2, pady=2, before=entry)

                # Truncate email panjang agar tag tidak terlalu lebar
                display = email if len(email) <= 22 else email[:19] + "…"
                lbl = tk.Label(tag_frame, text=display,
                               fg="white", bg=C["accent2"],
                               font=("Calibri", 8),
                               padx=4, pady=2)
                lbl.pack(side="left")
                lbl.bind("<Enter>", lambda e, em=email: self._show_tooltip(e.widget, em))
                lbl.bind("<Leave>", lambda e: self._hide_tooltip())

                _email = email
                tk.Button(tag_frame, text="×",
                           fg="white", bg=C["accent2"],
                           font=("Calibri", 8),
                           relief="flat", bd=0,
                           padx=3, pady=1,
                           cursor="hand2",
                           command=lambda e=_email: _remove(e)
                           ).pack(side="left")

        def _remove(email):
            if email in tag_list:
                tag_list.remove(email)
            _refresh()

        entry.bind("<Return>", commit)
        entry.bind("<comma>", commit)
        entry.bind("<Tab>", commit)
        entry.bind("<FocusOut>", lambda e: commit() if entry_var.get().strip() else None)

        # Simpan referensi refresh ke instance agar bisa dipanggil dari luar
        setattr(self, f"_refresh_{attr_name}", _refresh)

        return outer

    def _get_email_to(self) -> str:
        """Return email To sebagai string pisah koma."""
        raw = self._email_to_txt.get("1.0", "end").strip()
        return ", ".join(e.strip() for e in raw.replace("\n", ",").split(",") if e.strip())

    def _get_email_cc(self) -> str:
        """Return email CC sebagai string pisah koma."""
        raw = self._email_cc_txt.get("1.0", "end").strip()
        return ", ".join(e.strip() for e in raw.replace("\n", ",").split(",") if e.strip())

    def _browse_plant_map(self):
        """Buka dialog pilih file Excel plant mapping."""
        from tkinter import filedialog
        current  = self.plant_map_var.get()
        init_dir = os.path.dirname(current) if current and os.path.exists(os.path.dirname(current)) else os.path.expanduser("~")
        path = filedialog.askopenfilename(
            title     = "Pilih file Excel Plant Mapping",
            initialdir= init_dir,
            filetypes = [("Excel File", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if path:
            self.plant_map_var.set(path)
            # Update Config supaya langsung aktif di session ini
            from config import Config
            Config.PLANT_MAPPING_FILE = path
            # Rebuild checklist plant dari file baru
            plants = self._load_plants_from_excel(path)
            self._rebuild_plant_checklist(self._plant_grid, plants)

            # Update label jumlah plant di header popup
            if hasattr(self, '_plant_count_lbl'):
                self._plant_count_lbl.config(
                    text=f"{len(self._plant_vars)} plant tersedia"
                )

            # Update label status — berhasil atau gagal baca Excel
            n = len(plants)
            is_fallback = (sorted(plants) == sorted(list(Config.PLANTS))
                           and n == len(Config.PLANTS))
            if not is_fallback:
                status_text  = f"  ✔ {n} plant dibaca dari Excel"
                status_color = C["success"]
            else:
                status_text  = "  ⚠ Gagal baca Excel — cek path & sheet 'Plant_CostCenter'"
                status_color = C["warning"]

            if hasattr(self, "_plant_map_status_lbl"):
                self._plant_map_status_lbl.config(
                    text=status_text, fg=status_color
                )

            # Log ke Activity Log
            self._write_log(
                f"Plant Map: {path}  →  {n} plant dimuat" if not is_fallback
                else f"⚠ Plant Map gagal dibaca — menggunakan fallback hardcode",
                "OK" if not is_fallback else "WARN"
            )

    def _browse_u2c_path(self):
        """Buka dialog pilih lokasi + nama file U2C."""
        from tkinter import filedialog
        current = self.u2c_path_var.get()
        init_dir  = os.path.dirname(current) if current else r"C:\Users\User\Documents\PGD\EOD"
        init_file = os.path.basename(current) if current else "U2C.txt"
        path = filedialog.asksaveasfilename(
            title        = "Pilih lokasi & nama file U2C",
            initialdir   = init_dir,
            initialfile  = init_file,
            defaultextension = ".txt",
            filetypes    = [("Text File", "*.txt"), ("All Files", "*.*")],
        )
        if path:
            self.u2c_path_var.set(path)
            self._u2c_save_lbl.config(text="")

    def _save_u2c_path(self):
        """Simpan path file U2C ke config."""
        path = self.u2c_path_var.get().strip()
        if not path:
            messagebox.showwarning("Peringatan", "Path file U2C tidak boleh kosong!")
            return
        try:
            from u2c_upload import save_u2c_config
            save_u2c_config(path)
            self._u2c_save_lbl.config(text="✔ Tersimpan", fg=C["success"])
            self.root.after(3000, lambda: self._u2c_save_lbl.config(text=""))
        except Exception as e:
            messagebox.showerror("Error", f"Gagal simpan config U2C:\n{e}")

    def _on_run(self):
        # Pastikan URL portal sudah ter-set sebelum run
        self._on_portal_change()

        if self._running:
            return

        plants = self.get_selected_plants()
        posting_date = self.date_var.get().strip()
        email_to     = self._get_email_to().strip()
        email_cc     = self._get_email_cc().strip()
        mode         = self.mode_var.get()

        if mode == "list_plants" and not plants:
            messagebox.showwarning("Peringatan", "Masukkan minimal 1 plant!")
            return
        if not self._get_email_to():
            messagebox.showwarning("Peringatan", "Email To tidak boleh kosong!")
            return

        try:
            from email_config_ui import load_credentials
            load_credentials()
        except FileNotFoundError:
            if messagebox.askyesno("Setup Email",
                                   "Kredensial email belum diset.\n"
                                   "Buka konfigurasi email sekarang?"):
                self._open_email_config()
            return
        except Exception as e:
            messagebox.showerror("Error", f"Gagal baca kredensial:\n{e}")
            return

        stop_event.clear()
        self._set_running(True)
        self._write_log(f"Run RPA  ·  {posting_date}  ·  mode={mode}", "OK")
        self._write_log(f"Email → {email_to}" + (f"  CC: {email_cc}" if email_cc else ""), "INFO")

        # Log T-code SAP yang akan dipakai
        from config import Config
        self._write_log(
            f"SAP T-code: /{Config.ACTIVE_TCODE_SAPSTK}  +  /{Config.ACTIVE_TCODE_U2C}  "
            f"(portal: {self._portal_var.get()})",
            "INFO"
        )

        threading.Thread(
            target=run_robot,
            args=(plants, posting_date, email_to, email_cc, mode),
            daemon=True
        ).start()

    def _on_stop(self):
        stop_event.set()
        self._write_log("=" * 40, "WARN")
        self._write_log("Stop diklik -- menghentikan robot...", "WARN")
        self._write_log("Tunggu proses saat ini selesai.", "WARN")
        self._write_log("=" * 40, "WARN")
        self.stop_btn.config(state="disabled")

    def _poll(self):
        try:
            while True:
                level, msg = log_queue.get_nowait()
                if level == "DONE":
                    self._set_running(False)
                    self._write_log("━" * 50, "INFO")
                    self._write_log("RPA selesai  —  cek inbox email.", "OK")
                elif level == "WAIT_LOGIN":
                    self.root.after(0, self._ask_login_confirm)
                elif level == "SAP_WARNING":
                    self.root.after(0, self._show_sap_warning)
                elif level == "SAP_WAIT":
                    self.root.after(0, self._ask_sap_confirm)
                else:
                    self._write_log(msg, level)
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _ask_sap_confirm(self):
        """Dialog popup minta user konfirmasi SAP sudah terbuka dan login."""
        self._write_log("━" * 50, "WARN")
        self._write_log("⚠  SAP belum terdeteksi!", "WARN")
        self._write_log("   Buka SAP GUI dan login terlebih dahulu,", "WARN")
        self._write_log("   lalu klik OK untuk mulai download SAPSTK.", "WARN")
        self._write_log("━" * 50, "WARN")

        popup = tk.Toplevel(self.root)
        popup.title("Buka SAP")
        popup.geometry("420x200")
        popup.configure(bg=C["surface"])
        popup.resizable(False, False)
        popup.grab_set()
        popup.lift()
        popup.focus_force()

        tk.Label(popup,
                 text="🖥  Buka SAP GUI",
                 font=("Calibri", 12, "bold"),
                 fg=C["text"], bg=C["surface"]
                 ).pack(pady=(20, 8))

        tk.Label(popup,
                 text="Silakan buka SAP GUI dan login.\n\nKlik OK setelah SAP terbuka\ndan siap digunakan.",
                 font=("Calibri", 10),
                 fg=C["text2"], bg=C["surface"],
                 justify="center"
                 ).pack(pady=(0, 16))

        def on_ok():
            popup.destroy()
            sap_event.set()

        HBtn(popup, C["run_bg"], C["run_hov"],
             text="✓  SAP Sudah Terbuka — Lanjutkan",
             font=("Calibri", 10, "bold"), fg="white",
             activeforeground="white", relief="flat", bd=0,
             padx=20, pady=8, cursor="hand2",
             command=on_ok
             ).pack(pady=(0, 16))

    def _show_sap_warning(self):
        """Tampilkan banner di log bahwa SAP sedang dikontrol robot."""
        self._write_log("━" * 50, "WARN")
        self._write_log("⚠  SAP sedang dikontrol robot!", "WARN")
        self._write_log("   Jangan sentuh keyboard / mouse", "WARN")
        self._write_log("   sampai log menampilkan 'Download SAP selesai'", "WARN")
        self._write_log("━" * 50, "WARN")
        # Update status pill
        self._dot.config(fg=C["warning"])
        self._stat.config(text="SAP Running", fg=C["warning"])

    def _ask_login_confirm(self):
        """Dialog popup minta user konfirmasi sudah login ke portal."""
        self._write_log("━" * 50, "WARN")
        self._write_log("⚠  Chrome terbuka — silakan LOGIN di browser!", "WARN")
        self._write_log("   Masukkan username & password di Chrome,", "WARN")
        self._write_log("   lalu klik OK di sini untuk melanjutkan RPA.", "WARN")
        self._write_log("━" * 50, "WARN")

        # Popup di atas semua window
        popup = tk.Toplevel(self.root)
        popup.title("Konfirmasi Login")
        popup.geometry("420x200")
        popup.configure(bg=C["surface"])
        popup.resizable(False, False)
        popup.grab_set()  # modal
        popup.lift()
        popup.focus_force()

        tk.Label(popup,
                 text="🔐  Login Portal Mayora",
                 font=("Calibri", 12, "bold"),
                 fg=C["text"], bg=C["surface"]
                 ).pack(pady=(20, 8))

        tk.Label(popup,
                 text="Silakan masukkan username & password\ndi Chrome yang sudah terbuka.\n\nKlik OK setelah berhasil login.",
                 font=("Calibri", 10),
                 fg=C["text2"], bg=C["surface"],
                 justify="center"
                 ).pack(pady=(0, 16))

        def on_ok():
            popup.destroy()
            login_event.set()  # kirim sinyal ke background thread

        HBtn(popup, C["run_bg"], C["run_hov"],
             text="✓  Sudah Login — Lanjutkan RPA",
             font=("Calibri", 10, "bold"), fg="white",
             activeforeground="white", relief="flat", bd=0,
             padx=20, pady=8, cursor="hand2",
             command=on_ok
             ).pack(pady=(0, 16))


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.iconbitmap("icon.ico")
    except Exception:
        pass
    RpaGui(root)
    root.mainloop()