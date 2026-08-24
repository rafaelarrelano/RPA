"""
email_config_ui.py
UI untuk simpan & ubah kredensial email — "Email Configuration"
(desain baru, tema terang, sesuai mockup)

CATATAN: Fungsi penyimpanan/pembacaan kredensial (save_credentials,
load_credentials) dan pembuatan draft Thunderbird TIDAK diubah — hanya
tampilannya yang diperbarui.
"""

import os
import json

# CRITICAL: DPI awareness harus diaktifkan SEBELUM tkinter di-import,
# supaya teks tidak blur/pecah di layar dengan scaling 125%/150%/200%.
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        from ctypes import windll
        windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import messagebox
from cryptography.fernet import Fernet

# ─────────────────────────────────────────────
# PATH FILE KREDENSIAL
# Dashboard dan settings harus pakai path yang sama agar tidak ada file
# kredensial yang terpisah antara folder project dan folder legacy C:\RPA_StockRecon.
# ─────────────────────────────────────────────
_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
LEGACY_BASE_DIR = r"C:\RPA_StockRecon"
BASE_DIR = _PROJECT_ROOT

CRED_FILE = os.path.join(BASE_DIR, "config", "email_cred.enc")
KEY_FILE  = os.path.join(BASE_DIR, "config", "email_key.key")
SYNC_FILE = os.path.join(BASE_DIR, "config", "email_sync.json")
LEGACY_CRED_FILE = os.path.join(LEGACY_BASE_DIR, "config", "email_cred.enc")
LEGACY_KEY_FILE  = os.path.join(LEGACY_BASE_DIR, "config", "email_key.key")
LEGACY_SYNC_FILE = os.path.join(LEGACY_BASE_DIR, "config", "email_sync.json")

# ── Tema terang (sesuai desain baru) ─────────
BG        = "#F5F6F8"
CARD      = "#FFFFFF"
BORDER    = "#EAECEF"
INPUT_BG  = "#F7F8FA"
INPUT_BD  = "#E3E5E9"
TEXT      = "#16181D"
TEXT_MUT  = "#6B7280"
TEXT_HNT  = "#9AA1AC"
ACCENT    = "#E5342E"
ACCENT_BG = "#FDEAEA"
BLUE      = "#2563EB"
BLUE_HOV  = "#1D4ED8"
GREEN     = "#10B981"
GREEN_HOV = "#059669"
AMBER     = "#F59E0B"
AMBER_HOV = "#D97706"
SUCCESS   = "#16A34A"
DANGER    = "#DC2626"
FONT      = "Segoe UI"


# ─────────────────────────────────────────────
# ENKRIPSI / DEKRIPSI (tidak berubah)
# ─────────────────────────────────────────────

def _get_or_create_key() -> bytes:
    os.makedirs(os.path.dirname(KEY_FILE), exist_ok=True)
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, "rb") as f:
            return f.read()
    key = Fernet.generate_key()
    with open(KEY_FILE, "wb") as f:
        f.write(key)
    return key


def save_credentials(smtp_host: str, smtp_port: int,
                     email_from: str, password: str,
                     email_to: str, email_cc: str):
    payload = {
        "smtp_host":  smtp_host,
        "smtp_port":  smtp_port,
        "email_from": email_from,
        "password":   password,
        "email_to":   email_to,
        "email_cc":   email_cc,
    }

    def _write(path: str, key_path: str):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        key = _get_or_create_key_for(key_path)
        fern = Fernet(key)
        data = json.dumps(payload).encode()
        with open(path, "wb") as f:
            f.write(fern.encrypt(data))

    def _write_sync(path: str):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

    _write(CRED_FILE, KEY_FILE)
    _write_sync(SYNC_FILE)
    if os.path.isdir(os.path.dirname(LEGACY_CRED_FILE)) or os.path.exists(LEGACY_CRED_FILE):
        _write(LEGACY_CRED_FILE, LEGACY_KEY_FILE)
        _write_sync(LEGACY_SYNC_FILE)


def _get_or_create_key_for(key_file: str) -> bytes:
    os.makedirs(os.path.dirname(key_file), exist_ok=True)
    if os.path.exists(key_file):
        with open(key_file, "rb") as f:
            return f.read()
    key = Fernet.generate_key()
    with open(key_file, "wb") as f:
        f.write(key)
    return key


def _read_credentials_from(path: str, key_path: str) -> dict:
    if not os.path.exists(path):
        raise FileNotFoundError
    key = _get_or_create_key_for(key_path)
    fern = Fernet(key)
    with open(path, "rb") as f:
        token = f.read()
    return json.loads(fern.decrypt(token).decode())


def load_credentials() -> dict:
    local_exists = os.path.exists(CRED_FILE)
    legacy_exists = os.path.exists(LEGACY_CRED_FILE)

    if local_exists:
        return _read_credentials_from(CRED_FILE, KEY_FILE)
    if legacy_exists:
        return _read_credentials_from(LEGACY_CRED_FILE, LEGACY_KEY_FILE)

    raise FileNotFoundError(
        "Kredensial email belum dikonfigurasi.\n"
        "Jalankan python email_config_ui.py untuk setup."
    )


# ─────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────

class EmailConfigUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Configuration")
        self.root.geometry("680x760")
        self.root.minsize(560, 620)
        self.root.resizable(True, True)
        self.root.configure(bg=BG)
        self._build_ui()
        self._load_existing()

    def _section(self, parent, icon, text):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", pady=(16, 6))
        tk.Label(f, text=icon, font=(FONT, 11), fg=ACCENT, bg=BG).pack(side="left", padx=(0, 6))
        tk.Label(f, text=text, font=(FONT, 11, "bold"), fg=TEXT, bg=BG).pack(side="left")
        return f

    def _card(self, parent):
        outer = tk.Frame(parent, bg=BORDER)
        outer.pack(fill="x")
        inner = tk.Frame(outer, bg=CARD, padx=18, pady=14)
        inner.pack(fill="x", padx=1, pady=1)
        return inner

    def _wrapping_hint(self, parent, text):
        """
        Label kecil (mis. "(pisah koma atau Enter)") yang otomatis
        menyesuaikan lebar (wrap) mengikuti lebar card saat window
        di-resize, supaya tidak pernah terpotong oleh tepi jendela.
        """
        lbl = tk.Label(parent, text=text, fg=ACCENT, bg=CARD, font=(FONT, 8),
                       justify="right", anchor="e")
        lbl.pack(fill="x", anchor="e")

        def _update_wrap(event, _lbl=lbl):
            # padx kiri-kanan card = 18+18, sisakan sedikit margin
            _lbl.config(wraplength=max(120, event.width - 4))

        parent.bind("<Configure>", _update_wrap, add="+")
        return lbl

    def _build_ui(self):
        # ── Header ────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=CARD, pady=16)
        hdr.pack(fill="x")
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

        title_row = tk.Frame(hdr, bg=CARD)
        title_row.pack(fill="x", padx=24)
        tk.Label(title_row, text="✉", font=(FONT, 16), fg="#FFFFFF", bg=ACCENT,
                 width=2, height=1).pack(side="left", padx=(0, 12))
        tcol = tk.Frame(title_row, bg=CARD)
        tcol.pack(side="left")
        tk.Label(tcol, text="Email Configuration", font=(FONT, 14, "bold"),
                 fg=TEXT, bg=CARD).pack(anchor="w")
        tk.Label(tcol, text="Konfigurasi Email — Thunderbird Draft (RPA Stock Recon)",
                 font=(FONT, 9), fg=TEXT_HNT, bg=CARD).pack(anchor="w")

        # ── Scrollable body ────────────────────────────────
        body_outer = tk.Frame(self.root, bg=BG)
        body_outer.pack(fill="both", expand=True)

        body_canvas = tk.Canvas(body_outer, bg=BG, highlightthickness=0)
        body_vsb = tk.Scrollbar(body_outer, orient="vertical", command=body_canvas.yview)
        body_canvas.configure(yscrollcommand=body_vsb.set)
        body_vsb.pack(side="right", fill="y")
        body_canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(body_canvas, bg=BG, padx=24)
        body_win = body_canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_body_configure(event):
            body_canvas.configure(scrollregion=body_canvas.bbox("all"))

        def _on_canvas_configure(event):
            body_canvas.itemconfig(body_win, width=event.width)

        body.bind("<Configure>", _on_body_configure)
        body_canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            body_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        body_canvas.bind("<Enter>", lambda e: body_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        body_canvas.bind("<Leave>", lambda e: body_canvas.unbind_all("<MouseWheel>"))

        # SMTP section
        self._section(body, "🗄", "SMTP")
        smtp_card = self._card(body)

        self.vars = {}

        def field_row(parent, label, key, show="", row_pad=(0, 10)):
            f = tk.Frame(parent, bg=CARD)
            f.pack(fill="x", pady=row_pad)
            tk.Label(f, text=label, fg=TEXT_MUT, bg=CARD, font=(FONT, 9)).pack(anchor="w")
            var = tk.StringVar()
            self.vars[key] = var
            wrap = tk.Frame(f, bg=INPUT_BD)
            wrap.pack(fill="x", pady=(4, 0))
            inner = tk.Frame(wrap, bg=INPUT_BG)
            inner.pack(fill="x", padx=1, pady=1)
            e = tk.Entry(inner, textvariable=var, show=show, bg=INPUT_BG, fg=TEXT,
                         insertbackground=ACCENT, font=(FONT, 10), relief="flat", bd=0)
            e.pack(fill="x", ipady=7, padx=10)
            return e, wrap

        field_row(smtp_card, "SMTP Host", "smtp_host", row_pad=(0, 10))
        field_row(smtp_card, "SMTP Port", "smtp_port", row_pad=(0, 10))
        field_row(smtp_card, "Email Pengirim", "email_from", row_pad=(0, 10))

        # Password (with show/hide)
        pf = tk.Frame(smtp_card, bg=CARD)
        pf.pack(fill="x")
        tk.Label(pf, text="Password", fg=TEXT_MUT, bg=CARD, font=(FONT, 9)).pack(anchor="w")
        self.vars["password"] = tk.StringVar()
        pw_wrap = tk.Frame(pf, bg=INPUT_BD)
        pw_wrap.pack(fill="x", pady=(4, 0))
        pw_inner = tk.Frame(pw_wrap, bg=INPUT_BG)
        pw_inner.pack(fill="x", padx=1, pady=1)
        self._pw_entry = tk.Entry(pw_inner, textvariable=self.vars["password"], show="*",
                                   bg=INPUT_BG, fg=TEXT, insertbackground=ACCENT,
                                   font=(FONT, 10), relief="flat", bd=0)
        self._pw_entry.pack(side="left", fill="x", expand=True, ipady=7, padx=(10, 0))
        self._pw_shown = False
        tk.Button(pw_inner, text="👁", font=(FONT, 9), bg=INPUT_BG, fg=TEXT_MUT,
                  relief="flat", bd=0, cursor="hand2", command=self._toggle_pw,
                  padx=8).pack(side="right")

        tk.Label(smtp_card, text="> Opsional — disimpan untuk referensi (tidak lagi digunakan untuk SMTP)",
                 fg=TEXT_HNT, bg=CARD, font=(FONT, 8)).pack(anchor="w", pady=(6, 0))

        # Email To
        self._section(body, "✉", "Email To")
        to_card = self._card(body)
        self._wrapping_hint(to_card, "(pisah koma atau Enter)")
        self._email_to_txt = tk.Text(to_card, height=3, wrap="word", bg=INPUT_BG, fg=TEXT,
                                      insertbackground=ACCENT, font=(FONT, 10), relief="flat",
                                      bd=0, highlightthickness=1, highlightbackground=INPUT_BD,
                                      highlightcolor=ACCENT)
        self._email_to_txt.pack(fill="x", pady=(4, 0), ipady=4)

        # Email CC
        self._section(body, "✉", "Email CC")
        cc_card = self._card(body)
        self._wrapping_hint(cc_card, "(pisah koma atau Enter, opsional)")
        self._email_cc_txt = tk.Text(cc_card, height=3, wrap="word", bg=INPUT_BG, fg=TEXT,
                                      insertbackground=ACCENT, font=(FONT, 10), relief="flat",
                                      bd=0, highlightthickness=1, highlightbackground=INPUT_BD,
                                      highlightcolor=ACCENT)
        self._email_cc_txt.pack(fill="x", pady=(4, 0), ipady=4)

        tk.Frame(body, bg=BG, height=10).pack()

        # ── Buttons ───────────────────────────────────────
        btn_frame = tk.Frame(self.root, bg=BG, pady=14, padx=24)
        btn_frame.pack(fill="x", side="bottom")

        tk.Button(btn_frame, text="💾  Simpan", font=(FONT, 10, "bold"), fg="#FFFFFF",
                  bg=BLUE, activebackground=BLUE_HOV, activeforeground="#FFFFFF",
                  relief="flat", padx=16, pady=8, cursor="hand2",
                  command=self._on_save).pack(side="left", padx=(0, 8))

        tk.Button(btn_frame, text="✉  Test Draft", font=(FONT, 10, "bold"), fg="#FFFFFF",
                  bg=GREEN, activebackground=GREEN_HOV, activeforeground="#FFFFFF",
                  relief="flat", padx=14, pady=8, cursor="hand2",
                  command=self._on_test).pack(side="left", padx=(0, 8))

        tk.Button(btn_frame, text="ⓘ  Info", font=(FONT, 10, "bold"), fg="#FFFFFF",
                  bg=AMBER, activebackground=AMBER_HOV, activeforeground="#FFFFFF",
                  relief="flat", padx=14, pady=8, cursor="hand2",
                  command=self._on_diagnose).pack(side="left")

        tk.Button(btn_frame, text="✕  Tutup", font=(FONT, 10), fg=TEXT_MUT, bg=CARD,
                  activebackground="#F3F4F6", relief="solid", bd=1,
                  padx=12, pady=8, cursor="hand2",
                  command=self.root.destroy).pack(side="right")

        # Status label
        self.status_lbl = tk.Label(self.root, text="", font=(FONT, 9),
                                    fg=SUCCESS, bg=BG, wraplength=580)
        self.status_lbl.pack(pady=(0, 8), side="bottom")

    # ── HELPER: baca/tulis Text widget ───────────────────────

    def _get_emails(self, widget: tk.Text) -> str:
        raw = widget.get("1.0", "end").strip()
        parts = []
        for part in raw.replace("\n", ",").split(","):
            part = part.strip()
            if part:
                parts.append(part)
        return ", ".join(parts)

    def _set_emails(self, widget: tk.Text, value: str):
        widget.delete("1.0", "end")
        if not value:
            return
        parts = [e.strip() for e in value.split(",") if e.strip()]
        widget.insert("1.0", "\n".join(parts))

    # ── HANDLERS ─────────────────────────────────────────────

    def _toggle_pw(self):
        self._pw_shown = not self._pw_shown
        self._pw_entry.config(show="" if self._pw_shown else "*")

    def _load_existing(self):
        try:
            cred = load_credentials()
            self.vars["smtp_host"].set(cred.get("smtp_host", ""))
            self.vars["smtp_port"].set(str(cred.get("smtp_port", "25")))
            self.vars["email_from"].set(cred.get("email_from", ""))
            self.vars["password"].set(cred.get("password", ""))
            self._set_emails(self._email_to_txt, cred.get("email_to", ""))
            self._set_emails(self._email_cc_txt, cred.get("email_cc", ""))
            self.status_lbl.config(text="✔ Kredensial sudah tersimpan sebelumnya", fg=SUCCESS)
        except FileNotFoundError:
            self.status_lbl.config(text="⚠ Belum ada kredensial — isi form lalu klik Simpan", fg=AMBER)
        except Exception:
            self.status_lbl.config(text="⚠ Gagal baca kredensial lama", fg=DANGER)

    def _on_save(self):
        try:
            smtp_port = int(self.vars["smtp_port"].get().strip())
        except ValueError:
            messagebox.showerror("Error", "SMTP Port harus angka!")
            return

        email_to = self._get_emails(self._email_to_txt)
        email_cc = self._get_emails(self._email_cc_txt)

        if not email_to:
            messagebox.showwarning("Peringatan", "Email To tidak boleh kosong!")
            return

        try:
            save_credentials(
                smtp_host  = self.vars["smtp_host"].get().strip(),
                smtp_port  = smtp_port,
                email_from = self.vars["email_from"].get().strip(),
                password   = self.vars["password"].get().strip(),
                email_to   = email_to,
                email_cc   = email_cc,
            )
            self.status_lbl.config(text="✔ Kredensial berhasil disimpan!", fg=SUCCESS)
            messagebox.showinfo("Berhasil", "Kredensial email berhasil disimpan.")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal simpan:\n{e}")

    def _on_test(self):
        self._on_save()
        try:
            cred     = load_credentials()
            email_to = cred["email_to"]
            email_cc = cred.get("email_cc", "")

            if not email_to:
                messagebox.showwarning("Peringatan", "Email To tidak boleh kosong!")
                return

            from send_email_report import _create_thunderbird_draft
            draft_path = _create_thunderbird_draft(
                cred      = cred,
                subject   = "[RPA] Test Email — Konfigurasi Thunderbird Draft",
                body_html = (
                    "<p>Test email draft dari <b>RPA Stock Reconciliation</b> berhasil dibuat.</p>"
                    "<p style='color:#64748B;font-size:12px'>"
                    f"Dari    : {cred['email_from']}<br>"
                    f"Tujuan  : {email_to}"
                    + (f"<br>CC: {email_cc}" if email_cc else "")
                    + "<br><br>File draft dapat dibuka di Thunderbird atau email client lainnya."
                    + "</p>"
                ),
                to        = email_to,
                cc        = email_cc,
            )
            self.status_lbl.config(text=f"✔ Draft email test dibuat untuk {email_to}", fg=SUCCESS)
            messagebox.showinfo(
                "Berhasil",
                f"Draft email test berhasil dibuat!\n\n"
                f"Path: {draft_path}\n\n"
                f"Double-click file .eml untuk membuka di Thunderbird"
            )
        except Exception as e:
            self.status_lbl.config(text=f"✗ Gagal: {e}", fg=DANGER)
            messagebox.showerror("Gagal", f"Gagal buat draft email:\n\n{e}")

    def _on_diagnose(self):
        messagebox.showinfo(
            "Info",
            "Sistem email menggunakan mode Thunderbird Draft.\n\n"
            "Email tidak dikirim langsung via SMTP, tetapi disimpan sebagai file .eml "
            "yang dapat dibuka di Mozilla Thunderbird untuk direview sebelum dikirim.\n\n"
            "Kredensial email dan SMTP masih disimpan untuk referensi.\n"
            "Anda dapat mengubah atau menghapusnya jika sudah tidak diperlukan."
        )


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()

    # Samakan skala Tk dengan DPI monitor sebenarnya. Tanpa ini, walau
    # proses sudah "DPI aware", Tk tetap menghitung ukuran font/widget
    # seolah layar 96 DPI — hasilnya teks bisa tetap terlihat kecil/blur
    # di layar dengan Windows scaling 125%/150%/200%.
    try:
        from ctypes import windll
        actual_dpi = windll.user32.GetDpiForWindow(root.winfo_id())
        root.tk.call('tk', 'scaling', actual_dpi / 72)
    except Exception:
        pass

    app  = EmailConfigUI(root)
    root.mainloop()