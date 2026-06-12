"""
email_config_ui.py
UI untuk simpan & ubah kredensial email SMTP.
Password disimpan terenkripsi di file lokal (tidak hardcode).

Perbaikan:
- Email To dan CC pakai Text widget multi-baris (bisa tampung banyak email)
- Password boleh dikosongkan (untuk Zimbra internal relay tanpa auth)
- Tambah tombol Diagnosa: cek port 25/465/587 mana yang bisa konek
"""

import os
import json
import tkinter as tk
from tkinter import messagebox, scrolledtext
from cryptography.fernet import Fernet

# ─────────────────────────────────────────────
# PATH FILE KREDENSIAL
# ─────────────────────────────────────────────
BASE_DIR  = r"C:\RPA_StockRecon"
CRED_FILE = os.path.join(BASE_DIR, "config", "email_cred.enc")
KEY_FILE  = os.path.join(BASE_DIR, "config", "email_key.key")

BG       = "#1A1F2E"
BG_DARK  = "#0F172A"
BG_PANEL = "#1E293B"
TEXT_PRI = "#E2E8F0"
TEXT_MUT = "#94A3B8"
TEXT_HNT = "#475569"
ACCENT   = "#60A5FA"
BORDER   = "#334155"


# ─────────────────────────────────────────────
# ENKRIPSI / DEKRIPSI
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
    os.makedirs(os.path.dirname(CRED_FILE), exist_ok=True)
    key  = _get_or_create_key()
    fern = Fernet(key)
    data = json.dumps({
        "smtp_host":  smtp_host,
        "smtp_port":  smtp_port,
        "email_from": email_from,
        "password":   password,
        "email_to":   email_to,
        "email_cc":   email_cc,
    }).encode()
    with open(CRED_FILE, "wb") as f:
        f.write(fern.encrypt(data))


def load_credentials() -> dict:
    if not os.path.exists(CRED_FILE):
        raise FileNotFoundError(
            "Kredensial email belum dikonfigurasi.\n"
            "Jalankan python email_config_ui.py untuk setup."
        )
    key  = _get_or_create_key()
    fern = Fernet(key)
    with open(CRED_FILE, "rb") as f:
        token = f.read()
    return json.loads(fern.decrypt(token).decode())


# ─────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────

class EmailConfigUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Konfigurasi Email — Thunderbird Draft (RPA Stock Recon)")
        self.root.geometry("600x620")
        self.root.resizable(False, True)
        self.root.configure(bg=BG)
        self._build_ui()
        self._load_existing()

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg="#111827", pady=14)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="⚙  Konfigurasi Email — Thunderbird Draft",
            font=("Consolas", 13, "bold"), fg=ACCENT, bg="#111827"
        ).pack(padx=20, anchor="w")
        tk.Label(
            hdr,
            text="Email disimpan sebagai draft di folder report — buka di Thunderbird sebelum mengirim",
            font=("Consolas", 8), fg=TEXT_HNT, bg="#111827"
        ).pack(padx=20, anchor="w")

        # Form
        form = tk.Frame(self.root, bg=BG, padx=24, pady=14)
        form.pack(fill="both", expand=True)
        form.columnconfigure(1, weight=1)

        def lbl(row, text):
            tk.Label(
                form, text=text, fg=TEXT_MUT, bg=BG,
                font=("Consolas", 10), anchor="w"
            ).grid(row=row, column=0, sticky="nw", pady=(10, 0), padx=(0, 12))

        def entry(row, var, show="", width=38):
            e = tk.Entry(
                form, textvariable=var, show=show, width=width,
                bg=BG_DARK, fg=TEXT_PRI, insertbackground=ACCENT,
                font=("Consolas", 10), relief="flat", bd=5
            )
            e.grid(row=row, column=1, sticky="ew", pady=(10, 0))
            return e

        # ── SMTP Host ─────────────────────────────────────────
        lbl(0, "SMTP Host")
        self.vars = {}
        self.vars["smtp_host"] = tk.StringVar()
        entry(0, self.vars["smtp_host"])

        # ── SMTP Port ─────────────────────────────────────────
        lbl(1, "SMTP Port")
        self.vars["smtp_port"] = tk.StringVar()
        entry(1, self.vars["smtp_port"], width=10)

        # ── Email Pengirim ────────────────────────────────────
        lbl(2, "Email Pengirim")
        self.vars["email_from"] = tk.StringVar()
        entry(2, self.vars["email_from"])

        # ── Password ──────────────────────────────────────────
        lbl(3, "Password")
        self.vars["password"] = tk.StringVar()
        pw_frame = tk.Frame(form, bg=BG)
        pw_frame.grid(row=3, column=1, sticky="ew", pady=(10, 0))
        pw_frame.columnconfigure(0, weight=1)

        self._pw_entry = tk.Entry(
            pw_frame, textvariable=self.vars["password"],
            show="*", bg=BG_DARK, fg=TEXT_PRI,
            insertbackground=ACCENT,
            font=("Consolas", 10), relief="flat", bd=5
        )
        self._pw_entry.grid(row=0, column=0, sticky="ew")
        self._pw_shown = False

        tk.Button(
            pw_frame, text="👁", font=("Consolas", 9),
            bg=BG_PANEL, fg=TEXT_MUT, relief="flat",
            cursor="hand2", command=self._toggle_pw, padx=6
        ).grid(row=0, column=1, padx=(4, 0))

        tk.Label(
            form,
            text="* Opsional — disimpan untuk referensi (tidak lagi digunakan untuk SMTP)",
            fg=TEXT_HNT, bg=BG, font=("Consolas", 8)
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(2, 4))

        # ── Email To (multi-baris) ────────────────────────────
        lbl(5, "Email To")
        tk.Label(
            form, text="(pisah koma atau Enter)",
            fg=TEXT_HNT, bg=BG, font=("Consolas", 8)
        ).grid(row=5, column=1, sticky="ne", pady=(10, 0))

        self._email_to_txt = tk.Text(
            form, height=3, wrap="word",
            bg=BG_DARK, fg=TEXT_PRI, insertbackground=ACCENT,
            font=("Consolas", 10), relief="flat", bd=5,
            selectbackground=ACCENT, selectforeground=BG_DARK
        )
        self._email_to_txt.grid(row=6, column=0, columnspan=2,
                                 sticky="ew", pady=(2, 0))

        # ── Email CC (multi-baris) ────────────────────────────
        lbl(7, "Email CC")
        tk.Label(
            form, text="(pisah koma atau Enter, opsional)",
            fg=TEXT_HNT, bg=BG, font=("Consolas", 8)
        ).grid(row=7, column=1, sticky="ne", pady=(10, 0))

        self._email_cc_txt = tk.Text(
            form, height=3, wrap="word",
            bg=BG_DARK, fg=TEXT_PRI, insertbackground=ACCENT,
            font=("Consolas", 10), relief="flat", bd=5,
            selectbackground=ACCENT, selectforeground=BG_DARK
        )
        self._email_cc_txt.grid(row=8, column=0, columnspan=2,
                                 sticky="ew", pady=(2, 0))

        # ── Buttons ───────────────────────────────────────────
        btn_frame = tk.Frame(self.root, bg=BG, pady=10, padx=24)
        btn_frame.pack(fill="x")

        tk.Button(
            btn_frame, text="💾  Simpan",
            font=("Consolas", 11, "bold"), fg="#0F172A", bg="#3B82F6",
            activebackground="#2563EB", relief="flat",
            padx=18, pady=7, cursor="hand2",
            command=self._on_save
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btn_frame, text="✉  Test Draft",
            font=("Consolas", 11), fg="#0F172A", bg="#10B981",
            activebackground="#059669", relief="flat",
            padx=14, pady=7, cursor="hand2",
            command=self._on_test
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btn_frame, text="ℹ  Info",
            font=("Consolas", 11), fg="#0F172A", bg="#F59E0B",
            activebackground="#D97706", relief="flat",
            padx=14, pady=7, cursor="hand2",
            command=self._on_diagnose
        ).pack(side="left")

        tk.Button(
            btn_frame, text="✕  Tutup",
            font=("Consolas", 10), fg=TEXT_MUT, bg=BG_PANEL,
            activebackground=BORDER, relief="flat",
            padx=12, pady=7, cursor="hand2",
            command=self.root.destroy
        ).pack(side="right")

        # Status label
        self.status_lbl = tk.Label(
            self.root, text="", font=("Consolas", 9),
            fg="#34D399", bg=BG, wraplength=560
        )
        self.status_lbl.pack(pady=(0, 8))

    # ── HELPER: baca/tulis Text widget ───────────────────────

    def _get_emails(self, widget: tk.Text) -> str:
        """
        Baca isi Text widget, bersihkan newline & spasi,
        kembalikan sebagai string pisah koma.
        Contoh: 'a@b.com\\nc@d.com' → 'a@b.com, c@d.com'
        """
        raw = widget.get("1.0", "end").strip()
        # Split by koma atau newline, bersihkan tiap item
        parts = []
        for part in raw.replace("\n", ",").split(","):
            part = part.strip()
            if part:
                parts.append(part)
        return ", ".join(parts)

    def _set_emails(self, widget: tk.Text, value: str):
        """
        Isi Text widget dari string pisah koma.
        Tiap email ditaruh di baris sendiri agar mudah dibaca.
        """
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
            self.status_lbl.config(
                text="✔ Kredensial sudah tersimpan sebelumnya", fg="#34D399"
            )
        except FileNotFoundError:
            self.status_lbl.config(
                text="⚠ Belum ada kredensial — isi form lalu klik Simpan",
                fg="#FBBF24"
            )
        except Exception:
            self.status_lbl.config(
                text="⚠ Gagal baca kredensial lama", fg="#F87171"
            )

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
            self.status_lbl.config(
                text="✔ Kredensial berhasil disimpan!", fg="#34D399"
            )
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
            self.status_lbl.config(
                text=f"✔ Draft email test dibuat untuk {email_to}", fg="#34D399"
            )
            messagebox.showinfo(
                "Berhasil",
                f"Draft email test berhasil dibuat!\n\n"
                f"Path: {draft_path}\n\n"
                f"Double-click file .eml untuk membuka di Thunderbird"
            )
        except Exception as e:
            self.status_lbl.config(text=f"✗ Gagal: {e}", fg="#F87171")
            messagebox.showerror("Gagal", f"Gagal buat draft email:\n\n{e}")

    def _on_diagnose(self):
        messagebox.showinfo(
            "Info",
            "Sistem email telah diubah ke mode Thunderbird Draft.\n\n"
            "Email tidak lagi dikirim via SMTP, tetapi disimpan sebagai file .eml "
            "yang dapat dibuka di Mozilla Thunderbird.\n\n"
            "Kredensial email dan SMTP masih disimpan untuk referensi.\n"
            "Anda dapat mengubah atau menghapusnya jika sudah tidak diperlukan."
        )


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app  = EmailConfigUI(root)
    root.mainloop()