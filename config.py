"""
config.py - Semua konfigurasi robot
Ganti nilai di sini sesuai environment kamu
"""
import os

class Config:
    # ── SAP ──────────────────────────────────────────────────
    SAP_CLIENT       = "120"
    SAP_USER         = "ITAPPS6.HC "
    SAP_PASSWORD     = os.environ.get("SAP_PASSWORD", "harco2026")
    SAP_TCODE_EXPORT = "ZPGD_SAPSTK"   # default fallback

    # ── Mapping Portal → T-code SAP ──────────────────────────
    # PGDMTX → /NZPGD_SAPSTK  +  /NZPGD_U2C
    # CMIS   → /NZCNS_SAPSTK  +  /NZCNS_U2C
    PORTAL_TCODE_MAP = {
        "PGDMTX": {"sapstk": "ZPGD_SAPSTK", "u2c": "ZPGD_U2C"},
        "CMIS":   {"sapstk": "ZCNS_SAPSTK", "u2c": "ZCNS_U2C"},
    }

    # T-code aktif — di-set otomatis GUI saat user pilih portal
    ACTIVE_TCODE_SAPSTK = "ZPGD_SAPSTK"
    ACTIVE_TCODE_U2C    = "ZPGD_U2C"

    # Folder otomatis hasil export ZPGD_SAPSTK
    SAP_DOWNLOAD_DIR = r"C:\MATRIX\DOWNLOAD"

    # Daftar semua plant yang perlu diproses tiap hari
    PLANTS = [
        "4502", "4503", "4504", "4505", "4506",
        "4507", "4508", "4509", "4510", "4511",
        "4513", "4516",
    ]

    # ── Portal ───────────────────────────────────────────────
    PORTAL_URL      = "https://portal.mayora.co.id"
    PORTAL_USER     = os.environ.get("PORTAL_USER", "mg134941")
    PORTAL_PASSWORD = os.environ.get("PORTAL_PASSWORD", "Rafaell22")

    # ── Folder robot ─────────────────────────────────────────
    BASE_DIR           = r"C:\RPA_StockRecon"
    FOLDER_INPUT       = os.path.join(BASE_DIR, "Input")
    FOLDER_OUTPUT      = os.path.join(BASE_DIR, "Output")
    FOLDER_SCREENSHOTS = os.path.join(BASE_DIR, "Screenshots")
    FOLDER_LOGS        = os.path.join(BASE_DIR, "Logs")

    # ── Mapping plant → cost center ──────────────────────────
    PLANT_MAPPING_FILE = r"C:\Users\user\source\repos\RPA\config\plant_mapping.xlsx"

    # ── File limit adjustment decimal ────────────────────────
    LIMIT_ADJ_FILE = r"C:\Users\user\source\repos\RPA\config\List Limit Adj. Material SAP.xlsx"

    # ── Toleransi default ────────────────────────────────────
    DIFF_THRESHOLD = 0.0

    # ── Email via Thunderbird ────────────────────────────────
    # Path instalasi Thunderbird (sesuaikan jika berbeda)
    THUNDERBIRD_PATH = r"C:\Program Files\Mozilla Thunderbird\thunderbird.exe"

    # Penerima email (pisah koma jika lebih dari satu)
    EMAIL_TO_ACCOUNTING = "steven.pedro@mayora.co.id"

    # CC (kosongkan string jika tidak perlu)
    EMAIL_CC = "yadi.it@mayora.co.id"

    # Subject prefix — tanggal akan ditambahkan otomatis
    EMAIL_SUBJECT_PREFIX = "[RPA] Selisih Stok Matrix vs SAP"

    # Folder simpan file Excel laporan sebelum dikirim
    FOLDER_REPORT = os.path.join(BASE_DIR, "Report")