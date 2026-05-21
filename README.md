# RPA Stock Reconciliation - Python + pywin32

Otomasi rekonsiliasi stok Matrix vs SAP via SAP GUI Scripting.

## Setup di VS Code

### 1. Buat virtual environment
```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
playwright install chromium
```

### 3. Set environment variables (JANGAN hardcode password)
```bash
# Windows PowerShell
$env:SAP_PASSWORD   = "password_sap_kamu"
$env:EMAIL_PASSWORD = "password_email_kamu"
```

### 4. Buat folder struktur
```
C:\RPA_StockRecon\
  Input\         <- file .txt dari SAP
  Output\        <- backup notif selisih dari portal
  Screenshots\   <- bukti posting MIGO
  Logs\          <- log harian robot
```

### 5. Sesuaikan config.py
Edit `config.py` dengan nilai yang sesuai:
- SAP_CLIENT, SAP_USER, SAP_TCODE_EXPORT
- PORTAL_URL dan selector elemen portal
- EMAIL_FROM, EMAIL_TO, SMTP_HOST

### 6. Sesuaikan selector SAP di main.py
Selector SAP GUI (contoh: `wnd[0]/usr/...`) bisa berbeda
tergantung versi SAP. Gunakan SAP GUI Scripting Recorder
untuk mendapatkan ID yang tepat:
- Di SAP, klik Tools > Macro > Record Script
- Lakukan langkah manual yang ingin direkam
- Salin ID dari hasil rekaman ke main.py

### 7. Sesuaikan selector portal di main.py
Ganti selector `input[type='file']`, `button[type='submit']`,
dan `#notif-result` di fungsi `get_notif_from_portal()`
dengan selector yang sesuai portal kamu.
Gunakan Playwright Codegen untuk auto-detect:
```bash
playwright codegen https://portal.company.com
```

## Jalankan robot
```bash
python main.py
```

## Jalankan unit tests
```bash
pytest test_rpa.py -v
```

## Jalankan otomatis (Task Scheduler Windows)
1. Buka Task Scheduler
2. Create Basic Task > Daily > jam 18:00
3. Action: Start a Program
   Program: C:\RPA_StockRecon\.venv\Scripts\python.exe
   Arguments: C:\RPA_StockRecon\main.py
   Start in: C:\RPA_StockRecon\

## Struktur file
```
rpa_stock_recon/
  main.py          <- robot utama + semua logika
  config.py        <- konfigurasi
  logger.py        <- setup logging
  test_rpa.py      <- unit tests (pytest)
  requirements.txt <- dependencies
  README.md        <- panduan ini
```
