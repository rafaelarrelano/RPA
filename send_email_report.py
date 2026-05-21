"""
send_email_report.py
Buat laporan Excel selisih stok per plant, lalu kirim otomatis via SMTP.
Kredensial dibaca dari file terenkripsi (diatur via email_config_ui.py).

Perubahan:
- Subject per plant: "Req. Adj. EOD Plant {code} {name} Tanggal {tanggal}"
- Body email per plant: format standar sesuai template (Posting Date, Material,
  Selisih2, UOM, Gudang, Plant, Adj, Plant Name)
- Kirim SATU EMAIL PER PLANT (bukan satu email untuk semua plant)
- Load Plant Name dari Excel plant_mapping (kolom B=Plant Code, C=Plant Name)
- Auto-detect apakah server support SMTP AUTH atau tidak
- Support semua port: 587, 465, 25, dan port custom lainnya
"""

import os
import smtplib
import ssl
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import Config
from logger import setup_logger
from email_config_ui import load_credentials

log = setup_logger()


# ─────────────────────────────────────────────
# LOAD PLANT NAME MAPPING DARI EXCEL
# ─────────────────────────────────────────────

def load_plant_name_mapping(filepath: str = None) -> dict:
    """
    Baca mapping Plant Code → Plant Name dari Excel plant_mapping.
    Struktur Excel (sheet Plant_CostCenter):
      Kolom B = Plant Code, Kolom C = Plant Name
      Data mulai baris 5

    Return: { "4502": "PGD Surabaya", "4503": "PGD Jakarta GT", ... }
    """
    if filepath is None:
        filepath = Config.PLANT_MAPPING_FILE

    mapping = {}
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb["Plant_CostCenter"]
        for row in ws.iter_rows(min_row=5, values_only=True):
            # row[1] = kolom B (Plant Code), row[2] = kolom C (Plant Name)
            if row[1] and row[2]:
                plant_code = str(row[1]).strip()
                plant_name = str(row[2]).strip()
                mapping[plant_code] = plant_name
        log.info(f"[MAPPING] {len(mapping)} plant name berhasil dibaca")
    except Exception as e:
        log.warning(f"[MAPPING] Gagal baca plant name mapping: {e} — akan pakai kode plant saja")

    return mapping


# ─────────────────────────────────────────────
# BUAT FILE EXCEL LAPORAN
# ─────────────────────────────────────────────

def build_excel_report(items_per_plant: dict, filepath: str,
                       plant_name_map: dict = None):
    """
    Buat file Excel laporan selisih stok.
    plant_name_map: { plant_code: plant_name } untuk judul sheet & summary
    """
    if plant_name_map is None:
        plant_name_map = {}

    wb        = openpyxl.Workbook()
    wb.remove(wb.active)
    tanggal   = datetime.now().strftime("%d/%m/%Y")
    hdr_fill  = PatternFill("solid", fgColor="1F5C99")
    hdr_font  = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    ttl_font  = Font(bold=True, color="1F5C99", name="Calibri", size=12)
    even_fill = PatternFill("solid", fgColor="EBF3FB")
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    thin      = Side(style="thin", color="CCCCCC")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    center    = Alignment(horizontal="center", vertical="center")

    # Header kolom sesuai template standar (cocok dengan gambar referensi)
    HEADERS = ["No", "Posting Date", "Material", "Selisih2",
               "UOM", "Gudang", "Plant", "Adj", "Plant Name"]
    COL_W   = [5, 14, 14, 13, 8, 10, 8, 8, 20]

    summary_rows = []

    for plant, items in sorted(items_per_plant.items()):
        plant_name = plant_name_map.get(plant, "")
        ws = wb.create_sheet(title=f"Plant {plant}")
        ws.merge_cells("A1:I1")
        display_name = f"{plant} {plant_name}".strip()
        ws["A1"] = f"Req. Adj. EOD Plant {display_name} — {tanggal}"
        ws["A1"].font      = ttl_font
        ws["A1"].alignment = center
        ws.row_dimensions[1].height = 22
        ws.row_dimensions[3].height = 18

        for col_idx, (h, w) in enumerate(zip(HEADERS, COL_W), start=1):
            cell            = ws.cell(row=3, column=col_idx, value=h)
            cell.fill       = hdr_fill
            cell.font       = hdr_font
            cell.alignment  = center
            cell.border     = border
            ws.column_dimensions[get_column_letter(col_idx)].width = w

        total_917 = total_918 = 0

        # Urutkan: Gudang (sloc) dulu, lalu material
        sorted_items = sorted(items, key=lambda x: (x.sloc, x.material))

        for i, item in enumerate(sorted_items, start=1):
            row  = 3 + i
            fill = even_fill if i % 2 == 0 else PatternFill()

            vals = [
                i,
                item.posting_date,          # Posting Date (dd.mm.yyyy)
                item.material,              # Material
                abs(item.diff),             # Selisih2 = nilai absolut selisih
                "CAR",                      # UOM default
                item.sloc,                  # Gudang = SLoc
                item.plant,                 # Plant
                item.mvt_type,              # Adj = movement type (917/918)
                plant_name,                 # Plant Name
            ]

            for col_idx, val in enumerate(vals, start=1):
                cell           = ws.cell(row=row, column=col_idx, value=val)
                cell.border    = border
                cell.alignment = center
                if fill.fill_type:
                    cell.fill = fill
                if col_idx == 4 and abs(item.diff) >= 1:
                    cell.fill = warn_fill
                if col_idx == 4:
                    cell.number_format = "#,##0.000"

            total_917 += item.mvt_type == "917"
            total_918 += item.mvt_type == "918"

        fr = 3 + len(sorted_items) + 2
        for label, val in [
            ("Total item",            len(sorted_items)),
            ("Mvt 917 (kurangi SAP)", total_917),
            ("Mvt 918 (tambah SAP)",  total_918),
        ]:
            ws.cell(row=fr, column=1, value=label).font = Font(bold=True)
            ws.cell(row=fr, column=2, value=val)
            fr += 1

        summary_rows.append({
            "plant":      plant,
            "plant_name": plant_name,
            "total":      len(sorted_items),
            "mvt_917":    total_917,
            "mvt_918":    total_918,
        })

    # ── Sheet Summary ────────────────────────────────────────
    ws_s = wb.create_sheet(title="Summary", index=0)
    ws_s.merge_cells("A1:G1")
    ws_s["A1"]           = f"Summary Req. Adj. EOD — {tanggal}"
    ws_s["A1"].font      = ttl_font
    ws_s["A1"].alignment = center
    ws_s.row_dimensions[1].height = 24

    sum_hdrs = ["No", "Plant", "Plant Name", "Total Selisih",
                "Mvt 917 (Kurangi SAP)", "Mvt 918 (Tambah SAP)", "Catatan"]
    sum_w    = [5, 8, 20, 14, 22, 22, 35]

    for col_idx, (h, w) in enumerate(zip(sum_hdrs, sum_w), start=1):
        cell            = ws_s.cell(row=3, column=col_idx, value=h)
        cell.fill       = hdr_fill
        cell.font       = hdr_font
        cell.alignment  = center
        cell.border     = border
        ws_s.column_dimensions[get_column_letter(col_idx)].width = w

    total_all = 0
    for i, r in enumerate(summary_rows, start=1):
        row  = 3 + i
        fill = even_fill if i % 2 == 0 else PatternFill()
        vals = [i, r["plant"], r["plant_name"], r["total"],
                r["mvt_917"], r["mvt_918"],
                "Perlu review & approval sebelum posting manual"]
        for col_idx, val in enumerate(vals, start=1):
            cell           = ws_s.cell(row=row, column=col_idx, value=val)
            cell.border    = border
            cell.alignment = center
            if fill.fill_type:
                cell.fill = fill
        total_all += r["total"]

    gt = 3 + len(summary_rows) + 1
    ws_s.cell(row=gt, column=1, value="TOTAL").font = Font(bold=True, color="1F5C99")
    ws_s.cell(row=gt, column=4, value=total_all).font = Font(bold=True, color="1F5C99")

    wb.save(filepath)
    log.info(f"[REPORT] Excel disimpan: {os.path.basename(filepath)}")


# ─────────────────────────────────────────────
# BUILD HTML BODY EMAIL PER PLANT
# ─────────────────────────────────────────────

def _build_body_html_per_plant(plant: str, plant_name: str,
                                items: list, posting_date: str) -> str:
    """
    Buat body HTML email untuk satu plant.
    Format tabel: Posting Date | Material | Selisih2 | UOM | Gudang | Plant | Adj | Plant Name
    Sesuai template standar dari gambar referensi.
    """
    sorted_items = sorted(items, key=lambda x: (x.sloc, x.material))

    th_style = (
        "padding:7px 12px;background:#1F5C99;color:white;"
        "font-family:Calibri,Arial;font-size:12px;border:1px solid #CBD5E1;"
    )
    td_base = (
        "padding:6px 12px;font-family:Calibri,Arial;font-size:12px;"
        "border:1px solid #CBD5E1;text-align:center;"
    )

    header_row = (
        f"<tr>"
        f"<th style='{th_style}'>Posting Date</th>"
        f"<th style='{th_style}'>Material</th>"
        f"<th style='{th_style}'>Selisih2</th>"
        f"<th style='{th_style}'>UOM</th>"
        f"<th style='{th_style}'>Gudang</th>"
        f"<th style='{th_style}'>Plant</th>"
        f"<th style='{th_style}'>Adj</th>"
        f"<th style='{th_style}'>Plant Name</th>"
        f"</tr>"
    )

    data_rows = ""
    for i, item in enumerate(sorted_items):
        bg  = "#EBF3FB" if i % 2 == 0 else "#FFFFFF"
        td  = td_base + f"background:{bg};"
        # Format selisih 3 desimal pakai koma (standar Indonesia)
        selisih_fmt = f"{abs(item.diff):,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
        data_rows += (
            f"<tr>"
            f"<td style='{td}'>{item.posting_date}</td>"
            f"<td style='{td}'>{item.material}</td>"
            f"<td style='{td}'>{selisih_fmt}</td>"
            f"<td style='{td}'>CAR</td>"
            f"<td style='{td}'>{item.sloc}</td>"
            f"<td style='{td}'>{item.plant}</td>"
            f"<td style='{td}'>{item.mvt_type}</td>"
            f"<td style='{td}'>{plant_name}</td>"
            f"</tr>"
        )

    total_917 = len([x for x in sorted_items if x.mvt_type == "917"])
    total_918 = len([x for x in sorted_items if x.mvt_type == "918"])

    body_html = f"""<html>
<body style="font-family:Calibri,Arial;font-size:13px;color:#1E293B;margin:0;padding:0">

<p style="margin:0 0 8px 0">Dear Team Accounting,</p>

<p style="margin:0 0 16px 0">
  Mohon dibantu untuk dilakukan Adjustment atas Selisih Endstock EOD:
</p>

<table border="0" cellpadding="0" cellspacing="0"
       style="border-collapse:collapse;margin-bottom:4px">
  {header_row}
  {data_rows}
</table>

<p style="margin:12px 0 4px 0;font-size:12px;color:#475569">
  Total item: <b>{len(sorted_items)}</b> &nbsp;|&nbsp;
  Mvt 917 (Kurangi SAP): <b>{total_917}</b> &nbsp;|&nbsp;
  Mvt 918 (Tambah SAP): <b>{total_918}</b>
</p>

<p style="margin:16px 0 0 0;font-size:13px;color:#1E293B">
  Terima Kasih,<br>
  <b>RPA</b>
</p>

</body>
</html>"""
    return body_html


# ─────────────────────────────────────────────
# KIRIM EMAIL VIA SMTP
# ─────────────────────────────────────────────

def _smtp_send(cred: dict, subject: str, body_html: str,
               to: str, cc: str, attachment_path: str = None):
    """
    Kirim email via SMTP.
    Auto-detect apakah server support AUTH atau tidak.
    - Port 465 : SSL langsung
    - Port 587 : STARTTLS, lalu coba login jika password ada & AUTH supported
    - Port lain : Plain/STARTTLS, skip login jika AUTH tidak didukung (Zimbra relay)
    """
    msg            = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = cred["email_from"]
    msg["To"]      = to
    if cc:
        msg["Cc"] = cc
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEBase(
                "application",
                "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{os.path.basename(attachment_path)}"'
        )
        msg.attach(part)

    recipients = [e.strip() for e in to.split(",") if e.strip()]
    if cc:
        recipients += [e.strip() for e in cc.split(",") if e.strip()]

    host     = cred["smtp_host"]
    port     = int(cred["smtp_port"])
    password = cred.get("password", "").strip()
    sender   = cred["email_from"]

    log.info(f"[EMAIL] Konek ke {host}:{port} ...")

    try:
        if port == 465:
            ctx = ssl.create_default_context()
            with smtplib.SMTP_SSL(host, port, context=ctx, timeout=15) as server:
                server.ehlo()
                if password:
                    server.login(sender, password)
                server.sendmail(sender, recipients, msg.as_string())

        elif port == 587:
            with smtplib.SMTP(host, port, timeout=15) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                auth_supported = "auth" in server.esmtp_features
                if password and auth_supported:
                    server.login(sender, password)
                elif password and not auth_supported:
                    log.warning("[EMAIL] Server tidak support AUTH di port 587 — kirim tanpa login")
                server.sendmail(sender, recipients, msg.as_string())

        else:
            with smtplib.SMTP(host, port, timeout=15) as server:
                server.ehlo()
                try:
                    server.starttls()
                    server.ehlo()
                    log.info("[EMAIL] STARTTLS berhasil")
                except (smtplib.SMTPNotSupportedError, smtplib.SMTPException):
                    log.info("[EMAIL] STARTTLS tidak didukung — lanjut plain")

                auth_supported = "auth" in server.esmtp_features
                if password and auth_supported:
                    server.login(sender, password)
                    log.info("[EMAIL] Login berhasil")
                else:
                    log.info("[EMAIL] Relay tanpa auth (internal Zimbra/Exchange)")

                server.sendmail(sender, recipients, msg.as_string())

        log.info(f"[EMAIL] ✓ Terkirim ke: {to}" + (f" | CC: {cc}" if cc else ""))

    except smtplib.SMTPAuthenticationError as e:
        log.error(f"[EMAIL] Autentikasi gagal: {e}")
        raise
    except smtplib.SMTPConnectError as e:
        log.error(f"[EMAIL] Gagal konek ke {host}:{port}: {e}")
        raise
    except smtplib.SMTPRecipientsRefused as e:
        log.error(f"[EMAIL] Penerima ditolak: {e}")
        raise
    except Exception as e:
        log.error(f"[EMAIL] Gagal kirim: {e}")
        raise


def diagnose_smtp(host: str, ports: list = None) -> dict:
    """
    Cek port mana yang bisa konek ke SMTP server.
    Return: { port: status_string }
    """
    if ports is None:
        ports = [25, 465, 587]

    results = {}
    for port in ports:
        try:
            if port == 465:
                ctx = ssl.create_default_context()
                with smtplib.SMTP_SSL(host, port, context=ctx, timeout=8) as s:
                    banner = s.ehlo()[1].decode(errors="ignore")
                    auth   = "auth" in s.esmtp_features
                    results[port] = f"✓ SSL OK | AUTH={'Ya' if auth else 'Tidak'} | {banner[:50]}"
            else:
                with smtplib.SMTP(host, port, timeout=8) as s:
                    s.ehlo()
                    tls_ok = False
                    try:
                        s.starttls(); s.ehlo(); tls_ok = True
                    except Exception:
                        pass
                    auth = "auth" in s.esmtp_features
                    results[port] = (
                        f"✓ OK | TLS={'Ya' if tls_ok else 'Tidak'} | "
                        f"AUTH={'Ya' if auth else 'Tidak (relay mode)'}"
                    )
        except Exception as e:
            results[port] = f"✗ Gagal: {e}"

    return results


# ─────────────────────────────────────────────
# FUNGSI UTAMA — KIRIM SATU EMAIL PER PLANT
# ─────────────────────────────────────────────

def send_stock_diff_report(
    items_per_plant: dict,
    override_to: str = None,
    override_cc: str = None,
) -> str:
    """
    Buat laporan Excel (semua plant dalam satu file) dan kirim
    SATU EMAIL PER PLANT via SMTP.

    Hanya item FSTKGD yang dikirim via email.
    FSTKVN tetap diproses untuk compare & U2C, tapi tidak masuk laporan email.

    Subject per plant:
        Req. Adj. EOD Plant {code} {name} Tanggal {tanggal posting}

    override_to / override_cc: override nilai dari kredensial tersimpan
    Return: path file Excel yang dibuat (atau "" jika tidak ada selisih FSTKGD)
    """
    # Filter: hanya FSTKGD yang masuk email
    items_email = {
        plant: [i for i in items if i.param == "FSTKGD"]
        for plant, items in items_per_plant.items()
    }
    items_email = {p: v for p, v in items_email.items() if v}   # hapus plant kosong

    if not items_email:
        log.info("[REPORT] Tidak ada selisih FSTKGD — email tidak dikirim")
        return ""

    cred      = load_credentials()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if override_to:
        cred["email_to"] = override_to
    if override_cc is not None:
        cred["email_cc"] = override_cc

    # Load mapping plant name dari Excel
    plant_name_map = load_plant_name_mapping()

    os.makedirs(Config.FOLDER_REPORT, exist_ok=True)
    excel_path = os.path.join(Config.FOLDER_REPORT, f"SelisihStok_{timestamp}.xlsx")

    log.info("[REPORT] Membuat file Excel laporan...")
    build_excel_report(items_email, excel_path, plant_name_map)

    # ── Kirim satu email per plant ───────────────────────────
    sent_count = 0
    for plant, items in sorted(items_email.items()):
        if not items:
            continue

        plant_name   = plant_name_map.get(plant, "")
        display_name = f"{plant} {plant_name}".strip()

        # Posting date dari item pertama (format dd.mm.yyyy)
        posting_date = items[0].posting_date

        # Subject sesuai template standar
        subject = f"Req. Adj. EOD Plant {display_name} Tanggal {posting_date}"

        # Body email per plant dengan tabel adjustment
        body_html = _build_body_html_per_plant(
            plant        = plant,
            plant_name   = plant_name,
            items        = items,
            posting_date = posting_date,
        )

        log.info(f"[EMAIL] Kirim plant {display_name} → {cred['email_to']}")
        log.info(f"[EMAIL] Subject: {subject}")

        try:
            _smtp_send(
                cred            = cred,
                subject         = subject,
                body_html       = body_html,
                to              = cred["email_to"],
                cc              = cred.get("email_cc", ""),
                attachment_path = excel_path,
            )
            sent_count += 1
            log.info(f"[EMAIL] ✓ Plant {plant} terkirim")
        except Exception as e:
            log.error(f"[EMAIL] ✗ Plant {plant} gagal: {e}")
            continue  # Lanjut ke plant berikutnya meski ada error

    total_plant = len(items_email)
    total_item  = sum(len(v) for v in items_email.values())
    fstkvn_skip = sum(
        len([i for i in v if i.param == "FSTKVN"])
        for v in items_per_plant.values()
    )
    log.info(
        f"[REPORT] Selesai | {sent_count}/{total_plant} plant email terkirim | "
        f"{total_item} item FSTKGD | {fstkvn_skip} item FSTKVN dilewati (tidak dikirim email)"
    )
    return excel_path