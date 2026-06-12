"""
test_draft_thunderbird.py
Test file untuk memverifikasi draft email Thunderbird berfungsi dengan baik.

Membuat sample data dan menjalankan send_stock_diff_report() untuk generate:
  1. File Excel laporan
  2. File .eml (draft email) per plant

Tidak memerlukan data dari SAP atau portal - cukup test data lokal.
"""

import os
from datetime import datetime
from dataclasses import dataclass

# Import dari module yang ada
from main import StockDiff
from send_email_report import send_stock_diff_report
from config import Config
from logger import setup_logger

log = setup_logger()


def test_draft_email_generation():
    """
    Test: Buat sample data dan generate draft email.
    """
    print("=" * 70)
    print("TEST: Draft Email Generation untuk Thunderbird")
    print("=" * 70)
    print()

    # ─────────────────────────────────────────────
    # STEP 1: Buat sample data StockDiff
    # ─────────────────────────────────────────────
    print("▶ STEP 1: Membuat sample data StockDiff...")
    print()

    sample_items = [
        # Plant 4502 - FSTKGD items (akan masuk email)
        StockDiff(
            param="FSTKGD",
            plant="4502",
            sloc="S401",
            posting_date="08.06.2026",
            material="101001",
            qty_matrix=1000.500,
            qty_sap=1001.000,
            diff=-0.500,
            status=1,
            mvt_type="917",
            qty_adjust=0.500,
        ),
        StockDiff(
            param="FSTKGD",
            plant="4502",
            sloc="S402",
            posting_date="08.06.2026",
            material="202002",
            qty_matrix=500.250,
            qty_sap=499.000,
            diff=1.250,
            status=1,
            mvt_type="918",
            qty_adjust=1.250,
        ),
        StockDiff(
            param="FSTKGD",
            plant="4502",
            sloc="S401",
            posting_date="08.06.2026",
            material="303003",
            qty_matrix=2000.000,
            qty_sap=2010.000,
            diff=-10.000,
            status=1,
            mvt_type="917",
            qty_adjust=10.000,
        ),
        # Plant 4502 - FSTKVN item (tidak akan masuk email, hanya untuk compare)
        StockDiff(
            param="FSTKVN",
            plant="4502",
            sloc="S401",
            posting_date="08.06.2026",
            material="404004",
            qty_matrix=750.000,
            qty_sap=750.000,
            diff=0.000,
            status=1,
            mvt_type="",
            qty_adjust=0.0,
        ),
        # Plant 4503 - FSTKGD items (akan masuk email)
        StockDiff(
            param="FSTKGD",
            plant="4503",
            sloc="S501",
            posting_date="08.06.2026",
            material="101001",
            qty_matrix=5000.000,
            qty_sap=4950.000,
            diff=50.000,
            status=1,
            mvt_type="918",
            qty_adjust=50.000,
        ),
        StockDiff(
            param="FSTKGD",
            plant="4503",
            sloc="S502",
            posting_date="08.06.2026",
            material="505005",
            qty_matrix=300.500,
            qty_sap=300.750,
            diff=-0.250,
            status=1,
            mvt_type="917",
            qty_adjust=0.250,
        ),
    ]

    print(f"✓ {len(sample_items)} sample items dibuat")
    print(f"  - FSTKGD items: {len([x for x in sample_items if x.param == 'FSTKGD'])}")
    print(f"  - FSTKVN items: {len([x for x in sample_items if x.param == 'FSTKVN'])}")
    print()

    # ─────────────────────────────────────────────
    # STEP 2: Organize per plant
    # ─────────────────────────────────────────────
    print("▶ STEP 2: Mengorganisir items per plant...")
    print()

    items_per_plant = {}
    for item in sample_items:
        if item.plant not in items_per_plant:
            items_per_plant[item.plant] = []
        items_per_plant[item.plant].append(item)

    for plant, items in sorted(items_per_plant.items()):
        print(f"  Plant {plant}: {len(items)} items")
    print()

    # ─────────────────────────────────────────────
    # STEP 3: Panggil send_stock_diff_report()
    # ─────────────────────────────────────────────
    print("▶ STEP 3: Menjalankan send_stock_diff_report()...")
    print()

    try:
        excel_path = send_stock_diff_report(items_per_plant)
        print()
        if excel_path:
            print(f"✓ Excel laporan dibuat: {excel_path}")
        else:
            print("⚠ Tidak ada laporan Excel (mungkin tidak ada FSTKGD items)")
    except Exception as e:
        print(f"✗ ERROR di send_stock_diff_report(): {e}")
        import traceback
        traceback.print_exc()
        return False

    # ─────────────────────────────────────────────
    # STEP 4: Verifikasi file yang dibuat
    # ─────────────────────────────────────────────
    print()
    print("▶ STEP 4: Memverifikasi file yang dibuat...")
    print()

    report_folder = Config.FOLDER_REPORT
    os.makedirs(report_folder, exist_ok=True)

    # Cek file Excel
    if excel_path and os.path.exists(excel_path):
        file_size = os.path.getsize(excel_path)
        print(f"✓ File Excel ada: {os.path.basename(excel_path)}")
        print(f"  Size: {file_size:,} bytes")
    else:
        print(f"✗ File Excel tidak ditemukan: {excel_path}")

    # Cek file .eml
    eml_files = [f for f in os.listdir(report_folder) if f.endswith(".eml")]
    print()
    if eml_files:
        print(f"✓ {len(eml_files)} file draft email (.eml) ditemukan:")
        for eml_file in sorted(eml_files):
            eml_path = os.path.join(report_folder, eml_file)
            file_size = os.path.getsize(eml_path)
            print(f"  - {eml_file}")
            print(f"    Size: {file_size:,} bytes")

            # Baca preview isi .eml
            try:
                with open(eml_path, "r", encoding="utf-8") as f:
                    content = f.read()
                    # Cari Subject line
                    for line in content.split("\n"):
                        if line.startswith("Subject:"):
                            print(f"    Subject: {line[8:].strip()}")
                            break
            except Exception as e:
                print(f"    (Gagal baca preview: {e})")
    else:
        print(f"✗ Tidak ada file .eml yang ditemukan di {report_folder}")

    # ─────────────────────────────────────────────
    # STEP 5: Summary dan instruksi
    # ─────────────────────────────────────────────
    print()
    print("=" * 70)
    print("✓ TEST SELESAI")
    print("=" * 70)
    print()
    print("Instruksi selanjutnya:")
    print(f"1. Buka folder report: {report_folder}")
    print("2. Double-click salah satu file .eml untuk membukanya di Thunderbird")
    print("3. Verifikasi isi email:")
    print("   - Subject sesuai dengan plant")
    print("   - Body menampilkan tabel adjustment")
    print("   - File Excel sudah di-attach")
    print("4. Jika semuanya ok, klik 'Send' untuk mengirim (atau batalkan)")
    print()

    return True


def test_draft_with_custom_email():
    """
    Test: Buat draft email dengan override email penerima.
    Berguna untuk testing dengan email development.
    """
    print()
    print("=" * 70)
    print("TEST: Draft Email dengan Email Override")
    print("=" * 70)
    print()

    print("▶ Membuat test data...")
    sample_items = [
        StockDiff(
            param="FSTKGD",
            plant="9999",
            sloc="TEST",
            posting_date="08.06.2026",
            material="TEST001",
            qty_matrix=100.000,
            qty_sap=150.000,
            diff=-50.000,
            status=1,
            mvt_type="917",
            qty_adjust=50.000,
        ),
    ]

    items_per_plant = {"9999": sample_items}

    print("▶ Menjalankan dengan email override...")
    try:
        excel_path = send_stock_diff_report(
            items_per_plant,
            override_to="test@mayora.co.id",
            override_cc="test-cc@mayora.co.id"
        )
        print("✓ Draft email dibuat dengan email override")
        print(f"  To: test@mayora.co.id")
        print(f"  CC: test-cc@mayora.co.id")
    except Exception as e:
        print(f"✗ ERROR: {e}")
        return False

    return True


if __name__ == "__main__":
    print()
    print("╔" + "═" * 68 + "╗")
    print("║" + " " * 68 + "║")
    print("║" + "  TEST DRAFT EMAIL — MOZILLA THUNDERBIRD".center(68) + "║")
    print("║" + " " * 68 + "║")
    print("╚" + "═" * 68 + "╝")
    print()

    # Test 1: Basic generation
    success1 = test_draft_email_generation()

    # Test 2: With email override
    success2 = test_draft_with_custom_email()

    # Summary
    print()
    print("╔" + "═" * 68 + "╗")
    if success1 and success2:
        print("║" + "  ✓ SEMUA TEST PASSED".ljust(68) + "║")
    else:
        print("║" + "  ✗ BEBERAPA TEST GAGAL".ljust(68) + "║")
    print("╚" + "═" * 68 + "╝")
    print()
