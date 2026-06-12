# Perubahan Sistem Email: SMTP → Thunderbird Draft

## 📋 Ringkasan Perubahan

Sistem pengiriman email telah diubah dari **SMTP (mengirim langsung)** menjadi **Draft Thunderbird (.eml files)**.

**Sebelum:** Email dikirim otomatis via SMTP ke server email
**Sesudah:** Email disimpan sebagai file draft yang dapat dibuka di Mozilla Thunderbird untuk review sebelum mengirim

---

## 📝 File yang Diubah

### 1. **send_email_report.py** (File Utama)
- ❌ Dihapus: `_smtp_send()` - fungsi mengirim via SMTP
- ❌ Dihapus: `diagnose_smtp()` - fungsi diagnosa port SMTP
- ❌ Dihapus: Import `smtplib`, `ssl`
- ✅ Ditambah: `_create_thunderbird_draft()` - fungsi membuat file .eml
- ✅ Ditambah: Import `mimetypes`, `base64`
- ✅ Diubah: `send_stock_diff_report()` - sekarang membuat draft, bukan mengirim
- ✅ Diubah: Docstring dan log messages untuk menjelaskan proses draft

**Perubahan Logika:**
```python
# SEBELUM (SMTP)
_smtp_send(cred, subject, body, to, cc, attachment)  # Langsung kirim

# SESUDAH (Thunderbird Draft)
_create_thunderbird_draft(cred, subject, body, to, cc, attachment)  # Buat file .eml
```

**Output:** File `.eml` disimpan di folder `Config.FOLDER_REPORT` dengan nama:
```
Draft_20260608_143025_123_Req_Adj_EOD_Plant.eml
```

---

### 2. **debug_email.py** (File Test)
- ❌ Dihapus: Import `_smtp_send`
- ✅ Ditambah: Import `_create_thunderbird_draft`
- ✅ Diubah: Membuat draft email test daripada mengirim
- ✅ Diubah: Output message untuk menjelaskan cara membuka file di Thunderbird

---

### 3. **email_config_ui.py** (GUI Konfigurasi)
- ✅ Diubah: Header UI menjadi "Konfigurasi Email — Thunderbird Draft"
- ✅ Diubah: Deskripsi dari "Kredensial disimpan terenkripsi..." menjadi "Email disimpan sebagai draft..."
- ✅ Diubah: Label password field dari "Kosongkan password jika server pakai relay..." menjadi "Opsional — disimpan untuk referensi..."
- ✅ Diubah: Button "✉ Test Kirim" menjadi "✉ Test Draft"
- ✅ Diubah: Button "🔍 Diagnosa" menjadi "ℹ Info"
- ✅ Diubah: Fungsi `_on_test()` - membuat draft email test
- ✅ Diubah: Fungsi `_on_diagnose()` - menampilkan info tentang perubahan mode email

---

## 🔄 Alur Kerja Baru

### Sebelumnya (SMTP - Automatic Sending)
```
RPA berjalan
    ↓
Export data dari SAP → Compare → Analisis
    ↓
Buat laporan Excel
    ↓
Kirim email via SMTP secara otomatis
    ↓
Email langsung sampai ke inbox penerima
```

### Sekarang (Thunderbird Draft - Manual Review)
```
RPA berjalan
    ↓
Export data dari SAP → Compare → Analisis
    ↓
Buat laporan Excel
    ↓
Buat file draft email (.eml) di folder report
    ↓
User membuka file .eml di Thunderbird/email client
    ↓
User review email, bisa edit sebelum mengirim
    ↓
User klik tombol "Send" di Thunderbird
    ↓
Email terkirim
```

---

## 📂 Folder Draft Email

Semua file draft email disimpan di:
```
<Folder Report>/Draft_*.eml
```

Contoh path lengkap:
```
C:\RPA_StockRecon\Report\Draft_20260608_143025_123_Req_Adj_EOD_Plant_4502.eml
```

---

## 🚀 Cara Menggunakan

### 1. Menjalankan RPA Seperti Biasa
```bash
python main.py
```

Output di log akan menunjukkan:
```
[DRAFT] Buat draft plant 4502 → steven.pedro@mayora.co.id
[DRAFT] Subject: Req. Adj. EOD Plant 4502 PGD Surabaya Tgl 08.06.2026
[DRAFT] ✓ Draft dibuat: Draft_20260608_143025_123_Req_Adj_EOD...eml
[DRAFT] Path: C:\RPA_StockRecon\Report\Draft_20260608_143025_123_Req_Adj_EOD...eml
[REPORT] Draft email tersimpan di: C:\RPA_StockRecon\Report\
[REPORT] Buka folder dan double-click file .eml untuk membuka di Thunderbird
```

### 2. Membuka Draft Email
- Buka folder report (lihat path di log)
- Double-click file `.eml`
- Thunderbird akan membuka draft email secara otomatis
- Review isi email
- Bisa edit jika diperlukan
- Klik "Send" untuk mengirim

### 3. Mengubah Konfigurasi Email
```bash
python email_config_ui.py
```
- Gunakan tombol "✉ Test Draft" untuk membuat email test
- Gunakan tombol "ℹ Info" untuk melihat informasi mode baru

---

## ⚙️ Kredensial Email (Tetap Digunakan)
Meski sudah tidak mengirim via SMTP, kredensial email masih disimpan untuk:
- Field "From" (pengirim email)
- Field "To" (penerima email)
- Field "CC" (penerima CC)
- Field "Password" (referensi, tidak lagi digunakan)

**File konfigurasi:** `config/email_cred.enc` (terenkripsi)

---

## 📋 Checklist Implementasi

- [x] Ganti `_smtp_send()` dengan `_create_thunderbird_draft()`
- [x] Hapus import SMTP dan SSL
- [x] Tambah import `mimetypes` untuk deteksi MIME type attachment
- [x] Update `send_stock_diff_report()` untuk membuat draft
- [x] Update `debug_email.py`
- [x] Update UI di `email_config_ui.py`
- [x] Update log messages untuk menjelaskan proses draft
- [x] Pertahankan semua logika bisnis (Excel, filtering FSTKGD, attachment, dll)
- [x] Support attachment Excel di file draft

---

## ✅ Keuntungan Perubahan

1. **Kontrol Lebih Baik** - User dapat review email sebelum mengirim
2. **Menghindari Kesalahan** - Jika data salah, bisa dibatalkan
3. **Tidak Perlu Autentikasi SMTP** - Tidak bergantung konfigurasi port/AUTH server
4. **Aman** - Email tidak terkirim otomatis, mengurangi kesalahan pengirim
5. **Fleksibel** - User bisa edit subject, penerima, atau isi email di Thunderbird

---

## ⚠️ Catatan Penting

- **Fungsi bisnis tidak berubah** - Semua logika Excel, filtering, dan formatting tetap sama
- **Data attachment tetap disertakan** - File Excel laporan tetap di-attach ke draft email
- **Kompatibilitas** - File .eml kompatibel dengan semua email client (Thunderbird, Outlook, Gmail, dll)
- **Tidak ada breaking changes** - Kode yang memanggil `send_stock_diff_report()` tetap bisa digunakan tanpa perubahan

---

## 🔧 Troubleshooting

### Masalah: File .eml tidak bisa dibuka
**Solusi:**
1. Pastikan Mozilla Thunderbird sudah terinstall
2. Right-click file .eml → Open With → Pilih Thunderbird
3. Atau drag-drop file .eml ke Thunderbird

### Masalah: Attachment tidak terlihat di Thunderbird
**Solusi:**
1. File attachment harus ada di path yang disimpan
2. Check log untuk memastikan attachment path benar
3. Coba buat draft baru dengan attachment yang sama

### Masalah: Email To/CC kosong
**Solusi:**
1. Jalankan `python email_config_ui.py`
2. Isi Email To dan CC
3. Klik "Simpan"
4. Coba lagi

---

## 📞 Support

Jika ada pertanyaan atau masalah dengan perubahan ini, silakan cek:
1. File log di console
2. Folder report untuk melihat file .eml yang dibuat
3. Pastikan Thunderbird sudah terinstall

---

**Last Updated:** 8 Juni 2026
**Status:** ✅ Implementasi Selesai
