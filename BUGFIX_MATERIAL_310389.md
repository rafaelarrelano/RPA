## MASALAH & SOLUSI: Material 310389 Lewat Limit Tapi Tetap Dikirim Email/U2C

### AKAR PENYEBAB MASALAH

File Excel limit adjustment (`List Limit Adj. Material SAP.xlsx`) di kolom material ada beberapa row 
dengan data **invalid** (berisi whitespace/spasi alih-alih material code).

#### Flow Masalah:
1. `load_limit_adjustment()` mencoba parse row dengan material = '      ' (spasi)
2. Saat convert ke int: `int('      ')` → **ValueError exception**
3. Exception di-catch di `test_compare.py` line 1317-1318
4. `limits = {}` (dict kosong) → semua material "tidak ditemukan" dalam limit
5. Dengan `limits` kosong, logic `is_within_limit()` return `True` untuk ALL materials
6. **Hasil: Material 310389 dianggap "lolos limit" dan dikirim email + U2C**

### DETAIL MATERIAL 310389

```
Material:      310389 (ROMA MALKIST CHOCOLATE 12BDX10SCX18G)
Row di Excel:  Baris 2425
Limit Plus:    0.008
Limit Minus:   -0.01

Dari attachment email user:
- Selisih di-report: 0.017 (positif)
- Karena 0.017 > 0.008 (limit_plus) → SEHARUSNYA LEWAT BATAS
```

Tapi karena file limit tidak di-load dengan benar → material 310389 tidak di-check terhadap limit 
→ dikirim email & U2C padahal SEHARUSNYA SKIP.

### SOLUSI (SUDAH DITERAPKAN)

Edit `limit_adjustment.py` line 27-55 untuk:
- Check jika row[1] hanya whitespace → skip
- Tambah try-except untuk ValueError saat convert ke int
- Skip invalid row daripada raise exception
- Add debug log untuk track row yang di-skip

Sekarang:
✓ File limit berhasil di-load: **2709 material** (tanpa error)
✓ Material 310389 ter-recognize dengan benar
✓ Material 310389 diff=0.017 → return False (LEWAT BATAS) → SKIP dari email & U2C

### VERIFIKASI

Setelah fix, test menunjukkan:
- load_limit_adjustment() berhasil load 2709 material tanpa exception
- Material 310389 dengan diff=0.017 → is_within_limit() return False (LEWAT BATAS)
- Material 310389 akan masuk ke items_skip → TIDAK dikirim email & U2C

### CATATAN

Ada minor bug di docstring `is_within_limit()` line 71-75 (comment tidak match dengan return value),
tapi logic-nya sudah benar. Comment hanya dokumentasi saja.

---

File yang di-modify:
- `limit_adjustment.py` (line 27-55): Add error handling & validation
