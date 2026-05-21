"""
test_rpa.py - Unit tests untuk semua fungsi parsing & logika
Jalankan: pytest test_rpa.py -v
"""
import pytest
from main import (
    parse_decimal, convert_date, parse_line,
    get_valid_items, set_movement_type, format_qty_for_sap,
    StockDiff
)


# ─────────────────────────────────────────────
# TEST KONVERSI DESIMAL
# ─────────────────────────────────────────────

class TestParseDecimal:
    def test_koma_ke_titik(self):
        assert parse_decimal("6,375") == 6.375

    def test_negatif_koma(self):
        assert parse_decimal("-0,001") == -0.001

    def test_integer_tanpa_desimal(self):
        assert parse_decimal("65") == 65.0

    def test_nilai_besar(self):
        assert parse_decimal("2145,917") == 2145.917

    def test_nol(self):
        assert parse_decimal("0") == 0.0

    def test_sudah_pakai_titik(self):
        # Kalau portal sudah pakai titik, tidak boleh error
        assert parse_decimal("6.375") == 6.375


# ─────────────────────────────────────────────
# TEST KONVERSI TANGGAL
# ─────────────────────────────────────────────

class TestConvertDate:
    def test_format_normal(self):
        assert convert_date("20260401") == "01.04.2026"

    def test_awal_tahun(self):
        assert convert_date("20260101") == "01.01.2026"

    def test_akhir_tahun(self):
        assert convert_date("20261231") == "31.12.2026"

    def test_format_salah_raises(self):
        with pytest.raises(ValueError):
            convert_date("2026040")   # kurang 1 digit

    def test_format_kosong_raises(self):
        with pytest.raises(ValueError):
            convert_date("")


# ─────────────────────────────────────────────
# TEST MOVEMENT TYPE
# ─────────────────────────────────────────────

class TestMovementType:
    def _make_item(self, diff):
        return StockDiff("FSTKGD","4503","WH01","01.04.2026","377001",65,66,diff,0)

    def test_negatif_pakai_917(self):
        item = self._make_item(-1.0)
        result = set_movement_type(item)
        assert result is True
        assert item.mvt_type == "917"
        assert item.qty_adjust == 1.0

    def test_positif_pakai_918(self):
        item = self._make_item(5.0)
        result = set_movement_type(item)
        assert result is True
        assert item.mvt_type == "918"
        assert item.qty_adjust == 5.0

    def test_nol_di_skip(self):
        item = self._make_item(0.0)
        result = set_movement_type(item)
        assert result is False
        assert item.mvt_type == ""

    def test_desimal_kecil_negatif(self):
        item = self._make_item(-0.001)
        set_movement_type(item)
        assert item.mvt_type == "917"
        assert abs(item.qty_adjust - 0.001) < 1e-9

    def test_desimal_kecil_positif(self):
        item = self._make_item(0.417)
        set_movement_type(item)
        assert item.mvt_type == "918"


# ─────────────────────────────────────────────
# TEST PARSE LINE
# ─────────────────────────────────────────────

class TestParseLine:
    VALID_LINE = "FSTKGD|4503|WH01|20260401|377001|65|66|-1|0"

    def test_happy_path(self):
        item = parse_line(self.VALID_LINE)
        assert item is not None
        assert item.plant        == "4503"
        assert item.sloc         == "WH01"
        assert item.material     == "377001"
        assert item.posting_date == "01.04.2026"
        assert item.qty_matrix   == 65.0
        assert item.qty_sap      == 66.0
        assert item.diff         == -1.0
        assert item.status       == 0

    def test_desimal_koma(self):
        line = "FSTKGD|4503|WH01|20260401|378020|6,375|36,375|-30|0"
        item = parse_line(line)
        assert item.qty_matrix == 6.375
        assert item.qty_sap    == 36.375
        assert item.diff       == -30.0

    def test_baris_kosong(self):
        assert parse_line("") is None
        assert parse_line("   ") is None

    def test_kurang_field(self):
        assert parse_line("FSTKGD|4503|WH01|20260401|377001|65|66|-1") is None

    def test_lebih_field(self):
        assert parse_line("FSTKGD|4503|WH01|20260401|377001|65|66|-1|0|EXTRA") is None


# ─────────────────────────────────────────────
# TEST GET VALID ITEMS (FILTER)
# ─────────────────────────────────────────────

class TestGetValidItems:
    def test_filter_status_bukan_0(self):
        text = "FSTKGD|4503|WH01|20260401|377001|65|66|-1|1\n"  # status=1
        items = get_valid_items(text)
        assert len(items) == 0

    def test_filter_selisih_nol(self):
        text = "FSTKGD|4503|WH01|20260401|377001|65|65|0|0\n"  # diff=0
        items = get_valid_items(text)
        assert len(items) == 0

    def test_multi_baris_valid(self):
        text = (
            "FSTKGD|4503|WH01|20260401|377001|65|66|-1|0\n"
            "FSTKGD|4503|WH01|20260401|377002|159,917|159,917|0|0\n"   # skip: diff=0
            "FSTKGD|4503|WH01|20260401|378020|6,375|36,375|-30|0\n"
        )
        items = get_valid_items(text)
        assert len(items) == 2
        assert items[0].material == "377001"
        assert items[1].material == "378020"

    def test_semua_terfilter(self):
        text = "FSTKGD|4503|WH01|20260401|377001|65|65|0|1\n"
        items = get_valid_items(text)
        assert len(items) == 0

    def test_teks_kosong(self):
        items = get_valid_items("")
        assert len(items) == 0


# ─────────────────────────────────────────────
# TEST FORMAT QTY UNTUK SAP
# ─────────────────────────────────────────────

class TestFormatQtyForSAP:
    def test_integer(self):
        assert format_qty_for_sap(1.0) == "1"

    def test_desimal_3_digit(self):
        assert format_qty_for_sap(6.375) == "6.375"

    def test_desimal_trailing_zero(self):
        # 30.0 harus jadi "30" bukan "30.000"
        assert format_qty_for_sap(30.0) == "30"

    def test_desimal_kecil(self):
        assert format_qty_for_sap(0.001) == "0.001"

    def test_nilai_besar(self):
        assert format_qty_for_sap(2145.917) == "2145.917"
