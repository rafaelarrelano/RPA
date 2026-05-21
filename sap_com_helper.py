"""
sap_com_helper.py
Helper untuk SAP GUI Scripting via COM (win32com).
Gunakan ini sebagai pengganti pyautogui untuk navigasi T-code.
Import fungsi ini di main.py dan rpa_phase1_2.py.
"""
import time
import win32com.client
import win32gui
import win32con

_sap_session_cache = None


def get_sap_session(force_refresh: bool = False):
    """
    Ambil SAP session aktif via COM scripting.
    Cache session agar tidak reconnect tiap kali.
    """
    global _sap_session_cache
    if _sap_session_cache and not force_refresh:
        try:
            # Test apakah session masih valid
            _ = _sap_session_cache.Info.Transaction
            return _sap_session_cache
        except Exception:
            _sap_session_cache = None

    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app     = sap_gui.GetScriptingEngine
        conn    = app.Children(0)
        session = conn.Children(0)
        _sap_session_cache = session
        return session
    except Exception as e:
        raise Exception(
            f"Gagal konek SAP via COM: {e}\n"
            "Pastikan:\n"
            "  1. SAP GUI sudah terbuka dan login\n"
            "  2. SAP GUI Scripting diaktifkan:\n"
            "     SAP → ikon ⚙ (pojok kanan atas) → Options → "
            "Accessibility & Scripting → Scripting → centang Enable Scripting"
        )


def sap_tcode(tcode: str, session=None, wait: float = 2.0):
    """
    Navigasi ke T-code via SAP COM scripting.
    Jauh lebih reliable dibanding pyautogui + Ctrl+/.
    
    Fallback otomatis:
    1. session.StartTransaction(tcode)   ← paling clean
    2. Isi command bar via findById      ← fallback
    """
    if session is None:
        session = get_sap_session()

    # Fokus window SAP dulu
    focus_sap()
    time.sleep(0.3)

    # Cara 1: StartTransaction
    try:
        session.StartTransaction(tcode)
        time.sleep(wait)
        return
    except Exception:
        pass

    # Cara 2: isi field command bar via scripting
    try:
        cmd_bar = session.findById("wnd[0]/tbar[0]/okcd[0]")
        cmd_bar.text = f"/{tcode}"
        session.findById("wnd[0]").sendVKey(0)  # VKey 0 = Enter
        time.sleep(wait)
        return
    except Exception as e:
        raise Exception(
            f"Gagal navigasi T-code {tcode!r} via COM: {e}\n"
            "Pastikan SAP GUI Scripting aktif."
        )


def focus_sap(title_keyword: str = "SAP"):
    """Fokuskan window SAP ke depan via win32gui."""
    SAP_KEYWORDS = [
        "SAP Easy Access", "SAP R/3", "SAP NetWeaver",
        "Program transfer", "MIGO", "ZPGD",
        "Display Material", "Goods Issue",
    ]
    SKIP = ["Chrome", "Firefox", "Edge", "Visual Studio",
            "Notepad", "Code", "Claude", "Thunderbird"]

    priority = []
    fallback = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return
        for kw in SAP_KEYWORDS:
            if kw in title:
                priority.append(hwnd)
                return
        if title_keyword in title and not any(s in title for s in SKIP):
            fallback.append(hwnd)

    win32gui.EnumWindows(cb, None)

    hwnd = priority[0] if priority else (fallback[0] if fallback else None)
    if not hwnd:
        raise Exception(f"SAP window tidak ditemukan!")

    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.6)
    return hwnd


def sap_send_vkey(vkey: int, session=None):
    """
    Kirim virtual key ke SAP.
    VKey umum: 0=Enter, 8=F8(Execute), 3=F3(Back), 15=F15(Save/Ctrl+S)
    """
    if session is None:
        session = get_sap_session()
    session.findById("wnd[0]").sendVKey(vkey)


def sap_set_field(field_id: str, value: str, session=None):
    """
    Set nilai field SAP via scripting ID.
    Contoh field_id: "wnd[0]/usr/ctxtMATNR-LOW"
    """
    if session is None:
        session = get_sap_session()
    session.findById(field_id).text = str(value)


def sap_press_button(button_id: str, session=None):
    """Klik tombol SAP via scripting ID."""
    if session is None:
        session = get_sap_session()
    session.findById(button_id).press()


def sap_get_statusbar(session=None) -> str:
    """Ambil teks status bar SAP (untuk baca doc number setelah posting)."""
    if session is None:
        session = get_sap_session()
    try:
        return session.findById("wnd[0]/sbar").text
    except Exception:
        return ""


# ─────────────────────────────────────────────
# TEST
# ─────────────────────────────────────────────
if __name__ == "__main__":
    import re

    print("Test SAP COM Helper")
    print("=" * 40)

    print("[1] Konek ke SAP...")
    try:
        sess = get_sap_session()
        print(f"  System  : {sess.Info.SystemName}")
        print(f"  User    : {sess.Info.User}")
        print(f"  T-code  : {sess.Info.Transaction}")
        print("  Koneksi OK!")
    except Exception as e:
        print(f"  GAGAL: {e}")
        exit()

    print("\n[2] Test navigasi T-code dalam 3 detik...")
    time.sleep(3)

    try:
        sap_tcode("ZPGD_SAPSTK", sess)
        print("  Navigasi OK! Cek SAP.")
    except Exception as e:
        print(f"  GAGAL: {e}")

    print("\n[3] Status bar:", sap_get_statusbar(sess))
