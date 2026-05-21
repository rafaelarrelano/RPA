"""
test_sap_connection.py
Test koneksi Python ke SAP GUI - READ ONLY, tidak ada perubahan di SAP
Jalankan: python test_sap_connection.py
"""
import win32com.client

def test_koneksi_sap():
    try:
        # 1. Coba konek ke SAP GUI yang sedang terbuka
        print("Mencoba konek ke SAP GUI...")
        sap_gui = win32com.client.GetObject("SAPGUI")
        print("SAP GUI ditemukan!")

        # 2. Ambil scripting engine
        app = sap_gui.GetScriptingEngine
        print(f"Jumlah koneksi aktif: {app.Children.Count}")

        # 3. Ambil koneksi pertama
        conn = app.Children(0)
        print(f"Koneksi: {conn.Description}")

        # 4. Ambil session pertama
        session = conn.Children(0)
        print(f"Session aktif: {session.Info.SystemName}")
        print(f"User login   : {session.Info.User}")
        print(f"T-code aktif : {session.Info.Transaction}")

        print("\nKoneksi SAP berhasil!")
        return True

    except Exception as e:
        print(f"\nGagal konek ke SAP: {e}")
        print("\nPastikan:")
        print("  1. SAP GUI sudah terbuka dan sudah login")
        print("  2. SAP GUI Scripting sudah diaktifkan")
        print("  3. Tidak ada popup/dialog yang terbuka di SAP")
        return False

if __name__ == "__main__":
    test_koneksi_sap()