# jalankan ini
import socket
host = "mail.mayora.co.id"
for port in [993, 143]:
    try:
        s = socket.create_connection((host, port), timeout=3)
        s.close()
        print(f"Port {port}: TERBUKA ✓")
    except Exception as e:
        print(f"Port {port}: GAGAL — {e}")