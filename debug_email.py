from email_config_ui import load_credentials
from send_email_report import _smtp_send

cred = load_credentials()
print("Konfigurasi:")
print(f"  Host : {cred['smtp_host']}")
print(f"  Port : {cred['smtp_port']}")
print(f"  From : {cred['email_from']}")
print(f"  To   : {cred['email_to']}")
print(f"  CC   : {cred.get('email_cc', '')}")
print()
print("Mengirim email test...")

_smtp_send(
    cred      = cred,
    subject   = "[RPA] DEBUG Test Email",
    body_html = "<p>Debug test dari RPA.</p>",
    to        = cred["email_to"],
    cc        = cred.get("email_cc", ""),
)
print("Selesai — tidak ada error!")
