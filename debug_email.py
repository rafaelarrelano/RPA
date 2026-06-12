from email_config_ui import load_credentials
from send_email_report import _create_thunderbird_draft

cred = load_credentials()
print("Konfigurasi:")
print(f"  From : {cred['email_from']}")
print(f"  To   : {cred['email_to']}")
print(f"  CC   : {cred.get('email_cc', '')}")
print()
print("Membuat draft email test...")

draft_path = _create_thunderbird_draft(
    cred      = cred,
    subject   = "[RPA] DEBUG Test Email",
    body_html = "<p>Debug test dari RPA — ini adalah email draft untuk Thunderbird.</p>",
    to        = cred["email_to"],
    cc        = cred.get("email_cc", ""),
)
print(f"✓ Draft email dibuat: {draft_path}")
print("Buka file .eml di Thunderbird untuk melihat draft email.")
