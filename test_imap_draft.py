# test_imap_draft.py — ganti isinya dengan ini
from email_config_ui import load_credentials
from send_email_report import _create_thunderbird_draft

cred = load_credentials()

result = _create_thunderbird_draft(
    cred      = cred,
    subject   = "[RPA TEST] Draft .eml test",
    body_html = "<p>Test draft .eml — Thunderbird harusnya buka otomatis.</p>",
    to        = cred["email_to"],
    cc        = cred.get("email_cc", ""),
)

print(f"File .eml: {result}")