#!/usr/bin/env python3
"""
Test fitur draft email untuk items lewat limit
"""
import os
import tempfile
from datetime import datetime

# Mock config
class MockConfig:
    FOLDER_DRAFT = os.path.join(tempfile.gettempdir(), "test_draft")

# Mock logger
class MockLogger:
    def info(self, msg):
        print(f"[INFO] {msg}")
    def error(self, msg):
        print(f"[ERROR] {msg}")

def save_email_draft_test(plant: str, subject: str, body_html: str, 
                         email_to: str, email_cc: str = "") -> str:
    """Test version of save_email_draft"""
    
    os.makedirs(MockConfig.FOLDER_DRAFT, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename_base = f"Draft_Plant_{plant}_{timestamp}"
    
    # Save HTML body
    html_path = os.path.join(MockConfig.FOLDER_DRAFT, f"{filename_base}.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(body_html)
    
    # Save metadata
    meta_path = os.path.join(MockConfig.FOLDER_DRAFT, f"{filename_base}.txt")
    meta_info = f"""DRAFT EMAIL — LEWAT LIMIT ADJUSTMENT
{'='*60}
Tanggal: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}
Plant: {plant}
Status: DISIMPAN (Tidak Dikirim)

TO: {email_to}
CC: {email_cc if email_cc else '(tidak ada)'}
Subject: {subject}

{'='*60}
File HTML: {os.path.basename(html_path)}
Untuk kirim: buka di browser, copy isi, paste di email client
{'='*60}
"""
    
    with open(meta_path, "w", encoding="utf-8") as f:
        f.write(meta_info)
    
    return html_path


def test_save_draft():
    """Test save draft functionality"""
    
    # Sample data
    plant = "4503"
    subject = "Req. Adj. EOD Plant 4503 PT Sukses Tgl 01.04.2026"
    body_html = """
    <html>
    <body>
    <h2>Laporan Adjustment Stok</h2>
    <table border="1">
      <tr><th>Material</th><th>Selisih</th></tr>
      <tr><td>377001</td><td>1.000</td></tr>
    </table>
    </body>
    </html>
    """
    email_to = "test@example.com"
    email_cc = "cc@example.com"
    
    # Test
    print("Testing save_email_draft()...")
    html_path = save_email_draft_test(plant, subject, body_html, email_to, email_cc)
    
    # Verify files exist
    meta_path = html_path.replace(".html", ".txt")
    
    assert os.path.exists(html_path), f"HTML file not created: {html_path}"
    assert os.path.exists(meta_path), f"Meta file not created: {meta_path}"
    
    print(f"✓ HTML file created: {html_path}")
    print(f"✓ Meta file created: {meta_path}")
    
    # Check content
    with open(html_path) as f:
        html_content = f.read()
        assert "Material" in html_content
        print(f"✓ HTML content valid")
    
    with open(meta_path) as f:
        meta_content = f.read()
        assert "DRAFT EMAIL" in meta_content
        assert plant in meta_content
        assert subject in meta_content
        print(f"✓ Meta content valid")
    
    # Print meta content for inspection
    print("\n--- META CONTENT ---")
    print(meta_content)
    
    # Cleanup
    os.remove(html_path)
    os.remove(meta_path)
    print("\n✓ Test passed!")


if __name__ == "__main__":
    test_save_draft()
