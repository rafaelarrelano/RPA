import rpa_phase1_2


def test_run_zpgd_sapstk_uses_new_multiple_selection_flow(monkeypatch, tmp_path):
    calls = []
    sleep_calls = []
    sample_file = tmp_path / "4502_SAPSTK.TXT"
    sample_file.write_text(
        "FSTKGD|4502|WH01|20260729|377001|100,000\n",
        encoding="utf-8",
    )

    monkeypatch.setattr(rpa_phase1_2, "sap_tcode", lambda tcode: calls.append(("sap_tcode", tcode)))
    monkeypatch.setattr(rpa_phase1_2, "_wait_sap_window", lambda *args, **kwargs: 123)
    monkeypatch.setattr(rpa_phase1_2, "_wait_sapstk_file", lambda *args, **kwargs: str(sample_file))
    monkeypatch.setattr(rpa_phase1_2.win32gui, "ShowWindow", lambda *args, **kwargs: None)
    monkeypatch.setattr(rpa_phase1_2.win32gui, "SetForegroundWindow", lambda *args, **kwargs: None)
    monkeypatch.setattr(rpa_phase1_2, "_interruptible_sleep", lambda seconds: sleep_calls.append(seconds))
    monkeypatch.setattr(rpa_phase1_2, "_copy_text_to_clipboard", lambda text: calls.append(("clipboard", text)))
    monkeypatch.setattr(rpa_phase1_2.pyautogui, "press", lambda key: calls.append(("press", key)))
    monkeypatch.setattr(rpa_phase1_2.pyautogui, "hotkey", lambda *keys: calls.append(("hotkey", keys)))
    monkeypatch.setattr(rpa_phase1_2, "type_field", lambda value: calls.append(("type_field", value)))
    monkeypatch.setattr(rpa_phase1_2, "tab_to", lambda n, delay=0.15: calls.append(("tab_to", n)))

    result = rpa_phase1_2.run_zpgd_sapstk(["4502", "4504"], send_log=lambda *args: None)

    assert result.endswith("4502_SAPSTK.TXT")
    assert ("type_field", "2") in calls
    assert ("tab_to", 3) in calls
    assert ("press", "enter") in calls
    assert ("hotkey", ("shift", "f4")) in calls
    assert ("hotkey", ("shift", "f12")) in calls
    assert calls.count(("press", "f8")) >= 2
    assert ("clipboard", "4502\r\n4504") in calls

    assert sleep_calls.count(1.5) >= 4
    assert 0.4 in sleep_calls


def test_parse_sapstk_file_preserves_plant_in_keys(tmp_path):
    sample = (
        "FSTKGD|4502|WH01|20260729|377001|233,333\n"
        "FSTKGD|4503|WH01|20260729|377002|100,000\n"
    )
    file_path = tmp_path / "sapstk.txt"
    file_path.write_text(sample, encoding="utf-8")

    result = rpa_phase1_2.parse_sapstk_file(str(file_path))

    assert result[("4502", "377001", "WH01", "FSTKGD")] == 233.333
    assert result[("4503", "377002", "WH01", "FSTKGD")] == 100.0
