import win32gui

def check_windows():
    windows = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                windows.append((hwnd, title))
    win32gui.EnumWindows(cb, None)
    
    print("Semua window yang terbuka:")
    for hwnd, title in windows:
        print(f"  [{hwnd}] {title!r}")
    
    print("\nWindow yang mengandung 'SAP':")
    for hwnd, title in windows:
        if 'SAP' in title:
            print(f"  [{hwnd}] {title!r}")

check_windows()
