"""
theme_manager.py
Centralized theme management for dark/light mode across RPA modules
"""

# CRITICAL: Enable DPI awareness BEFORE any tkinter operations
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

import tkinter as tk
import json
import os

THEME_CONFIG_FILE = "theme_config.json"


class ThemeManager:
    """Manages dark/light mode themes across all modules"""

    # Light theme colors
    LIGHT_THEME = {
        "bg": "#F5F5F5",
        "bg_card": "#FFFFFF",
        "bg_hover": "#F0F0F0",
        "bg_active": "#1A1A1A",
        "text_dark": "#1A1A1A",
        "text_mid": "#555555",
        "text_light": "#999999",
        "text_white": "#FFFFFF",
        "border": "#E0E0E0",
        "shadow": "#D0D0D0",
        "accent": "#1A1A1A",
        "input_bg": "#FFFFFF",
        "scrollbar": "#CCCCCC",
        "success": "#2E7D32",
        "warning": "#E65100",
        "danger": "#C62828",
    }

    # Dark theme colors
    DARK_THEME = {
        "bg": "#1A1A1A",
        "bg_card": "#2A2A2A",
        "bg_hover": "#333333",
        "bg_active": "#F5F5F5",
        "text_dark": "#E8E8E8",
        "text_mid": "#A8A8A8",
        "text_light": "#666666",
        "text_white": "#1A1A1A",
        "border": "#404040",
        "shadow": "#0A0A0A",
        "accent": "#E8E8E8",
        "input_bg": "#2A2A2A",
        "scrollbar": "#444444",
        "success": "#66BB6A",
        "warning": "#FFA726",
        "danger": "#EF5350",
    }

    def __init__(self):
        self.is_dark = self._load_preference()
        self.current_theme = self.DARK_THEME if self.is_dark else self.LIGHT_THEME

    def _load_preference(self) -> bool:
        """Load saved theme preference from config file"""
        try:
            if os.path.exists(THEME_CONFIG_FILE):
                with open(THEME_CONFIG_FILE, "r") as f:
                    config = json.load(f)
                    return config.get("dark_mode", False)
        except Exception:
            pass
        return False

    def _save_preference(self):
        """Save theme preference to config file"""
        try:
            with open(THEME_CONFIG_FILE, "w") as f:
                json.dump({"dark_mode": self.is_dark}, f)
        except Exception:
            pass

    def toggle_theme(self):
        """Toggle between dark and light themes"""
        self.is_dark = not self.is_dark
        self.current_theme = self.DARK_THEME if self.is_dark else self.LIGHT_THEME
        self._save_preference()

    def get_color(self, key: str) -> str:
        """Get color value from current theme"""
        return self.current_theme.get(key, "#000000")

    def get_theme_dict(self) -> dict:
        """Get all colors from current theme"""
        return self.current_theme.copy()


# Global theme manager instance
_theme_manager = None


def get_theme_manager() -> ThemeManager:
    """Get or create the global theme manager"""
    global _theme_manager
    if _theme_manager is None:
        _theme_manager = ThemeManager()
    return _theme_manager