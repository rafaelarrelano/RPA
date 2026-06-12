#!/usr/bin/env python3
"""
launcher_sharp.py
IMPROVED launcher with enhanced rendering for crystal clear text

This is the main entry point - use this instead of launcher.py for best results
"""

# CRITICAL: Must be first - enables high-DPI rendering
import os
os.environ['TKINTER_ENABLE_DPI'] = '1'

try:
    from ctypes import windll
    # Enable Windows 10+ high DPI awareness
    windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

import tkinter as tk
from tkinter import font as tkfont
import importlib
import sys

from theme_manager import get_theme_manager

# Get theme
theme = get_theme_manager()

# ─────────────────────────────────────────────
# CRYSTAL CLEAR FONTS
# ─────────────────────────────────────────────

# Segoe UI is the Windows native font optimized for high DPI
# These sizes are empirically tested for clarity
FONT_TINY   = ("Segoe UI", 9)
FONT_SMALL  = ("Segoe UI", 10)
FONT_MEDIUM = ("Segoe UI", 11)
FONT_LARGE  = ("Segoe UI", 12, "bold")
FONT_XLARGE = ("Segoe UI", 14, "bold")
FONT_HEADER = ("Segoe UI", 26, "bold")


# ─────────────────────────────────────────────
# MODULE REGISTRY
# ─────────────────────────────────────────────

MODULES = [
    (
        "Compare Stock",
        "Reconcile Matrix portal vs SAP stock — generate U2C & email report",
        "rpa_gui",
        "RpaGui",
        True,
    ),
    (
        "Maintain Material",
        "Add, edit, or deactivate materials in the SAP master data",
        "maintain_material",
        "MaintainMaterialGui",
        True,
    ),
    (
        "Extend Material",
        "Extend existing materials to new plants or storage locations",
        "extend_material",
        "ExtendMaterialGui",
        False,
    ),
]


# ─────────────────────────────────────────────
# SHARP LAUNCHER
# ─────────────────────────────────────────────

class SharpLauncher:
    """Main launcher with crystal clear rendering"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("RPA — PT Mayora Indah Tbk")

        # Optimal size for modern monitors
        self.root.geometry("950x720")
        self.root.minsize(850, 620)
        self.root.resizable(True, True)

        # Main container
        self._container = tk.Frame(self.root)
        self._container.pack(fill="both", expand=True)

        self._active_module_frame = None
        self._refresh_colors()
        self._build_menu()

    def _refresh_colors(self):
        """Get current theme colors"""
        self.BG          = theme.get_color("bg")
        self.BG_CARD     = theme.get_color("bg_card")
        self.BG_HOVER    = theme.get_color("bg_hover")
        self.BG_ACTIVE   = theme.get_color("bg_active")
        self.TEXT_DARK   = theme.get_color("text_dark")
        self.TEXT_MID    = theme.get_color("text_mid")
        self.TEXT_LIGHT  = theme.get_color("text_light")
        self.BORDER      = theme.get_color("border")
        self.SHADOW      = theme.get_color("shadow")

    def _build_menu(self):
        """Build main menu"""
        # Clear
        for w in self._container.winfo_children():
            w.destroy()

        self._refresh_colors()
        self.root.configure(bg=self.BG)
        self._container.configure(bg=self.BG)

        # ── Top bar with theme toggle ─────────────────────
        top_bar = tk.Frame(self._container, bg=self.BG)
        top_bar.pack(fill="x", padx=48, pady=(16, 0))

        theme_text = "🌙 Dark" if not theme.is_dark else "☀️  Light"
        tk.Button(
            top_bar,
            text=theme_text,
            font=FONT_SMALL,
            fg=self.TEXT_LIGHT,
            bg=self.BG,
            activeforeground=self.TEXT_DARK,
            activebackground=self.BG_HOVER,
            relief="flat",
            bd=0,
            cursor="hand2",
            command=self._toggle_theme,
        ).pack(side="right")

        # ── Header ────────────────────────────────────────
        header = tk.Frame(self._container, bg=self.BG)
        header.pack(fill="x", padx=48, pady=(32, 0))

        tk.Label(
            header,
            text="PT MAYORA INDAH TBK",
            font=FONT_SMALL,
            fg=self.TEXT_LIGHT,
            bg=self.BG,
        ).pack(anchor="w")

        tk.Label(
            header,
            text="RPA Automation Suite",
            font=FONT_HEADER,
            fg=self.TEXT_DARK,
            bg=self.BG,
        ).pack(anchor="w", pady=(4, 0))

        tk.Label(
            header,
            text="Select a module to get started",
            font=FONT_SMALL,
            fg=self.TEXT_MID,
            bg=self.BG,
        ).pack(anchor="w", pady=(6, 0))

        # Divider
        tk.Frame(self._container, bg=self.BORDER, height=1).pack(
            fill="x", padx=48, pady=(24, 28)
        )

        # ── Module buttons ────────────────────────────────
        btn_frame = tk.Frame(self._container, bg=self.BG)
        btn_frame.pack(fill="both", expand=True, padx=48)

        for label, subtitle, module_file, class_name, enabled in MODULES:
            self._make_button(btn_frame, label, subtitle, module_file, class_name, enabled)

        # ── Footer ────────────────────────────────────────
        footer = tk.Frame(self._container, bg=self.BG)
        footer.pack(fill="x", padx=48, pady=(20, 28))

        tk.Label(
            footer,
            text=f"v2.0  ·  {len([m for m in MODULES if m[4]])} modules active",
            font=FONT_TINY,
            fg=self.TEXT_LIGHT,
            bg=self.BG,
        ).pack(side="left")

        tk.Label(
            footer,
            text="IT Department — Mayora",
            font=FONT_TINY,
            fg=self.TEXT_LIGHT,
            bg=self.BG,
        ).pack(side="right")

    def _make_button(self, parent, label, subtitle, module_file, class_name, enabled):
        """Create module button"""
        # Shadow
        shadow = tk.Frame(parent, bg=self.SHADOW, bd=0)
        shadow.pack(fill="x", pady=(0, 12))

        # Card
        card = tk.Frame(
            shadow,
            bg=self.BG_CARD,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.BORDER,
        )
        card.pack(fill="x", padx=(0, 2), pady=(0, 2))

        # Content
        inner = tk.Frame(card, bg=self.BG_CARD, padx=32, pady=24)
        inner.pack(fill="x")
        inner.columnconfigure(0, weight=1)

        # Text
        text_frame = tk.Frame(inner, bg=self.BG_CARD)
        text_frame.pack(side="left", fill="x", expand=True)

        lbl = tk.Label(
            text_frame,
            text=label,
            font=FONT_LARGE,
            fg=self.TEXT_DARK if enabled else self.TEXT_LIGHT,
            bg=self.BG_CARD,
        )
        lbl.pack(anchor="w")

        sub = tk.Label(
            text_frame,
            text=subtitle if enabled else subtitle + "  —  Coming Soon",
            font=FONT_SMALL,
            fg=self.TEXT_MID if enabled else self.TEXT_LIGHT,
            bg=self.BG_CARD,
            wraplength=240,
        )
        sub.pack(anchor="w", pady=(4, 0))

        # Arrow
        arrow = tk.Label(
            inner,
            text="→" if enabled else "○",
            font=("Segoe UI", 18),
            fg=self.TEXT_DARK if enabled else self.TEXT_LIGHT,
            bg=self.BG_CARD,
        )
        arrow.pack(side="right", padx=(16, 0))

        if enabled:
            def on_enter(e, c=card, i=inner, tf=text_frame, l=lbl, s=sub, a=arrow):
                c.configure(bg=self.BG_HOVER, highlightbackground=self.TEXT_DARK)
                i.configure(bg=self.BG_HOVER)
                tf.configure(bg=self.BG_HOVER)
                l.configure(bg=self.BG_HOVER)
                s.configure(bg=self.BG_HOVER)
                a.configure(bg=self.BG_HOVER)

            def on_leave(e, c=card, i=inner, tf=text_frame, l=lbl, s=sub, a=arrow):
                c.configure(bg=self.BG_CARD, highlightbackground=self.BORDER)
                i.configure(bg=self.BG_CARD)
                tf.configure(bg=self.BG_CARD)
                l.configure(bg=self.BG_CARD)
                s.configure(bg=self.BG_CARD)
                a.configure(bg=self.BG_CARD)

            def on_click(e, mf=module_file, cn=class_name, lbl=label):
                self._load_module(mf, cn, lbl)

            for w in [card, inner, text_frame, lbl, sub, arrow]:
                w.bind("<Enter>", on_enter)
                w.bind("<Leave>", on_leave)
                w.bind("<Button-1>", on_click)
                w.configure(cursor="hand2")

    def _load_module(self, module_file: str, class_name: str, label: str):
        """Load module into container"""
        for w in self._container.winfo_children():
            w.destroy()

        # Back bar
        back_bar_bg = self.BG_ACTIVE if theme.is_dark else "#1A1A1A"
        back_bar_fg = "#FFFFFF"
        back_bar = tk.Frame(self._container, bg=back_bar_bg, height=44)
        back_bar.pack(fill="x")
        back_bar.pack_propagate(False)

        back_btn = tk.Label(
            back_bar,
            text="← Menu",
            font=("Segoe UI", 11, "bold"),
            fg=back_bar_fg,
            bg=back_bar_bg,
            cursor="hand2",
            padx=20,
        )
        back_btn.pack(side="left", pady=10)
        back_btn.bind("<Button-1>", lambda e: self._build_menu())
        back_btn.bind("<Enter>", lambda e: back_btn.configure(fg="#AAAAAA"))
        back_btn.bind("<Leave>", lambda e: back_btn.configure(fg=back_bar_fg))

        tk.Label(
            back_bar,
            text=label.upper(),
            font=("Segoe UI", 10),
            fg="#888888",
            bg=back_bar_bg,
        ).pack(side="left", pady=10)

        # Module frame
        module_container = tk.Frame(self._container, bg=self.BG)
        module_container.pack(fill="both", expand=True)

        try:
            if module_file in sys.modules:
                mod = importlib.reload(sys.modules[module_file])
            else:
                mod = importlib.import_module(module_file)

            cls = getattr(mod, class_name)

            try:
                instance = cls(
                    module_container,
                    back_callback=self._build_menu,
                    theme_manager=theme
                )
            except TypeError:
                instance = cls(module_container, back_callback=self._build_menu)

            self._active_module_frame = instance

        except Exception as e:
            self._show_error(module_container, str(e))

    def _show_error(self, parent, message: str):
        """Show error"""
        self._refresh_colors()
        tk.Label(
            parent,
            text="⚠  Module failed to load",
            font=("Segoe UI", 14, "bold"),
            fg=self.TEXT_DARK,
            bg=self.BG,
        ).pack(pady=(60, 12))

        tk.Label(
            parent,
            text=message,
            font=FONT_SMALL,
            fg=self.TEXT_MID,
            bg=self.BG,
            wraplength=500,
        ).pack(pady=(0, 24))

    def _toggle_theme(self):
        """Toggle dark/light mode"""
        theme.toggle_theme()
        self._refresh_colors()
        self._build_menu()


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.iconbitmap("icon.ico")
    except Exception:
        pass
    app = SharpLauncher(root)
    root.mainloop()