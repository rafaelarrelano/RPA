"""
launcher.py
Main entry point for RPA Stock Reconciliation System
PT Mayora Indah Tbk

Run: python launcher.py

Menu-based launcher — one window, content swaps dynamically.
Each module loads inside the same window and can return to this menu.
"""

import tkinter as tk
from tkinter import font as tkfont
import importlib
import sys
import os

# ─────────────────────────────────────────────
# ENABLE DPI AWARENESS - MUST BE FIRST
# ─────────────────────────────────────────────

try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

from theme_manager import get_theme_manager


# ─────────────────────────────────────────────
# THEME MANAGER
# ─────────────────────────────────────────────

theme = get_theme_manager()


# ─────────────────────────────────────────────
# CONSTANTS - CRISP FONTS
# ─────────────────────────────────────────────

# Use Segoe UI for crisp, clear rendering on Windows
FONT_TITLE  = ("Segoe UI", 12, "bold")
FONT_BTN    = ("Segoe UI", 11, "bold")
FONT_SUB    = ("Segoe UI", 10)
FONT_HEADER = ("Segoe UI", 24, "bold")
FONT_SMALL  = ("Segoe UI", 9)


# ─────────────────────────────────────────────
# MODULE REGISTRY
# Each entry: (label, subtitle, module_file, class_name, enabled)
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
        True,
    ),
]


# ─────────────────────────────────────────────
# LAUNCHER CLASS
# ─────────────────────────────────────────────

class Launcher:
    def __init__(self, root: tk.Tk):
        self.root        = root
        self.root.title("RPA — PT Mayora Indah Tbk")
        self.root.geometry("920x700")
        self.root.minsize(820, 600)
        self.root.resizable(True, True)

        # Configure for sharp rendering
        self.root.tk.call('tk', 'scaling', 2.0)

        # Container that holds either the menu OR a loaded module
        self._container = tk.Frame(self.root)
        self._container.pack(fill="both", expand=True)

        # Currently loaded module frame (if any)
        self._active_module_frame = None
        self._theme_toggle_btn = None

        self._refresh_colors()
        self._build_menu()

    def _refresh_colors(self):
        """Update all colors from current theme"""
        self.BG          = theme.get_color("bg")
        self.BG_CARD     = theme.get_color("bg_card")
        self.BG_HOVER    = theme.get_color("bg_hover")
        self.BG_ACTIVE   = theme.get_color("bg_active")
        self.TEXT_DARK   = theme.get_color("text_dark")
        self.TEXT_MID    = theme.get_color("text_mid")
        self.TEXT_LIGHT  = theme.get_color("text_light")
        self.TEXT_WHITE  = theme.get_color("text_white")
        self.BORDER      = theme.get_color("border")
        self.SHADOW      = theme.get_color("shadow")
        self.ACCENT      = theme.get_color("accent")

    # ── BUILD MENU ───────────────────────────────────────────

    def _build_menu(self):
        """Build the main menu view."""
        # Clear container
        for w in self._container.winfo_children():
            w.destroy()
        self._active_module_frame = None

        self._refresh_colors()
        self.root.configure(bg=self.BG)
        self._container.configure(bg=self.BG)

        # ── Header with Theme Toggle ─────────────────────────
        header_top = tk.Frame(self._container, bg=self.BG)
        header_top.pack(fill="x", padx=48, pady=(16, 0))

        # Theme toggle button (top-right)
        theme_text = "🌙 Dark" if not theme.is_dark else "☀️  Light"
        self._theme_toggle_btn = tk.Button(
            header_top,
            text=theme_text,
            font=FONT_SUB,
            fg=self.TEXT_LIGHT,
            bg=self.BG,
            activeforeground=self.TEXT_DARK,
            activebackground=self.BG_HOVER,
            relief="flat",
            bd=0,
            cursor="hand2",
            command=self._toggle_theme,
        )
        self._theme_toggle_btn.pack(side="right")

        # ── Header ───────────────────────────────────────────
        header = tk.Frame(self._container, bg=self.BG)
        header.pack(fill="x", padx=48, pady=(28, 0))

        # Top label
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
        ).pack(anchor="w", pady=(2, 0))

        tk.Label(
            header,
            text="Select a module to get started",
            font=FONT_SUB,
            fg=self.TEXT_MID,
            bg=self.BG,
        ).pack(anchor="w", pady=(4, 0))

        # Divider
        tk.Frame(
            self._container,
            bg=self.BORDER,
            height=1,
        ).pack(fill="x", padx=48, pady=(20, 24))

        # ── Module Buttons ────────────────────────────────────
        btn_frame = tk.Frame(self._container, bg=self.BG)
        btn_frame.pack(fill="both", expand=True, padx=48)

        for label, subtitle, module_file, class_name, enabled in MODULES:
            self._make_module_button(
                btn_frame, label, subtitle,
                module_file, class_name, enabled
            )

        # ── Footer ────────────────────────────────────────────
        footer = tk.Frame(self._container, bg=self.BG)
        footer.pack(fill="x", padx=48, pady=(16, 28))

        tk.Label(
            footer,
            text=f"v2.0  ·  {len([m for m in MODULES if m[4]])} modules active",
            font=FONT_SMALL,
            fg=self.TEXT_LIGHT,
            bg=self.BG,
        ).pack(side="left")

        tk.Label(
            footer,
            text="IT Department — Mayora",
            font=FONT_SMALL,
            fg=self.TEXT_LIGHT,
            bg=self.BG,
        ).pack(side="right")

    def _make_module_button(self, parent, label, subtitle,
                            module_file, class_name, enabled):
        """Create a single large rounded module button card."""

        # Outer shadow layer
        shadow = tk.Frame(parent, bg=self.SHADOW, bd=0)
        shadow.pack(fill="x", pady=(0, 10))

        # Card (main button area)
        card = tk.Frame(
            shadow,
            bg=self.BG_CARD,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.BORDER,
            highlightcolor=self.BORDER,
        )
        card.pack(fill="x", padx=(0, 2), pady=(0, 2))

        # Inner content row
        inner = tk.Frame(card, bg=self.BG_CARD, padx=28, pady=20)
        inner.pack(fill="x")
        inner.columnconfigure(0, weight=1)

        # Left side: text
        text_frame = tk.Frame(inner, bg=self.BG_CARD)
        text_frame.pack(side="left", fill="x", expand=True)

        lbl = tk.Label(
            text_frame,
            text=label,
            font=FONT_BTN,
            fg=self.TEXT_DARK if enabled else self.TEXT_LIGHT,
            bg=self.BG_CARD,
            anchor="w",
        )
        lbl.pack(anchor="w")

        sub = tk.Label(
            text_frame,
            text=subtitle if enabled else subtitle + "  —  Coming Soon",
            font=FONT_SUB,
            fg=self.TEXT_MID if enabled else self.TEXT_LIGHT,
            bg=self.BG_CARD,
            anchor="w",
        )
        sub.pack(anchor="w", pady=(2, 0))

        # Right side: arrow indicator
        arrow = tk.Label(
            inner,
            text="→" if enabled else "○",
            font=("Georgia", 16),
            fg=self.TEXT_DARK if enabled else self.TEXT_LIGHT,
            bg=self.BG_CARD,
        )
        arrow.pack(side="right", padx=(12, 0))

        # ── Hover + Click effects ─────────────────────────────
        all_widgets = [card, inner, text_frame, lbl, sub, arrow, shadow]

        if enabled:
            def on_enter(e, c=card, i=inner, tf=text_frame,
                         l=lbl, s=sub, a=arrow):
                c.configure(bg=self.BG_HOVER, highlightbackground=self.TEXT_DARK)
                i.configure(bg=self.BG_HOVER)
                tf.configure(bg=self.BG_HOVER)
                l.configure(bg=self.BG_HOVER)
                s.configure(bg=self.BG_HOVER)
                a.configure(bg=self.BG_HOVER)

            def on_leave(e, c=card, i=inner, tf=text_frame,
                         l=lbl, s=sub, a=arrow):
                c.configure(bg=self.BG_CARD, highlightbackground=self.BORDER)
                i.configure(bg=self.BG_CARD)
                tf.configure(bg=self.BG_CARD)
                l.configure(bg=self.BG_CARD)
                s.configure(bg=self.BG_CARD)
                a.configure(bg=self.BG_CARD)

            def on_click(e, mf=module_file, cn=class_name, lbl=label):
                self._load_module(mf, cn, lbl)

            for w in [card, inner, text_frame, lbl, sub, arrow]:
                w.bind("<Enter>",   on_enter)
                w.bind("<Leave>",   on_leave)
                w.bind("<Button-1>", on_click)
                w.configure(cursor="hand2")

    # ── LOAD MODULE ──────────────────────────────────────────

    def _load_module(self, module_file: str, class_name: str, label: str):
        """
        Dynamically load a module and display it in the container.
        The module's GUI class receives a special `back_callback` so it
        can return to this menu.
        """
        # Clear the container
        for w in self._container.winfo_children():
            w.destroy()

        # ── Back bar at the top ───────────────────────────────
        back_bar_bg = theme.get_color("bg_active") if theme.is_dark else "#1A1A1A"
        back_bar_fg = theme.get_color("text_white") if theme.is_dark else "#FFFFFF"
        back_bar = tk.Frame(self._container, bg=back_bar_bg, height=40)
        back_bar.pack(fill="x")
        back_bar.pack_propagate(False)

        back_btn = tk.Label(
            back_bar,
            text="← Menu",
            font=("Courier", 10, "bold"),
            fg=back_bar_fg,
            bg=back_bar_bg,
            cursor="hand2",
            padx=18,
        )
        back_btn.pack(side="left", pady=8)
        back_btn.bind("<Button-1>", lambda e: self._build_menu())
        back_btn.bind("<Enter>",
                      lambda e: back_btn.configure(fg="#AAAAAA"))
        back_btn.bind("<Leave>",
                      lambda e: back_btn.configure(fg=back_bar_fg))

        # Module name in back bar
        tk.Label(
            back_bar,
            text=label.upper(),
            font=("Courier", 9),
            fg="#888888",
            bg=back_bar_bg,
        ).pack(side="left", pady=8)

        # ── Module frame area ─────────────────────────────────
        module_container = tk.Frame(self._container, bg=self.BG)
        module_container.pack(fill="both", expand=True)

        # Try to import and instantiate the module
        try:
            # Reload in case file changed
            if module_file in sys.modules:
                mod = importlib.reload(sys.modules[module_file])
            else:
                mod = importlib.import_module(module_file)

            cls = getattr(mod, class_name)

            # Instantiate — pass back_callback and theme_manager
            try:
                instance = cls(
                    module_container,
                    back_callback=self._build_menu,
                    theme_manager=theme
                )
            except TypeError:
                # Module class doesn't accept theme_manager — try without it
                instance = cls(module_container, back_callback=self._build_menu)

            self._active_module_frame = instance

        except ImportError as e:
            self._show_module_error(
                module_container,
                f"Module file '{module_file}.py' not found.\n\n{e}"
            )
        except Exception as e:
            self._show_module_error(module_container, str(e))

    def _show_module_error(self, parent, message: str):
        """Show a friendly error if a module fails to load."""
        self._refresh_colors()
        tk.Label(
            parent,
            text="⚠  Module failed to load",
            font=("Georgia", 13, "bold"),
            fg=self.TEXT_DARK,
            bg=self.BG,
        ).pack(pady=(60, 8))

        tk.Label(
            parent,
            text=message,
            font=FONT_SUB,
            fg=self.TEXT_MID,
            bg=self.BG,
            wraplength=480,
            justify="center",
        ).pack(pady=(0, 24))

    def _toggle_theme(self):
        """Toggle between dark and light mode"""
        theme.toggle_theme()
        self._refresh_colors()
        self._build_menu()


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.iconbitmap("icon.ico")
    except Exception:
        pass
    Launcher(root)
    root.mainloop()