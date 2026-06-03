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
# CONSTANTS
# ─────────────────────────────────────────────

BG          = "#F5F5F5"       # off-white background
BG_CARD     = "#FFFFFF"       # white card/button background
BG_HOVER    = "#F0F0F0"       # hover state
BG_ACTIVE   = "#1A1A1A"       # active/pressed
TEXT_DARK   = "#1A1A1A"       # primary text
TEXT_MID    = "#555555"       # secondary text
TEXT_LIGHT  = "#999999"       # hint text
TEXT_WHITE  = "#FFFFFF"       # white text
BORDER      = "#E0E0E0"       # card border
SHADOW      = "#D0D0D0"       # shadow effect
ACCENT      = "#1A1A1A"       # accent (black)

FONT_TITLE  = ("Georgia", 13, "bold")
FONT_BTN    = ("Georgia", 12, "bold")
FONT_SUB    = ("Courier", 9)
FONT_HEADER = ("Georgia", 18, "bold")
FONT_SMALL  = ("Courier", 8)


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
        False,   # Coming soon
    ),
]


# ─────────────────────────────────────────────
# LAUNCHER CLASS
# ─────────────────────────────────────────────

class Launcher:
    def __init__(self, root: tk.Tk):
        self.root        = root
        self.root.title("RPA — PT Mayora Indah Tbk")
        self.root.geometry("680x580")
        self.root.minsize(560, 460)
        self.root.configure(bg=BG)
        self.root.resizable(True, True)

        # Container that holds either the menu OR a loaded module
        self._container = tk.Frame(self.root, bg=BG)
        self._container.pack(fill="both", expand=True)

        # Currently loaded module frame (if any)
        self._active_module_frame = None

        self._build_menu()

    # ── BUILD MENU ───────────────────────────────────────────

    def _build_menu(self):
        """Build the main menu view."""
        # Clear container
        for w in self._container.winfo_children():
            w.destroy()
        self._active_module_frame = None

        # ── Header ───────────────────────────────────────────
        header = tk.Frame(self._container, bg=BG)
        header.pack(fill="x", padx=48, pady=(44, 0))

        # Top label
        tk.Label(
            header,
            text="PT MAYORA INDAH TBK",
            font=FONT_SMALL,
            fg=TEXT_LIGHT,
            bg=BG,
        ).pack(anchor="w")

        tk.Label(
            header,
            text="RPA Automation Suite",
            font=FONT_HEADER,
            fg=TEXT_DARK,
            bg=BG,
        ).pack(anchor="w", pady=(2, 0))

        tk.Label(
            header,
            text="Select a module to get started",
            font=FONT_SUB,
            fg=TEXT_MID,
            bg=BG,
        ).pack(anchor="w", pady=(4, 0))

        # Divider
        tk.Frame(
            self._container,
            bg=BORDER,
            height=1,
        ).pack(fill="x", padx=48, pady=(20, 24))

        # ── Module Buttons ────────────────────────────────────
        btn_frame = tk.Frame(self._container, bg=BG)
        btn_frame.pack(fill="both", expand=True, padx=48)

        for label, subtitle, module_file, class_name, enabled in MODULES:
            self._make_module_button(
                btn_frame, label, subtitle,
                module_file, class_name, enabled
            )

        # ── Footer ────────────────────────────────────────────
        footer = tk.Frame(self._container, bg=BG)
        footer.pack(fill="x", padx=48, pady=(16, 28))

        tk.Label(
            footer,
            text=f"v2.0  ·  {len([m for m in MODULES if m[4]])} modules active",
            font=FONT_SMALL,
            fg=TEXT_LIGHT,
            bg=BG,
        ).pack(side="left")

        tk.Label(
            footer,
            text="IT Department — Mayora",
            font=FONT_SMALL,
            fg=TEXT_LIGHT,
            bg=BG,
        ).pack(side="right")

    def _make_module_button(self, parent, label, subtitle,
                            module_file, class_name, enabled):
        """Create a single large rounded module button card."""

        # Outer shadow layer
        shadow = tk.Frame(parent, bg=SHADOW, bd=0)
        shadow.pack(fill="x", pady=(0, 10))

        # Card (main button area)
        card = tk.Frame(
            shadow,
            bg=BG_CARD,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=BORDER,
        )
        card.pack(fill="x", padx=(0, 2), pady=(0, 2))

        # Inner content row
        inner = tk.Frame(card, bg=BG_CARD, padx=28, pady=20)
        inner.pack(fill="x")
        inner.columnconfigure(0, weight=1)

        # Left side: text
        text_frame = tk.Frame(inner, bg=BG_CARD)
        text_frame.pack(side="left", fill="x", expand=True)

        lbl = tk.Label(
            text_frame,
            text=label,
            font=FONT_BTN,
            fg=TEXT_DARK if enabled else TEXT_LIGHT,
            bg=BG_CARD,
            anchor="w",
        )
        lbl.pack(anchor="w")

        sub = tk.Label(
            text_frame,
            text=subtitle if enabled else subtitle + "  —  Coming Soon",
            font=FONT_SUB,
            fg=TEXT_MID if enabled else TEXT_LIGHT,
            bg=BG_CARD,
            anchor="w",
        )
        sub.pack(anchor="w", pady=(2, 0))

        # Right side: arrow indicator
        arrow = tk.Label(
            inner,
            text="→" if enabled else "○",
            font=("Georgia", 16),
            fg=TEXT_DARK if enabled else TEXT_LIGHT,
            bg=BG_CARD,
        )
        arrow.pack(side="right", padx=(12, 0))

        # ── Hover + Click effects ─────────────────────────────
        all_widgets = [card, inner, text_frame, lbl, sub, arrow, shadow]

        if enabled:
            def on_enter(e, c=card, i=inner, tf=text_frame,
                         l=lbl, s=sub, a=arrow):
                c.configure(bg=BG_HOVER, highlightbackground=TEXT_DARK)
                i.configure(bg=BG_HOVER)
                tf.configure(bg=BG_HOVER)
                l.configure(bg=BG_HOVER)
                s.configure(bg=BG_HOVER)
                a.configure(bg=BG_HOVER)

            def on_leave(e, c=card, i=inner, tf=text_frame,
                         l=lbl, s=sub, a=arrow):
                c.configure(bg=BG_CARD, highlightbackground=BORDER)
                i.configure(bg=BG_CARD)
                tf.configure(bg=BG_CARD)
                l.configure(bg=BG_CARD)
                s.configure(bg=BG_CARD)
                a.configure(bg=BG_CARD)

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
        back_bar = tk.Frame(self._container, bg="#1A1A1A", height=40)
        back_bar.pack(fill="x")
        back_bar.pack_propagate(False)

        back_btn = tk.Label(
            back_bar,
            text="← Menu",
            font=("Courier", 10, "bold"),
            fg="#FFFFFF",
            bg="#1A1A1A",
            cursor="hand2",
            padx=18,
        )
        back_btn.pack(side="left", pady=8)
        back_btn.bind("<Button-1>", lambda e: self._build_menu())
        back_btn.bind("<Enter>",
                      lambda e: back_btn.configure(fg="#AAAAAA"))
        back_btn.bind("<Leave>",
                      lambda e: back_btn.configure(fg="#FFFFFF"))

        # Module name in back bar
        tk.Label(
            back_bar,
            text=label.upper(),
            font=("Courier", 9),
            fg="#888888",
            bg="#1A1A1A",
        ).pack(side="left", pady=8)

        # ── Module frame area ─────────────────────────────────
        module_container = tk.Frame(self._container, bg=BG)
        module_container.pack(fill="both", expand=True)

        # Try to import and instantiate the module
        try:
            # Reload in case file changed
            if module_file in sys.modules:
                mod = importlib.reload(sys.modules[module_file])
            else:
                mod = importlib.import_module(module_file)

            cls = getattr(mod, class_name)

            # Instantiate — pass back_callback so module can return to menu
            instance = cls(module_container, back_callback=self._build_menu)
            self._active_module_frame = instance

        except ImportError as e:
            self._show_module_error(
                module_container,
                f"Module file '{module_file}.py' not found.\n\n{e}"
            )
        except TypeError:
            # Module class doesn't accept back_callback — try without it
            try:
                instance = cls(module_container)
                self._active_module_frame = instance
            except Exception as e2:
                self._show_module_error(module_container, str(e2))
        except Exception as e:
            self._show_module_error(module_container, str(e))

    def _show_module_error(self, parent, message: str):
        """Show a friendly error if a module fails to load."""
        tk.Label(
            parent,
            text="⚠  Module failed to load",
            font=("Georgia", 13, "bold"),
            fg=TEXT_DARK,
            bg=BG,
        ).pack(pady=(60, 8))

        tk.Label(
            parent,
            text=message,
            font=FONT_SUB,
            fg=TEXT_MID,
            bg=BG,
            wraplength=480,
            justify="center",
        ).pack(pady=(0, 24))


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