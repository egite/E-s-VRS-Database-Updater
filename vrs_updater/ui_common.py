"""
Shared dark-theme widget helpers for the editor windows.

Keeps the Rules and Sils editors visually identical to the main window and
to each other.
"""

import os
import tkinter as tk
from tkinter import ttk


# Palette shared with the main window / search dialog
BG = "#1E1E1E"
PANEL = "#263238"
ENTRY = "#37474F"
ACCENT = "#1565C0"
ACCENT_HOVER = "#1976D2"
BTN = "#333333"
BTN_HOVER = "#555555"
TEXT = "#D4D4D4"
LABEL = "#B0BEC5"
CYAN = "#4FC3F7"
ORANGE = "#FFB74D"
RED = "#EF5350"
GREEN = "#81C784"

DB_COLUMNS = ("Icao", "Registration", "Country", "Manufacturer", "Model",
              "ModelIcao", "Operator", "OperatorIcao")


# ---------------------------------------------------------------------------
# Small widget helpers
# ---------------------------------------------------------------------------

def dark_window(win, parent, title, geometry, data_dir=None):
    win.title(title)
    win.geometry(geometry)
    win.configure(bg=BG)
    win.transient(parent)
    if data_dir:
        try:
            icon_path = os.path.join(data_dir, "database.ico")
            if os.path.exists(icon_path):
                win.iconbitmap(icon_path)
        except Exception:
            pass


def dark_button(master, text, command, primary=False, width=None):
    return tk.Button(
        master, text="  %s  " % text, command=command,
        bg=ACCENT if primary else BTN, fg="white",
        font=("Segoe UI", 10, "bold" if primary else "normal"),
        activebackground=ACCENT_HOVER if primary else BTN_HOVER,
        activeforeground="white", relief="flat", width=width,
        disabledforeground="#777777")


def dark_entry(master, textvariable, width=24):
    return tk.Entry(master, textvariable=textvariable, width=width,
                    font=("Consolas", 10), bg=ENTRY, fg="#FFFFFF",
                    insertbackground="#FFFFFF", relief="flat")


def dark_label(master, text, bg=BG, fg=LABEL, font=("Segoe UI", 9), **kwargs):
    return tk.Label(master, text=text, bg=bg, fg=fg, font=font, **kwargs)


def style_tree(prefix):
    """Register the dark Treeview style used by the editor windows."""
    style = ttk.Style()
    style.configure("%s.Treeview" % prefix,
                    background=BG, foreground=TEXT, fieldbackground=BG,
                    font=("Consolas", 9), rowheight=20)
    style.configure("%s.Treeview.Heading" % prefix,
                    background=ENTRY, foreground=LABEL,
                    font=("Segoe UI", 9, "bold"))
    style.map("%s.Treeview" % prefix,
              background=[("selected", "#264F78")],
              foreground=[("selected", "#FFFFFF")])
    return "%s.Treeview" % prefix


def center_on(win, parent):
    win.update_idletasks()
    try:
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h = win.winfo_width(), win.winfo_height()
        win.geometry("+%d+%d" % (px + (pw - w) // 2, py + (ph - h) // 3))
    except Exception:
        pass
