"""
E's VRS Database Updater - Tkinter GUI
Mirrors the original VB.NET WinForms layout.
"""

import csv
import json
import os
import sqlite3
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

from .config import Settings, load_settings
from .faa import download_faa, parse_faa
from .ccar import parse_ccar
from .nz_caa import download_nz_caa, parse_nz_caa
from .casa import download_casa, parse_casa
from .opensky import download_opensky, parse_opensky
from .vrs_merge import update_vrs

# When frozen (PyInstaller), bundled data files are in sys._MEIPASS; settings persist next to the exe
if getattr(sys, 'frozen', False):
    _DATA_DIR = sys._MEIPASS
    _SETTINGS_JSON = os.path.join(os.path.dirname(sys.executable), "settings.json")
else:
    _DATA_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _SETTINGS_JSON = os.path.join(_DATA_DIR, "settings.json")


class VRSUpdaterApp:
    """Main GUI application matching the VB.NET Form1 layout."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("E's VRS Database Updater (Python Edition v2.0)")
        self.root.geometry("720x790")
        self.root.resizable(True, True)
        self.root.minsize(600, 640)

        # Set window icon
        icon_path = os.path.join(_DATA_DIR, "database.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except tk.TclError:
                pass

        self.settings = Settings()
        self.chk_log_errors_var = tk.BooleanVar(value=True)
        self.chk_log_activity_var = tk.BooleanVar(value=False)
        self._error_log_file = None   # Open file handle during a run
        self._activity_log_file = None
        self._in_merge_phase = False  # suppress non-error activity logging during merge
        self.running = False
        self.want_exit = False
        # Per-task progress tracking for concurrent downloads
        self._task_progress = {}       # {task_name: pct}
        self._task_eta = {}            # {task_name: eta_seconds}
        self._aggregate_mode = False   # True during concurrent downloads
        self._thread_task_map = {}     # {thread_id: parent_task_name}
        self._cancelled = threading.Event()  # Set when user clicks Cancel
        self._detail_batch = []            # Batched detail messages
        self._detail_flush_scheduled = False
        self._disk_last_bytes = 0          # Last measured total file sizes
        self._disk_last_time = 0.0         # Time of last measurement
        self._disk_poll_id = None          # after() ID for polling

        self._build_menu()
        self._build_ui()
        self._load_settings()
        self._countdown_id = None  # after() ID for auto-start countdown
        self._auto_run = False     # True when launched via "Run on Start"

        # Refresh DB status whenever the working folder changes (browse, edit, load)
        self.work_dir_var.trace_add("write", lambda *_: self._refresh_db_status())
        self._refresh_db_status()

        # Apply saved window geometry (after UI is built and settings loaded)
        self._apply_saved_geometry()

        # Auto-start countdown if "Run on Start" is checked
        if self.chk_run_on_start_var.get():
            self._start_countdown(10)

    # ------------------------------------------------------------------
    # Menu bar
    # ------------------------------------------------------------------
    def _build_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Settings...", command=self._show_settings)
        file_menu.add_command(label="Options...", command=self._show_options)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_exit)
        menubar.add_cascade(label="File", menu=file_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Search Database...", command=self._show_search)
        menubar.add_cascade(label="Tools", menu=tools_menu)

        help_menu = tk.Menu(menubar, tearoff=0)

        features_menu = tk.Menu(help_menu, tearoff=0)
        features_menu.add_command(label="View All Features...", command=self._show_features)
        features_menu.add_separator()
        features_menu.add_command(label="Concurrent Downloads", state="disabled")
        features_menu.add_command(label="Parallel Database Parsing", state="disabled")
        features_menu.add_command(label="O(1) Silhouette Lookups", state="disabled")
        features_menu.add_command(label="Dark Terminal Theme", state="disabled")
        features_menu.add_command(label="Cancel Support", state="disabled")
        features_menu.add_command(label="Disk Throughput Monitor", state="disabled")
        features_menu.add_command(label="Custom Rules Override", state="disabled")
        features_menu.add_command(label="Skip Processing Controls", state="disabled")
        features_menu.add_command(label="Phase Indicator", state="disabled")
        features_menu.add_command(label="Download Age Limits", state="disabled")
        features_menu.add_command(label="Database Priority Order", state="disabled")
        help_menu.add_cascade(label="Features", menu=features_menu)

        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    # ------------------------------------------------------------------
    # Main UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        # --- Title banner with icons ---
        banner = tk.Frame(self.root, bg="#1565C0", height=50)
        banner.pack(fill="x")
        banner.pack_propagate(False)

        # Load banner images (app icon -> red arrow -> VRS logo)
        base_dir = _DATA_DIR
        self._banner_images = []  # prevent garbage collection

        for fname in ("banner_sources.png", "banner_arrow.png", "banner_app_icon.png", "banner_arrow.png", "banner_vrs_icon.png"):
            img_path = os.path.join(base_dir, fname)
            if os.path.exists(img_path):
                try:
                    img = tk.PhotoImage(file=img_path)
                    self._banner_images.append(img)
                    if fname.startswith("banner_sources"):
                        padx = (10, 0)
                    elif fname.startswith("banner_app"):
                        padx = (6, 0)
                    elif fname.startswith("banner_arrow"):
                        padx = (11, 0)
                    else:
                        padx = (4, 0)
                    tk.Label(banner, image=img, bg="#1565C0").pack(side="left", padx=padx, pady=4)
                except tk.TclError:
                    pass

        # Centered title text (centered across full banner width, ignoring icons)
        tk.Label(banner, text="E's VRS Database Updater",
                 font=("Segoe UI", 18, "bold"), fg="white",
                 bg="#1565C0", pady=8).place(relx=0.5, rely=0.5, anchor="center")

        # --- Top frame: paths ---
        path_frame = ttk.LabelFrame(self.root, text="Paths", padding=8)
        path_frame.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(path_frame, text="Working Folder:").grid(row=0, column=0, sticky="w")
        self.work_dir_var = tk.StringVar()
        self.txt_work_dir = ttk.Entry(path_frame, textvariable=self.work_dir_var, width=60)
        self.txt_work_dir.grid(row=0, column=1, padx=4, sticky="ew")
        ttk.Button(path_frame, text="Browse...", command=self._browse_work_dir).grid(row=0, column=2, padx=4)

        ttk.Label(path_frame, text="VRS Folder:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.vrs_dir_var = tk.StringVar()
        self.txt_vrs_dir = ttk.Entry(path_frame, textvariable=self.vrs_dir_var, width=60)
        self.txt_vrs_dir.grid(row=1, column=1, padx=4, pady=(4, 0), sticky="ew")
        ttk.Button(path_frame, text="Browse...", command=self._browse_vrs_dir).grid(row=1, column=2, padx=4, pady=(4, 0))

        path_frame.columnconfigure(1, weight=1)

        # --- Options: Auto-Download (left) + stacked Skip/Options (right) ---
        opt_row = ttk.Frame(self.root)
        opt_row.pack(fill="x", padx=8, pady=4)

        # Auto-Download group — one row per source with date/age status
        dl_frame = ttk.LabelFrame(opt_row, text="Auto-Download", padding=8)
        dl_frame.pack(side="left", padx=(0, 4), fill="y")

        self.chk_faa_download_var = tk.BooleanVar(value=True)
        self.chk_ccar_download_var = tk.BooleanVar(value=True)
        self.chk_nzcaa_download_var = tk.BooleanVar(value=True)
        self.chk_casa_download_var = tk.BooleanVar(value=True)
        self.chk_opensky_download_var = tk.BooleanVar(value=True)

        # (label, var, db filename, settings attr for stale-days threshold)
        self._db_sources = [
            ("FAA",     self.chk_faa_download_var,     "FAADatabase.sqb",    "faa_max_age_days"),
            ("CCAR",    self.chk_ccar_download_var,    "CCARDatabase.sqb",   "ccar_max_age_days"),
            ("NZ CAA",  self.chk_nzcaa_download_var,   "NZCAADatabase.sqb",  "nz_caa_max_age_days"),
            ("CASA",    self.chk_casa_download_var,    "CASADatabase.sqb",   "casa_max_age_days"),
            ("OpenSky", self.chk_opensky_download_var, "OpenSkyDatabase.sqb","opensky_max_age_days"),
        ]
        self._db_status_labels = {}  # name -> (label widget, filename, age_attr)
        for r, (name, var, filename, age_attr) in enumerate(self._db_sources):
            ttk.Checkbutton(dl_frame, text=name, variable=var).grid(
                row=r, column=0, sticky="w", padx=(0, 12), pady=1)
            lbl = tk.Label(dl_frame, text="—", fg="#888888",
                           font=("Segoe UI", 9), anchor="w")
            lbl.grid(row=r, column=1, sticky="w", pady=1)
            self._db_status_labels[name] = (lbl, filename, age_attr)
        dl_frame.columnconfigure(1, weight=1)

        # Right column: Skip Processing on top, Options on bottom
        right_col = ttk.Frame(opt_row)
        right_col.pack(side="left", fill="both", expand=True, padx=(4, 0))

        # Skip Processing group (top)
        skip_frame = ttk.LabelFrame(right_col, text="Skip Processing", padding=8)
        skip_frame.pack(side="top", fill="x")

        self.chk_skip_faa_var = tk.BooleanVar()
        self.chk_skip_ccar_var = tk.BooleanVar()
        self.chk_skip_nzcaa_var = tk.BooleanVar()
        self.chk_skip_casa_var = tk.BooleanVar()
        self.chk_skip_opensky_var = tk.BooleanVar()
        self.chk_skip_rules_var = tk.BooleanVar()

        ttk.Checkbutton(skip_frame, text="FAA", variable=self.chk_skip_faa_var).grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Checkbutton(skip_frame, text="CCAR", variable=self.chk_skip_ccar_var).grid(row=0, column=1, sticky="w", padx=(0, 10))
        ttk.Checkbutton(skip_frame, text="NZ CAA", variable=self.chk_skip_nzcaa_var).grid(row=0, column=2, sticky="w", padx=(0, 10))
        ttk.Checkbutton(skip_frame, text="CASA", variable=self.chk_skip_casa_var).grid(row=0, column=3, sticky="w", padx=(0, 10))
        ttk.Checkbutton(skip_frame, text="OpenSky", variable=self.chk_skip_opensky_var).grid(row=0, column=4, sticky="w", padx=(0, 10))
        ttk.Checkbutton(skip_frame, text="Rules", variable=self.chk_skip_rules_var).grid(row=0, column=5, sticky="w")

        # Options group (bottom)
        opt_frame = ttk.LabelFrame(right_col, text="Options", padding=8)
        opt_frame.pack(side="top", fill="x", pady=(4, 0))

        self.chk_run_on_start_var = tk.BooleanVar()
        self.chk_backup_var = tk.BooleanVar(value=True)
        self.chk_complete_var = tk.BooleanVar()

        ttk.Checkbutton(opt_frame, text="Run on Start", variable=self.chk_run_on_start_var).grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Checkbutton(opt_frame, text="Backup VRS database", variable=self.chk_backup_var).grid(row=0, column=1, sticky="w", padx=(0, 16))
        ttk.Checkbutton(opt_frame, text="Build complete database", variable=self.chk_complete_var).grid(row=0, column=2, sticky="w")

        # Progress frame (dark themed) — bottom of right column so it bottom-
        # aligns with the Auto-Download box on the left.
        prog_frame = tk.Frame(right_col, bg="#263238", padx=8, pady=6)
        prog_frame.pack(side="top", fill="x", pady=(4, 0))

        tk.Label(prog_frame, text="Progress:", bg="#263238", fg="#B0BEC5",
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.percent_var = tk.StringVar(value=" - %")
        tk.Label(prog_frame, textvariable=self.percent_var, width=10, anchor="center",
                 font=("Segoe UI", 11, "bold"), bg="#263238", fg="#4CAF50").grid(row=0, column=1, padx=8)

        tk.Label(prog_frame, text="ETA:", bg="#263238", fg="#B0BEC5",
                 font=("Segoe UI", 9)).grid(row=0, column=2, sticky="w")
        self.eta_var = tk.StringVar(value=" - ")
        tk.Label(prog_frame, textvariable=self.eta_var, width=12, anchor="center",
                 font=("Segoe UI", 11, "bold"), bg="#263238", fg="#FFD54F").grid(row=0, column=3, padx=4)

        # Progress bar (custom green style)
        style = ttk.Style()
        style.configure("Green.Horizontal.TProgressbar",
                         troughcolor="#37474F", background="#4CAF50")
        self.progress_bar = ttk.Progressbar(prog_frame, mode="determinate", length=200,
                                            style="Green.Horizontal.TProgressbar")
        self.progress_bar.grid(row=0, column=4, padx=(16, 0), sticky="ew")
        prog_frame.columnconfigure(4, weight=1)

        # --- Start / Cancel buttons + throughput ---
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=8, pady=4)

        self.btn_start = tk.Button(btn_frame, text="  Start  ", command=self._on_start,
                                   bg="#4CAF50", fg="white", font=("Segoe UI", 11, "bold"),
                                   activebackground="#45a049", relief="raised", bd=2)
        self.btn_start.pack(side="left")

        self.btn_cancel = tk.Button(btn_frame, text="  Cancel  ", command=self._on_cancel,
                                    bg="#f44336", fg="white", font=("Segoe UI", 11, "bold"),
                                    activebackground="#d32f2f", relief="raised", bd=2,
                                    state="disabled")
        self.btn_cancel.pack(side="left", padx=(8, 0))

        # Phase indicator (right side)
        self.phase_var = tk.StringVar(value="Idle")
        tk.Label(btn_frame, textvariable=self.phase_var,
                 font=("Segoe UI", 11, "bold"), fg="#FF9800").pack(side="right", padx=(0, 4))
        tk.Label(btn_frame, text="Phase:",
                 font=("Segoe UI", 10), fg="#666666").pack(side="right", padx=(0, 4))

        # Throughput indicator (teal)
        self.throughput_var = tk.StringVar(value="")
        tk.Label(btn_frame, textvariable=self.throughput_var,
                 font=("Consolas", 10, "bold"), fg="#00BCD4").pack(side="left", padx=(16, 0))

        # --- Resizable split: Program Progress (top) + Database Update Status (bottom)
        # User can drag the sash; ratio is preserved across window resizes and runs.
        self.paned = ttk.PanedWindow(self.root, orient="vertical")
        self.paned.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        status_frame = ttk.LabelFrame(self.paned, text="Program Progress", padding=4)
        self.txt_status = scrolledtext.ScrolledText(
            status_frame, height=10, wrap="word",
            font=("Consolas", 9), state="disabled",
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="#D4D4D4",
            selectbackground="#264F78", selectforeground="#FFFFFF")
        self.txt_status.pack(fill="both", expand=True)
        # weight=1 (top) vs weight=2 (bottom) gives a 1:2 default split before the
        # user has dragged the sash; after that, _on_paned_configure preserves
        # whatever ratio the user chose.
        self.paned.add(status_frame, weight=1)

        # Color tags for status log
        self.txt_status.tag_configure("step", foreground="#4FC3F7")       # cyan - step headers
        self.txt_status.tag_configure("separator", foreground="#616161")  # dim gray
        self.txt_status.tag_configure("warning", foreground="#FFD54F")    # yellow
        self.txt_status.tag_configure("error", foreground="#EF5350")      # red
        self.txt_status.tag_configure("success", foreground="#66BB6A")    # green
        self.txt_status.tag_configure("info", foreground="#B0BEC5")       # light gray
        self.txt_status.tag_configure("done", foreground="#66BB6A", font=("Consolas", 9, "bold"))

        # --- Detail log (Database Update Status) - dark terminal ---
        detail_frame = ttk.LabelFrame(self.paned, text="Database Update Status", padding=4)
        self.txt_details = scrolledtext.ScrolledText(
            detail_frame, height=20, wrap="word",
            font=("Consolas", 9), state="disabled",
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="#D4D4D4",
            selectbackground="#264F78", selectforeground="#FFFFFF")
        self.txt_details.pack(fill="both", expand=True)
        self.paned.add(detail_frame, weight=2)

        # Sash-position tracking: captured on user drag, re-applied on window resize
        self._pane_ratio = None              # status_height / total_height (0..1)
        self._sash_dragging = False
        self._suppress_pane_configure = False
        self.paned.bind("<ButtonPress-1>", self._on_sash_press)
        self.paned.bind("<ButtonRelease-1>", self._on_sash_release)
        self.paned.bind("<Configure>", self._on_paned_configure)

        # Color tags for detail log
        self.txt_details.tag_configure("rule", foreground="#FFB74D")       # orange - from rules
        self.txt_details.tag_configure("pii", foreground="#FFB74D")        # orange - PII removed/redacted
        self.txt_details.tag_configure("operator", foreground="#4FC3F7")   # light blue - operator format
        self.txt_details.tag_configure("registration", foreground="#CE93D8")  # light purple - registration
        self.txt_details.tag_configure("complete", foreground="#66BB6A")   # green - completion

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_exit)

    # ------------------------------------------------------------------
    # Settings persistence (JSON file)
    # ------------------------------------------------------------------
    def _load_settings_json(self) -> dict:
        """Load settings from JSON file, returns dict (empty if not found)."""
        if os.path.exists(_SETTINGS_JSON):
            try:
                with open(_SETTINGS_JSON, "r") as f:
                    return json.load(f)
            except (json.JSONDecodeError, OSError):
                pass
        return {}

    def _save_settings_json(self):
        """Save all options and window geometry to JSON."""
        # When the window is maximized, winfo_* reports the zoomed (full-screen)
        # geometry, not the un-zoomed restore size. Preserve the previous restore
        # geometry instead so un-maximizing on next launch lands somewhere sane.
        is_zoomed = (self.root.state() == "zoomed")
        if is_zoomed:
            prev = self._load_settings_json()
            win_x = prev.get("window_x", self.root.winfo_x())
            win_y = prev.get("window_y", self.root.winfo_y())
            win_w = prev.get("window_width", self.root.winfo_width())
            win_h = prev.get("window_height", self.root.winfo_height())
        else:
            win_x = self.root.winfo_x()
            win_y = self.root.winfo_y()
            win_w = self.root.winfo_width()
            win_h = self.root.winfo_height()

        data = {
            "work_dir": self.work_dir_var.get(),
            "vrs_dir": self.vrs_dir_var.get(),
            "download_faa": self.chk_faa_download_var.get(),
            "download_ccar": self.chk_ccar_download_var.get(),
            "download_opensky": self.chk_opensky_download_var.get(),
            "download_nz_caa": self.chk_nzcaa_download_var.get(),
            "download_casa": self.chk_casa_download_var.get(),
            "backup_vrs_db": self.chk_backup_var.get(),
            "build_complete": self.chk_complete_var.get(),
            "run_on_start": self.chk_run_on_start_var.get(),
            "skip_faa": self.chk_skip_faa_var.get(),
            "skip_ccar": self.chk_skip_ccar_var.get(),
            "skip_opensky": self.chk_skip_opensky_var.get(),
            "skip_nz_caa": self.chk_skip_nzcaa_var.get(),
            "skip_casa": self.chk_skip_casa_var.get(),
            "skip_rules": self.chk_skip_rules_var.get(),
            "log_errors": self.chk_log_errors_var.get(),
            "log_activity": self.chk_log_activity_var.get(),
            "faa_url": self.settings.faa_url,
            "ccar_url": self.settings.ccar_url,
            "opensky_url": self.settings.opensky_url,
            "nz_caa_url": self.settings.nz_caa_url,
            "casa_url": self.settings.casa_url,
            "merge_order": self.settings.merge_order,
            "faa_max_age_days": self.settings.faa_max_age_days,
            "ccar_max_age_days": self.settings.ccar_max_age_days,
            "opensky_max_age_days": self.settings.opensky_max_age_days,
            "nz_caa_max_age_days": self.settings.nz_caa_max_age_days,
            "casa_max_age_days": self.settings.casa_max_age_days,
            "window_x": win_x,
            "window_y": win_y,
            "window_width": win_w,
            "window_height": win_h,
            "window_zoomed": is_zoomed,
            "pane_ratio": self._pane_ratio,
        }
        try:
            with open(_SETTINGS_JSON, "w") as f:
                json.dump(data, f, indent=2)
        except OSError:
            pass

    def _apply_saved_geometry(self):
        """Apply saved window position and size from JSON.

        Negative coordinates are valid on multi-monitor setups (a secondary
        monitor positioned to the left of / above the primary), so we only
        sanity-check the size and clamp coordinates to a reasonable range.
        If the window was zoomed (maximized) at last exit, restore it to that
        state — but only after applying the un-zoomed geometry so un-maximizing
        lands at the previous restore size/position.
        """
        data = self._load_settings_json()
        if not data:
            return
        w = data.get("window_width", 0)
        h = data.get("window_height", 0)
        x = data.get("window_x", 0)
        y = data.get("window_y", 0)
        # Sanity bounds: size must be reasonable; coords accept negative values
        # for multi-monitor setups but reject obviously bogus values.
        if w > 100 and h > 100 and -10000 < x < 20000 and -10000 < y < 20000:
            self.root.geometry(f"{w}x{h}+{x}+{y}")
        if data.get("window_zoomed"):
            # Re-maximize after the geometry call has been processed.
            self.root.after(0, lambda: self.root.state("zoomed"))

        # Restore saved pane split ratio (status_height / total_height)
        ratio = data.get("pane_ratio")
        if isinstance(ratio, (int, float)) and 0.05 < ratio < 0.95:
            self._pane_ratio = float(ratio)
            # PanedWindow needs to be mapped before sashpos works; defer.
            self.root.after(50, self._apply_pane_ratio)

    # ------------------------------------------------------------------
    # Resizable split between Program Progress and Database Update Status
    # ------------------------------------------------------------------
    def _on_sash_press(self, _event):
        # Tk handles the actual sash drag; we just need to know we're in one
        # so window-resize <Configure> events don't fight the user's drag.
        self._sash_dragging = True

    def _on_sash_release(self, _event):
        if not self._sash_dragging:
            return
        self._sash_dragging = False
        try:
            total = self.paned.winfo_height()
            sash = self.paned.sashpos(0)
        except tk.TclError:
            return
        if total > 50 and 10 < sash < total - 10:
            self._pane_ratio = sash / total
            # Persist immediately so a crash before exit doesn't lose it
            self._save_settings_json()

    def _on_paned_configure(self, event):
        """When the PanedWindow itself resizes (e.g., window resize), preserve
        the user's chosen ratio by reapplying sashpos."""
        if self._sash_dragging or self._suppress_pane_configure:
            return
        if self._pane_ratio is None:
            return
        if event.height <= 50:
            return
        self.paned.after_idle(self._apply_pane_ratio)

    def _apply_pane_ratio(self):
        if self._pane_ratio is None:
            return
        try:
            total = self.paned.winfo_height()
        except tk.TclError:
            return
        if total <= 50:
            return
        new_sash = int(total * self._pane_ratio)
        # Keep both panes at least ~30 px tall
        new_sash = max(30, min(new_sash, total - 30))
        self._suppress_pane_configure = True
        try:
            self.paned.sashpos(0, new_sash)
        except tk.TclError:
            pass
        # Release the guard on the next idle so the resulting Configure
        # (caused by sashpos) doesn't trigger another reapplication.
        self.paned.after_idle(self._release_pane_configure_guard)

    def _release_pane_configure_guard(self):
        self._suppress_pane_configure = False

    def _load_settings(self):
        """Load settings from JSON (falls back to settings.sqb for migration)."""
        data = self._load_settings_json()

        if data:
            # Load from JSON
            self.work_dir_var.set(data.get("work_dir", os.getcwd()))
            self.vrs_dir_var.set(data.get("vrs_dir", ""))
            self.chk_faa_download_var.set(data.get("download_faa", True))
            self.chk_ccar_download_var.set(data.get("download_ccar", True))
            self.chk_opensky_download_var.set(data.get("download_opensky", True))
            self.chk_nzcaa_download_var.set(data.get("download_nz_caa", True))
            self.chk_casa_download_var.set(data.get("download_casa", True))
            self.chk_backup_var.set(data.get("backup_vrs_db", True))
            self.chk_complete_var.set(data.get("build_complete", False))
            self.chk_run_on_start_var.set(data.get("run_on_start", False))
            self.chk_skip_faa_var.set(data.get("skip_faa", False))
            self.chk_skip_ccar_var.set(data.get("skip_ccar", False))
            self.chk_skip_opensky_var.set(data.get("skip_opensky", False))
            self.chk_skip_nzcaa_var.set(data.get("skip_nz_caa", False))
            self.chk_skip_casa_var.set(data.get("skip_casa", False))
            self.chk_skip_rules_var.set(data.get("skip_rules", False))
            self.chk_log_errors_var.set(data.get("log_errors", True))
            self.chk_log_activity_var.set(data.get("log_activity", False))
            from .config import DEFAULT_FAA_URL, DEFAULT_CCAR_URL, DEFAULT_OPENSKY_URL, DEFAULT_NZCAA_URL, DEFAULT_CASA_URL
            self.settings.faa_url = data.get("faa_url", DEFAULT_FAA_URL)
            self.settings.ccar_url = data.get("ccar_url", DEFAULT_CCAR_URL)
            self.settings.opensky_url = data.get("opensky_url", DEFAULT_OPENSKY_URL)
            self.settings.nz_caa_url = data.get("nz_caa_url", DEFAULT_NZCAA_URL)
            self.settings.casa_url = data.get("casa_url", DEFAULT_CASA_URL)
            saved_order = data.get("merge_order", ["Rules.CSV (when present)", "FAA", "CCAR", "NZ CAA", "OpenSky"])
            # Ensure Rules entry is present (added after earlier versions)
            if not any(item.startswith("Rules") for item in saved_order):
                saved_order.insert(0, "Rules.CSV (when present)")
            # Ensure NZ CAA entry is present (added after earlier versions)
            if "NZ CAA" not in saved_order:
                # Insert before OpenSky if present, otherwise append before last
                try:
                    os_idx = saved_order.index("OpenSky")
                    saved_order.insert(os_idx, "NZ CAA")
                except ValueError:
                    saved_order.append("NZ CAA")
            # Ensure CASA entry is present (added after earlier versions)
            if "CASA" not in saved_order:
                try:
                    os_idx = saved_order.index("OpenSky")
                    saved_order.insert(os_idx, "CASA")
                except ValueError:
                    saved_order.append("CASA")
            self.settings.merge_order = saved_order
            self.settings.faa_max_age_days = data.get("faa_max_age_days", 0)
            self.settings.ccar_max_age_days = data.get("ccar_max_age_days", 0)
            self.settings.opensky_max_age_days = data.get("opensky_max_age_days", 0)
            self.settings.nz_caa_max_age_days = data.get("nz_caa_max_age_days", 0)
            self.settings.casa_max_age_days = data.get("casa_max_age_days", 0)
        else:
            # Fall back to settings.sqb (VB.NET migration)
            settings_path = os.path.join(os.getcwd(), "settings.sqb")
            if os.path.exists(settings_path):
                self.settings = load_settings(settings_path)

            if not self.settings.work_dir:
                self.settings.work_dir = os.getcwd()
            self.work_dir_var.set(self.settings.work_dir)

            if not self.settings.vrs_dir:
                username = os.environ.get("USERNAME", os.environ.get("USER", ""))
                default_vrs = os.path.join("C:\\Users", username, "AppData", "Local", "VirtualRadar")
                if os.path.exists(default_vrs):
                    self.settings.vrs_dir = default_vrs
            self.vrs_dir_var.set(self.settings.vrs_dir)

            self.chk_faa_download_var.set(self.settings.download_faa)
            self.chk_ccar_download_var.set(self.settings.download_ccar)
            self.chk_opensky_download_var.set(self.settings.download_opensky)
            self.chk_backup_var.set(self.settings.backup_vrs_db)
            self.chk_complete_var.set(self.settings.build_complete)

        # Default VRS dir if still empty
        if not self.vrs_dir_var.get():
            username = os.environ.get("USERNAME", os.environ.get("USER", ""))
            default_vrs = os.path.join("C:\\Users", username, "AppData", "Local", "VirtualRadar")
            if os.path.exists(default_vrs):
                self.vrs_dir_var.set(default_vrs)

    def _build_settings(self) -> Settings:
        """Build Settings from current UI state."""
        s = Settings()
        s.work_dir = os.path.abspath(self.work_dir_var.get())
        s.vrs_dir = self.vrs_dir_var.get().strip()
        if s.vrs_dir:
            s.vrs_dir = os.path.abspath(s.vrs_dir)
        s.download_faa = self.chk_faa_download_var.get()
        s.download_ccar = self.chk_ccar_download_var.get()
        s.download_opensky = self.chk_opensky_download_var.get()
        s.download_nz_caa = self.chk_nzcaa_download_var.get()
        s.download_casa = self.chk_casa_download_var.get()
        s.backup_vrs_db = self.chk_backup_var.get()
        s.build_complete = self.chk_complete_var.get()
        s.skip_faa = self.chk_skip_faa_var.get()
        s.skip_ccar = self.chk_skip_ccar_var.get()
        s.skip_opensky = self.chk_skip_opensky_var.get()
        s.skip_nz_caa = self.chk_skip_nzcaa_var.get()
        s.skip_casa = self.chk_skip_casa_var.get()
        s.skip_rules = self.chk_skip_rules_var.get()
        s.faa_url = self.settings.faa_url
        s.ccar_url = self.settings.ccar_url
        s.opensky_url = self.settings.opensky_url
        s.nz_caa_url = self.settings.nz_caa_url
        s.casa_url = self.settings.casa_url
        s.faa_max_age_days = self.settings.faa_max_age_days
        s.ccar_max_age_days = self.settings.ccar_max_age_days
        s.opensky_max_age_days = self.settings.opensky_max_age_days
        s.nz_caa_max_age_days = self.settings.nz_caa_max_age_days
        s.casa_max_age_days = self.settings.casa_max_age_days
        s.merge_order = self.settings.merge_order
        return s

    # ------------------------------------------------------------------
    # Logging helpers (thread-safe via root.after)
    # ------------------------------------------------------------------
    def _log_status(self, msg: str):
        """Append to status log with color coding."""
        tag = self._classify_status(msg)
        self.root.after(0, self._append_colored, self.txt_status, msg, tag)
        self._log_activity(msg)

    @staticmethod
    def _classify_status(msg: str) -> str:
        """Determine the color tag for a status message."""
        stripped = msg.strip()
        if not stripped:
            return ""
        if stripped.startswith("[Step"):
            return "step"
        if stripped.startswith("---") or stripped.startswith("==="):
            return "separator"
        if "WARNING" in stripped or "warning" in stripped:
            return "warning"
        if "ERROR" in stripped:
            return "error"
        if stripped.startswith("Done!") or "complete" in stripped.lower():
            return "done"
        if stripped.startswith("Started at") or stripped.startswith("Total time"):
            return "info"
        return ""

    def _log_detail(self, msg: str):
        """Append to detail log, batched for performance."""
        tag = self._classify_detail(msg)
        self._detail_batch.append((msg, tag))
        if not self._detail_flush_scheduled:
            self._detail_flush_scheduled = True
            self.root.after(500, self._flush_detail_batch)

    @staticmethod
    def _classify_detail(msg: str) -> str:
        """Determine the color tag for a detail message."""
        if "From rules" in msg:
            return "rule"
        if "PII" in msg:
            return "pii"
        if "Operator format" in msg:
            return "operator"
        if "Operator has registration" in msg:
            return "registration"
        if "processing completed" in msg:
            return "complete"
        return ""

    def _flush_detail_batch(self):
        """Flush batched detail messages to the widget."""
        self._detail_flush_scheduled = False
        if not self._detail_batch:
            return
        self.txt_details.config(state="normal")
        for msg, tag in self._detail_batch:
            if tag:
                self.txt_details.insert("end", msg + "\n", tag)
            else:
                self.txt_details.insert("end", msg + "\n")
        self.txt_details.see("end")
        self.txt_details.config(state="disabled")
        self._detail_batch.clear()

    def _append_colored(self, widget, msg, tag=""):
        """Append a colored line to a text widget."""
        widget.config(state="normal")
        if tag:
            widget.insert("end", msg + "\n", tag)
        else:
            widget.insert("end", msg + "\n")
        widget.see("end")
        widget.config(state="disabled")

    def _get_work_dir_bytes(self) -> int:
        """Sum sizes of .sqb files in the working directory."""
        work_dir = self.work_dir_var.get()
        if not work_dir or not os.path.isdir(work_dir):
            return 0
        total = 0
        try:
            for f in os.listdir(work_dir):
                if f.endswith(".sqb"):
                    fp = os.path.join(work_dir, f)
                    try:
                        total += os.path.getsize(fp)
                    except OSError:
                        pass
        except OSError:
            pass
        return total

    def _start_disk_poll(self):
        """Start polling disk throughput every second."""
        import time
        self._disk_last_bytes = self._get_work_dir_bytes()
        self._disk_last_time = time.time()
        self._poll_disk()

    def _poll_disk(self):
        """Periodic disk throughput measurement."""
        if not self.running:
            self.throughput_var.set("")
            self._disk_poll_id = None
            return
        import time
        now = time.time()
        current_bytes = self._get_work_dir_bytes()
        dt = now - self._disk_last_time
        if dt > 0.5:
            delta = current_bytes - self._disk_last_bytes
            rate = abs(delta) / dt
            if rate >= 1024 * 1024:
                self.throughput_var.set(f"{rate / (1024*1024):.1f} MB/s")
            elif rate >= 1024:
                self.throughput_var.set(f"{rate / 1024:.0f} KB/s")
            elif rate > 0:
                self.throughput_var.set(f"{rate:.0f} B/s")
            else:
                self.throughput_var.set("idle")
            self._disk_last_bytes = current_bytes
            self._disk_last_time = now
        self._disk_poll_id = self.root.after(1000, self._poll_disk)

    def _stop_disk_poll(self):
        """Stop the disk throughput polling."""
        if self._disk_poll_id:
            self.root.after_cancel(self._disk_poll_id)
            self._disk_poll_id = None

    def _set_phase(self, phase: str):
        """Update the phase indicator (thread-safe)."""
        self.root.after(0, self.phase_var.set, phase)

    def _set_progress(self, percent: float, eta_seconds: float,
                       task_name: str = ""):
        """Update progress display. In aggregate mode, combines per-task progress."""
        if self._aggregate_mode and task_name:
            # Map ProgressReporter task names back to the parent download task
            parent = self._thread_task_map.get(threading.current_thread().ident, task_name)
            self._task_progress[parent] = percent
            self._task_eta[parent] = eta_seconds if eta_seconds >= 0 else 0
            if self._task_progress:
                avg = sum(self._task_progress.values()) / len(self._task_progress)
            else:
                avg = 0
            # Overall ETA is the longest remaining task (they run in parallel)
            agg_eta = max(self._task_eta.values()) if self._task_eta else -1
            self.root.after(0, self._update_progress_ui, avg, agg_eta)
        else:
            self.root.after(0, self._update_progress_ui, percent, eta_seconds)

    def _update_progress_ui(self, percent, eta_seconds):
        self.percent_var.set(f"{percent:.0f}%")
        if eta_seconds >= 0:
            h = int(eta_seconds) // 3600
            m = (int(eta_seconds) % 3600) // 60
            s = int(eta_seconds) % 60
            self.eta_var.set(f"{h}:{m:02d}:{s:02d}")
        else:
            self.eta_var.set(" - ")
        self.progress_bar["value"] = percent

    def _reset_progress(self):
        self.root.after(0, self._do_reset_progress)

    def _do_reset_progress(self):
        self.percent_var.set(" - %")
        self.eta_var.set(" - ")
        self.progress_bar["value"] = 0

    # ------------------------------------------------------------------
    # Button / menu handlers
    # ------------------------------------------------------------------
    def _browse_work_dir(self):
        d = filedialog.askdirectory(title="Select Working Folder",
                                   initialdir=self.work_dir_var.get())
        if d:
            self.work_dir_var.set(d)

    def _browse_vrs_dir(self):
        d = filedialog.askdirectory(title="Select VRS Folder",
                                   initialdir=self.vrs_dir_var.get())
        if d:
            self.vrs_dir_var.set(d)

    def _refresh_db_status(self):
        """Update the date/age label next to each Auto-Download checkbox.

        Reads each *.sqb file's mtime in the current working folder and
        flags the entry as stale if it exceeds that source's max-age setting
        (0 = never stale).
        """
        if not hasattr(self, "_db_status_labels"):
            return
        work_dir = self.work_dir_var.get().strip()
        now = datetime.now()
        for name, (lbl, filename, age_attr) in self._db_status_labels.items():
            if not work_dir:
                lbl.config(text="—", fg="#888888")
                continue
            path = os.path.join(work_dir, filename)
            if os.path.exists(path):
                try:
                    mtime = datetime.fromtimestamp(os.path.getmtime(path))
                except OSError:
                    lbl.config(text="—", fg="#888888")
                    continue
                age_days = (now - mtime).days
                stale_days = getattr(self.settings, age_attr, 0) or 0
                date_str = mtime.strftime("%d %b %Y")
                if stale_days > 0 and age_days > stale_days:
                    lbl.config(text=f"✓ {date_str} ({age_days}d, stale)",
                               fg="#FF8F00")
                else:
                    lbl.config(text=f"✓ {date_str} ({age_days}d)",
                               fg="#2E7D32")
            else:
                lbl.config(text="✗ Missing", fg="#C62828")

    def _show_settings(self):
        """Show a dialog with editable download URLs."""
        from .config import DEFAULT_FAA_URL, DEFAULT_OPENSKY_URL, DEFAULT_CCAR_URL, DEFAULT_NZCAA_URL, DEFAULT_CASA_URL

        win = tk.Toplevel(self.root)
        win.title("Settings - Download URLs")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        frame = ttk.Frame(win, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="FAA Database URL:").grid(row=0, column=0, sticky="w", pady=(0, 2))
        faa_var = tk.StringVar(value=self.settings.faa_url)
        ttk.Entry(frame, textvariable=faa_var, width=70).grid(row=1, column=0, pady=(0, 8), sticky="ew")

        ttk.Label(frame, text="CCAR Database URL:").grid(row=2, column=0, sticky="w", pady=(0, 2))
        ccar_var = tk.StringVar(value=self.settings.ccar_url)
        ttk.Entry(frame, textvariable=ccar_var, width=70).grid(row=3, column=0, pady=(0, 8), sticky="ew")

        ttk.Label(frame, text="OpenSky Database URL:").grid(row=4, column=0, sticky="w", pady=(0, 2))
        opensky_var = tk.StringVar(value=self.settings.opensky_url)
        ttk.Entry(frame, textvariable=opensky_var, width=70).grid(row=5, column=0, pady=(0, 8), sticky="ew")

        ttk.Label(frame, text="NZ CAA Register URL:").grid(row=6, column=0, sticky="w", pady=(0, 2))
        nzcaa_var = tk.StringVar(value=self.settings.nz_caa_url)
        ttk.Entry(frame, textvariable=nzcaa_var, width=70).grid(row=7, column=0, pady=(0, 8), sticky="ew")

        ttk.Label(frame, text="CASA Register URL:").grid(row=8, column=0, sticky="w", pady=(0, 2))
        casa_var = tk.StringVar(value=self.settings.casa_url)
        ttk.Entry(frame, textvariable=casa_var, width=70).grid(row=9, column=0, pady=(0, 12), sticky="ew")

        # Download age limits
        ttk.Label(frame, text="Re-download if local copy older than (0 = always):").grid(
            row=10, column=0, sticky="w", pady=(0, 4))

        age_frame = ttk.Frame(frame)
        age_frame.grid(row=11, column=0, sticky="w", pady=(0, 12))

        faa_age_var = tk.IntVar(value=self.settings.faa_max_age_days)
        ccar_age_var = tk.IntVar(value=self.settings.ccar_max_age_days)
        opensky_age_var = tk.IntVar(value=self.settings.opensky_max_age_days)
        nzcaa_age_var = tk.IntVar(value=self.settings.nz_caa_max_age_days)
        casa_age_var = tk.IntVar(value=self.settings.casa_max_age_days)

        ttk.Label(age_frame, text="FAA:").grid(row=0, column=0, sticky="w", padx=(0, 4))
        ttk.Spinbox(age_frame, from_=0, to=30, width=4, textvariable=faa_age_var).grid(row=0, column=1, padx=(0, 4))
        ttk.Label(age_frame, text="days").grid(row=0, column=2, padx=(0, 12))

        ttk.Label(age_frame, text="CCAR:").grid(row=0, column=3, sticky="w", padx=(0, 4))
        ttk.Spinbox(age_frame, from_=0, to=30, width=4, textvariable=ccar_age_var).grid(row=0, column=4, padx=(0, 4))
        ttk.Label(age_frame, text="days").grid(row=0, column=5, padx=(0, 12))

        ttk.Label(age_frame, text="NZ CAA:").grid(row=0, column=6, sticky="w", padx=(0, 4))
        ttk.Spinbox(age_frame, from_=0, to=30, width=4, textvariable=nzcaa_age_var).grid(row=0, column=7, padx=(0, 4))
        ttk.Label(age_frame, text="days").grid(row=0, column=8, padx=(0, 12))

        ttk.Label(age_frame, text="CASA:").grid(row=0, column=9, sticky="w", padx=(0, 4))
        ttk.Spinbox(age_frame, from_=0, to=30, width=4, textvariable=casa_age_var).grid(row=0, column=10, padx=(0, 4))
        ttk.Label(age_frame, text="days").grid(row=0, column=11, padx=(0, 12))

        ttk.Label(age_frame, text="OpenSky:").grid(row=0, column=12, sticky="w", padx=(0, 4))
        ttk.Spinbox(age_frame, from_=0, to=30, width=4, textvariable=opensky_age_var).grid(row=0, column=13, padx=(0, 4))
        ttk.Label(age_frame, text="days").grid(row=0, column=14)

        # Merge order
        ttk.Label(frame, text="Database Priority (most trusted source at top):").grid(
            row=12, column=0, sticky="w", pady=(0, 2))

        order_frame = ttk.Frame(frame)
        order_frame.grid(row=13, column=0, sticky="w", pady=(0, 12))

        order_listbox = tk.Listbox(order_frame, height=6, width=25,
                                   font=("Segoe UI", 10), selectmode="single")
        order_listbox.pack(side="left", padx=(0, 8))
        for item in self.settings.merge_order:
            order_listbox.insert("end", item)
        order_listbox.select_set(0)

        btn_order_frame = ttk.Frame(order_frame)
        btn_order_frame.pack(side="left")

        def move_up():
            sel = order_listbox.curselection()
            if not sel or sel[0] == 0:
                return
            idx = sel[0]
            val = order_listbox.get(idx)
            order_listbox.delete(idx)
            order_listbox.insert(idx - 1, val)
            order_listbox.select_set(idx - 1)

        def move_down():
            sel = order_listbox.curselection()
            if not sel or sel[0] >= order_listbox.size() - 1:
                return
            idx = sel[0]
            val = order_listbox.get(idx)
            order_listbox.delete(idx)
            order_listbox.insert(idx + 1, val)
            order_listbox.select_set(idx + 1)

        ttk.Button(btn_order_frame, text="\u25B2 Up", width=8, command=move_up).pack(pady=(0, 4))
        ttk.Button(btn_order_frame, text="\u25BC Down", width=8, command=move_down).pack()

        def save_and_close():
            self.settings.faa_url = faa_var.get().strip()
            self.settings.ccar_url = ccar_var.get().strip()
            self.settings.opensky_url = opensky_var.get().strip()
            self.settings.nz_caa_url = nzcaa_var.get().strip()
            self.settings.casa_url = casa_var.get().strip()
            self.settings.faa_max_age_days = faa_age_var.get()
            self.settings.ccar_max_age_days = ccar_age_var.get()
            self.settings.opensky_max_age_days = opensky_age_var.get()
            self.settings.nz_caa_max_age_days = nzcaa_age_var.get()
            self.settings.casa_max_age_days = casa_age_var.get()
            self.settings.merge_order = list(order_listbox.get(0, "end"))
            self._save_settings_json()
            self._refresh_db_status()
            win.destroy()

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=14, column=0, sticky="e")
        ttk.Button(btn_frame, text="OK", command=save_and_close).pack(side="right", padx=(4, 0))
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side="right")

    def _show_options(self):
        """Show options dialog."""
        win = tk.Toplevel(self.root)
        win.title("Options")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        frame = ttk.Frame(win, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Logging", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 4))

        log_var = tk.BooleanVar(value=self.chk_log_errors_var.get())
        ttk.Checkbutton(frame, text="Write errors to log file", variable=log_var).grid(
            row=1, column=0, sticky="w", padx=(8, 0))

        activity_var = tk.BooleanVar(value=self.chk_log_activity_var.get())
        ttk.Checkbutton(frame, text="Write activity to log file", variable=activity_var).grid(
            row=2, column=0, sticky="w", padx=(8, 0))

        ttk.Label(frame, text="Log files are saved to the working folder",
                  font=("Segoe UI", 8)).grid(row=3, column=0, sticky="w", padx=(24, 0), pady=(0, 12))

        def save_and_close():
            self.chk_log_errors_var.set(log_var.get())
            self.chk_log_activity_var.set(activity_var.get())
            self._save_settings_json()
            win.destroy()

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=4, column=0, sticky="e")
        ttk.Button(btn_frame, text="OK", command=save_and_close).pack(side="right", padx=(4, 0))
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side="right")

    def _open_error_log(self):
        """Open an error log file for this run if logging is enabled."""
        if not self.chk_log_errors_var.get():
            return
        from datetime import datetime
        timestamp = datetime.now().strftime("%d%m%Y-%H%M")
        log_dir = self.work_dir_var.get()
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, f"error_log_{timestamp}.txt")
        try:
            self._error_log_file = open(log_path, "w", encoding="utf-8")
            self._error_log_file.write(f"E's VRS Database Updater - Error Log\n")
            self._error_log_file.write(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            self._error_log_file.write("=" * 60 + "\n\n")
            self._log_status(f"  Error log: {log_path}")
        except OSError:
            self._error_log_file = None

    def _log_error(self, msg: str):
        """Write an error message to the log file if open."""
        if self._error_log_file:
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            try:
                self._error_log_file.write(f"[{timestamp}] {msg}\n")
                self._error_log_file.flush()
            except OSError:
                pass

    def _close_error_log(self):
        """Close the error log file."""
        if self._error_log_file:
            try:
                from datetime import datetime
                self._error_log_file.write(f"\nCompleted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                self._error_log_file.close()
            except OSError:
                pass
            self._error_log_file = None

    def _open_activity_log(self):
        """Open an activity log file for this run if logging is enabled."""
        if not self.chk_log_activity_var.get():
            return
        from datetime import datetime
        timestamp = datetime.now().strftime("%d%m%Y-%H%M")
        log_dir = self.work_dir_var.get()
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, f"activity_log_{timestamp}.txt")
        try:
            self._activity_log_file = open(log_path, "w", encoding="utf-8")
            self._activity_log_file.write(f"E's VRS Database Updater - Activity Log\n")
            self._activity_log_file.write(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            self._activity_log_file.write("=" * 60 + "\n\n")
            self._log_status(f"  Activity log: {log_path}")
        except OSError:
            self._activity_log_file = None

    def _log_activity(self, msg: str):
        """Write a message to the activity log if open."""
        if not self._activity_log_file:
            return
        # During merge phase, only log errors/warnings
        if self._in_merge_phase:
            upper = msg.strip().upper()
            if not ("ERROR" in upper or "WARNING" in upper or "FAILED" in upper):
                return
        try:
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            self._activity_log_file.write(f"[{timestamp}] {msg}\n")
            self._activity_log_file.flush()
        except OSError:
            pass

    def _close_activity_log(self):
        """Close the activity log file."""
        if self._activity_log_file:
            try:
                from datetime import datetime
                self._activity_log_file.write(f"\nCompleted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                self._activity_log_file.close()
            except OSError:
                pass
            self._activity_log_file = None

    def _show_about(self):
        messagebox.showinfo(
            "About",
            "E's VRS Database Updater\n"
            "Python Edition v2.0\n\n"
            "https://github.com/egite/E-s-VRS-Database-Updater"
        )

    def _show_features(self):
        win = tk.Toplevel(self.root)
        win.title("Program Features")
        win.geometry("520x540")
        win.resizable(False, False)
        win.configure(bg="#1E1E1E")
        try:
            icon_path = os.path.join(_DATA_DIR, "database.ico")
            if os.path.exists(icon_path):
                win.iconbitmap(icon_path)
        except Exception:
            pass

        header = tk.Label(win, text="E's VRS Database Updater — Features",
                          font=("Segoe UI", 12, "bold"), bg="#1565C0", fg="white",
                          pady=8)
        header.pack(fill="x")

        txt = tk.Text(win, wrap="word", bg="#1E1E1E", fg="#D4D4D4",
                      font=("Consolas", 10), bd=0, padx=12, pady=8,
                      highlightthickness=0, cursor="arrow")
        txt.pack(fill="both", expand=True)

        txt.tag_configure("cat", foreground="#4FC3F7", font=("Consolas", 10, "bold"))
        txt.tag_configure("item", foreground="#D4D4D4")

        features = [
            ("Data Sources", [
                "FAA aircraft registration database (DEREG + MASTER + ACFTREF)",
                "CCAR Canadian Civil Aircraft Register (ASP.NET postback download)",
                "NZ CAA New Zealand aircraft register",
                "CASA Australian aircraft register (registration-based matching)",
                "OpenSky Network aircraft database",
                "Custom Rules.csv for operator/registration overrides",
                "Sils.csv for ICAO type code to silhouette mapping",
            ]),
            ("Downloads & Parsing", [
                "Concurrent FAA, CCAR, and OpenSky downloads",
                "Parallel database parsing with ThreadPoolExecutor",
                "Configurable database download URLs via Settings dialog",
                "Configurable download age limits (0-30 days) per database",
                "Automatic CCAR download via ASP.NET form postback",
                "Streaming CSV parsing (low memory footprint)",
            ]),
            ("Performance Optimizations", [
                "O(1) dict-based silhouette lookups (vs linear scan)",
                "Pre-loaded VRS database dict for O(1) ICAO lookups",
                "Parameterized SQL queries (no string concatenation)",
                "Batched transactions (10,000 records per commit)",
                "SQLite WAL mode and indexed ICAO columns",
                "Pre-compiled regex patterns for text cleaning",
                "CCAR owner lookup via dict (O(1) vs linear scan)",
            ]),
            ("User Interface", [
                "Dark terminal theme with color-coded output",
                "Two-panel layout: Program Progress + Database Update Status",
                "Banner with data source icons (FAA, Transport Canada, OpenSky)",
                "Real-time disk throughput indicator",
                "Aggregate progress bar for concurrent operations",
                "ETA display in h:mm:ss format",
                "Cancel button to stop processing mid-stream",
                "Phase indicator showing current processing step",
                "Batched detail log updates (500ms flush) for UI performance",
                "Search VRS database by ICAO, registration, operator, etc.",
                "Export search results to CSV",
            ]),
            ("Settings & Persistence", [
                "JSON-based settings persistence (window size, position, all options)",
                "Configurable FAA, CCAR, and OpenSky download URLs",
                "Configurable download age limits (0-30 days) per database",
                "Skip Processing: selectively skip FAA, CCAR, OpenSky, or Rules",
                "Configurable database priority order (FAA, CCAR, OpenSky, Rules)",
                "Remembers working directory and VRS directory paths",
                "Saves/restores all checkboxes and options between sessions",
                "Optional error log file (DDMMYYYY-HHMM timestamped, saved to working folder)",
                "Optional activity log file (captures download and build steps)",
            ]),
            ("Data Processing", [
                "Title case normalization with abbreviation handling (LLC, LLP, etc.)",
                "Operator name cleaning (Mc prefix, & patterns, embedded registrations)",
                "Non-ASCII character removal and SQL escaping",
                "CCAR binary-to-hex ICAO address conversion",
                "Merges FAA + CCAR + OpenSky + Rules into VRS AircraftOnlineLookupCache.sqb",
            ]),
        ]

        for cat, items in features:
            txt.insert("end", f"\n  {cat}\n", "cat")
            for item in items:
                txt.insert("end", f"    \u2022 {item}\n", "item")

        txt.config(state="disabled")

        btn = tk.Button(win, text="Close", command=win.destroy,
                        bg="#333333", fg="white", activebackground="#555555",
                        activeforeground="white", font=("Segoe UI", 10),
                        relief="flat", padx=20, pady=4)
        btn.pack(pady=10)

    # ------------------------------------------------------------------
    # Search / Export dialog
    # ------------------------------------------------------------------
    def _show_search(self):
        """Open a search dialog for the VRS database."""
        vrs_db = os.path.join(self.work_dir_var.get(), "AircraftOnlineLookupCache.sqb")
        if not os.path.exists(vrs_db):
            messagebox.showwarning("Search", "VRS database not found in the working folder.")
            return

        win = tk.Toplevel(self.root)
        win.title("Search VRS Database")
        win.geometry("920x560")
        win.resizable(True, True)
        win.minsize(700, 400)
        win.configure(bg="#1E1E1E")
        win.transient(self.root)
        try:
            icon_path = os.path.join(_DATA_DIR, "database.ico")
            if os.path.exists(icon_path):
                win.iconbitmap(icon_path)
        except Exception:
            pass

        # --- Search bar ---
        search_frame = tk.Frame(win, bg="#263238", padx=8, pady=8)
        search_frame.pack(fill="x")

        tk.Label(search_frame, text="Search:", bg="#263238", fg="#B0BEC5",
                 font=("Segoe UI", 10)).pack(side="left")

        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30,
                                font=("Consolas", 11), bg="#37474F", fg="#FFFFFF",
                                insertbackground="#FFFFFF", relief="flat")
        search_entry.pack(side="left", padx=(8, 4), fill="x", expand=True)

        field_var = tk.StringVar(value="Any")
        field_combo = ttk.Combobox(search_frame, textvariable=field_var, width=14,
                                   values=["Any", "ICAO", "Registration", "Operator",
                                           "Manufacturer", "Model", "Country"],
                                   state="readonly")
        field_combo.pack(side="left", padx=4)

        result_count_var = tk.StringVar(value="")
        count_label = tk.Label(search_frame, textvariable=result_count_var,
                               bg="#263238", fg="#4FC3F7", font=("Segoe UI", 9))
        count_label.pack(side="right", padx=(8, 0))

        # --- Results treeview ---
        columns = ("Icao", "Registration", "Country", "Manufacturer", "Model",
                   "ModelIcao", "Operator", "OperatorIcao", "YearBuilt")
        col_widths = {
            "Icao": 65, "Registration": 85, "Country": 90, "Manufacturer": 130,
            "Model": 110, "ModelIcao": 75, "Operator": 170, "OperatorIcao": 80,
            "YearBuilt": 55,
        }

        tree_frame = tk.Frame(win, bg="#1E1E1E")
        tree_frame.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        style = ttk.Style()
        style.configure("Search.Treeview",
                         background="#1E1E1E", foreground="#D4D4D4",
                         fieldbackground="#1E1E1E", font=("Consolas", 9),
                         rowheight=20)
        style.configure("Search.Treeview.Heading",
                         background="#37474F", foreground="#B0BEC5",
                         font=("Segoe UI", 9, "bold"))
        style.map("Search.Treeview",
                   background=[("selected", "#264F78")],
                   foreground=[("selected", "#FFFFFF")])

        tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                            style="Search.Treeview")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths.get(col, 80), minwidth=40)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        # --- Bottom buttons ---
        btn_frame = tk.Frame(win, bg="#1E1E1E", pady=6)
        btn_frame.pack(fill="x", padx=8)

        export_btn = tk.Button(btn_frame, text="  Export CSV  ", command=lambda: self._export_results(tree, columns),
                               bg="#1565C0", fg="white", font=("Segoe UI", 10, "bold"),
                               activebackground="#1976D2", relief="flat")
        export_btn.pack(side="left")

        tk.Button(btn_frame, text="  Close  ", command=win.destroy,
                  bg="#333333", fg="white", font=("Segoe UI", 10),
                  activebackground="#555555", relief="flat").pack(side="right")

        # --- Search logic ---
        _search_results = []  # store for export

        def do_search(*_args):
            query = search_var.get().strip()
            field = field_var.get()
            if not query:
                tree.delete(*tree.get_children())
                result_count_var.set("")
                _search_results.clear()
                return

            try:
                conn = sqlite3.connect(vrs_db)
                conn.row_factory = sqlite3.Row

                if field == "Any":
                    sql = """SELECT Icao, Registration, Country, Manufacturer, Model,
                                    ModelIcao, Operator, OperatorIcao, YearBuilt
                             FROM AircraftDetail
                             WHERE Icao LIKE ? OR Registration LIKE ? OR Operator LIKE ?
                                OR Manufacturer LIKE ? OR Model LIKE ? OR Country LIKE ?
                             LIMIT 500"""
                    like = f"%{query}%"
                    rows = conn.execute(sql, (like, like, like, like, like, like)).fetchall()
                else:
                    col_map = {"ICAO": "Icao", "Registration": "Registration",
                               "Operator": "Operator", "Manufacturer": "Manufacturer",
                               "Model": "Model", "Country": "Country"}
                    col_name = col_map.get(field, "Icao")
                    sql = f"""SELECT Icao, Registration, Country, Manufacturer, Model,
                                     ModelIcao, Operator, OperatorIcao, YearBuilt
                              FROM AircraftDetail
                              WHERE {col_name} LIKE ?
                              LIMIT 500"""
                    rows = conn.execute(sql, (f"%{query}%",)).fetchall()

                conn.close()

                tree.delete(*tree.get_children())
                _search_results.clear()
                for row in rows:
                    vals = tuple(str(row[c] or "") for c in columns)
                    tree.insert("", "end", values=vals)
                    _search_results.append(vals)

                n = len(rows)
                result_count_var.set(f"{n:,} result{'s' if n != 1 else ''}" +
                                     (" (limit 500)" if n == 500 else ""))
            except Exception as e:
                result_count_var.set(f"Error: {e}")

        search_var.trace_add("write", do_search)
        field_combo.bind("<<ComboboxSelected>>", do_search)
        search_entry.focus_set()

    @staticmethod
    def _export_results(tree, columns):
        """Export current search results to CSV."""
        items = tree.get_children()
        if not items:
            messagebox.showinfo("Export", "No results to export.")
            return

        path = filedialog.asksaveasfilename(
            title="Export Results",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"vrs_search_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not path:
            return

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(columns)
                for item in items:
                    writer.writerow(tree.item(item, "values"))
            messagebox.showinfo("Export", f"Exported {len(items):,} records to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def _on_cancel(self):
        if self._countdown_id:
            # Cancel auto-start countdown
            self._cancelled.set()
            return
        if self.running:
            self._cancelled.set()
            self._log_status("")
            self._log_status("*** Cancelling... please wait ***")
            self.btn_cancel.config(state="disabled")

    def _on_exit(self):
        self._save_settings_json()
        self._cancelled.set()
        self.want_exit = True
        self.root.destroy()

    def _start_countdown(self, seconds: int):
        """Show a countdown in the Program Progress box on a single updating line."""
        self._countdown_remaining = seconds
        self.btn_cancel.config(state="normal")
        self.phase_var.set("Countdown...")
        self.txt_status.config(state="normal")
        self.txt_status.delete("1.0", "end")
        self.txt_status.insert("1.0", "Run on Start is enabled", "info")
        self.txt_status.insert("end", "\n\n")
        self.txt_status.insert("end", "Click Cancel to abort", "error")
        self.txt_status.insert("end", "\n\n")
        self.txt_status.mark_set("countdown_start", "end-1c linestart")
        self.txt_status.mark_gravity("countdown_start", "left")
        self.txt_status.insert("countdown_start", f"Starting in {seconds}...", "info")
        self.txt_status.see("end")
        self.txt_status.config(state="disabled")
        self._countdown_tick()

    def _countdown_tick(self):
        """Update the countdown number on the same line each second."""
        if self._cancelled.is_set():
            self.txt_status.config(state="normal")
            self.txt_status.delete("countdown_start", "end-1c")
            self.txt_status.insert("countdown_start", "Auto-start cancelled.", "warning")
            self.txt_status.config(state="disabled")
            self.btn_cancel.config(state="disabled")
            self._cancelled.clear()
            self._countdown_id = None
            self.phase_var.set("Idle")
            return
        if self._countdown_remaining <= 0:
            self.txt_status.config(state="normal")
            self.txt_status.delete("countdown_start", "end-1c")
            self.txt_status.insert("countdown_start", "Starting now!", "success")
            self.txt_status.config(state="disabled")
            self.btn_cancel.config(state="disabled")
            self._countdown_id = None
            self._auto_run = True
            self._on_start()
            return
        # Replace the countdown line in place
        self.txt_status.config(state="normal")
        self.txt_status.delete("countdown_start", "end-1c")
        self.txt_status.insert("countdown_start", f"Starting in {self._countdown_remaining}...", "info")
        self.txt_status.config(state="disabled")
        self._countdown_remaining -= 1
        self._countdown_id = self.root.after(1000, self._countdown_tick)

    def _on_start(self):
        if self.running:
            return
        # Cancel any pending countdown
        if self._countdown_id:
            self.root.after_cancel(self._countdown_id)
            self._countdown_id = None

        # Validate
        settings = self._build_settings()
        if settings.vrs_dir and settings.vrs_dir == settings.work_dir:
            messagebox.showerror("Error",
                                 "Working folder and VRS folder cannot be the same location.\n"
                                 "Fix this and run again.")
            return

        os.makedirs(settings.work_dir, exist_ok=True)

        # Disable controls during run
        self.running = True
        self._cancelled.clear()
        self.btn_start.config(state="disabled", bg="#999999")
        self.btn_cancel.config(state="normal")
        self._clear_logs()

        # Start disk throughput polling
        self._start_disk_poll()

        # Run in background thread
        thread = threading.Thread(target=self._run_updater, args=(settings,), daemon=True)
        thread.start()

    def _clear_logs(self):
        for w in (self.txt_status, self.txt_details):
            w.config(state="normal")
            w.delete("1.0", "end")
            w.config(state="disabled")

    # ------------------------------------------------------------------
    # Main processing (runs in background thread)
    # ------------------------------------------------------------------
    def _run_updater(self, settings: Settings):
        """Execute the full update pipeline, posting UI updates via root.after."""
        import time
        from concurrent.futures import ThreadPoolExecutor, as_completed
        from .ccar import download_ccar
        import os

        # Monkey-patch the progress reporting so it updates the GUI
        self._patch_progress()

        time_started = datetime.now().strftime("%H:%M")
        start_time = time.time()
        self._open_error_log()
        self._open_activity_log()
        self._log_status(f"Started at: {time_started}")

        # ---- Concurrent Downloads ----
        self._log_status("")
        self._set_phase("Downloading databases")
        self._log_status("[Step 1/4] Downloading databases...")
        self._log_status("-" * 40)

        download_tasks = {}
        self._task_progress = {}
        self._task_eta = {}
        self._thread_task_map = {}
        self._aggregate_mode = True

        def _tracked_call(task_name, func, *args):
            """Wrapper that registers the thread ID to the parent task name."""
            self._thread_task_map[threading.current_thread().ident] = task_name
            return func(*args)

        def _db_is_fresh(db_path, max_age_days):
            """Return True if the local DB exists and is younger than max_age_days."""
            if max_age_days <= 0 or not os.path.exists(db_path):
                return False
            age_days = (time.time() - os.path.getmtime(db_path)) / 86400
            return age_days < max_age_days

        fresh_skipped = set()   # databases skipped because local copy is still valid

        with ThreadPoolExecutor(max_workers=4) as pool:
            if settings.download_faa and not settings.skip_faa:
                if _db_is_fresh(settings.faa_db_path, settings.faa_max_age_days):
                    age = (time.time() - os.path.getmtime(settings.faa_db_path)) / 86400
                    remaining = settings.faa_max_age_days - age
                    self._log_status(f"  FAA database has {remaining:.0f} days of validity remaining, skipping download.")
                    fresh_skipped.add("FAA")
                else:
                    self._task_progress["FAA"] = 0
                    download_tasks["FAA"] = pool.submit(_tracked_call, "FAA", download_faa, settings)
            elif settings.skip_faa and settings.download_faa:
                self._log_status("  FAA download skipped.")
            if settings.download_ccar and not settings.skip_ccar:
                if _db_is_fresh(settings.ccar_db_path, settings.ccar_max_age_days):
                    age = (time.time() - os.path.getmtime(settings.ccar_db_path)) / 86400
                    remaining = settings.ccar_max_age_days - age
                    self._log_status(f"  CCAR database has {remaining:.0f} days of validity remaining, skipping download.")
                    fresh_skipped.add("CCAR")
                else:
                    self._task_progress["CCAR"] = 0
                    zip_path = os.path.join(settings.work_dir, "ccarcsdb.zip")
                    download_tasks["CCAR"] = pool.submit(_tracked_call, "CCAR", download_ccar, zip_path)
            elif settings.skip_ccar and settings.download_ccar:
                self._log_status("  CCAR download skipped.")
            if settings.download_nz_caa and not settings.skip_nz_caa:
                if _db_is_fresh(settings.nz_caa_db_path, settings.nz_caa_max_age_days):
                    age = (time.time() - os.path.getmtime(settings.nz_caa_db_path)) / 86400
                    remaining = settings.nz_caa_max_age_days - age
                    self._log_status(f"  NZ CAA database has {remaining:.0f} days of validity remaining, skipping download.")
                    fresh_skipped.add("NZ CAA")
                else:
                    self._task_progress["NZ CAA"] = 0
                    download_tasks["NZ CAA"] = pool.submit(_tracked_call, "NZ CAA", download_nz_caa, settings)
            elif settings.skip_nz_caa and settings.download_nz_caa:
                self._log_status("  NZ CAA download skipped.")
            if settings.download_casa and not settings.skip_casa:
                if _db_is_fresh(settings.casa_db_path, settings.casa_max_age_days):
                    age = (time.time() - os.path.getmtime(settings.casa_db_path)) / 86400
                    remaining = settings.casa_max_age_days - age
                    self._log_status(f"  CASA database has {remaining:.0f} days of validity remaining, skipping download.")
                    fresh_skipped.add("CASA")
                else:
                    self._task_progress["CASA"] = 0
                    download_tasks["CASA"] = pool.submit(_tracked_call, "CASA", download_casa, settings)
            elif settings.skip_casa and settings.download_casa:
                self._log_status("  CASA download skipped.")
            if settings.download_opensky and not settings.skip_opensky:
                if _db_is_fresh(settings.opensky_db_path, settings.opensky_max_age_days):
                    age = (time.time() - os.path.getmtime(settings.opensky_db_path)) / 86400
                    remaining = settings.opensky_max_age_days - age
                    self._log_status(f"  OpenSky database has {remaining:.0f} days of validity remaining, skipping download.")
                    fresh_skipped.add("OpenSky")
                else:
                    self._task_progress["OpenSky"] = 0
                    download_tasks["OpenSky"] = pool.submit(_tracked_call, "OpenSky", download_opensky, settings)
            elif settings.skip_opensky and settings.download_opensky:
                self._log_status("  OpenSky download skipped.")

            if not download_tasks:
                self._log_status("  No downloads requested, using existing data.")

            for future in as_completed(download_tasks.values()):
                # Results are logged via print patch as they complete
                pass

        self._aggregate_mode = False
        self._task_progress = {}
        self._task_eta = {}
        self._thread_task_map = {}

        if self._cancelled.is_set():
            self._log_status(""); self._log_status("*** Cancelled ***")
            self._log_error("Run cancelled by user")
            self._close_error_log()
            self._close_activity_log()
            self._reset_progress(); self.root.after(0, self._finish_run); return

        # Collect results (use exception handling in case cancel interrupted a download)
        def _safe_result(key):
            f = download_tasks.get(key)
            if not f:
                return False
            try:
                return f.result()
            except Exception as e:
                # Catch-all safety net: anything a download task raised that
                # wasn't already logged with its reason (extraction, move, an
                # unexpected error) is surfaced here instead of vanishing.
                self._log_status(f"  ERROR: {key} download/extract failed: {e}")
                self._log_error(f"Download: {key} exception: {e}")
                return False

        faa_downloaded = _safe_result("FAA")
        ccar_downloaded = _safe_result("CCAR")
        nzcaa_downloaded = _safe_result("NZ CAA")
        casa_downloaded = _safe_result("CASA")
        opensky_downloaded = _safe_result("OpenSky")

        if settings.download_faa and not faa_downloaded and "FAA" not in fresh_skipped:
            self._log_status("  WARNING: FAA download failed. Will try existing data.")
            self._log_error("Download: FAA download failed")
        if settings.download_ccar and not ccar_downloaded and "CCAR" not in fresh_skipped:
            self._log_status("  WARNING: CCAR download failed. Will try existing data.")
            self._log_error("Download: CCAR download failed")
        if settings.download_nz_caa and not nzcaa_downloaded and "NZ CAA" not in fresh_skipped:
            self._log_status("  WARNING: NZ CAA download failed. Will try existing data.")
            self._log_error("Download: NZ CAA download failed")
        if settings.download_casa and not casa_downloaded and "CASA" not in fresh_skipped:
            self._log_status("  WARNING: CASA download failed. Will try existing data.")
            self._log_error("Download: CASA download failed")
        if settings.download_opensky and not opensky_downloaded and "OpenSky" not in fresh_skipped:
            self._log_status("  WARNING: OpenSky download failed. Will try existing data.")
            self._log_error("Download: OpenSky download failed")

        if self._cancelled.is_set():
            self._log_status(""); self._log_status("*** Cancelled ***")
            self._log_error("Run cancelled by user")
            self._close_error_log()
            self._close_activity_log()
            self._reset_progress(); self.root.after(0, self._finish_run); return

        # ---- Step 2: Parse databases (concurrent) ----
        self._log_status("")
        self._set_phase("Building info databases")
        self._log_status("[Step 2/4] Parsing databases...")
        self._log_status("-" * 40)

        # CCAR download already handled above, don't re-download
        settings.download_ccar = False

        parse_tasks = {}
        self._task_progress = {}
        self._task_eta = {}
        self._thread_task_map = {}
        self._aggregate_mode = True

        with ThreadPoolExecutor(max_workers=4) as pool:
            if settings.skip_faa:
                self._log_status("  FAA - skipped.")
            elif settings.download_faa or not os.path.exists(settings.faa_db_path):
                self._task_progress["FAA"] = 0
                parse_tasks["FAA"] = pool.submit(_tracked_call, "FAA", parse_faa, settings)
            else:
                self._log_status("  FAA - using existing database.")

            if settings.skip_ccar:
                self._log_status("  CCAR - skipped.")
            else:
                self._task_progress["CCAR"] = 0
                parse_tasks["CCAR"] = pool.submit(_tracked_call, "CCAR", parse_ccar, settings)

            if settings.skip_nz_caa:
                self._log_status("  NZ CAA - skipped.")
            elif settings.download_nz_caa or not os.path.exists(settings.nz_caa_db_path):
                self._task_progress["NZ CAA"] = 0
                parse_tasks["NZ CAA"] = pool.submit(_tracked_call, "NZ CAA", parse_nz_caa, settings)
            else:
                self._log_status("  NZ CAA - using existing database.")

            if settings.skip_casa:
                self._log_status("  CASA - skipped.")
            elif settings.download_casa or not os.path.exists(settings.casa_db_path):
                self._task_progress["CASA"] = 0
                parse_tasks["CASA"] = pool.submit(_tracked_call, "CASA", parse_casa, settings)
            else:
                self._log_status("  CASA - using existing database.")

            if settings.skip_opensky:
                self._log_status("  OpenSky - skipped.")
            elif settings.download_opensky or not os.path.exists(settings.opensky_db_path):
                self._task_progress["OpenSky"] = 0
                parse_tasks["OpenSky"] = pool.submit(_tracked_call, "OpenSky", parse_opensky, settings)
            else:
                self._log_status("  OpenSky - using existing database.")

            for future in as_completed(parse_tasks.values()):
                pass

        self._aggregate_mode = False
        self._task_progress = {}
        self._task_eta = {}
        self._thread_task_map = {}

        if parse_tasks.get("FAA"):
            try:
                if not parse_tasks["FAA"].result():
                    self._log_status("  WARNING: FAA parse failed.")
                    self._log_error("Parse: FAA parse failed")
            except Exception as e:
                self._log_status("  WARNING: FAA parse failed.")
                self._log_error(f"Parse: FAA parse exception: {e}")
        if parse_tasks.get("NZ CAA"):
            try:
                if not parse_tasks["NZ CAA"].result():
                    self._log_status("  WARNING: NZ CAA parse failed.")
                    self._log_error("Parse: NZ CAA parse failed")
            except Exception as e:
                self._log_status("  WARNING: NZ CAA parse failed.")
                self._log_error(f"Parse: NZ CAA parse exception: {e}")
        if parse_tasks.get("CASA"):
            try:
                if not parse_tasks["CASA"].result():
                    self._log_status("  WARNING: CASA parse failed.")
                    self._log_error("Parse: CASA parse failed")
            except Exception as e:
                self._log_status("  WARNING: CASA parse failed.")
                self._log_error(f"Parse: CASA parse exception: {e}")
        if parse_tasks.get("OpenSky"):
            try:
                if not parse_tasks["OpenSky"].result():
                    self._log_status("  WARNING: OpenSky parse failed.")
                    self._log_error("Parse: OpenSky parse failed")
            except Exception as e:
                self._log_status("  WARNING: OpenSky parse failed.")
                self._log_error(f"Parse: OpenSky parse exception: {e}")

        if self._cancelled.is_set():
            self._log_status(""); self._log_status("*** Cancelled ***")
            self._log_error("Run cancelled by user")
            self._close_error_log()
            self._close_activity_log()
            self._reset_progress(); self.root.after(0, self._finish_run); return

        # ---- Step 3: VRS Merge ----
        self._log_status("")
        self._set_phase("Processing: Merge")
        self._log_status("[Step 3/4] VRS Database Merge")
        self._log_status("-" * 40)
        self._in_merge_phase = True
        try:
            update_vrs(settings)
        except Exception as e:
            self._log_status(f"  ERROR: VRS merge failed: {e}")
            self._log_error(f"Merge: VRS merge exception: {e}")
        self._in_merge_phase = False

        # Done
        elapsed = time.time() - start_time
        if elapsed > 60:
            elapsed_str = f"{elapsed / 60:.1f} minutes"
        else:
            elapsed_str = f"{elapsed:.1f} seconds"

        self._log_status("")
        self._log_status("=" * 50)
        self._log_status(f"Done! Started: {time_started}  Ended: {datetime.now().strftime('%H:%M')}")
        self._log_status(f"Total time: {elapsed_str}")
        self._log_status("=" * 50)
        self._log_detail("Done.")
        self._close_error_log()
        self._close_activity_log()

        self._reset_progress()
        self.root.after(0, self._finish_run)

    def _finish_run(self):
        self.running = False
        self._stop_disk_poll()
        self.throughput_var.set("")
        self.phase_var.set("Idle")
        self.btn_start.config(state="normal", bg="#4CAF50")
        self.btn_cancel.config(state="disabled")
        if not self._cancelled.is_set():
            self.percent_var.set("100%")
        self._refresh_db_status()
        # Auto-exit countdown if this was an auto-run
        if self._auto_run and not self._cancelled.is_set():
            self._auto_run = False
            self._start_exit_countdown(10)

    def _start_exit_countdown(self, seconds: int):
        """Show an exit countdown in Program Progress, then close the app."""
        self._exit_countdown_remaining = seconds
        self._append_colored(self.txt_status, "")
        self._append_colored(self.txt_status, "")
        self.txt_status.config(state="normal")
        self.txt_status.mark_set("exit_countdown", "end-1c")
        self.txt_status.mark_gravity("exit_countdown", "left")
        self.txt_status.insert("end", f"Exiting in {seconds}...\n", "info")
        self.txt_status.see("end")
        self.txt_status.config(state="disabled")
        self._exit_countdown_id = self.root.after(1000, self._exit_countdown_tick)

    def _exit_countdown_tick(self):
        """Update exit countdown each second, then close."""
        self._exit_countdown_remaining -= 1
        if self._exit_countdown_remaining <= 0:
            self.txt_status.config(state="normal")
            self.txt_status.delete("exit_countdown", "end")
            self.txt_status.insert("end", "Exiting now.\n", "info")
            self.txt_status.config(state="disabled")
            self._save_settings_json()
            self.root.destroy()
            return
        self.txt_status.config(state="normal")
        self.txt_status.delete("exit_countdown", "end")
        self.txt_status.insert("end", f"Exiting in {self._exit_countdown_remaining}...\n", "info")
        self.txt_status.see("end")
        self.txt_status.config(state="disabled")
        self._exit_countdown_id = self.root.after(1000, self._exit_countdown_tick)

    def _patch_progress(self):
        """Redirect print(), ProgressReporter, and detail_log to the GUI."""
        import builtins
        from .utils import ProgressReporter
        from . import vrs_merge

        original_print = builtins.print
        gui = self

        # Patch detail_log to route to Details tab
        def gui_detail_log(msg):
            gui._log_detail(msg)
        vrs_merge.detail_log = gui_detail_log

        # Patch phase_callback to update phase indicator
        def gui_phase(phase):
            gui._set_phase(phase)
        vrs_merge.phase_callback = gui_phase

        def gui_print(*args, **kwargs):
            msg = " ".join(str(a) for a in args)
            # Still print to console
            original_print(*args, **kwargs)
            # Also send to GUI status log
            if msg.strip():
                gui._log_status(msg)
                # Log errors/warnings to file
                upper = msg.strip().upper()
                if "ERROR" in upper or "WARNING" in upper or "FAILED" in upper:
                    gui._log_error(msg.strip())

        builtins.print = gui_print

        # Patch ProgressReporter to update GUI progress
        original_update = ProgressReporter.update
        original_done = ProgressReporter.done

        def gui_update(self_pr, current, total):
            if gui._cancelled.is_set():
                raise InterruptedError("Cancelled by user")
            if total <= 0:
                return
            import time as _time
            pct = 100 * current / total
            if hasattr(self_pr, '_last_gui_pct') and int(pct) == self_pr._last_gui_pct:
                return
            self_pr._last_gui_pct = int(pct)
            elapsed = _time.time() - self_pr._start_time
            if pct > 0:
                eta_sec = elapsed / (pct / 100) * (1 - pct / 100)
            else:
                eta_sec = -1
            gui._set_progress(pct, eta_sec, task_name=self_pr.task_name)

        def gui_done(self_pr):
            original_done(self_pr)
            # Mark this task as 100% in aggregate mode, otherwise reset
            if gui._aggregate_mode:
                parent = gui._thread_task_map.get(threading.current_thread().ident, self_pr.task_name)
                gui._task_progress[parent] = 100
            else:
                gui._reset_progress()

        ProgressReporter.update = gui_update
        ProgressReporter.done = gui_done


def run_gui():
    """Launch the GUI application."""
    root = tk.Tk()
    app = VRSUpdaterApp(root)
    root.mainloop()
