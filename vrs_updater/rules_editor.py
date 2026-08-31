"""
Integrated Rules.csv editor.

Replaces hand-editing Rules.csv in Excel. Adds three things a spreadsheet
cannot do: it understands the two-block match/change layout, it validates
rules against the format's real semantics (first-match-wins, "!" negation),
and it can test a rule against the live VRS database before you save it.
"""

import os
import queue
import sqlite3
import threading
import tkinter as tk
from tkinter import ttk, messagebox

from .rules import (Rule, FIELD_NAMES, MSG_FIELD_CHOICES, RULE_TO_DB,
                    load_rules, save_rules, validate_rule, validate_all)
from .ui_common import (BG, PANEL, ENTRY, ACCENT, BTN, TEXT, LABEL, CYAN,
                        DB_COLUMNS,
                        ORANGE, RED, GREEN, dark_window, dark_button,
                        dark_entry, dark_label, style_tree, center_on)

# ---------------------------------------------------------------------------
# Matching a rule against the live VRS database
# ---------------------------------------------------------------------------

def count_matches(db_path, rule, sample_limit=200):
    """Return (match_count, sample_rows) for `rule` against the VRS database.

    The match is expressed in SQL (case-insensitive, NULL treated as "")
    so the count is exact rather than a sampled estimate. Rows come back as
    dicts keyed by DB column name.
    """
    positives = {k: v for k, v in rule.match_fields.items() if not v.startswith('!')}
    if not positives:
        return 0, []  # a rule with no positive match can never fire

    where, params = [], []
    for name, val in rule.match_fields.items():
        col = RULE_TO_DB.get(name)
        if not col:
            continue
        if val.startswith('!'):
            where.append("IFNULL(%s,'') <> ? COLLATE NOCASE" % col)
            params.append(val[1:])
        else:
            where.append("IFNULL(%s,'') = ? COLLATE NOCASE" % col)
            params.append(val)

    clause = " AND ".join(where)
    conn = sqlite3.connect(db_path)
    try:
        conn.row_factory = sqlite3.Row
        count = conn.execute(
            "SELECT COUNT(*) FROM AircraftDetail WHERE %s" % clause,
            params).fetchone()[0]
        rows = conn.execute(
            "SELECT %s FROM AircraftDetail WHERE %s LIMIT %d"
            % (", ".join(DB_COLUMNS), clause, sample_limit),
            params).fetchall()
        return count, [dict(r) for r in rows]
    finally:
        conn.close()


def row_to_record(row):
    """Convert a VRS DB row dict into a rule-field record."""
    return {rk: (row.get(dk) or "") for rk, dk in RULE_TO_DB.items()}


# ---------------------------------------------------------------------------
# Rule edit dialog
# ---------------------------------------------------------------------------

class RuleDialog:
    """Modal editor for a single rule. `result` is a Rule, or None if cancelled."""

    def __init__(self, parent, rule=None, vrs_db=None, data_dir=None,
                 source_record=None, title="Edit Rule"):
        self.result = None
        self.vrs_db = vrs_db
        self.data_dir = data_dir

        self.win = tk.Toplevel(parent)
        dark_window(self.win, parent, title, "860x560", data_dir)
        self.win.minsize(760, 500)

        rule = rule or Rule()

        self.match_vars, self.not_vars, self.change_vars = {}, {}, {}

        if source_record:
            info = "  \u2022  ".join(
                str(source_record.get(c)) for c in
                ("Registration", "Icao", "Manufacturer", "Model", "Operator")
                if source_record.get(c))
            bar = tk.Frame(self.win, bg=PANEL, padx=10, pady=6)
            bar.pack(fill="x")
            dark_label(bar, "From aircraft:", bg=PANEL).pack(side="left")
            dark_label(bar, info, bg=PANEL, fg=CYAN,
                   font=("Consolas", 10)).pack(side="left", padx=(8, 0))

        body = tk.Frame(self.win, bg=BG, padx=10, pady=10)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1, uniform="cols")
        body.columnconfigure(1, weight=1, uniform="cols")
        body.rowconfigure(0, weight=1)

        # --- When (match) -------------------------------------------------
        when = tk.LabelFrame(body, text=" When the record matches ", bg=BG,
                             fg=CYAN, font=("Segoe UI", 10, "bold"),
                             padx=8, pady=8, relief="groove", bd=1)
        when.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        when.columnconfigure(1, weight=1)

        dark_label(when, "not").grid(row=0, column=2, padx=(6, 0))
        for i, name in enumerate(FIELD_NAMES):
            raw = rule.match_fields.get(name, "")
            negated = raw.startswith("!")
            var = tk.StringVar(value=raw[1:] if negated else raw)
            nvar = tk.BooleanVar(value=negated)
            self.match_vars[name] = var
            self.not_vars[name] = nvar

            dark_label(when, name).grid(row=i + 1, column=0, sticky="w", pady=3)
            e = dark_entry(when, var)
            e.grid(row=i + 1, column=1, sticky="ew", padx=(6, 0), pady=3)
            cb = tk.Checkbutton(when, variable=nvar, bg=BG, activebackground=BG,
                                selectcolor=ENTRY, relief="flat", bd=0,
                                highlightthickness=0)
            cb.grid(row=i + 1, column=2, padx=(6, 0))
            var.trace_add("write", self._revalidate)
            nvar.trace_add("write", self._revalidate)

        # --- Then (set) ---------------------------------------------------
        then = tk.LabelFrame(body, text=" Set these values ", bg=BG,
                             fg=GREEN, font=("Segoe UI", 10, "bold"),
                             padx=8, pady=8, relief="groove", bd=1)
        then.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        then.columnconfigure(1, weight=1)

        dark_label(then, " ").grid(row=0, column=1)
        for i, name in enumerate(FIELD_NAMES):
            var = tk.StringVar(value=rule.change_fields.get(name, ""))
            self.change_vars[name] = var
            dark_label(then, name).grid(row=i + 1, column=0, sticky="w", pady=3)
            e = dark_entry(then, var)
            e.grid(row=i + 1, column=1, sticky="ew", padx=(6, 0), pady=3)
            var.trace_add("write", self._revalidate)

        # --- Message ------------------------------------------------------
        msg = tk.Frame(self.win, bg=BG, padx=10)
        msg.pack(fill="x")
        dark_label(msg, "Log message").grid(row=0, column=0, sticky="w")
        self.msg_var = tk.StringVar(value=rule.msg_text)
        dark_entry(msg, self.msg_var, width=50).grid(row=0, column=1, sticky="ew", padx=6)
        dark_label(msg, "append field").grid(row=0, column=2, sticky="w", padx=(10, 0))
        self.msg_field_var = tk.StringVar(value=rule.msg_field)
        ttk.Combobox(msg, textvariable=self.msg_field_var, width=14,
                     values=MSG_FIELD_CHOICES, state="readonly").grid(row=0, column=3, padx=6)
        msg.columnconfigure(1, weight=1)

        # --- Validation feedback -----------------------------------------
        self.issue_var = tk.StringVar(value="")
        self.issue_label = tk.Label(self.win, textvariable=self.issue_var, bg=BG,
                                    fg=ORANGE, font=("Segoe UI", 9), anchor="w",
                                    justify="left", wraplength=820)
        self.issue_label.pack(fill="x", padx=12, pady=(8, 0))

        # --- Buttons ------------------------------------------------------
        btns = tk.Frame(self.win, bg=BG, padx=10, pady=10)
        btns.pack(fill="x")
        self.test_btn = dark_button(btns, "Test Against Database", self._on_test)
        self.test_btn.pack(side="left")
        if not (vrs_db and os.path.exists(vrs_db)):
            self.test_btn.config(state="disabled")
        dark_button(btns, "Cancel", self.win.destroy).pack(side="right")
        dark_button(btns, "OK", self._on_ok, primary=True).pack(side="right", padx=(0, 8))

        self.win.bind("<Escape>", lambda _e: self.win.destroy())
        self._revalidate()

        center_on(self.win, parent)
        self.win.grab_set()
        self.win.focus_set()
        parent.wait_window(self.win)

    # -- helpers -----------------------------------------------------------
    def _build_rule(self):
        match = {}
        for name in FIELD_NAMES:
            val = self.match_vars[name].get().strip()
            if val:
                match[name] = ("!" + val) if self.not_vars[name].get() else val
        change = {name: self.change_vars[name].get().strip()
                  for name in FIELD_NAMES if self.change_vars[name].get().strip()}
        return Rule(match, change, self.msg_field_var.get().strip(),
                    self.msg_var.get().strip())

    def _revalidate(self, *_args):
        issues = validate_rule(self._build_rule())
        if not issues:
            self.issue_var.set("")
            return
        has_error = any(sev == "error" for sev, _ in issues)
        self.issue_label.config(fg=RED if has_error else ORANGE)
        self.issue_var.set("  ".join(
            ("\u2716 " if sev == "error" else "\u26a0 ") + msg for sev, msg in issues))

    def _on_test(self):
        rule = self._build_rule()
        if any(sev == "error" for sev, _ in validate_rule(rule)):
            messagebox.showwarning("Test Rule",
                                   "Fix the errors shown below the form first.",
                                   parent=self.win)
            return
        TestResultDialog(self.win, rule, self.vrs_db, self.data_dir)

    def _on_ok(self):
        rule = self._build_rule()
        errors = [m for sev, m in validate_rule(rule) if sev == "error"]
        if errors:
            messagebox.showerror("Invalid Rule", "\n".join(errors), parent=self.win)
            return
        self.result = rule
        self.win.destroy()


# ---------------------------------------------------------------------------
# Test-against-database result window
# ---------------------------------------------------------------------------

class TestResultDialog:
    def __init__(self, parent, rule, vrs_db, data_dir=None):
        win = tk.Toplevel(parent)
        dark_window(win, parent, "Test Rule", "940x520", data_dir)
        win.minsize(700, 380)

        try:
            count, rows = count_matches(vrs_db, rule)
        except Exception as e:
            messagebox.showerror("Test Rule", "Could not query the database:\n%s" % e,
                                 parent=parent)
            win.destroy()
            return

        head = tk.Frame(win, bg=PANEL, padx=10, pady=8)
        head.pack(fill="x")

        if count:
            summary, color = "%d matching aircraft" % count, GREEN
        else:
            summary, color = "No aircraft match this rule", ORANGE
        dark_label(head, summary, bg=PANEL, fg=color,
               font=("Segoe UI", 11, "bold")).pack(anchor="w")
        dark_label(head, "When:  %s" % (rule.describe_match() or "(nothing)"),
               bg=PANEL, font=("Consolas", 9)).pack(anchor="w", pady=(4, 0))
        dark_label(head, "Set:   %s" % (rule.describe_change() or "(nothing)"),
               bg=PANEL, font=("Consolas", 9)).pack(anchor="w")
        if count > len(rows):
            dark_label(head, "Showing the first %d." % len(rows), bg=PANEL,
                   fg=CYAN).pack(anchor="w", pady=(4, 0))
        dark_label(head, "Rules run in order and the first match wins, so a rule "
                     "earlier in the list may claim some of these.",
               bg=PANEL, fg=LABEL).pack(anchor="w", pady=(4, 0))

        frame = tk.Frame(win, bg=BG)
        frame.pack(fill="both", expand=True, padx=4, pady=4)

        tree = ttk.Treeview(frame, columns=DB_COLUMNS, show="headings",
                            style=style_tree("RuleTest"))
        widths = {"Icao": 65, "Registration": 90, "Country": 95,
                  "Manufacturer": 130, "Model": 110, "ModelIcao": 75,
                  "Operator": 175, "OperatorIcao": 85}
        for col in DB_COLUMNS:
            tree.heading(col, text=col)
            tree.column(col, width=widths.get(col, 90), minwidth=40)

        # Highlight the columns this rule rewrites
        changed_cols = {RULE_TO_DB.get(f) for f in rule.change_fields}
        for col in DB_COLUMNS:
            if col in changed_cols:
                tree.heading(col, text="%s \u2192" % col)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        for row in rows:
            tree.insert("", "end", values=tuple(str(row.get(c) or "") for c in DB_COLUMNS))

        btns = tk.Frame(win, bg=BG, pady=8)
        btns.pack(fill="x", padx=10)
        dark_button(btns, "Close", win.destroy).pack(side="right")

        win.bind("<Escape>", lambda _e: win.destroy())
        center_on(win, parent)
        win.grab_set()


# ---------------------------------------------------------------------------
# Validation report window
# ---------------------------------------------------------------------------

class CheckReportDialog:
    """Validation report. Static checks appear at once; the database pass,
    which scans the whole table per rule, runs in the background."""

    SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}

    def __init__(self, parent, rules, vrs_db=None, data_dir=None):
        self.rules = rules
        self.vrs_db = vrs_db
        self.alive = True
        self.queue = queue.Queue()

        self.win = win = tk.Toplevel(parent)
        dark_window(win, parent, "Check Rules", "900x520", data_dir)
        win.minsize(650, 380)

        issues = validate_all(rules)

        head = tk.Frame(win, bg=PANEL, padx=10, pady=8)
        head.pack(fill="x")
        errors = sum(1 for _, s, _ in issues if s == "error")
        warns = sum(1 for _, s, _ in issues if s == "warning")
        if not issues:
            summary, color = ("%d rules checked - nothing to report." % len(rules), GREEN)
        else:
            summary = ("%d rules checked:  %d errors, %d warnings"
                       % (len(rules), errors, warns))
            color = RED if errors else ORANGE
        self.summary_var = tk.StringVar(value=summary)
        tk.Label(head, textvariable=self.summary_var, bg=PANEL, fg=color,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w")

        self.status_var = tk.StringVar(value="")
        tk.Label(head, textvariable=self.status_var, bg=PANEL, fg=CYAN,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))

        frame = tk.Frame(win, bg=BG)
        frame.pack(fill="both", expand=True, padx=4, pady=4)

        cols = ("rule", "severity", "detail")
        self.tree = tree = ttk.Treeview(frame, columns=cols, show="headings",
                                        style=style_tree("RuleCheck"))
        tree.heading("rule", text="Rule")
        tree.heading("severity", text="Severity")
        tree.heading("detail", text="Detail")
        tree.column("rule", width=60, minwidth=50, anchor="e")
        tree.column("severity", width=80, minwidth=60)
        tree.column("detail", width=700, minwidth=200)
        tree.tag_configure("error", foreground=RED)
        tree.tag_configure("warning", foreground=ORANGE)
        tree.tag_configure("info", foreground=CYAN)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Errors first, then warnings, then the database notes as they arrive
        for idx, sev, msg in sorted(
                issues, key=lambda t: (self.SEVERITY_ORDER.get(t[1], 9), t[0])):
            self._insert(idx, sev, msg)

        btns = tk.Frame(win, bg=BG, pady=8)
        btns.pack(fill="x", padx=10)
        dark_button(btns, "Close", win.destroy).pack(side="right")

        win.bind("<Escape>", lambda _e: win.destroy())
        win.bind("<Destroy>", self._on_destroy)
        center_on(win, parent)

        if vrs_db and os.path.exists(vrs_db):
            self.status_var.set("Checking against the database…")
            threading.Thread(target=self._scan_db, daemon=True).start()
            win.after(120, self._drain)

    def _insert(self, idx, sev, msg):
        summary = self.rules[idx].describe_match() if idx < len(self.rules) else ""
        self.tree.insert("", "end", tags=(sev,),
                         values=(idx + 1, sev, "%s   —   %s" % (msg, summary)))

    def _on_destroy(self, event):
        if event.widget is self.win:
            self.alive = False

    def _scan_db(self):
        """Worker thread: find rules that match nothing in the current database."""
        for i, rule in enumerate(self.rules):
            if not self.alive:
                return
            try:
                count, _ = count_matches(self.vrs_db, rule, sample_limit=1)
            except Exception as e:
                self.queue.put(("error", str(e)))
                return
            self.queue.put(("progress", (i + 1, count)))
        self.queue.put(("done", None))

    def _drain(self):
        """UI thread: apply whatever the worker has produced so far."""
        if not self.alive:
            return
        done = False
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "progress":
                    i, count = payload
                    if count == 0:
                        self._insert(i - 1, "info",
                                     "Matches no aircraft currently in the database.")
                    self.status_var.set("Checking against the database… %d/%d"
                                        % (i, len(self.rules)))
                elif kind == "error":
                    self.status_var.set("Database check failed: %s" % payload)
                    done = True
                else:
                    self.status_var.set(
                        "Database check complete - %d rule(s) match nothing in the "
                        "current database." % sum(
                            1 for iid in self.tree.get_children()
                            if self.tree.item(iid, "values")[1] == "info"))
                    done = True
        except queue.Empty:
            pass
        except tk.TclError:
            return  # window closed mid-update
        if not done:
            self.win.after(120, self._drain)




# ---------------------------------------------------------------------------
# Main rules editor
# ---------------------------------------------------------------------------

class RulesEditor:
    """Rules.csv editor window."""

    def __init__(self, parent, rules_path, vrs_db=None, data_dir=None,
                 on_saved=None, prefill_record=None, save_path=None):
        self.parent = parent
        self.rules_path = rules_path
        # Loading may come from the bundled copy while saving must land in the
        # user's working folder (a frozen build's bundle dir is temporary).
        self.save_path = save_path or rules_path
        self.vrs_db = vrs_db
        self.data_dir = data_dir
        self.on_saved = on_saved
        self.dirty = False

        try:
            self.rules = load_rules(rules_path, quiet=True)
        except Exception as e:
            messagebox.showerror("Rules Editor",
                                 "Could not read %s:\n%s" % (rules_path, e),
                                 parent=parent)
            self.rules = []

        self.win = tk.Toplevel(parent)
        dark_window(self.win, parent, "Rules Editor", "1040x620", data_dir)
        self.win.minsize(820, 480)

        self._build_ui()
        self._refresh()

        self.win.protocol("WM_DELETE_WINDOW", self._on_close)
        self.win.bind("<Control-s>", lambda _e: self._on_save())
        self.win.bind("<Delete>", lambda _e: self._on_delete())

        if prefill_record:
            self.win.after(100, lambda: self._on_new(prefill_record))

    # -- UI ---------------------------------------------------------------
    def _build_ui(self):
        head = tk.Frame(self.win, bg=PANEL, padx=10, pady=6)
        head.pack(fill="x")
        dark_label(head, "Rules are applied in order and the first match wins.",
               bg=PANEL, font=("Segoe UI", 9)).pack(side="left")
        dark_label(head, self.save_path, bg=PANEL, fg=CYAN,
               font=("Consolas", 8)).pack(side="right")

        frame = tk.Frame(self.win, bg=BG)
        frame.pack(fill="both", expand=True, padx=4, pady=4)

        cols = ("num", "when", "then", "message")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                 style=style_tree("RulesEdit"))
        self.tree.heading("num", text="#")
        self.tree.heading("when", text="When the record matches")
        self.tree.heading("then", text="Set these values")
        self.tree.heading("message", text="Log message")
        self.tree.column("num", width=45, minwidth=40, anchor="e")
        self.tree.column("when", width=420, minwidth=150)
        self.tree.column("then", width=260, minwidth=120)
        self.tree.column("message", width=260, minwidth=100)
        self.tree.tag_configure("error", foreground=RED)
        self.tree.tag_configure("warning", foreground=ORANGE)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda _e: self._on_edit())
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # Status line - shows issues for the selected rule
        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(self.win, textvariable=self.status_var, bg=BG,
                                     fg=ORANGE, font=("Segoe UI", 9), anchor="w",
                                     wraplength=1000, justify="left")
        self.status_label.pack(fill="x", padx=12)

        btns = tk.Frame(self.win, bg=BG, padx=8, pady=8)
        btns.pack(fill="x")

        dark_button(btns, "New", lambda: self._on_new(), primary=True).pack(side="left")
        dark_button(btns, "Edit", self._on_edit).pack(side="left", padx=(6, 0))
        dark_button(btns, "Duplicate", self._on_duplicate).pack(side="left", padx=(6, 0))
        dark_button(btns, "Delete", self._on_delete).pack(side="left", padx=(6, 0))
        dark_button(btns, "\u25b2", self._on_move_up).pack(side="left", padx=(14, 0))
        dark_button(btns, "\u25bc", self._on_move_down).pack(side="left", padx=(4, 0))
        self.test_btn = dark_button(btns, "Test", self._on_test)
        self.test_btn.pack(side="left", padx=(14, 0))
        if not (self.vrs_db and os.path.exists(self.vrs_db)):
            self.test_btn.config(state="disabled")
        dark_button(btns, "Check Rules", self._on_check).pack(side="left", padx=(6, 0))

        dark_button(btns, "Close", self._on_close).pack(side="right")
        self.save_btn = dark_button(btns, "Save", self._on_save, primary=True)
        self.save_btn.pack(side="right", padx=(0, 8))

        self.count_var = tk.StringVar(value="")
        tk.Label(btns, textvariable=self.count_var, bg=BG, fg=CYAN,
                 font=("Segoe UI", 9)).pack(side="right", padx=(0, 14))

    # -- state -------------------------------------------------------------
    def _refresh(self, select_index=None):
        """Rebuild the list from self.rules."""
        self.tree.delete(*self.tree.get_children())

        self._issues_by_rule = {}
        for idx, sev, msg in validate_all(self.rules):
            self._issues_by_rule.setdefault(idx, []).append((sev, msg))

        for i, rule in enumerate(self.rules):
            issues = self._issues_by_rule.get(i, [])
            tag = ("error",) if any(s == "error" for s, _ in issues) else \
                  ("warning",) if issues else ()
            self.tree.insert("", "end", iid=str(i), tags=tag,
                             values=(i + 1, rule.describe_match(),
                                     rule.describe_change(), rule.msg_text))

        n_err = sum(1 for v in self._issues_by_rule.values()
                    if any(s == "error" for s, _ in v))
        n_warn = len(self._issues_by_rule) - n_err
        parts = ["%d rules" % len(self.rules)]
        if n_err:
            parts.append("%d with errors" % n_err)
        if n_warn:
            parts.append("%d with warnings" % n_warn)
        if self.dirty:
            parts.append("unsaved changes")
        self.count_var.set("   \u2022   ".join(parts))
        self.save_btn.config(state="normal" if self.dirty else "disabled")

        if select_index is not None and 0 <= select_index < len(self.rules):
            iid = str(select_index)
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)
        self._on_select()

    def _selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def _on_select(self, *_args):
        idx = self._selected()
        issues = self._issues_by_rule.get(idx, []) if idx is not None else []
        if not issues:
            self.status_var.set("")
            return
        has_error = any(s == "error" for s, _ in issues)
        self.status_label.config(fg=RED if has_error else ORANGE)
        self.status_var.set("Rule %d:  " % (idx + 1) + "   ".join(
            ("\u2716 " if s == "error" else "\u26a0 ") + m for s, m in issues))

    def _mark_dirty(self):
        self.dirty = True

    # -- actions -----------------------------------------------------------
    def _on_new(self, source_record=None):
        prefill = None
        if source_record:
            match = {}
            if source_record.get("Registration"):
                match["Registration"] = str(source_record["Registration"])
            elif source_record.get("Icao"):
                match["ICAO"] = str(source_record["Icao"])
            prefill = Rule(match)

        dlg = RuleDialog(self.win, prefill, self.vrs_db, self.data_dir,
                         source_record, title="New Rule")
        if dlg.result:
            idx = self._selected()
            pos = len(self.rules) if idx is None else idx + 1
            self.rules.insert(pos, dlg.result)
            self._mark_dirty()
            self._refresh(pos)

    def _on_edit(self):
        idx = self._selected()
        if idx is None:
            return
        dlg = RuleDialog(self.win, self.rules[idx], self.vrs_db, self.data_dir,
                         title="Edit Rule %d" % (idx + 1))
        if dlg.result:
            self.rules[idx] = dlg.result
            self._mark_dirty()
            self._refresh(idx)

    def _on_duplicate(self):
        idx = self._selected()
        if idx is None:
            return
        self.rules.insert(idx + 1, self.rules[idx].copy())
        self._mark_dirty()
        self._refresh(idx + 1)

    def _on_delete(self):
        idx = self._selected()
        if idx is None:
            return
        rule = self.rules[idx]
        if not messagebox.askyesno(
                "Delete Rule",
                "Delete rule %d?\n\n%s\n%s" % (idx + 1, rule.describe_match(),
                                               rule.describe_change()),
                parent=self.win):
            return
        del self.rules[idx]
        self._mark_dirty()
        self._refresh(min(idx, len(self.rules) - 1))

    def _on_move_up(self):
        idx = self._selected()
        if idx is None or idx == 0:
            return
        self.rules[idx - 1], self.rules[idx] = self.rules[idx], self.rules[idx - 1]
        self._mark_dirty()
        self._refresh(idx - 1)

    def _on_move_down(self):
        idx = self._selected()
        if idx is None or idx >= len(self.rules) - 1:
            return
        self.rules[idx + 1], self.rules[idx] = self.rules[idx], self.rules[idx + 1]
        self._mark_dirty()
        self._refresh(idx + 1)

    def _on_test(self):
        idx = self._selected()
        if idx is None:
            messagebox.showinfo("Test", "Select a rule to test.", parent=self.win)
            return
        TestResultDialog(self.win, self.rules[idx], self.vrs_db, self.data_dir)

    def _on_check(self):
        CheckReportDialog(self.win, self.rules, self.vrs_db, self.data_dir)

    def _on_save(self):
        errors = [(i + 1, m) for i, s, m in validate_all(self.rules) if s == "error"]
        if errors:
            detail = "\n".join("Rule %d: %s" % e for e in errors[:10])
            if not messagebox.askyesno(
                    "Save With Errors",
                    "%d rule(s) have errors and will not work:\n\n%s\n\nSave anyway?"
                    % (len(errors), detail), parent=self.win):
                return
        try:
            save_rules(self.rules_path, self.rules)
        except Exception as e:
            messagebox.showerror("Save Failed",
                                 "Could not write %s:\n%s" % (self.rules_path, e),
                                 parent=self.win)
            return
        self.dirty = False
        self._refresh(self._selected())
        if self.on_saved:
            self.on_saved(len(self.rules))
        messagebox.showinfo(
            "Rules Saved",
            "%d rules written to:\n%s\n\nThe previous file was kept as Rules.csv.bak.\n"
            "Rules take effect on the next update." % (len(self.rules), self.save_path),
            parent=self.win)

    def _on_close(self):
        if self.dirty:
            answer = messagebox.askyesnocancel(
                "Unsaved Changes",
                "Save changes to Rules.csv before closing?", parent=self.win)
            if answer is None:
                return
            if answer:
                self._on_save()
                if self.dirty:   # save failed or was declined
                    return
        self.win.destroy()


def open_rules_editor(parent, rules_path, vrs_db=None, data_dir=None,
                      on_saved=None, prefill_record=None, save_path=None):
    """Open the rules editor. Returns the RulesEditor instance."""
    return RulesEditor(parent, rules_path, vrs_db, data_dir, on_saved,
                       prefill_record, save_path)
