"""
Integrated Sils.csv editor.

Replaces hand-editing Sils.csv in Excel. The file's two wide columns each
hold a comma-separated alias list, which a spreadsheet shows as one
unreadable cell; here each alias gets its own line. The editor also knows
the lookup's real semantics - first match wins, "*" is a wildcard, a row
with no manufacturer or model is skipped - and can run a manufacturer /
model through the actual silhouette resolver before you save.
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox

from .sils import (SilEntry, load_sils, save_sils,
                   validate_entry, validate_all)
from .ui_common import (BG, PANEL, ENTRY, CYAN, ORANGE, RED, GREEN, LABEL,
                        TEXT,
                        dark_window, dark_button, dark_entry, dark_label,
                        style_tree, center_on)

SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}


def _alias_box(master, values, height=8):
    """A dark multi-line box holding one alias per line."""
    txt = tk.Text(master, height=height, width=30, font=("Consolas", 10),
                  bg=ENTRY, fg="#FFFFFF", insertbackground="#FFFFFF",
                  relief="flat", wrap="none", undo=True)
    txt.insert("1.0", "\n".join(values))
    return txt


def _read_alias_box(txt):
    """Read the box as an alias list.

    Commas split as well as newlines: a comma is the file's own separator and
    cannot appear inside an alias, so treating one as a separator is the only
    reading that round-trips. It also makes pasting a comma-separated list
    from the old spreadsheet work as expected.
    """
    aliases = []
    for line in txt.get("1.0", "end").splitlines():
        for part in line.split(","):
            part = part.strip()
            if part:
                aliases.append(part)
    return aliases


# ---------------------------------------------------------------------------
# Row edit dialog
# ---------------------------------------------------------------------------

class SilRowDialog:
    """Modal editor for one Sils row. `result` is a SilEntry, or None."""

    def __init__(self, parent, entry=None, entries=None, data_dir=None,
                 title="Edit Row"):
        self.result = None
        self.entries = entries or []
        self.data_dir = data_dir

        self.win = tk.Toplevel(parent)
        dark_window(self.win, parent, title, "820x520", data_dir)
        self.win.minsize(700, 460)

        entry = entry or SilEntry()

        body = tk.Frame(self.win, bg=BG, padx=10, pady=10)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1, uniform="cols")
        body.columnconfigure(1, weight=1, uniform="cols")
        body.rowconfigure(0, weight=1)

        mfr_frame = tk.LabelFrame(body, text=" Manufacturer aliases ", bg=BG,
                                  fg=CYAN, font=("Segoe UI", 10, "bold"),
                                  padx=8, pady=6, relief="groove", bd=1)
        mfr_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        dark_label(mfr_frame,
                   "One per line (commas also split).  *  matches any.\n"
                   "Write names without commas - the lookup strips them first.",
                   justify="left").pack(anchor="w")
        self.mfr_box = _alias_box(mfr_frame, entry.manufacturers)
        self.mfr_box.pack(fill="both", expand=True, pady=(4, 0))

        mdl_frame = tk.LabelFrame(body, text=" Model aliases ", bg=BG,
                                  fg=GREEN, font=("Segoe UI", 10, "bold"),
                                  padx=8, pady=6, relief="groove", bd=1)
        mdl_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        dark_label(mdl_frame,
                   "One per line (commas also split).  *  matches any.\n"
                   "Models match on the original text, dashes included.",
                   justify="left").pack(anchor="w")
        self.mdl_box = _alias_box(mdl_frame, entry.models)
        self.mdl_box.pack(fill="both", expand=True, pady=(4, 0))

        codes = tk.Frame(self.win, bg=BG, padx=10)
        codes.pack(fill="x")
        dark_label(codes, "Type").grid(row=0, column=0, sticky="w")
        self.type_var = tk.StringVar(value=entry.type_code)
        dark_entry(codes, self.type_var, width=10).grid(row=0, column=1, padx=(6, 20))
        dark_label(codes, "Remap").grid(row=0, column=2, sticky="w")
        self.remap_var = tk.StringVar(value=entry.remap)
        dark_entry(codes, self.remap_var, width=10).grid(row=0, column=3, padx=(6, 20))
        dark_label(codes, "Remap wins when set; otherwise Type is used.").grid(
            row=0, column=4, sticky="w")

        self.issue_var = tk.StringVar(value="")
        self.issue_label = tk.Label(self.win, textvariable=self.issue_var, bg=BG,
                                    fg=ORANGE, font=("Segoe UI", 9), anchor="w",
                                    justify="left", wraplength=780)
        self.issue_label.pack(fill="x", padx=12, pady=(8, 0))

        btns = tk.Frame(self.win, bg=BG, padx=10, pady=10)
        btns.pack(fill="x")
        dark_button(btns, "Test Lookup", self._on_test).pack(side="left")
        dark_button(btns, "Cancel", self.win.destroy).pack(side="right")
        dark_button(btns, "OK", self._on_ok, primary=True).pack(side="right", padx=(0, 8))

        for box in (self.mfr_box, self.mdl_box):
            box.bind("<KeyRelease>", self._revalidate)
        self.type_var.trace_add("write", self._revalidate)
        self.remap_var.trace_add("write", self._revalidate)

        self.win.bind("<Escape>", lambda _e: self.win.destroy())
        self._revalidate()

        center_on(self.win, parent)
        self.win.grab_set()
        self.mfr_box.focus_set()
        parent.wait_window(self.win)

    def _build_entry(self):
        return SilEntry(_read_alias_box(self.mfr_box), _read_alias_box(self.mdl_box),
                        self.remap_var.get().strip(), self.type_var.get().strip())

    def _revalidate(self, *_args):
        issues = validate_entry(self._build_entry())
        if not issues:
            self.issue_var.set("")
            return
        has_error = any(sev == "error" for sev, _ in issues)
        self.issue_label.config(fg=RED if has_error else ORANGE)
        self.issue_var.set("   ".join(
            ("\u2716 " if sev == "error" else "\u26a0 ") + msg for sev, msg in issues))

    def _on_test(self):
        entry = self._build_entry()
        if not entry.usable:
            messagebox.showwarning("Test Lookup",
                                   "Add at least one manufacturer and one model first.",
                                   parent=self.win)
            return
        TestLookupDialog(self.win, self.entries, entry, self.data_dir)

    def _on_ok(self):
        entry = self._build_entry()
        errors = [m for sev, m in validate_entry(entry) if sev == "error"]
        if errors:
            messagebox.showerror("Invalid Row", "\n".join(errors), parent=self.win)
            return
        self.result = entry
        self.win.destroy()


# ---------------------------------------------------------------------------
# Lookup test
# ---------------------------------------------------------------------------

class TestLookupDialog:
    """Run a manufacturer / model through the real silhouette resolver."""

    def __init__(self, parent, entries, focus_entry=None, data_dir=None):
        self.entries = entries
        self.focus_entry = focus_entry

        self.win = win = tk.Toplevel(parent)
        dark_window(win, parent, "Test Lookup", "760x420", data_dir)
        win.minsize(620, 360)

        head = tk.Frame(win, bg=PANEL, padx=10, pady=8)
        head.pack(fill="x")
        dark_label(head, "Resolve a manufacturer and model the way the merge does.",
                   bg=PANEL, font=("Segoe UI", 10, "bold"), fg=CYAN).pack(anchor="w")
        dark_label(head, "Sils.csv is layered over the program's built-in table: "
                         "your rows win, and anything you do not cover still "
                         "resolves from the built-in mappings.",
                   bg=PANEL, wraplength=720,
                   justify="left").pack(anchor="w", pady=(4, 0))

        form = tk.Frame(win, bg=BG, padx=10, pady=10)
        form.pack(fill="x")
        dark_label(form, "Manufacturer").grid(row=0, column=0, sticky="w")
        self.mfr_var = tk.StringVar(
            value=focus_entry.manufacturers[0] if focus_entry and focus_entry.manufacturers else "")
        dark_entry(form, self.mfr_var, width=32).grid(row=0, column=1, padx=6, sticky="ew")
        dark_label(form, "Model").grid(row=0, column=2, sticky="w", padx=(12, 0))
        self.mdl_var = tk.StringVar(
            value=focus_entry.models[0] if focus_entry and focus_entry.models else "")
        dark_entry(form, self.mdl_var, width=24).grid(row=0, column=3, padx=6, sticky="ew")
        dark_button(form, "Look Up", self._run, primary=True).grid(row=0, column=4, padx=(10, 0))
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        self.result_var = tk.StringVar(value="")
        tk.Label(win, textvariable=self.result_var, bg=BG, fg=TEXT,
                 font=("Consolas", 11), anchor="w", justify="left",
                 wraplength=720).pack(fill="x", padx=12)

        self.detail_var = tk.StringVar(value="")
        tk.Label(win, textvariable=self.detail_var, bg=BG, fg=LABEL,
                 font=("Segoe UI", 9), anchor="w", justify="left",
                 wraplength=720).pack(fill="x", padx=12, pady=(6, 0))

        btns = tk.Frame(win, bg=BG, pady=10)
        btns.pack(fill="x", padx=10, side="bottom")
        dark_button(btns, "Close", win.destroy).pack(side="right")

        win.bind("<Return>", lambda _e: self._run())
        win.bind("<Escape>", lambda _e: win.destroy())
        center_on(win, parent)
        win.grab_set()
        if self.mfr_var.get() or self.mdl_var.get():
            self._run()

    def _run(self):
        from .silhouette import SilhouetteLookup

        mfr = self.mfr_var.get().strip()
        model = self.mdl_var.get().strip()
        if not model:
            self.result_var.set("Enter a model to look up.")
            self.detail_var.set("")
            return

        # Resolve against the rows currently in the editor, layered over the
        # built-in table exactly as the merge does.
        lookup = SilhouetteLookup()
        lookup.load_entries(self.entries)

        code, found = lookup.determine_silhouette(model, mfr)
        row = lookup.last_match_index
        from_builtin = row >= lookup.builtin_start

        if found and code:
            self.result_var.set("\u2192  %s" % code)
        elif code:
            self.result_var.set("\u2192  %s   (no confirmed match - resolver fell "
                                "through to a rewritten code)" % code)
        else:
            self.result_var.set("\u2192  no silhouette resolved")

        parts = []
        if row < 0:
            parts.append("Nothing matched in Sils.csv or the built-in table; the "
                         "code above (if any) came from the resolver's "
                         "manufacturer-specific rules.")
        elif from_builtin:
            parts.append("Matched a BUILT-IN mapping, not a Sils.csv row:  %s  |  %s"
                         % (lookup.manufacturers[row][:55],
                            lookup.models[row][:55]))
            parts.append("Add a row here with the same manufacturer and model to "
                         "override it, or one with an empty Type and Remap to "
                         "suppress it.")
        else:
            matched = self.entries[row]
            parts.append("Matched Sils.csv row %d:  %s  |  %s  \u2192  %s"
                         % (row + 1, matched.describe_manufacturers()[:55],
                            matched.describe_models()[:55], matched.resolved))
            if self.focus_entry is not None and matched is not self.focus_entry:
                parts.append("Note: the row you are editing did not win - "
                             "row %d is earlier in the file." % (row + 1))
        self.detail_var.set("\n".join(parts))



# ---------------------------------------------------------------------------
# Validation report
# ---------------------------------------------------------------------------

class SilCheckDialog:
    def __init__(self, parent, editor, data_dir=None):
        self.editor = editor
        entries = editor.entries

        win = self.win = tk.Toplevel(parent)
        dark_window(win, parent, "Check Rows", "920x520", data_dir)
        win.minsize(680, 380)

        issues = sorted(validate_all(entries),
                        key=lambda t: (SEVERITY_ORDER.get(t[1], 9), t[0]))

        head = tk.Frame(win, bg=PANEL, padx=10, pady=8)
        head.pack(fill="x")
        usable = sum(1 for e in entries if e.usable)
        unused = len(entries) - usable
        errors = sum(1 for _, s, _ in issues if s == "error")
        warns = sum(1 for _, s, _ in issues if s == "warning")
        if not issues:
            dark_label(head, "%d rows checked - nothing to report." % usable,
                       bg=PANEL, fg=GREEN, font=("Segoe UI", 11, "bold")).pack(anchor="w")
        else:
            dark_label(head, "%d rows checked:  %d errors, %d warnings"
                       % (usable, errors, warns), bg=PANEL,
                       fg=RED if errors else ORANGE,
                       font=("Segoe UI", 11, "bold")).pack(anchor="w")
        dark_label(head, "%d unused rows (no manufacturer or no model) are kept in "
                         "the file but never match." % unused,
                   bg=PANEL, fg=CYAN).pack(anchor="w", pady=(4, 0))

        frame = tk.Frame(win, bg=BG)
        frame.pack(fill="both", expand=True, padx=4, pady=4)

        cols = ("row", "severity", "detail")
        tree = ttk.Treeview(frame, columns=cols, show="headings",
                            style=style_tree("SilCheck"))
        tree.heading("row", text="Row")
        tree.heading("severity", text="Severity")
        tree.heading("detail", text="Detail")
        tree.column("row", width=60, minwidth=50, anchor="e")
        tree.column("severity", width=80, minwidth=60)
        tree.column("detail", width=740, minwidth=200)
        tree.tag_configure("error", foreground=RED)
        tree.tag_configure("warning", foreground=ORANGE)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        for idx, sev, msg in issues:
            e = entries[idx]
            tree.insert("", "end", tags=(sev,),
                        values=(idx + 1, sev, "%s   \u2014   %s | %s"
                                % (msg, e.describe_manufacturers()[:40],
                                   e.describe_models()[:40])))
        self.tree = tree

        btns = tk.Frame(win, bg=BG, pady=8)
        btns.pack(fill="x", padx=10)
        if unused:
            dark_button(btns, "Remove %d Unused Rows" % unused,
                        lambda: self._remove_unused(unused)).pack(side="left")
        dark_button(btns, "Close", win.destroy).pack(side="right")

        win.bind("<Escape>", lambda _e: win.destroy())
        center_on(win, parent)

    def _remove_unused(self, unused):
        if not messagebox.askyesno(
                "Remove Unused Rows",
                "Remove %d rows that have no manufacturer or no model?\n\n"
                "The lookup already skips them, so this changes nothing except "
                "the size of the file. Nothing is written until you save."
                % unused, parent=self.win):
            return
        self.editor.remove_unused()
        self.win.destroy()


# ---------------------------------------------------------------------------
# Main editor
# ---------------------------------------------------------------------------

class SilsEditor:
    """Sils.csv editor window."""

    def __init__(self, parent, sils_path, data_dir=None, on_saved=None,
                 save_path=None):
        self.parent = parent
        self.sils_path = sils_path
        self.save_path = save_path or sils_path
        self.data_dir = data_dir
        self.on_saved = on_saved
        self.dirty = False
        self._filter_job = None

        try:
            self.entries = load_sils(sils_path, quiet=True)
        except Exception as e:
            messagebox.showerror("Sils Editor",
                                 "Could not read %s:\n%s" % (sils_path, e),
                                 parent=parent)
            self.entries = []

        self.win = tk.Toplevel(parent)
        dark_window(self.win, parent, "Silhouette Editor", "1060x640", data_dir)
        self.win.minsize(860, 500)

        self._build_ui()
        self._refresh()

        self.win.protocol("WM_DELETE_WINDOW", self._on_close)
        self.win.bind("<Control-s>", lambda _e: self._on_save())

    # -- UI ---------------------------------------------------------------
    def _build_ui(self):
        head = tk.Frame(self.win, bg=PANEL, padx=10, pady=6)
        head.pack(fill="x")
        dark_label(head, "Search:", bg=PANEL, font=("Segoe UI", 10)).pack(side="left")

        self.search_var = tk.StringVar()
        search = dark_entry(head, self.search_var, width=28)
        search.pack(side="left", padx=(8, 4))
        self.search_var.trace_add("write", self._schedule_filter)

        self.field_var = tk.StringVar(value="Any")
        combo = ttk.Combobox(head, textvariable=self.field_var, width=14,
                             values=["Any", "Manufacturer", "Model", "Type"],
                             state="readonly")
        combo.pack(side="left", padx=4)
        combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh(keep_selection=True))

        dark_label(head, self.save_path, bg=PANEL, fg=CYAN,
                   font=("Consolas", 8)).pack(side="right")

        info = tk.Frame(self.win, bg=BG, padx=10, pady=3)
        info.pack(fill="x")
        dark_label(info, "Your rows are scanned first and override the program's "
                         "built-in table; mappings you do not cover still come "
                         "from it. A row with no Type and no Remap suppresses "
                         "the built-in mapping.",
                   wraplength=1020, justify="left").pack(anchor="w")

        frame = tk.Frame(self.win, bg=BG)
        frame.pack(fill="both", expand=True, padx=4, pady=4)

        cols = ("num", "mfr", "model", "remap", "type")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                 style=style_tree("SilsEdit"))
        self.tree.heading("num", text="#")
        self.tree.heading("mfr", text="Manufacturer aliases")
        self.tree.heading("model", text="Model aliases")
        self.tree.heading("remap", text="Remap")
        self.tree.heading("type", text="Type")
        self.tree.column("num", width=60, minwidth=50, anchor="e")
        self.tree.column("mfr", width=330, minwidth=140)
        self.tree.column("model", width=420, minwidth=160)
        self.tree.column("remap", width=70, minwidth=55)
        self.tree.column("type", width=70, minwidth=55)
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

        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(self.win, textvariable=self.status_var, bg=BG,
                                     fg=ORANGE, font=("Segoe UI", 9), anchor="w",
                                     wraplength=1020, justify="left")
        self.status_label.pack(fill="x", padx=12)

        btns = tk.Frame(self.win, bg=BG, padx=8, pady=8)
        btns.pack(fill="x")
        dark_button(btns, "New", self._on_new, primary=True).pack(side="left")
        dark_button(btns, "Edit", self._on_edit).pack(side="left", padx=(6, 0))
        dark_button(btns, "Duplicate", self._on_duplicate).pack(side="left", padx=(6, 0))
        dark_button(btns, "Delete", self._on_delete).pack(side="left", padx=(6, 0))
        dark_button(btns, "\u25b2", self._on_move_up).pack(side="left", padx=(14, 0))
        dark_button(btns, "\u25bc", self._on_move_down).pack(side="left", padx=(4, 0))
        dark_button(btns, "Test Lookup", self._on_test).pack(side="left", padx=(14, 0))
        dark_button(btns, "Check Rows", self._on_check).pack(side="left", padx=(6, 0))

        dark_button(btns, "Close", self._on_close).pack(side="right")
        self.save_btn = dark_button(btns, "Save", self._on_save, primary=True)
        self.save_btn.pack(side="right", padx=(0, 8))

        self.count_var = tk.StringVar(value="")
        tk.Label(btns, textvariable=self.count_var, bg=BG, fg=CYAN,
                 font=("Segoe UI", 9)).pack(side="right", padx=(0, 14))

    # -- state -------------------------------------------------------------
    def _schedule_filter(self, *_args):
        """Debounce the search box - refiltering 1,900 rows per keystroke lags."""
        if self._filter_job:
            self.win.after_cancel(self._filter_job)
        self._filter_job = self.win.after(200, lambda: self._refresh(keep_selection=True))

    def _matches_filter(self, entry):
        query = self.search_var.get().strip().lower()
        if not query:
            return True
        field = self.field_var.get()
        haystacks = []
        if field in ("Any", "Manufacturer"):
            haystacks.append(entry.describe_manufacturers().lower())
        if field in ("Any", "Model"):
            haystacks.append(entry.describe_models().lower())
        if field in ("Any", "Type"):
            haystacks.append((entry.type_code + " " + entry.remap).lower())
        return any(query in h for h in haystacks)

    def _refresh(self, select_index=None, keep_selection=False):
        """Rebuild the visible list. Unusable rows stay in self.entries but hidden."""
        self._filter_job = None
        if keep_selection and select_index is None:
            select_index = self._selected()

        self.tree.delete(*self.tree.get_children())

        self._issues_by_row = {}
        for idx, sev, msg in validate_all(self.entries):
            self._issues_by_row.setdefault(idx, []).append((sev, msg))

        shown = 0
        for i, entry in enumerate(self.entries):
            if not entry.usable or not self._matches_filter(entry):
                continue
            issues = self._issues_by_row.get(i, [])
            tag = ("error",) if any(s == "error" for s, _ in issues) else \
                  ("warning",) if issues else ()
            self.tree.insert("", "end", iid=str(i), tags=tag,
                             values=(i + 1, entry.describe_manufacturers(),
                                     entry.describe_models(), entry.remap,
                                     entry.type_code))
            shown += 1

        usable = sum(1 for e in self.entries if e.usable)
        unused = len(self.entries) - usable
        parts = ["%d rows" % usable] if shown == usable else \
                ["%d of %d rows" % (shown, usable)]
        if unused:
            parts.append("%d unused (hidden)" % unused)
        if self.dirty:
            parts.append("unsaved changes")
        self.count_var.set("   \u2022   ".join(parts))
        self.save_btn.config(state="normal" if self.dirty else "disabled")

        if select_index is not None:
            iid = str(select_index)
            if self.tree.exists(iid):
                self.tree.selection_set(iid)
                self.tree.focus(iid)
                self.tree.see(iid)
        self._on_select()

    def _selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def _on_select(self, *_args):
        idx = self._selected()
        issues = self._issues_by_row.get(idx, []) if idx is not None else []
        if not issues:
            self.status_var.set("")
            return
        has_error = any(s == "error" for s, _ in issues)
        self.status_label.config(fg=RED if has_error else ORANGE)
        self.status_var.set("Row %d:  " % (idx + 1) + "   ".join(
            ("\u2716 " if s == "error" else "\u26a0 ") + m for s, m in issues))

    def _mark_dirty(self):
        self.dirty = True

    def _visible_indices(self):
        return [int(iid) for iid in self.tree.get_children()]

    # -- actions -----------------------------------------------------------
    def _on_new(self):
        dlg = SilRowDialog(self.win, None, self.entries, self.data_dir,
                           title="New Row")
        if dlg.result:
            idx = self._selected()
            pos = len(self.entries) if idx is None else idx + 1
            self.entries.insert(pos, dlg.result)
            self._mark_dirty()
            self._refresh(pos)

    def _on_edit(self):
        idx = self._selected()
        if idx is None:
            return
        dlg = SilRowDialog(self.win, self.entries[idx], self.entries, self.data_dir,
                           title="Edit Row %d" % (idx + 1))
        if dlg.result:
            self.entries[idx] = dlg.result
            self._mark_dirty()
            self._refresh(idx)

    def _on_duplicate(self):
        idx = self._selected()
        if idx is None:
            return
        self.entries.insert(idx + 1, self.entries[idx].copy())
        self._mark_dirty()
        self._refresh(idx + 1)

    def _on_delete(self):
        idx = self._selected()
        if idx is None:
            return
        e = self.entries[idx]
        if not messagebox.askyesno(
                "Delete Row",
                "Delete row %d?\n\n%s\n%s\n\n\u2192 %s"
                % (idx + 1, e.describe_manufacturers(), e.describe_models(),
                   e.resolved), parent=self.win):
            return
        del self.entries[idx]
        self._mark_dirty()
        self._refresh()

    def _swap_visible(self, offset):
        """Swap the selected row with the previous/next *visible* row.

        Unusable rows sit between the visible ones and are left untouched.
        """
        idx = self._selected()
        if idx is None:
            return
        visible = self._visible_indices()
        try:
            pos = visible.index(idx)
        except ValueError:
            return
        target = pos + offset
        if not (0 <= target < len(visible)):
            return
        other = visible[target]
        self.entries[idx], self.entries[other] = self.entries[other], self.entries[idx]
        self._mark_dirty()
        self._refresh(other)

    def _on_move_up(self):
        self._swap_visible(-1)

    def _on_move_down(self):
        self._swap_visible(1)

    def _on_test(self):
        idx = self._selected()
        focus = self.entries[idx] if idx is not None else None
        TestLookupDialog(self.win, self.entries, focus, self.data_dir)

    def _on_check(self):
        SilCheckDialog(self.win, self, self.data_dir)

    def remove_unused(self):
        """Drop every row the lookup skips. Called from the check report."""
        before = len(self.entries)
        self.entries = [e for e in self.entries if e.usable]
        if len(self.entries) != before:
            self._mark_dirty()
        self._refresh()

    def _on_save(self):
        errors = [(i + 1, m) for i, s, m in validate_all(self.entries) if s == "error"]
        if errors:
            detail = "\n".join("Row %d: %s" % e for e in errors[:10])
            if not messagebox.askyesno(
                    "Save With Errors",
                    "%d row(s) have errors and will not work:\n\n%s\n\nSave anyway?"
                    % (len(errors), detail), parent=self.win):
                return
        try:
            save_sils(self.save_path, self.entries)
        except Exception as e:
            messagebox.showerror("Save Failed",
                                 "Could not write %s:\n%s" % (self.save_path, e),
                                 parent=self.win)
            return
        self.dirty = False
        self._refresh(keep_selection=True)
        if self.on_saved:
            self.on_saved(len(self.entries))
        messagebox.showinfo(
            "Sils Saved",
            "%d rows written to:\n%s\n\nThe previous file was kept as Sils.csv.bak.\n"
            "Changes take effect on the next update."
            % (len(self.entries), self.save_path), parent=self.win)

    def _on_close(self):
        if self.dirty:
            answer = messagebox.askyesnocancel(
                "Unsaved Changes",
                "Save changes to Sils.csv before closing?", parent=self.win)
            if answer is None:
                return
            if answer:
                self._on_save()
                if self.dirty:
                    return
        self.win.destroy()


def open_sils_editor(parent, sils_path, data_dir=None, on_saved=None,
                     save_path=None):
    """Open the Sils editor. Returns the SilsEditor instance."""
    return SilsEditor(parent, sils_path, data_dir, on_saved, save_path)
