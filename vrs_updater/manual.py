"""
Built-in user manual: content plus the viewer window (Help -> User Manual).

The text lives here rather than in a data file so that a frozen build carries
it automatically without another entry in the .spec datas list.

Block types used below:
    ("h",    text)  sub-heading
    ("p",    text)  paragraph
    ("b",    text)  bullet
    ("m",    text)  monospace / literal example
    ("note", text)  callout worth not missing
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .ui_common import (BG, PANEL, ENTRY, CYAN, ORANGE, GREEN, LABEL, TEXT,
                        dark_window, dark_button, dark_entry, dark_label,
                        center_on)

APP_TITLE = "E's VRS Database Updater"


MANUAL = [

("Overview", [
    ("p", "This program updates Virtual Radar Server's aircraft database "
          "(AircraftOnlineLookupCache.sqb) from public aircraft registers, so "
          "that aircraft shown in VRS carry the correct owner, manufacturer, "
          "model, and ICAO type code - and therefore the correct silhouette "
          "and icon."),
    ("h", "Where the data comes from"),
    ("b", "FAA - United States registration database"),
    ("b", "CCAR - Canadian Civil Aircraft Register"),
    ("b", "NZ CAA - New Zealand aircraft register"),
    ("b", "CASA - Australian aircraft register"),
    ("b", "OpenSky Network - community aircraft database"),
    ("h", "What you can control"),
    ("b", "Which registers are downloaded, and how often"),
    ("b", "Which registers are merged, and which one wins when they disagree"),
    ("b", "Your own corrections, through Rules.csv (see Custom Rules)"),
    ("b", "Which silhouette each aircraft type gets, through Sils.csv "
          "(see Silhouettes)"),
    ("note", "Your VRS database is copied and backed up before anything is "
             "changed. Nothing is written to VRS until the merge has finished "
             "successfully."),
]),

("Getting Started", [
    ("h", "1. Set the two folders"),
    ("p", "On the main window, set the Working Folder and the VRS Folder."),
    ("b", "Working Folder - scratch space. Downloaded registers, backups, and "
          "logs are kept here. Pick a drive with a few gigabytes free."),
    ("b", "VRS Folder - the folder containing your VRS "
          "AircraftOnlineLookupCache.sqb file."),
    ("note", "These two must not be the same folder. The program refuses to "
             "run if they are."),
    ("h", "2. Choose which registers to download"),
    ("p", "Tick the registers you want under Auto-Download. If you only track "
          "aircraft in one country, there is no need to download the others. "
          "Each entry shows the date of the copy you already have."),
    ("h", "3. Close VRS"),
    ("p", "Shut down Virtual Radar Server before running an update, so that "
          "the database file is not in use while it is being replaced."),
    ("h", "4. Press Start"),
    ("p", "The first run downloads everything and takes the longest. Later "
          "runs reuse recent downloads unless they are older than the age "
          "limit you set in Settings."),
    ("note", "Once a week is plenty. Running more often mostly just burdens "
             "the register servers."),
]),

("The Main Window", [
    ("h", "Paths"),
    ("p", "The Working Folder and VRS Folder, each with a Browse button."),
    ("h", "Auto-Download"),
    ("p", "One checkbox per register. Next to each is the date of the copy in "
          "your working folder, turning amber when it is older than that "
          "register's age limit."),
    ("h", "Skip Processing"),
    ("p", "Skips merging a source even though you have a copy of it. Use this "
          "to leave one register out of a run without losing its downloaded "
          "data. The Rules checkbox skips the Rules.csv pass in the same way."),
    ("h", "Options"),
    ("b", "Run on Start - begins an update ten seconds after launch, and "
          "closes the program when it finishes. Intended for scheduled tasks."),
    ("b", "Backup VRS database - saves a dated copy of your VRS database into "
          "the working folder before making changes. Leave this on."),
    ("b", "Build complete database - also adds aircraft VRS has never seen, "
          "instead of only updating ones already in your database. Makes the "
          "database much larger, and is what you want for offline use."),
    ("b", "Apply Rules - applies Rules.csv to your VRS database on its own, "
          "with no downloading or merging. Greyed out when no rules file is "
          "found. See Custom Rules."),
    ("h", "Progress"),
    ("p", "Shows percent complete, an estimated time remaining, current phase, "
          "and live disk throughput. Program Progress lists the major steps; "
          "Database Update Status lists individual record changes."),
    ("h", "Start and Cancel"),
    ("p", "Cancel stops the run at the next safe point. Because all changes "
          "are written in a single step at the very end, cancelling partway "
          "through leaves your VRS database untouched."),
]),

("Running an Update", [
    ("h", "What happens, in order"),
    ("b", "1. Your VRS database is copied into the working folder, and backed "
          "up if that option is on."),
    ("b", "2. The registers you selected are downloaded, several at once. A "
          "register still within its age limit is skipped."),
    ("b", "3. Each register is parsed into its own intermediate database."),
    ("b", "4. The registers are merged, least trusted first, so the most "
          "trusted source has the last word."),
    ("b", "5. Rules.csv is applied at its position in the priority list."),
    ("b", "6. Every changed record is written in one transaction, and the "
          "database is copied back to your VRS folder."),
    ("h", "Which source wins"),
    ("p", "Settings has a Database Priority list, most trusted at the top. "
          "Sources lower in the list are merged first and can be overwritten "
          "by sources above them. Put Rules.CSV at the top and your own "
          "corrections will survive everything else."),
    ("h", "Why every touched aircraft is rewritten"),
    ("p", "VRS keeps its own timestamp on each cached aircraft and re-queries "
          "its online lookup service for anything it has not written in 28 "
          "days - replacing the operator, model, and type code with whatever "
          "that lookup returns. The update therefore rewrites every aircraft "
          "it touched, even ones whose details did not change, so the "
          "timestamp stays current and your merged data and rules are not "
          "overwritten later."),
    ("h", "Aircraft already in your database"),
    ("p", "Existing values are generally kept when a register has nothing "
          "better to offer, so a source with a gap will not blank a field that "
          "another source already filled in."),
]),

("Settings", [
    ("p", "File -> Settings."),
    ("h", "Download URLs"),
    ("p", "The address each register is downloaded from. These only need "
          "changing if a register moves its file, which does happen. Leave "
          "them alone otherwise."),
    ("h", "Re-download if local copy older than"),
    ("p", "An age limit in days, per register. If your copy is younger than "
          "this, the download is skipped and the existing copy is reused. Set "
          "a register to 0 to always download it."),
    ("p", "Higher numbers are kinder to the register servers. The FAA data "
          "changes daily but a week-old copy is perfectly serviceable for most "
          "people."),
    ("h", "Database Priority"),
    ("p", "Drag the sources so the one you trust most is at the top. See "
          "Running an Update for what this affects."),
]),

("Options and Logs", [
    ("p", "File -> Options."),
    ("b", "Write errors to log file - records warnings and failures."),
    ("b", "Write activity to log file - records everything that scrolls past "
          "in the status panes. These files get large."),
    ("h", "Where the logs go"),
    ("p", "Both are written into the working folder, one file per run:"),
    ("m", "error_log_DDMMYYYY-HHMM.txt\nactivity_log_DDMMYYYY-HHMM.txt"),
    ("p", "They are never deleted automatically, so clear out old ones "
          "occasionally."),
    ("note", "If a run does something you did not expect, turn on the activity "
             "log and run it again. The log names every record that changed "
             "and which source changed it."),
]),

("Custom Rules", [
    ("p", "Rules let you correct or override what the registers say. A rule "
          "matches aircraft on one or more fields, then sets other fields on "
          "the aircraft it matched."),
    ("h", "What rules are good for"),
    ("b", "Giving local operators their own ICAO code so they get their own "
          "flag in VRS"),
    ("b", "Fixing capitalisation, such as \"Atp Aircraft\" to \"ATP Aircraft\""),
    ("b", "Assigning a fleet-wide operator code to aircraft the register left "
          "blank"),
    ("b", "Naming operators the register hides behind a holding company"),
    ("h", "Opening the editor"),
    ("p", "Tools -> Rules Editor. Each row is one rule: what it matches on the "
          "left, what it sets in the middle, and the log message it prints on "
          "the right."),
    ("h", "Order matters"),
    ("p", "Rules are checked from the top down and the FIRST rule that matches "
          "an aircraft wins - later rules that would also have matched are "
          "skipped. Use the arrow buttons to move a rule up or down. If a rule "
          "never seems to fire, an earlier and broader rule is probably "
          "claiming those aircraft first; Check Rules will tell you."),
    ("h", "Excluding aircraft"),
    ("p", "The \"not\" checkbox beside a match field means \"only if this field "
          "is NOT that value\". The usual pattern is to match a fleet by "
          "operator name while excluding aircraft that already have the right "
          "code, so you are not rewriting values that are already correct."),
    ("note", "A rule made only of \"not\" conditions can never fire. It needs "
             "at least one positive match. The editor flags this."),
    ("h", "Testing before you save"),
    ("p", "Test Against Database runs the rule against your actual VRS "
          "database and lists the aircraft it would affect, with a count. "
          "Check Rules validates every rule at once and reports rules that "
          "cannot fire, rules hidden by an earlier rule, and rules that match "
          "nothing in your database."),
    ("h", "Building a rule from an aircraft"),
    ("p", "Open Tools -> Search Database, find the aircraft that is wrong, "
          "right-click it, and choose Create Rule from This Aircraft. The "
          "registration is filled in for you."),
    ("h", "Applying rules without a full update"),
    ("p", "The Apply Rules button on the main window applies Rules.csv to your "
          "VRS database immediately - no downloads, no merging. Use it after "
          "editing rules to see the effect in seconds rather than after a full "
          "update. It backs up and copies the database back exactly as a full "
          "run does."),
    ("h", "File format"),
    ("p", "Rules.csv is a plain CSV you can still edit by hand. Columns 2-9 "
          "are the fields to match, columns 13-20 the fields to set, and the "
          "last two columns the log message and the field name to append to "
          "it. The first column is a rule number for readability only."),
    ("note", "Saving from the editor keeps the previous file as Rules.csv.bak."),
]),

("Silhouettes", [
    ("p", "VRS shows a silhouette for an aircraft when it knows the aircraft's "
          "ICAO type code. Registers often leave that code out, which is why "
          "so many aircraft show a generic icon. This program fills the code "
          "in by matching each aircraft's manufacturer and model against a "
          "mapping table."),
    ("h", "Two sources of mappings"),
    ("p", "There is a large mapping table built into the program, and an "
          "optional Sils.csv file. Sils.csv is an OVERLAY on the built-in "
          "table, not a replacement:"),
    ("b", "Rows in your file are checked first, so they win"),
    ("b", "Anything your file does not cover still resolves from the built-in "
          "table"),
    ("b", "So your file only needs rows for what you actually want to change"),
    ("h", "The three things you can do"),
    ("b", "CHANGE a mapping - add a row with the manufacturer and model, and "
          "the code you want instead. Yours wins."),
    ("b", "ADD a mapping - add a row for a manufacturer and model the built-in "
          "table does not cover."),
    ("b", "DISABLE a mapping - add a row with the manufacturer and model, and "
          "leave BOTH the Type and Remap boxes empty. Matching stops at your "
          "row and no code is assigned."),
    ("note", "Deleting a row does NOT disable a mapping. If the built-in table "
             "has the same mapping, it simply takes over. To switch a mapping "
             "off you need the empty-code row described above."),
    ("p", "An empty-code row stops a code from being ASSIGNED; it does not "
          "erase a code already stored. Aircraft that already carry that type "
          "code in VRS keep it, so you will only see the effect on aircraft "
          "the program would otherwise have set. Clear an already-stored code "
          "in VRS itself."),
    ("h", "Opening the editor"),
    ("p", "Tools -> Silhouette Editor. Each row holds a list of manufacturer "
          "names and a list of model names, shown one per line instead of "
          "crammed into a single spreadsheet cell. An asterisk on its own "
          "means \"any\" - useful for kit aircraft, where the manufacturer is "
          "whoever built it."),
    ("note", "Do not use an asterisk for the manufacturer AND the model on the "
             "same row. That row would match every aircraft and hide "
             "everything below it. The editor warns about this."),
    ("h", "Type and Remap"),
    ("p", "Type is the ICAO type designator. Remap is used instead of Type "
          "when it is set, which lets you point an aircraft at a different "
          "silhouette than its official designator. They are usually the same."),
    ("h", "Testing a mapping"),
    ("p", "Test Lookup resolves a manufacturer and model exactly the way an "
          "update does, and tells you which row won - including whether the "
          "answer came from your file or from the built-in table. This is the "
          "quickest way to find out why an aircraft is getting the wrong "
          "silhouette."),
    ("h", "Rows that never match"),
    ("p", "A row with no manufacturer, or no model, can never match anything. "
          "The shipped file contains thousands of these. The editor hides them "
          "and keeps them in the file untouched; Check Rows offers to remove "
          "them if you want a smaller file."),
    ("h", "What the file cannot change"),
    ("p", "Some manufacturers are handled by logic built into the program "
          "rather than by the mapping table. Boeing, Cessna, Airbus, Piper, "
          "Pilatus, and Mooney model strings are rewritten before the table is "
          "consulted - Boeing \"78710\" becomes \"B789\", Cessna \"T182\" "
          "becomes \"C182\", and so on. Those rewrites cannot be changed from "
          "Sils.csv."),
    ("note", "Saving from the editor keeps the previous file as Sils.csv.bak."),
]),

("Searching the Database", [
    ("p", "Tools -> Search Database searches your VRS database."),
    ("b", "Type any text to search across ICAO, registration, operator, "
          "manufacturer, model, and country at once, or pick a single field "
          "from the dropdown"),
    ("b", "Results are limited to 500 rows; narrow the search if you hit that"),
    ("b", "Export CSV writes the current results to a file"),
    ("b", "Right-click a row to create a rule from that aircraft"),
    ("note", "This searches the copy of the database in your working folder, "
             "so run an update at least once before expecting results."),
]),

("Files in the Working Folder", [
    ("m", "AircraftOnlineLookupCache.sqb\n"
          "    The working copy of your VRS database."),
    ("m", "AircraftOnlineLookupCache-YYMMDD-HHMM.sqb\n"
          "    A backup taken before a run. Keep a few; delete the rest."),
    ("m", "FAADatabase.sqb, CCARDatabase.sqb, OpenSkyDatabase.sqb,\n"
          "NZCAADatabase.sqb, CASADatabase.sqb\n"
          "    Registers parsed into a usable form. Safe to delete - they are\n"
          "    downloaded again next run."),
    ("m", "FAADatabase DDMMYYYY.sqb\n"
          "    A dated FAA snapshot. The FAA removes owner details from its\n"
          "    files over time, so keeping these preserves data you already\n"
          "    have."),
    ("m", "Rules.csv, Sils.csv\n"
          "    Your customizations. These are the files worth backing up.\n"
          "    A copy in the working folder is used ahead of the one shipped\n"
          "    with the program."),
    ("m", "Rules.csv.bak, Sils.csv.bak\n"
          "    The previous version, kept automatically whenever you save\n"
          "    from an editor."),
    ("m", "error_log_*.txt, activity_log_*.txt\n"
          "    Logs, if enabled in Options."),
    ("m", "settings.json\n"
          "    Your settings, stored next to the program."),
]),

("Troubleshooting", [
    ("h", "The program says the folders cannot be the same"),
    ("p", "The Working Folder and VRS Folder must be different. The working "
          "folder is scratch space; the VRS folder is your live installation."),
    ("h", "VRS shows none of the changes"),
    ("p", "Check that VRS was closed during the update, and that the VRS "
          "Folder points at the folder actually holding the "
          "AircraftOnlineLookupCache.sqb file that VRS uses."),
    ("h", "An aircraft still has the wrong operator"),
    ("p", "Write a rule for it, then use Test Against Database to confirm the "
          "rule matches it. If the rule matches but nothing changes, another "
          "rule higher up the list is claiming the aircraft first - Check "
          "Rules will name it."),
    ("h", "An aircraft has no silhouette, or the wrong one"),
    ("p", "Open the Silhouette Editor and use Test Lookup with that aircraft's "
          "manufacturer and model. It will tell you which row answered and "
          "where that row came from. If nothing matched, add a row."),
    ("h", "A download fails"),
    ("p", "Registers do move their files and go offline. The run continues "
          "with the copy you already have, and says so. If a register has "
          "moved for good, update its address in Settings."),
    ("p", "One failure is usually a blip. The program counts consecutive "
          "failures per register, and after 14 in a row it tells you the "
          "address is probably wrong and names the URL. That warning waits "
          "until you acknowledge it, so an unattended run cannot lose it - "
          "and an unattended run still closes itself as normal. It clears "
          "itself as soon as the download works again."),
    ("h", "A run is slower than usual"),
    ("p", "The first run of the day downloads everything. Build complete "
          "database also makes runs considerably longer, since it adds every "
          "aircraft in the registers rather than only the ones VRS has seen."),
    ("h", "Something went wrong and you want the old database back"),
    ("p", "Copy the most recent AircraftOnlineLookupCache-YYMMDD-HHMM.sqb from "
          "your working folder over the AircraftOnlineLookupCache.sqb in your "
          "VRS folder, with VRS closed."),
]),

]


# ---------------------------------------------------------------------------
# Viewer
# ---------------------------------------------------------------------------

class ManualWindow:
    """Help -> User Manual: contents list on the left, text on the right."""

    def __init__(self, parent, data_dir=None):
        self.win = tk.Toplevel(parent)
        dark_window(self.win, parent, "%s - User Manual" % APP_TITLE,
                    "980x680", data_dir)
        self.win.minsize(760, 480)

        head = tk.Frame(self.win, bg=PANEL, padx=10, pady=6)
        head.pack(fill="x")
        dark_label(head, "User Manual", bg=PANEL, fg=CYAN,
                   font=("Segoe UI", 12, "bold")).pack(side="left")
        dark_label(head, "Search:", bg=PANEL).pack(side="left", padx=(20, 4))
        self.search_var = tk.StringVar()
        dark_entry(head, self.search_var, width=26).pack(side="left")
        self.search_var.trace_add("write", self._on_search)
        self.hits_var = tk.StringVar(value="")
        tk.Label(head, textvariable=self.hits_var, bg=PANEL, fg=ORANGE,
                 font=("Segoe UI", 9)).pack(side="left", padx=(8, 0))

        body = tk.Frame(self.win, bg=BG)
        body.pack(fill="both", expand=True, padx=4, pady=4)

        # contents
        left = tk.Frame(body, bg=BG)
        left.pack(side="left", fill="y")
        self.toc = tk.Listbox(left, width=26, bg=BG, fg=TEXT,
                              font=("Segoe UI", 10), relief="flat",
                              selectbackground="#264F78", selectforeground="#FFFFFF",
                              highlightthickness=0, activestyle="none")
        self.toc.pack(side="left", fill="y", expand=True)
        toc_sb = ttk.Scrollbar(left, orient="vertical", command=self.toc.yview)
        self.toc.configure(yscrollcommand=toc_sb.set)
        toc_sb.pack(side="left", fill="y")
        self.toc.bind("<<ListboxSelect>>", self._on_pick)

        # content
        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True, padx=(8, 0))
        self.text = tk.Text(right, wrap="word", bg=BG, fg=TEXT,
                            font=("Segoe UI", 10), bd=0, padx=16, pady=12,
                            highlightthickness=0, cursor="arrow", spacing1=2,
                            spacing3=6)
        self.text.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(right, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")

        self.text.tag_configure("title", foreground=CYAN,
                                font=("Segoe UI", 15, "bold"), spacing3=10)
        self.text.tag_configure("h", foreground=GREEN,
                                font=("Segoe UI", 11, "bold"),
                                spacing1=12, spacing3=4)
        self.text.tag_configure("p", foreground=TEXT, spacing3=8)
        self.text.tag_configure("b", foreground=TEXT, lmargin1=18, lmargin2=32,
                                spacing3=3)
        self.text.tag_configure("m", foreground="#CE9178", font=("Consolas", 9),
                                lmargin1=18, lmargin2=18, spacing1=4, spacing3=8)
        self.text.tag_configure("note", foreground=ORANGE, lmargin1=14,
                                lmargin2=14, spacing1=8, spacing3=8,
                                font=("Segoe UI", 10, "italic"))
        self.text.tag_configure("hit", background="#5A4500", foreground="#FFFFFF")

        btns = tk.Frame(self.win, bg=BG, padx=8, pady=8)
        btns.pack(fill="x")
        dark_button(btns, "Save to File...", self._on_save).pack(side="left")
        dark_button(btns, "Close", self.win.destroy).pack(side="right")

        self.win.bind("<Escape>", lambda _e: self.win.destroy())
        self._fill_toc()
        self._show(0)
        center_on(self.win, parent)

    # -- contents ---------------------------------------------------------
    def _fill_toc(self, only=None):
        self.toc.delete(0, "end")
        self.visible = []
        for i, (title, _blocks) in enumerate(MANUAL):
            if only is not None and i not in only:
                continue
            self.visible.append(i)
            self.toc.insert("end", "  " + title)

    def _on_pick(self, *_args):
        sel = self.toc.curselection()
        if sel:
            self._show(self.visible[sel[0]])

    def _show(self, index):
        self.current = index
        title, blocks = MANUAL[index]

        self.text.config(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("end", title + "\n", "title")
        for kind, body in blocks:
            if kind == "b":
                self.text.insert("end", "•  " + body + "\n", "b")
            elif kind == "m":
                self.text.insert("end", body + "\n", "m")
            else:
                self.text.insert("end", body + "\n", kind)
        self.text.config(state="disabled")
        self.text.yview_moveto(0)

        if self.visible and index in self.visible:
            pos = self.visible.index(index)
            self.toc.selection_clear(0, "end")
            self.toc.selection_set(pos)

        self._highlight()

    # -- search -----------------------------------------------------------
    @staticmethod
    def _section_text(index):
        title, blocks = MANUAL[index]
        return (title + " " + " ".join(b for _k, b in blocks)).lower()

    def _on_search(self, *_args):
        query = self.search_var.get().strip().lower()
        if not query:
            self._fill_toc()
            self.hits_var.set("")
            self._show(self.current)
            return

        matches = [i for i in range(len(MANUAL)) if query in self._section_text(i)]
        if not matches:
            self.hits_var.set("no matches")
            self._fill_toc()
            return

        self.hits_var.set("%d section%s" % (len(matches), "" if len(matches) == 1 else "s"))
        self._fill_toc(only=set(matches))
        self._show(matches[0] if self.current not in matches else self.current)

    def _highlight(self):
        self.text.tag_remove("hit", "1.0", "end")
        query = self.search_var.get().strip()
        if not query:
            return
        start = "1.0"
        while True:
            pos = self.text.search(query, start, stopindex="end", nocase=True)
            if not pos:
                break
            end = "%s+%dc" % (pos, len(query))
            self.text.tag_add("hit", pos, end)
            start = end
        first = self.text.tag_ranges("hit")
        if first:
            self.text.see(first[0])

    # -- export -----------------------------------------------------------
    def _on_save(self):
        path = filedialog.asksaveasfilename(
            parent=self.win, title="Save User Manual",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile="VRS Database Updater - User Manual.txt")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("%s - User Manual\n" % APP_TITLE)
                f.write("=" * 60 + "\n\n")
                for title, blocks in MANUAL:
                    f.write(title + "\n" + "-" * len(title) + "\n\n")
                    for kind, body in blocks:
                        if kind == "h":
                            f.write(body + "\n\n")
                        elif kind == "b":
                            f.write("  * " + body + "\n")
                        elif kind == "m":
                            f.write("".join("    " + ln + "\n"
                                            for ln in body.splitlines()) + "\n")
                        elif kind == "note":
                            f.write("  NOTE: " + body + "\n\n")
                        else:
                            f.write(body + "\n\n")
                    f.write("\n")
        except OSError as e:
            messagebox.showerror("Save Manual", str(e), parent=self.win)
            return
        messagebox.showinfo("Save Manual", "Saved to:\n%s" % path, parent=self.win)


def open_manual(parent, data_dir=None):
    """Open the user manual window."""
    return ManualWindow(parent, data_dir)
