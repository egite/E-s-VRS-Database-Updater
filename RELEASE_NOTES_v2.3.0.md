# v2.3.0

Runs on Linux, a reworked main window with a summary at the end of every run,
and a much faster merge. Also fixes the command-line mode, which refused to
start.

**Downloads:** `VRS_Database_Updater_x64.exe`, `VRS_Database_Updater_x86.exe`,
and `Rules.csv`.

Upgrading is just replacing the executable. Your `settings.json`, working
folder and `Rules.csv` carry over unchanged.

---

## Runs on Linux

If you run Virtual Radar Server under Mono on Linux, you no longer need to copy
`AircraftOnlineLookupCache.sqb` to a Windows machine and back. You don't need
Mono or the .exe either: run the program straight from the Python source with
your distribution's Python. The README has a new **Running on Linux** section
with install commands for Debian, Ubuntu, Raspberry Pi OS, Fedora, Arch and
openSUSE, plus a cron line for scheduled runs. In short:

```bash
sudo apt install python3 python3-tk python3-requests git
git clone https://github.com/egite/E-s-VRS-Database-Updater.git ~/vrs-updater
python3 ~/vrs-updater/run.py
```

What changed to make that work:

- **The VRS folder is found on Linux.** Under Mono, VRS keeps its database in
  `~/.local/share/VirtualRadar` (or `$XDG_DATA_HOME/VirtualRadar`) rather than
  `AppData\Local\VirtualRadar`. The program now looks in the right place for
  the OS it is running on and fills in the VRS Folder itself. On Windows it now
  reads `%LOCALAPPDATA%` instead of assuming `C:\Users\<name>\AppData\Local`,
  which also covers profiles that live somewhere else.
- **Fonts fall back properly.** Segoe UI and Consolas only exist on Windows.
  Elsewhere the program uses the system's own sans-serif and monospace fonts,
  so the status log and the editor tables stay aligned.
- **A maximized window is remembered on Linux**, and a `settings.json` copied
  from Windows with the window maximized no longer prints a Tk error.

## Command-line mode works again

`run.py --no-gui` passed `--no-gui` on to the command-line parser, which did not
recognize it and quit with an error before doing anything. That affected every
OS. It is now accepted, so the command-line mode works for scheduled tasks and
cron jobs:

```bash
python3 run.py --no-gui --work-dir ~/vrs-updater-data --vrs-dir ~/.local/share/VirtualRadar
```

The command-line mode does not read `settings.json`; run `run.py --help` for
its options.

## Reworked main window

**Data Sources** replaces the separate Auto-Download and Skip Processing boxes.
Each register has one row with two toggles:

- **Download** fetches a fresh copy when the local one is older than that
  register's age limit
- **Process** merges the copy you have into VRS. Untick it to leave a register
  out of a run without losing its downloaded data

Each row also shows the date of your local copy, in amber once it is older
than the age limit. A Rules row shows how many rules `Rules.csv` holds and has
its own Process toggle for the rules pass. Your existing choices carry over;
`settings.json` is unchanged.

**Apply Rules** moves out of the Options checkboxes and sits beside Start and
Cancel as a full-size button.

**Every run ends with a summary** in Program Progress: records processed,
updated and added per source, rules applied, how redacted owners were handled,
any sources merged from a stale copy, what was written, and how many warnings
and errors appeared.

**Database Update Status keeps the last 5,000 lines on screen.** A complete
build logs well over 100,000 lines, and the window used to slow down long
before the end. Every line is still kept, and the new **Save Log** button
writes the whole run to a text file.

**Help → Features** is a single menu item again, and the Features list and the
user manual describe the new layout.

## Faster merge

The silhouette lookup scanned the whole table, splitting and lowercasing every
alias of every row, each time it met a manufacturer and model it had not seen
yet. With about 28,500 distinct pairs in the FAA data alone, that scan was 95%
of the merge time. The table is now indexed by manufacturer, and the rules are
prepared once per run instead of once per aircraft.

On a 301,000-record FAA snapshot:

| Step        | v2.2.0 | v2.3.0 |
|-------------|-------:|-------:|
| FAA merge   | 32.6 s |  1.5 s |
| Rules pass  |  4.7 s |  1.3 s |

The results are identical. Both were checked record for record against the
previous code, including the Silhouette Editor's overlay path and the rules
pass's log lines.
