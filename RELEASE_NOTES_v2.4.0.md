# v2.4.0

Checks your VRS configuration before every run, keeps your settings up to date
across versions, and shows progress through the last, slowest part of a run.
Also fixes a run that stopped just before handing the database back to VRS
while VRS was open.

**Downloads:** `VRS_Database_Updater_x64.exe` and `Rules.csv`. There is no
longer a 32-bit build; see below.

Upgrading is just replacing the executable. Your working folder and
`Rules.csv` carry over unchanged, and `settings.json` is updated in place the
first time v2.4.0 starts (see **Settings stay up to date**).

---

## VRS configuration is checked before a run

Some VRS settings stop VRS from reading the lookup cache, or stop this
program's updates from reaching it. The run still succeeds, so nothing looked
wrong; VRS just didn't show the new data. Before a run or **Apply Rules**, the
program now reads VRS's `Configuration.xml` and `PluginsConfiguration.txt` from
the VRS Folder and looks for:

- **VRS's BaseStation database set to `AircraftOnlineLookupCache.sqb`.** The
  cache isn't a BaseStation database, so VRS fails to open it ("no such table
  DBHistory", "disk I/O error" or a crash at start-up).
- **A plugin, usually the Database Writer, pointed at the cache** for the same
  reason.
- **Online lookup turned off in VRS** without **Build complete database** ticked
  here. VRS only adds new aircraft to the cache through its online lookup, and
  this program only refreshes aircraft already in it, so new aircraft never get
  details.
- **A VRS Folder that doesn't exist or isn't a VRS data folder**, which usually
  means the program is running as a different user than VRS, or the VRS install
  folder was picked instead of its data folder.
- **No permission to write to the VRS Folder.**

Each problem is shown with how to fix it, and the run waits for you to click OK
before carrying on, even on a **Run on Start** run, since these problems stop
VRS from seeing the update at all. They are also kept in the run log. In
command-line mode they are printed and the program waits for Enter, or exits
with code 2 when there is no console to acknowledge them.

## Settings stay up to date

`settings.json` used to save every download link, including the defaults, and
the saved copy always won. So when a new version shipped a new link, existing
users never got it. Anyone who started on v2.0 is still downloading the NZ CAA
register from the address it used before the NZ CAA renamed the file. Values
were also read without any checks, and a missing age limit meant "always
download" where a fresh install uses 30 days.

Every time the program starts it now checks `settings.json` and brings it up to
date:

- **A default link that has since changed is replaced with the current one.** A
  link you typed in yourself is never replaced. From now on the file also
  records which defaults were in effect, so future link changes reach you
  without anything more on your part.
- **Values of the wrong type or out of range are replaced with the default**,
  settings added in later versions are filled in, and settings this version no
  longer uses are dropped.
- **The Database Priority list loses unknown or repeated entries**, and any
  register missing from it is added.
- **A `settings.json` that can't be read is renamed** to
  `settings.json.bad-<date-time>` rather than overwritten, so you can still
  recover anything in it.

Anything the program changed is listed in yellow in Program Progress when it
starts, for example:

> The NZ CAA download link has changed in this version and has been updated to
> https://www.aviation.govt.nz/assets/aircraft/aircraft-register/Aircraft-Register-for-website.csv

**Settings** also has a new **Restore Default Links** button, which puts all
five download links back to the program's own. Nothing is saved until you
click OK.

## Progress through the end of a run

**The progress bar moves while the database is written.** All changes are still
written in a single transaction, so the write is all-or-nothing, but they are
fed to it in batches and the bar shows how far it has got. **Cancel** during
the write discards it, and VRS's database is left as it was.

**Copying the database back to VRS is its own step**, *Copying to VRS*, with
its own progress. Cancel is ignored during the copy, because stopping halfway
would leave VRS with a cut-off database.

**Local copy dates update during a run.** Each row in **Data Sources** now
shows its new date as soon as that register has been rebuilt, instead of
waiting until the whole run has finished.

## Fixes

**A run no longer stops just before it reaches VRS when VRS is open.** Before
copying the database back, the program removes any leftover
`AircraftOnlineLookupCache.sqb-wal` and `-shm` files beside it. Windows won't
delete a file that another program has open, so if VRS was holding them the
run stopped with an error, and the update never reached VRS. It now warns you
and carries on.

**The database is handed back in the format VRS itself uses.** The program
works on its copy in SQLite's WAL mode for speed, and that mode is stored in
the file. The copy now goes back to VRS in rollback-journal mode, VRS's own,
so VRS never has to open it in WAL mode or create `-wal`/`-shm` files beside
it. On Linux, when the program is run as root, the new file is given to the
VRS folder's owner so VRS can keep adding lookups to it.

## No 32-bit build

There is no longer a `VRS_Database_Updater_x86.exe`. Use the x64 build, or run
from the Python source on 32-bit systems. The build script also runs pip as
`python -m pip`, so it no longer stops at the dependency check on PCs where
Device Guard or WDAC blocks `pip.exe`.

## README

- **Which VRS database this updates**, a new section. The program updates
  the online lookup cache, `AircraftOnlineLookupCache.sqb`, never the
  BaseStation database. The section explains both files and two VRS behaviors
  that can hide the update: your own non-blank BaseStation entries win over
  the cache field by field, and the Database Writer's **Save online lookups in
  database** option makes VRS ignore the cache entirely.
- **3D traffic model coverage.** A new `model-coverage` folder, separate from
  the program, counts how many aircraft in `AircraftDatabase.sqb` each 3D
  model in the Flight Data Analysis Tool's 3D Flight Map draws, by type code,
  and ranks the types no model covers.
