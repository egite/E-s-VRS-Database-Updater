# v2.2.0

A data-safety release. On 2026-09-30 the FAA removed `ACFTREF.txt` from
`ReleasableAircraft.zip`, and that one missing file silently stripped the
manufacturer and ICAO type code from every US aircraft in the VRS database —
roughly 300,000 manufacturers and 247,000 type codes, which in Virtual Radar
Server means a generic marker for almost everything on the map.

This release makes that class of failure survivable. A source that stops
supplying a field no longer costs you the data you already had.

**Downloads:** `VRS_Database_Updater_x64.exe`, `VRS_Database_Updater_x86.exe`,
and `Rules.csv`.

Upgrading is just replacing the executable. Nothing in your working folder or
your `Rules.csv` needs to change.

---

## What went wrong

The FAA's releasable database ships `MASTER.txt` (one row per aircraft) and
shipped `ACFTREF.txt` (one row per aircraft *type*, carrying the manufacturer
and the model that the silhouette lookup turns into an ICAO type code). The
rebuild on 2026-09-29 at 23:23 UTC dropped `ACFTREF.txt` entirely, along with
`DEALER.txt`, `ENGINE.txt` and `ardata.pdf`. No replacement address was
published.

The updater did not notice. Four things went wrong at once:

- A missing `ACFTREF.txt` meant the parse block was skipped by an
  `if os.path.exists(...)` with no message of any kind
- A zero-row reference table was written and accepted as success
- The merge read that table inside `except Exception: pass`, so even a hard
  failure was invisible
- `Model` fell back to the value already in VRS when the source had none, but
  `Manufacturer` and `ModelIcao` did not — they were overwritten with blanks

The run reported success and the error log was empty. The result looked
baffling from the outside: models intact, manufacturers gone, silhouettes gone,
and only US registrations affected.

## Fixes

**Reference data is now cached between runs.** A good parse saves the FAA
reference table to `FAAReference.sqb` in your working folder; a release that
omits or mangles `ACFTREF.txt` is refilled from that copy and the run
continues. The reference data describes aircraft *types* rather than individual
aircraft, so it ages slowly — a copy from last week still resolves very nearly
every aircraft. The cache is written to a temporary file and renamed into
place, so an interrupted run cannot leave you with neither the old copy nor a
complete new one.

**No source can blank a field it simply failed to report.** `Manufacturer` now
falls back to the value already in VRS exactly as `Model` always did, and a
failed silhouette lookup keeps the existing `ModelIcao` instead of clearing it.
An absent value means "the source did not tell us", never "this aircraft has no
manufacturer". This is the change that matters most — it is not specific to the
FAA, and it protects the CCAR, CASA, NZ CAA and OpenSky passes the same way.

**Failures are loud.** A missing `ACFTREF.txt` is reported. A layout change is
reported with the count of lines that were too short to parse, so you can tell
a changed format from a missing file. An unreadable reference table prints the
actual error instead of being swallowed.

**The FAA step refuses to run rather than cause damage.** If the download has
no usable reference data *and* no cache exists, the FAA update is abandoned and
the database is left untouched, with an explanation of why. An empty reference
table likewise skips the FAA merge instead of writing blanks over 300,000
records.

## If you were hit by this

Rebuild with this version and run normally. Manufacturers and type codes are
restored from the cache, so no manual repair is needed.

If your working folder has no `FAAReference.sqb` yet and you keep dated
`FAADatabase <ddmmyyyy>.sqb` snapshots, the program will tell you it has
nothing to fall back on rather than proceeding — copy the `Aircraft_Reference`
table out of any snapshot from before 2026-09-30 into `FAAReference.sqb`, or
restore your VRS database from one of the dated backups and run again.

## Known limitation

Aircraft types certified after the cached reference data was captured are not
in it, and those aircraft will come through with no manufacturer and no
silhouette. If the FAA restores `ACFTREF.txt`, a normal run picks it up and
refreshes the cache automatically — no action needed.
