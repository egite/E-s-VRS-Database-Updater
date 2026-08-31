# v2.1.0

Adds built-in editors for the two customization files, a user manual, and a
working Australian register. Also fixes two silent data problems and a broken
download address.

**Downloads:** `VRS_Database_Updater_x64.exe`, `VRS_Database_Updater_x86.exe`,
and `Rules.csv`.

`Rules.csv` is an example ruleset. Put it in your working folder to start from
it, or leave it out and build your own in the new Rules Editor — it is no
longer carried inside the program, so none of it applies unless you place the
file there yourself.

`Sils.csv` is no longer included. The silhouette mappings are built into the
program, and the Silhouette Editor creates the file only when you override
something.

---

## Built-in editors — no spreadsheet needed

**Tools → Rules Editor** and **Tools → Silhouette Editor** replace hand-editing
the CSV files in Excel. Both understand what the columns actually mean, and
both keep a `.bak` of the previous file every time you save.

**Rules Editor**
- Match fields and change fields laid out side by side, with the `!` exclusion
  as a checkbox rather than punctuation you have to remember
- Reorder rules — the order matters, because the first rule that matches an
  aircraft wins
- **Test Against Database** lists exactly which aircraft a rule would affect,
  with a count, before you save it
- **Check Rules** finds rules that can never fire, rules hidden by an earlier
  rule, and rules that match nothing in your database
- Right-click any aircraft in **Tools → Search Database** to build a rule from it
- **Apply Rules** on the main window applies your rules to the VRS database on
  its own, with no downloading or merging — seconds instead of a full run

**Silhouette Editor**
- Each manufacturer and model alias on its own line instead of crammed into one
  spreadsheet cell — some rows hold more than twenty
- Search across 1,900+ mappings by manufacturer, model or type code
- **Test Lookup** resolves a manufacturer and model exactly the way an update
  does, including the built-in manufacturer rewrites, and tells you which row
  won and where it came from
- **Check Rows** reports duplicate aliases, unreachable rows, and rows that can
  never match

## Sils.csv now layers over the built-in mappings

Previously, supplying a `Sils.csv` **replaced** the program's built-in mapping
table entirely — so a file missing a mapping lost it. It is now an overlay:

- Your rows are checked first and override the built-in mapping
- Anything your file does not cover still resolves from the built-in table
- A row with an empty Type **and** Remap switches a built-in mapping off
- Your file therefore only needs rows for what you actually want to change

Measured against a real 884,838-aircraft database, this alone gave 353 aircraft
a silhouette they did not have and corrected 346 more, with no regressions.

## Built-in user manual

**Help → User Manual** — eleven sections covering the main window, how an
update runs, settings, rules, silhouettes, the working folder and
troubleshooting. Searchable, and can be saved out as a text file.

## Fixes

**CASA register produced nothing, silently.** The CSV carries a byte-order
mark, which attached itself to the first column name and made every row fail
the registration check. The parser reported success, wrote an empty database,
deleted its input, and logged no error — for every run since CASA support was
added. Fixed, and the parser now reports a missing column by name and refuses
to accept a zero-row result. Recovers 16,684 Australian aircraft, 8,443 of
which match aircraft already in a typical database. The same byte-order-mark
guard was applied to the NZ CAA and OpenSky parsers.

**NZ CAA download address was wrong.** CAA dropped a hyphen from the filename,
so the download had been failing with a 404 on every run. Address corrected.

**Repeated download failures are now reported.** A source that fails once is
usually a blip and the run carries on with the copy on disk. After 14
consecutive failures the program says the address is probably wrong and names
the URL. The warning waits until it is acknowledged, so an unattended run
cannot lose it — and an unattended run still closes itself as normal. It clears
itself as soon as the download works again.

**Rules.csv parsing.** Values containing commas are now quoted correctly, and
rows saved short by a spreadsheet are no longer silently dropped. The "From dB"
column now works — it had never taken effect in either the VB.NET original or
the Python port.

## Under the hood

- Silhouette lookups are memoized. The table scan is linear and runs once or
  twice per aircraft; on a real database this cut about four minutes of lookup
  time to under a second, which more than pays for the larger overlaid table.
- `rules.py` and `sils.py` are now the single parser for each file, shared by
  the merge and the editors, so they cannot drift apart.
- Loading `Sils.csv` no longer treats the header row as data.

## Notes

- `UpdatedUtc` is deliberately refreshed on every touched record. VRS treats it
  as a cache-expiry stamp and re-queries its own online lookup service —
  overwriting operator, model and type — for anything older than 28 days. Not
  refreshing it would let VRS undo the merge and your custom rules.
- Existing `Rules.csv` and `Sils.csv` files work unchanged; the formats have
  not changed and files remain editable by hand.
