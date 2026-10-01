"""
FAA database parser.

Downloads and parses the FAA Releasable Aircraft database (MASTER.txt and ACFTREF.txt)
into a local SQLite database.

Efficiency improvements over VB.NET version:
- Uses parameterized queries (no SQL injection, no manual quote-escaping)
- Uses executemany() with batches inside a single transaction
- Reads files line-by-line instead of loading entire file into memory
- Uses Python string slicing (equivalent to Mid$) for fixed-width fields
"""

import os
import sqlite3
from datetime import datetime
from .utils import (
    title_case, remove_non_ascii, download_file, extract_zip,
    safe_delete, ProgressReporter
)
from .config import Settings


# FAA MASTER.txt fixed-width column positions (0-based)
# These correspond to the VB.NET Mid$() calls with 1-based offsets
MASTER_REGISTRATION_END = 5          # Left("N" & line, 6) -> "N" + line[0:5]
MASTER_TYPE_START = 37               # Mid$(line, 38, 7)
MASTER_TYPE_LEN = 7
MASTER_YEAR_START = 51               # Mid$(line, 52, 4)
MASTER_YEAR_LEN = 4
MASTER_OWNER_START = 58              # Mid$(line, 59, 50)
MASTER_OWNER_LEN = 50
MASTER_MANU_KIT_START = 549          # Mid$(line, 550, 30)
MASTER_MANU_KIT_LEN = 30
MASTER_TYPE_KIT_START = 580          # Mid$(line, 581, 20)
MASTER_TYPE_KIT_LEN = 20
# ICAO hex is near end of line: Mid$(line, Len(line) - 10, 6)

# FAA ACFTREF.txt fixed-width column positions
REF_TYPE_END = 7                     # Left(line, 7)
REF_MANU_START = 8                   # Mid$(line, 9, 30)
REF_MANU_LEN = 30
REF_MODEL_START = 39                 # Mid$(line, 40, 20)
REF_MODEL_LEN = 20
REF_SPECIES_POS = 60                 # Mid$(line, 61, 1)
REF_ENGINE_TYPE_START = 62           # Mid$(line, 63, 2)
REF_ENGINE_TYPE_LEN = 2
REF_ENGINE_NUM_START = 69            # Mid$(line, 70, 2)
REF_ENGINE_NUM_LEN = 2
REF_WTC_START = 76                   # Mid$(line, 77, 7)
REF_WTC_LEN = 7

# WTC class mapping
WTC_MAP = {
    "CLASS 1": "L",
    "CLASS 2": "M",
    "CLASS 3": "L",
    "CLASS 4": "UAV",
}

BATCH_SIZE = 10000  # Increased from VB.NET's 1000 for better throughput


def download_faa(settings: Settings) -> bool:
    """Download and extract the FAA database ZIP."""
    zip_path = os.path.join(settings.work_dir, "ReleasableAircraft.zip")
    extract_dir = os.path.join(settings.work_dir, "_faa_extract")

    if not download_file(settings.faa_url, zip_path, "Downloading FAA database"):
        return False

    try:
        extract_zip(zip_path, extract_dir, "Extracting FAA database")
    except Exception as e:
        print(f"  ERROR: Failed to extract FAA database: {e}")
        return False
    safe_delete(zip_path)
    return True


def parse_faa(settings: Settings) -> bool:
    """Parse FAA MASTER.txt and ACFTREF.txt into FAADatabase.sqb.

    Returns True on success.
    """
    extract_dir = os.path.join(settings.work_dir, "_faa_extract")
    master_path = os.path.join(extract_dir, "MASTER.txt")
    acftref_path = os.path.join(extract_dir, "ACFTREF.txt")

    if not os.path.exists(master_path):
        print(f"  ERROR: {master_path} not found.")
        return False

    db_path = settings.faa_db_path
    safe_delete(db_path)

    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=OFF")

    # Create tables
    conn.execute("""
        CREATE TABLE Master (
            Registration TEXT,
            ICAO TEXT,
            Type TEXT,
            Owner_Name TEXT,
            Manufacturer_Kit TEXT,
            Type_Kit TEXT,
            Year TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE Aircraft_Reference (
            Type TEXT,
            Manufacturer TEXT,
            Model TEXT,
            Engine_Type TEXT,
            Species TEXT,
            Engine_Num TEXT,
            WTC TEXT
        )
    """)

    # ---- Parse MASTER.txt ----
    print("  Creating FAA SQL database...")
    print("    'Master' table...")
    prog = ProgressReporter("Master")

    # Count lines first for progress
    with open(master_path, 'r', encoding='latin-1', errors='replace') as f:
        total_lines = sum(1 for _ in f)

    batch = []
    with open(master_path, 'r', encoding='latin-1', errors='replace') as f:
        for line_num, line in enumerate(f):
            if line_num == 0:
                continue  # Skip header

            if len(line) < 590:
                continue  # Skip malformed lines

            registration = "N" + line[:5].rstrip()
            faa_type = line[MASTER_TYPE_START:MASTER_TYPE_START + MASTER_TYPE_LEN]
            year_built = line[MASTER_YEAR_START:MASTER_YEAR_START + MASTER_YEAR_LEN]

            # ICAO hex is near end: 6 chars starting at len-11
            icao = line[len(line.rstrip()) - 11: len(line.rstrip()) - 5]

            # Owner name
            owner = line[MASTER_OWNER_START:MASTER_OWNER_START + MASTER_OWNER_LEN].rstrip()
            owner = title_case(owner)

            # Manufacturer Kit
            manu_kit = line[MASTER_MANU_KIT_START:MASTER_MANU_KIT_START + MASTER_MANU_KIT_LEN].rstrip()
            if manu_kit:
                manu_kit = title_case(remove_non_ascii(manu_kit))
            else:
                manu_kit = None

            # Type Kit
            type_kit = line[MASTER_TYPE_KIT_START:MASTER_TYPE_KIT_START + MASTER_TYPE_KIT_LEN].rstrip()
            if type_kit:
                type_kit = remove_non_ascii(type_kit)
            else:
                type_kit = None

            batch.append((registration, icao, faa_type, owner, manu_kit, type_kit, year_built))

            if len(batch) >= BATCH_SIZE:
                conn.executemany(
                    "INSERT INTO Master (Registration, ICAO, Type, Owner_Name, Manufacturer_Kit, Type_Kit, Year) VALUES (?,?,?,?,?,?,?)",
                    batch
                )
                conn.commit()
                batch.clear()
                prog.update(line_num, total_lines)

    if batch:
        conn.executemany(
            "INSERT INTO Master (Registration, ICAO, Type, Owner_Name, Manufacturer_Kit, Type_Kit, Year) VALUES (?,?,?,?,?,?,?)",
            batch
        )
        conn.commit()
        batch.clear()
    prog.done()

    # ---- Parse ACFTREF.txt ----
    # The reference table is what gives every US aircraft its manufacturer and
    # its ICAO type code (and so its silhouette and map icon). Merging without
    # it blanks all three, so a missing or unparsable file must never pass
    # quietly - see _ensure_reference() below.
    ref_rows = 0
    if os.path.exists(acftref_path):
        print("    'Aircraft_Reference' table...")
        prog = ProgressReporter("Reference")

        with open(acftref_path, 'r', encoding='latin-1', errors='replace') as f:
            total_lines = sum(1 for _ in f)

        short_lines = 0
        batch = []
        with open(acftref_path, 'r', encoding='latin-1', errors='replace') as f:
            for line_num, line in enumerate(f):
                if line_num == 0:
                    continue

                if len(line) < 77:
                    short_lines += 1
                    continue

                faa_type = line[:REF_TYPE_END]
                manufacturer = title_case(line[REF_MANU_START:REF_MANU_START + REF_MANU_LEN].rstrip())
                model = line[REF_MODEL_START:REF_MODEL_START + REF_MODEL_LEN].rstrip()
                species = line[REF_SPECIES_POS]
                engine_type = line[REF_ENGINE_TYPE_START:REF_ENGINE_TYPE_START + REF_ENGINE_TYPE_LEN].rstrip()
                engine_num = line[REF_ENGINE_NUM_START:REF_ENGINE_NUM_START + REF_ENGINE_NUM_LEN]
                # Strip leading zero from engine number
                engine_num = engine_num.lstrip('0') or engine_num[-1:]

                wtc_raw = line[REF_WTC_START:REF_WTC_START + REF_WTC_LEN].rstrip()
                wtc = WTC_MAP.get(wtc_raw, "?")

                batch.append((faa_type, manufacturer, model, engine_type, species, engine_num, wtc))

                if len(batch) >= BATCH_SIZE:
                    conn.executemany(
                        "INSERT INTO Aircraft_Reference (Type, Manufacturer, Model, Engine_Type, Species, Engine_Num, WTC) VALUES (?,?,?,?,?,?,?)",
                        batch
                    )
                    conn.commit()
                    batch.clear()
                    prog.update(line_num, total_lines)

        if batch:
            conn.executemany(
                "INSERT INTO Aircraft_Reference (Type, Manufacturer, Model, Engine_Type, Species, Engine_Num, WTC) VALUES (?,?,?,?,?,?,?)",
                batch
            )
            conn.commit()
        prog.done()

        ref_rows = conn.execute("SELECT COUNT(*) FROM Aircraft_Reference").fetchone()[0]
        if ref_rows == 0 and short_lines:
            print(f"    WARNING: every one of the {short_lines:,} data lines in "
                  f"ACFTREF.txt was shorter than the 77 columns the parser "
                  f"expects - the file layout has probably changed.")
    else:
        print("    WARNING: ACFTREF.txt is not in the FAA download.")

    # Create indexes for faster lookups during merge
    print("    Creating indexes...")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_master_icao ON Master(ICAO)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_ref_type ON Aircraft_Reference(Type)")
    conn.commit()

    ok = _ensure_reference(conn, settings, ref_rows)

    conn.close()

    # Clean up extracted files
    safe_delete(extract_dir)
    if not ok:
        return False
    print("  FAA database creation complete.")
    return True


def _ensure_reference(conn: sqlite3.Connection, settings: Settings,
                      ref_rows: int) -> bool:
    """Guarantee FAADatabase.sqb has reference data, or fail the parse.

    A good parse refreshes the cache. A bad one is refilled from it, because
    the reference table describes types rather than individual aircraft and so
    goes stale very slowly - a copy from last week still resolves nearly every
    aircraft, which is worth far more than the empty table it replaces.

    Returns False when no reference data can be had from either source. The
    caller must then abandon the FAA update: merging a 0-row reference table
    silently blanks Manufacturer and ModelIcao for every US aircraft in VRS,
    which is how the 2026-09-30 release wiped 247,000 type codes.
    """
    cache = settings.faa_ref_cache_path

    if ref_rows > 0:
        print(f"    {ref_rows:,} reference records parsed.")
        try:
            # Build beside the cache and rename over it, so an interruption
            # mid-write cannot leave us with neither the old copy nor a
            # complete new one.
            tmp = cache + ".tmp"
            safe_delete(tmp)
            conn.execute("ATTACH DATABASE ? AS refcache", (tmp,))
            conn.execute("CREATE TABLE refcache.Aircraft_Reference AS "
                         "SELECT * FROM main.Aircraft_Reference")
            conn.commit()
            conn.execute("DETACH DATABASE refcache")
            os.replace(tmp, cache)
            print(f"    Cached reference data as {os.path.basename(cache)}.")
        except Exception as e:
            # A cache we failed to write is not worth failing the run over;
            # this release's own reference data is already in place.
            print(f"    WARNING: could not refresh the reference cache: {e}")
        return True

    if not os.path.exists(cache):
        print("  ERROR: the FAA release carries no usable aircraft reference "
              "data and no cached copy exists.")
        print("         Merging now would erase the manufacturer and ICAO type "
              "code of every US aircraft in the VRS database, so the FAA "
              "update has been abandoned.")
        print("         Fix: restore FAAReference.sqb from a backup, or build "
              "it from an older FAADatabase snapshot.")
        return False

    try:
        conn.execute("ATTACH DATABASE ? AS refcache", (cache,))
        conn.execute("INSERT INTO main.Aircraft_Reference "
                     "SELECT * FROM refcache.Aircraft_Reference")
        conn.commit()
        conn.execute("DETACH DATABASE refcache")
        restored = conn.execute("SELECT COUNT(*) FROM Aircraft_Reference").fetchone()[0]
    except Exception as e:
        print(f"  ERROR: cached reference data could not be read: {e}")
        return False

    if restored == 0:
        print("  ERROR: the cached reference data is itself empty.")
        return False

    age = ""
    try:
        days = (datetime.now() - datetime.fromtimestamp(os.path.getmtime(cache))).days
        age = f", cached {days} day{'s' if days != 1 else ''} ago"
    except OSError:
        pass
    print(f"    Restored {restored:,} reference records from "
          f"{os.path.basename(cache)}{age}.")
    print("    NOTE: manufacturers and type codes come from cached data until "
          "the FAA ships ACFTREF.txt again.")
    return True
