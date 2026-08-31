"""
CASA (Australian Civil Aviation Safety Authority) aircraft register parser.

Downloads the CASA aircraft register ZIP, extracts the CSV, and parses it
into a local SQLite database. The ZIP contains the complete register (~1.3MB)
while the direct CSV download is incomplete.

Unlike other sources, CASA has no ICAO hex/Mode S column. Australian hex codes
(7Cxxxx) are assigned sequentially with no public algorithm. Matching to VRS
records is done by registration (VH-xxx) during the merge phase.

CSV columns of interest:
  Mark, Manu, Model, Serial, Yearmanu, regholdname, ICAOtypedesig
"""

import csv
import os
import sqlite3
import zipfile
from .utils import title_case, download_file, safe_delete, ProgressReporter
from .config import Settings


BATCH_SIZE = 10000

# Column names parse_casa reads. "Mark" is mandatory - every row is gated on
# it, so if it is absent the parse yields nothing at all.
REQUIRED_COLUMNS = ("Mark", "Manu", "Model", "regholdname", "Serial",
                    "Yearmanu", "ICAOtypedesig")


def download_casa(settings: Settings) -> bool:
    """Download the CASA aircraft register ZIP and extract the CSV."""
    zip_dest = os.path.join(settings.work_dir, "casa_register.zip")
    csv_dest = os.path.join(settings.work_dir, "casa_register.csv")

    if not download_file(settings.casa_url, zip_dest, "Downloading CASA register"):
        return False

    # Extract CSV from ZIP
    try:
        with zipfile.ZipFile(zip_dest, 'r') as zf:
            # Find the CSV file inside the ZIP
            csv_names = [n for n in zf.namelist() if n.lower().endswith('.csv')]
            if not csv_names:
                print("  ERROR: No CSV file found inside CASA ZIP.")
                safe_delete(zip_dest)
                return False
            csv_name = csv_names[0]
            print(f"  Extracting {csv_name}...")
            with zf.open(csv_name) as src, open(csv_dest, 'wb') as dst:
                dst.write(src.read())
    except zipfile.BadZipFile:
        print("  ERROR: CASA download is not a valid ZIP file.")
        safe_delete(zip_dest)
        return False
    except OSError as e:
        print(f"  ERROR: Failed to extract CASA CSV: {e}")
        return False

    safe_delete(zip_dest)
    size_kb = os.path.getsize(csv_dest) / 1024
    print(f"  CASA download complete ({size_kb:.0f} KB).")
    return True


def parse_casa(settings: Settings) -> bool:
    """Parse CASA aircraft register CSV into CASADatabase.sqb.

    Returns True on success.
    """
    csv_path = os.path.join(settings.work_dir, "casa_register.csv")
    if not os.path.exists(csv_path):
        print("  CASA CSV not found. Skipping.")
        return False

    # Count lines for progress
    with open(csv_path, 'r', encoding='utf-8-sig', errors='replace') as f:
        total_lines = sum(1 for _ in f)

    db_path = settings.casa_db_path
    safe_delete(db_path)

    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=OFF")
    conn.execute("""
        CREATE TABLE Aircraft (
            Registration TEXT,
            Manufacturer TEXT,
            Model TEXT,
            Owner TEXT,
            Serial TEXT,
            YearBuilt TEXT,
            ICAOTypeDesig TEXT
        )
    """)

    print("  Creating CASA SQL database...")
    prog = ProgressReporter("CASA")

    batch = []
    with open(csv_path, 'r', encoding='utf-8-sig', errors='replace') as f:
        reader = csv.reader(f)
        header = None
        line_num = 0
        for row in reader:
            line_num += 1
            if header is None:
                # Map column names to indices. Strip any stray BOM as well:
                # every row is gated on "Mark", so a corrupted first column
                # name would silently discard the whole register.
                header = {col.strip().lstrip('\ufeff'): i
                          for i, col in enumerate(row)}
                missing = [c for c in REQUIRED_COLUMNS if c not in header]
                if missing:
                    print("  WARNING: CASA CSV is missing expected column(s): %s"
                          % ", ".join(missing))
                    print("           Columns found: %s"
                          % ", ".join(sorted(header)))
                    if "Mark" in missing:
                        print("  ERROR: without the Mark column no aircraft can be "
                              "read. The CASA register format has probably changed.")
                        conn.close()
                        safe_delete(db_path)
                        return False
                continue

            prog.update(line_num, total_lines)

            if len(row) < 5:
                continue

            mark = _col(row, header, "Mark", "").strip()
            if not mark:
                continue

            # Build full Australian registration
            registration = "VH-" + mark if not mark.upper().startswith("VH-") else mark

            manufacturer = _col(row, header, "Manu", "").strip()
            manufacturer = title_case(manufacturer) if manufacturer else ""

            model = _col(row, header, "Model", "").strip()

            owner = _col(row, header, "regholdname", "").strip()
            owner = title_case(owner) if owner else ""

            serial = _col(row, header, "Serial", "").strip()

            year_built = _col(row, header, "Yearmanu", "").strip()

            icao_type = _col(row, header, "ICAOtypedesig", "").strip().upper()

            def _or_none(s):
                return s if s else None

            batch.append((
                _or_none(registration),
                _or_none(manufacturer),
                _or_none(model),
                _or_none(owner),
                _or_none(serial),
                _or_none(year_built),
                _or_none(icao_type),
            ))

            if len(batch) >= BATCH_SIZE:
                conn.executemany(
                    "INSERT INTO Aircraft (Registration, Manufacturer, Model, Owner, Serial, YearBuilt, ICAOTypeDesig) VALUES (?,?,?,?,?,?,?)",
                    batch
                )
                conn.commit()
                batch.clear()

    if batch:
        conn.executemany(
            "INSERT INTO Aircraft (Registration, Manufacturer, Model, Owner, Serial, YearBuilt, ICAOTypeDesig) VALUES (?,?,?,?,?,?,?)",
            batch
        )
        conn.commit()

    conn.execute("CREATE INDEX IF NOT EXISTS idx_casa_reg ON Aircraft(Registration)")
    conn.commit()
    rows_written = conn.execute("SELECT COUNT(*) FROM Aircraft").fetchone()[0]
    conn.close()
    prog.done()

    if not rows_written:
        print("  ERROR: CASA register parsed to 0 aircraft - the downloaded file "
              "was empty or its format has changed. Keeping the CSV at %s for "
              "inspection." % csv_path)
        return False

    safe_delete(csv_path)
    print("  CASA database creation complete (%s aircraft)."
          % "{:,}".format(rows_written))
    return True


def _col(row, header, col_name, default=""):
    """Get a column value by name, with fallback."""
    idx = header.get(col_name)
    if idx is not None and idx < len(row):
        return row[idx]
    return default
