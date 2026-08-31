"""
NZ CAA (New Zealand Civil Aviation Authority) aircraft register parser.

Downloads and parses the NZ CAA aircraft register CSV into a local SQLite database.
The CSV is a direct download (~945KB), no ZIP extraction needed.

CSV columns (mapped by header name):
  Model Category, Registration Mark, Registered on, Manufacturer, Model,
  Serial No., MCTOW (Kg), Owner Name, Owner Address, Mode S Code HEX,
  Mode S Code Binary, Flight manual no.
"""

import csv
import os
import sqlite3
from .utils import title_case, download_file, safe_delete, ProgressReporter
from .config import Settings


BATCH_SIZE = 10000


def download_nz_caa(settings: Settings) -> bool:
    """Download the NZ CAA aircraft register CSV."""
    dest = os.path.join(settings.work_dir, "nz_caa_register.csv")
    if not download_file(settings.nz_caa_url, dest, "Downloading NZ CAA register"):
        return False

    # Verify we got a CSV and not an HTML bot-challenge page
    try:
        with open(dest, 'r', encoding='utf-8-sig', errors='replace') as f:
            first_line = f.read(512)
        if '<html' in first_line.lower():
            print("  ERROR: NZ CAA download returned an HTML page (bot protection).")
            print("  Download the CSV manually from the NZ CAA website and place it")
            print(f"  in: {dest}")
            safe_delete(dest)
            return False
    except OSError as e:
        print(f"  ERROR: Could not read downloaded NZ CAA file: {e}")
        return False

    size_kb = os.path.getsize(dest) / 1024
    print(f"  NZ CAA download complete ({size_kb:.0f} KB).")
    return True


def parse_nz_caa(settings: Settings) -> bool:
    """Parse NZ CAA aircraft register CSV into NZCAADatabase.sqb.

    Returns True on success.
    """
    csv_path = os.path.join(settings.work_dir, "nz_caa_register.csv")
    if not os.path.exists(csv_path):
        print("  NZ CAA CSV not found. Skipping.")
        return False

    # Count lines for progress
    with open(csv_path, 'r', encoding='utf-8-sig', errors='replace') as f:
        total_lines = sum(1 for _ in f)

    db_path = settings.nz_caa_db_path
    safe_delete(db_path)

    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=OFF")
    conn.execute("""
        CREATE TABLE Aircraft (
            ICAO TEXT,
            Registration TEXT,
            Manufacturer TEXT,
            Model TEXT,
            Owner TEXT,
            Serial TEXT,
            YearBuilt TEXT
        )
    """)

    print("  Creating NZ CAA SQL database...")
    prog = ProgressReporter("NZ CAA")

    batch = []
    with open(csv_path, 'r', encoding='utf-8-sig', errors='replace') as f:
        reader = csv.reader(f)
        header = None
        line_num = 0
        for row in reader:
            line_num += 1
            if header is None:
                # Map column names to indices
                header = {col.strip(): i for i, col in enumerate(row)}
                continue

            prog.update(line_num, total_lines)

            if len(row) < 5:
                continue

            # Extract fields by column name, falling back gracefully
            icao = _col(row, header, "Mode S Code HEX", "").strip().upper()
            if not icao:
                continue

            reg_mark = _col(row, header, "Registration Mark", "").strip()
            if reg_mark and not reg_mark.upper().startswith("ZK-"):
                registration = "ZK-" + reg_mark
            else:
                registration = reg_mark

            manufacturer = _col(row, header, "Manufacturer", "").strip()
            manufacturer = title_case(manufacturer) if manufacturer else ""

            model = _col(row, header, "Model", "").strip()

            owner = _col(row, header, "Owner Name", "").strip()
            owner = title_case(owner) if owner else ""

            serial = _col(row, header, "Serial No.", "").strip()

            # NZ CAA CSV has "Registered on" date but no year-built field;
            # leave YearBuilt empty so it can be filled from other sources
            year_built = ""

            def _or_none(s):
                return s if s else None

            batch.append((
                icao,
                _or_none(registration),
                _or_none(manufacturer),
                _or_none(model),
                _or_none(owner),
                _or_none(serial),
                _or_none(year_built),
            ))

            if len(batch) >= BATCH_SIZE:
                conn.executemany(
                    "INSERT INTO Aircraft (ICAO, Registration, Manufacturer, Model, Owner, Serial, YearBuilt) VALUES (?,?,?,?,?,?,?)",
                    batch
                )
                conn.commit()
                batch.clear()

    if batch:
        conn.executemany(
            "INSERT INTO Aircraft (ICAO, Registration, Manufacturer, Model, Owner, Serial, YearBuilt) VALUES (?,?,?,?,?,?,?)",
            batch
        )
        conn.commit()

    conn.execute("CREATE INDEX IF NOT EXISTS idx_nzcaa_icao ON Aircraft(ICAO)")
    conn.commit()
    conn.close()
    prog.done()

    safe_delete(csv_path)
    print("  NZ CAA database creation complete.")
    return True


def _col(row, header, col_name, default=""):
    """Get a column value by name, with fallback."""
    idx = header.get(col_name)
    if idx is not None and idx < len(row):
        return row[idx]
    return default
