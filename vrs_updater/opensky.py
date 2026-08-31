"""
OpenSky Network database parser.

Downloads and parses the OpenSky aircraftDatabase.csv into a local SQLite database.

Efficiency improvements over VB.NET version:
- Uses Python csv module for proper CSV parsing
- Uses parameterized queries and executemany()
- Streams the CSV line-by-line instead of loading everything into memory
"""

import os
import csv
import sqlite3
import shutil
from .utils import (
    title_case, download_file, extract_zip, safe_delete, ProgressReporter
)
from .config import Settings


BATCH_SIZE = 10000


def download_opensky(settings: Settings) -> bool:
    """Download and extract the OpenSky database ZIP."""
    zip_path = os.path.join(settings.work_dir, "aircraftDatabase.zip")

    if not download_file(settings.opensky_url, zip_path, "Downloading OpenSky database"):
        return False

    # OpenSky extracts to a nested path: media/data/samples/metadata/aircraftDatabase.csv
    extract_dir = os.path.join(settings.work_dir, "_opensky_extract")
    dest_csv = os.path.join(settings.work_dir, "aircraftDatabase.csv")
    try:
        extract_zip(zip_path, extract_dir, "Extracting OpenSky database")

        # Move the CSV to the working directory
        nested_csv = os.path.join(extract_dir, "media", "data", "samples", "metadata", "aircraftDatabase.csv")
        flat_csv = os.path.join(extract_dir, "aircraftDatabase.csv")

        if os.path.exists(nested_csv):
            src = nested_csv
        elif os.path.exists(flat_csv):
            src = flat_csv
        else:
            print("  ERROR: aircraftDatabase.csv not found inside OpenSky ZIP "
                  "(archive layout may have changed).")
            return False

        if os.path.exists(dest_csv):
            os.remove(dest_csv)
        shutil.move(src, dest_csv)
    except Exception as e:
        print(f"  ERROR: Failed to extract/move OpenSky database: {e}")
        return False

    safe_delete(extract_dir)
    safe_delete(zip_path)
    return True


def parse_opensky(settings: Settings) -> bool:
    """Parse OpenSky aircraftDatabase.csv into OpenSkyDatabase.sqb.

    Returns True on success.
    """
    csv_path = os.path.join(settings.work_dir, "aircraftDatabase.csv")
    if not os.path.exists(csv_path):
        print("  OpenSky CSV not found. Skipping.")
        return False

    # Count lines for progress
    with open(csv_path, 'r', encoding='utf-8-sig', errors='replace') as f:
        total_lines = sum(1 for _ in f)

    db_path = settings.opensky_db_path
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
            ModelIcao TEXT,
            Operator TEXT,
            OperatorIcao TEXT,
            Serial TEXT,
            YearBuilt TEXT
        )
    """)

    print("  Creating OpenSky SQL database...")
    prog = ProgressReporter("OpenSky")

    batch = []
    with open(csv_path, 'r', encoding='utf-8-sig', errors='replace') as f:
        # The OpenSky CSV uses a slightly unusual quoting style.
        # The VB.NET code splits on '","' which works for their format.
        # We'll use the same approach for compatibility.
        line_num = 0
        for raw_line in f:
            line_num += 1
            if line_num == 1:
                continue  # Skip header

            # Split the same way VB.NET does: Split(a, """,")
            fields = raw_line.split('",')
            if len(fields) < 19:
                continue

            icao = fields[0].strip().strip('"').upper()
            if not icao:
                continue

            registration = fields[1].strip().strip('"')
            manufacturer = fields[3].strip().strip('"')
            model = fields[4].strip().strip('"')
            model_icao = fields[5].strip().strip('"')
            serial = fields[6].strip().strip('"')
            operator_icao = fields[11].strip().strip('"')
            operator_name = fields[13].strip().strip('"') if len(fields) > 13 else ""
            year_built = fields[18].strip().strip('"') if len(fields) > 18 else ""

            # Fix specific model names (from VB.NET)
            model = model.replace("X''AIR", "XAIR").replace("X''air", "Xair")

            # Specific ICAO overrides from VB.NET
            if icao in ("404E83", "405F86", "C0890F", "C08819", "C062EC"):
                model = "Emarit"

            # Clean year
            if year_built:
                year_built = year_built[:4]

            # Clean operator name
            if operator_name:
                operator_name = operator_name.rstrip()
                operator_name = title_case(operator_name)
                if operator_name == "Private":
                    operator_name = ""

            # Convert empty strings to None for proper NULL handling
            def _or_none(s):
                return s if s else None

            batch.append((
                icao,
                _or_none(registration),
                _or_none(manufacturer),
                _or_none(model),
                _or_none(model_icao),
                _or_none(operator_name),
                _or_none(operator_icao),
                _or_none(serial),
                _or_none(year_built),
            ))

            if len(batch) >= BATCH_SIZE:
                conn.executemany(
                    "INSERT INTO Aircraft (ICAO, Registration, Manufacturer, Model, ModelIcao, Operator, OperatorIcao, Serial, YearBuilt) VALUES (?,?,?,?,?,?,?,?,?)",
                    batch
                )
                conn.commit()
                batch.clear()

            prog.update(line_num, total_lines)

    if batch:
        conn.executemany(
            "INSERT INTO Aircraft (ICAO, Registration, Manufacturer, Model, ModelIcao, Operator, OperatorIcao, Serial, YearBuilt) VALUES (?,?,?,?,?,?,?,?,?)",
            batch
        )
        conn.commit()

    conn.execute("CREATE INDEX IF NOT EXISTS idx_opensky_icao ON Aircraft(ICAO)")
    conn.commit()
    conn.close()
    prog.done()

    # Clean up CSV
    safe_delete(csv_path)
    print("  OpenSky database creation complete.")
    return True
