"""
CCAR (Canadian Civil Aircraft Register) database parser.

Parses carscurr.txt and carsownr.txt from the CCAR ZIP into a local SQLite database.

Efficiency improvements over VB.NET version:
- The manual Bin2Hex() function (16 if/else branches) is replaced with int(s, 2)
- Uses parameterized queries and executemany()
- Reads files efficiently with csv-aware splitting
"""

import os
import re
import sqlite3
from .utils import (
    title_case, download_file, extract_zip, safe_delete, ProgressReporter
)
from .config import Settings

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False


BATCH_SIZE = 10000
CCAR_URL = "https://wwwapps.tc.gc.ca/saf-sec-sur/2/ccarcs-riacc/DDZip.aspx"


def download_ccar(dest_path: str) -> bool:
    """Download the CCAR database ZIP via ASP.NET form postback.

    The CCAR website uses an ASP.NET form with __VIEWSTATE.
    We GET the page first to extract hidden fields, then POST
    with the Download button to receive the ZIP.

    Returns True on success.
    """
    if not HAS_REQUESTS:
        print("  ERROR: 'requests' package required for CCAR download.")
        return False

    print(f"  Downloading CCAR database: {CCAR_URL}")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/131.0.0.0 Safari/537.36"
    }

    try:
        session = requests.Session()
        session.headers.update(headers)

        # Step 1: GET the page to extract hidden form fields
        resp = session.get(CCAR_URL, timeout=60)
        resp.raise_for_status()

        viewstate = re.search(r'__VIEWSTATE[^G].*?value="([^"]*)"', resp.text)
        viewstate_gen = re.search(r'__VIEWSTATEGENERATOR.*?value="([^"]*)"', resp.text)
        event_validation = re.search(r'__EVENTVALIDATION.*?value="([^"]*)"', resp.text)

        data = {
            "__VIEWSTATE": viewstate.group(1) if viewstate else "",
            "__VIEWSTATEGENERATOR": viewstate_gen.group(1) if viewstate_gen else "",
            "__EVENTVALIDATION": event_validation.group(1) if event_validation else "",
            "ctl00$ContentPlaceHolder1$btnDownload": "Download",
        }

        # Step 2: POST to trigger the download (streamed for progress)
        resp2 = session.post(CCAR_URL, data=data, timeout=120, stream=True)
        resp2.raise_for_status()

        total = int(resp2.headers.get("content-length", 0))
        downloaded = 0
        prog = ProgressReporter("Downloading CCAR database")

        if os.path.exists(dest_path):
            os.remove(dest_path)

        with open(dest_path, "wb") as f:
            for chunk in resp2.iter_content(chunk_size=65536):
                f.write(chunk)
                downloaded += len(chunk)
                if total:
                    prog.update(downloaded, total)
        prog.done()

        # Verify we got a ZIP
        with open(dest_path, "rb") as f:
            magic = f.read(2)
        if magic != b"PK":
            print("  ERROR: CCAR download did not return a ZIP file.")
            safe_delete(dest_path)
            return False

        size_mb = os.path.getsize(dest_path) / (1024 * 1024)
        print(f"  CCAR download complete ({size_mb:.1f} MB).")
        return True

    except Exception as e:
        print(f"  ERROR downloading CCAR: {e}")
        return False


def bin_to_hex(binary_str: str) -> str:
    """Convert a 24-bit binary string to 6-char uppercase hex.

    Replaces the VB.NET Bin2Hex() that had 16 if/else branches per nibble.
    The CCAR data stores ICAO addresses as 24-bit binary strings in the CSV.

    Example: "110000001000100100001111" -> "C0890F"
    """
    # Pad to multiple of 4 if needed
    padded = binary_str.ljust(24, '0')
    value = int(padded, 2)
    return f"{value:06X}"


def _parse_ccar_csv_field(field: str) -> str:
    """Strip surrounding quotes from a CCAR CSV field."""
    return field.strip().strip('"')


def parse_ccar(settings: Settings) -> bool:
    """Parse CCAR carscurr.txt and carsownr.txt into CCARDatabase.sqb.

    Downloads the CCAR ZIP automatically if download_ccar is enabled.
    Returns True on success.
    """
    zip_path = os.path.join(settings.work_dir, "ccarcsdb.zip")
    if settings.download_ccar or not os.path.exists(zip_path):
        if not download_ccar(zip_path):
            if not os.path.exists(zip_path):
                print("  CCAR download failed. Skipping CCAR database.")
                return False
            print("  CCAR download failed. Using existing ZIP.")

    extract_dir = os.path.join(settings.work_dir, "_ccar_extract")
    extract_zip(zip_path, extract_dir, "Extracting CCAR database")

    curr_path = os.path.join(extract_dir, "carscurr.txt")
    owner_path = os.path.join(extract_dir, "carsownr.txt")

    if not os.path.exists(curr_path):
        print(f"  ERROR: {curr_path} not found in ZIP.")
        safe_delete(extract_dir)
        return False

    # Read owner data into a dict keyed by registration for fast lookup
    # (VB.NET uses a linear scan with index tracking - this is O(1) per lookup)
    print("  Building owner lookup table...")
    owners = {}
    if os.path.exists(owner_path):
        with open(owner_path, 'r', encoding='latin-1', errors='replace') as f:
            for line in f:
                fields = line.split('","')
                if len(fields) >= 2:
                    reg_key = fields[0]  # Keep raw (with quotes) for matching
                    owner_name = fields[1].strip().rstrip('"')
                    owner_name = owner_name.rstrip()
                    if reg_key not in owners:  # First owner entry wins
                        owners[reg_key] = owner_name

    # Parse registration data
    db_path = settings.ccar_db_path
    safe_delete(db_path)

    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=OFF")
    conn.execute("""
        CREATE TABLE Master (
            Registration TEXT,
            ICAO TEXT,
            CCAR_Type TEXT,
            Owner_Name TEXT,
            Manufacturer TEXT,
            Serial TEXT,
            Year TEXT
        )
    """)

    print("  Creating CCAR SQL database...")
    prog = ProgressReporter("CCAR")

    # Count lines for progress (stream instead of loading all into memory)
    with open(curr_path, 'r', encoding='latin-1', errors='replace') as f:
        total_lines = sum(1 for _ in f)

    batch = []
    line_num = 0
    with open(curr_path, 'r', encoding='latin-1', errors='replace') as f:
        for line in f:
            line_num += 1
            fields = line.split('","')
            if len(fields) < 43:
                continue

            reg_orig = fields[0]  # Raw field with quotes, for owner lookup
            reg_clean = reg_orig.strip().strip('"')

            # Registration format: if starts with space -> "CF-xxx", else "C-xxxx"
            if reg_clean and reg_clean[0] == ' ':
                registration = "CF-" + reg_clean.strip()
            else:
                registration = "C-" + reg_clean

            # ICAO from binary field (index 42)
            icao_binary = fields[42].strip().strip('"')
            if len(icao_binary) >= 24:
                icao = bin_to_hex(icao_binary[:24])
            else:
                continue  # Skip if no valid ICAO

            # Aircraft type (index 4)
            ccar_type = _parse_ccar_csv_field(fields[4])

            # Manufacturer (index 7)
            manufacturer = _parse_ccar_csv_field(fields[7])
            manufacturer = title_case(manufacturer)
            manufacturer = manufacturer.replace(" Of ", " of ").replace(" And ", " and ")

            # Serial number (index 5)
            serial = _parse_ccar_csv_field(fields[5])

            # Year built (index 31)
            year_built = fields[31][:4] if len(fields) > 31 else ""

            # Owner lookup (O(1) dict lookup vs VB.NET's linear scan)
            owner_name = owners.get(reg_orig, "")
            owner_name = owner_name.rstrip()
            owner_name = owner_name.replace(" Of ", " of ")

            batch.append((registration, icao, ccar_type, owner_name, manufacturer, serial, year_built))

            if len(batch) >= BATCH_SIZE:
                conn.executemany(
                    "INSERT INTO Master (Registration, ICAO, CCAR_Type, Owner_Name, Manufacturer, Serial, Year) VALUES (?,?,?,?,?,?,?)",
                    batch
                )
                conn.commit()
                batch.clear()

            prog.update(line_num, total_lines)

    if batch:
        conn.executemany(
            "INSERT INTO Master (Registration, ICAO, CCAR_Type, Owner_Name, Manufacturer, Serial, Year) VALUES (?,?,?,?,?,?,?)",
            batch
        )
        conn.commit()

    conn.execute("CREATE INDEX IF NOT EXISTS idx_ccar_icao ON Master(ICAO)")
    conn.commit()
    conn.close()
    prog.done()

    safe_delete(extract_dir)
    safe_delete(zip_path)
    print("  CCAR database creation complete.")
    return True
