"""
Configuration and settings management for VRS Database Updater.
Settings are stored in a settings.sqb SQLite database (compatible with the VB.NET version).
"""

import os
import sys
import sqlite3
from dataclasses import dataclass, field
from pathlib import Path


def _resolve_data_file(work_dir: str, filename: str) -> str:
    """Return path to a data file, checking work_dir first, then the PyInstaller bundle."""
    work_path = os.path.join(work_dir, filename)
    if os.path.exists(work_path):
        return work_path
    # Fall back to bundled copy (PyInstaller _MEIPASS or source tree)
    if getattr(sys, 'frozen', False):
        bundled = os.path.join(sys._MEIPASS, filename)
    else:
        bundled = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), filename)
    if os.path.exists(bundled):
        return bundled
    return work_path  # Return work_dir path even if missing, so caller gets a clear error


# Default download URLs
DEFAULT_FAA_URL = "https://registry.faa.gov/database/ReleasableAircraft.zip"
DEFAULT_OPENSKY_URL = "https://s3.opensky-network.org/data-samples/metadata/aircraftDatabase.zip"
DEFAULT_CCAR_URL = "https://wwwapps.tc.gc.ca/saf-sec-sur/2/ccarcs-riacc/DDZip.aspx"
DEFAULT_NZCAA_URL = "https://www.aviation.govt.nz/assets/aircraft/aircraft-register/Aircraft-Register-for-website-.csv"
DEFAULT_CASA_URL = "https://services.casa.gov.au/CSV/acrftreg.zip"


@dataclass
class Settings:
    """Runtime settings for the updater."""
    work_dir: str = ""                # Working folder (where intermediate DBs go)
    vrs_dir: str = ""                 # VRS install folder (AircraftOnlineLookupCache.sqb)
    faa_url: str = DEFAULT_FAA_URL
    ccar_url: str = DEFAULT_CCAR_URL
    opensky_url: str = DEFAULT_OPENSKY_URL
    nz_caa_url: str = DEFAULT_NZCAA_URL
    casa_url: str = DEFAULT_CASA_URL
    download_faa: bool = True
    download_ccar: bool = True
    download_opensky: bool = True
    download_nz_caa: bool = True
    download_casa: bool = True
    backup_vrs_db: bool = True
    build_complete: bool = False      # If True, update every record even if it exists
    skip_faa: bool = False
    skip_ccar: bool = False
    skip_opensky: bool = False
    skip_nz_caa: bool = False
    skip_casa: bool = False
    skip_rules: bool = False
    faa_max_age_days: int = 30        # 0 = always download
    ccar_max_age_days: int = 30
    opensky_max_age_days: int = 30
    nz_caa_max_age_days: int = 30
    casa_max_age_days: int = 30
    merge_order: list = None  # Default: ["Rules.CSV (when present)", "FAA", "CCAR", "CASA", "NZ CAA", "OpenSky"]

    def __post_init__(self):
        if not self.work_dir:
            self.work_dir = str(Path.cwd())
        if self.merge_order is None:
            self.merge_order = ["Rules.CSV (when present)", "FAA", "CCAR", "CASA", "NZ CAA", "OpenSky"]

    @property
    def faa_db_path(self) -> str:
        return os.path.join(self.work_dir, "FAADatabase.sqb")

    @property
    def ccar_db_path(self) -> str:
        return os.path.join(self.work_dir, "CCARDatabase.sqb")

    @property
    def opensky_db_path(self) -> str:
        return os.path.join(self.work_dir, "OpenSkyDatabase.sqb")

    @property
    def nz_caa_db_path(self) -> str:
        return os.path.join(self.work_dir, "NZCAADatabase.sqb")

    @property
    def casa_db_path(self) -> str:
        return os.path.join(self.work_dir, "CASADatabase.sqb")

    @property
    def vrs_db_path(self) -> str:
        return os.path.join(self.work_dir, "AircraftOnlineLookupCache.sqb")

    @property
    def rules_path(self) -> str:
        return _resolve_data_file(self.work_dir, "Rules.csv")

    @property
    def sils_path(self) -> str:
        return _resolve_data_file(self.work_dir, "Sils.csv")


def load_settings(settings_db_path: str) -> Settings:
    """Load settings from settings.sqb (VB.NET compatible format)."""
    s = Settings()
    if not os.path.exists(settings_db_path):
        return s
    try:
        conn = sqlite3.connect(settings_db_path)
        cur = conn.execute("SELECT * FROM settings LIMIT 1")
        row = cur.fetchone()
        if row:
            cols = [desc[0] for desc in cur.description]
            data = dict(zip(cols, row))
            s.work_dir = data.get("dbPath", s.work_dir)
            s.vrs_dir = data.get("VRSdbPath", s.vrs_dir)
            s.faa_url = data.get("FAA_URL", s.faa_url)
            s.opensky_url = data.get("OpenSky_URL", s.opensky_url)
            s.download_faa = data.get("FAA_db_Download", "Y") == "Y"
            s.download_opensky = data.get("OpenSky_Download", "Y") == "Y"
            s.backup_vrs_db = data.get("BackupVRSdb", "Y") == "Y"
        conn.close()
    except Exception:
        pass
    return s
