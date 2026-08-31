"""
E's VRS Database Updater - Python Edition

Main entry point with command-line interface.

Usage:
    python -m vrs_updater                     # Run with defaults (interactive prompts)
    python -m vrs_updater --work-dir D:\\VRS   # Specify working directory
    python -m vrs_updater --no-faa            # Skip FAA download
    python -m vrs_updater --no-opensky        # Skip OpenSky download
    python -m vrs_updater --complete          # Full rebuild (update all records)
    python -m vrs_updater --help              # Show all options
"""

import argparse
import os
import sys
import time
from datetime import datetime

from .config import Settings, load_settings
from .faa import download_faa, parse_faa
from .ccar import parse_ccar
from .nz_caa import download_nz_caa, parse_nz_caa
from .casa import download_casa, parse_casa
from .opensky import download_opensky, parse_opensky
from .vrs_merge import update_vrs
from . import display_version


def print_banner():
    print("=" * 60)
    print("  E's VRS Database Updater  (Python Edition %s)" % display_version())
    print("=" * 60)
    print()


def main():
    parser = argparse.ArgumentParser(
        description="E's VRS Database Updater - Merges FAA, CCAR, NZ CAA, CASA, and OpenSky data into VRS.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m vrs_updater --work-dir "D:\\VRS Data"
  python -m vrs_updater --no-faa --no-opensky     # Use existing downloads
  python -m vrs_updater --complete                  # Full database rebuild
  python -m vrs_updater --settings settings.sqb     # Load settings from VB.NET app
        """
    )
    parser.add_argument('--work-dir', '-w', default=os.getcwd(),
                        help='Working directory for intermediate files (default: current dir)')
    parser.add_argument('--vrs-dir', '-v', default='',
                        help='VRS install directory containing AircraftOnlineLookupCache.sqb')
    parser.add_argument('--no-faa', action='store_true',
                        help='Skip downloading FAA database (use existing)')
    parser.add_argument('--no-ccar', action='store_true',
                        help='Skip downloading CCAR database (use existing)')
    parser.add_argument('--no-opensky', action='store_true',
                        help='Skip downloading OpenSky database (use existing)')
    parser.add_argument('--no-nz-caa', action='store_true',
                        help='Skip downloading NZ CAA register (use existing)')
    parser.add_argument('--no-casa', action='store_true',
                        help='Skip downloading CASA register (use existing)')
    parser.add_argument('--no-backup', action='store_true',
                        help='Skip backing up VRS database')
    parser.add_argument('--complete', action='store_true',
                        help='Full rebuild: update all records, not just existing ones')
    parser.add_argument('--faa-url', default=None,
                        help='Override FAA download URL')
    parser.add_argument('--opensky-url', default=None,
                        help='Override OpenSky download URL')
    parser.add_argument('--nz-caa-url', default=None,
                        help='Override NZ CAA register download URL')
    parser.add_argument('--casa-url', default=None,
                        help='Override CASA register download URL')
    parser.add_argument('--settings', default=None,
                        help='Load settings from VB.NET settings.sqb file')

    args = parser.parse_args()

    print_banner()

    # Build settings
    if args.settings and os.path.exists(args.settings):
        settings = load_settings(args.settings)
        print(f"  Loaded settings from {args.settings}")
    else:
        settings = Settings()

    # CLI args override loaded settings
    settings.work_dir = os.path.abspath(args.work_dir)
    if args.vrs_dir:
        settings.vrs_dir = os.path.abspath(args.vrs_dir)
    settings.download_faa = not args.no_faa
    settings.download_ccar = not args.no_ccar
    settings.download_opensky = not args.no_opensky
    settings.download_nz_caa = not args.no_nz_caa
    settings.download_casa = not args.no_casa
    settings.backup_vrs_db = not args.no_backup
    settings.build_complete = args.complete
    if args.faa_url:
        settings.faa_url = args.faa_url
    if args.opensky_url:
        settings.opensky_url = args.opensky_url
    if args.nz_caa_url:
        settings.nz_caa_url = args.nz_caa_url
    if args.casa_url:
        settings.casa_url = args.casa_url

    print(f"  Working directory: {settings.work_dir}")
    if settings.vrs_dir:
        print(f"  VRS directory:     {settings.vrs_dir}")
    print(f"  Download FAA:      {'Yes' if settings.download_faa else 'No'}")
    print(f"  Download CCAR:     {'Yes' if settings.download_ccar else 'No'}")
    print(f"  Download OpenSky:  {'Yes' if settings.download_opensky else 'No'}")
    print(f"  Download NZ CAA:   {'Yes' if settings.download_nz_caa else 'No'}")
    print(f"  Download CASA:     {'Yes' if settings.download_casa else 'No'}")
    print(f"  Full rebuild:      {'Yes' if settings.build_complete else 'No'}")
    print()

    # Validate
    if settings.vrs_dir and settings.vrs_dir == settings.work_dir:
        print("ERROR: Working directory and VRS directory cannot be the same.")
        sys.exit(1)

    os.makedirs(settings.work_dir, exist_ok=True)
    start_time = time.time()
    time_started = datetime.now().strftime("%H:%M")
    print(f"  Started at: {time_started}")
    print()

    # ---- Step 1: FAA ----
    if settings.download_faa:
        print("[Step 1/6] FAA Database")
        print("-" * 40)
        if not download_faa(settings):
            print("  WARNING: FAA download failed. Will try to use existing data.")
        if not parse_faa(settings):
            print("  WARNING: FAA parse failed.")
        print()
    else:
        print("[Step 1/6] FAA Database - Skipped (using existing)")
        if not os.path.exists(settings.faa_db_path):
            # Try to parse from existing extracted files
            extract_dir = os.path.join(settings.work_dir, "_faa_extract")
            if os.path.exists(os.path.join(extract_dir, "MASTER.txt")):
                parse_faa(settings)
        print()

    # ---- Step 2: CCAR ----
    print("[Step 2/6] CCAR Database")
    print("-" * 40)
    parse_ccar(settings)
    print()

    # ---- Step 3: NZ CAA ----
    if settings.download_nz_caa:
        print("[Step 3/6] NZ CAA Register")
        print("-" * 40)
        if not download_nz_caa(settings):
            print("  WARNING: NZ CAA download failed. Will try to use existing data.")
        if not parse_nz_caa(settings):
            print("  WARNING: NZ CAA parse failed.")
        print()
    else:
        print("[Step 3/6] NZ CAA Register - Skipped (using existing)")
        csv_path = os.path.join(settings.work_dir, "nz_caa_register.csv")
        if os.path.exists(csv_path) and not os.path.exists(settings.nz_caa_db_path):
            parse_nz_caa(settings)
        print()

    # ---- Step 4: CASA ----
    if settings.download_casa:
        print("[Step 4/6] CASA Register")
        print("-" * 40)
        if not download_casa(settings):
            print("  WARNING: CASA download failed. Will try to use existing data.")
        if not parse_casa(settings):
            print("  WARNING: CASA parse failed.")
        print()
    else:
        print("[Step 4/6] CASA Register - Skipped (using existing)")
        csv_path = os.path.join(settings.work_dir, "casa_register.csv")
        if os.path.exists(csv_path) and not os.path.exists(settings.casa_db_path):
            parse_casa(settings)
        print()

    # ---- Step 5: OpenSky ----
    if settings.download_opensky:
        print("[Step 5/6] OpenSky Database")
        print("-" * 40)
        if not download_opensky(settings):
            print("  WARNING: OpenSky download failed. Will try to use existing data.")
        if not parse_opensky(settings):
            print("  WARNING: OpenSky parse failed.")
        print()
    else:
        print("[Step 5/6] OpenSky Database - Skipped (using existing)")
        csv_path = os.path.join(settings.work_dir, "aircraftDatabase.csv")
        if os.path.exists(csv_path) and not os.path.exists(settings.opensky_db_path):
            parse_opensky(settings)
        print()

    # ---- Step 6: VRS Merge ----
    print("[Step 6/6] VRS Database Merge")
    print("-" * 40)
    update_vrs(settings)
    print()

    # Done
    elapsed = time.time() - start_time
    if elapsed > 60:
        elapsed_str = f"{elapsed / 60:.1f} minutes"
    else:
        elapsed_str = f"{elapsed:.1f} seconds"

    print("=" * 60)
    print(f"  Done! Started: {time_started}  Ended: {datetime.now().strftime('%H:%M')}")
    print(f"  Total time: {elapsed_str}")
    print("=" * 60)


if __name__ == "__main__":
    main()
