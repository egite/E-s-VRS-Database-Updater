"""
Read settings.json and bring it up to date with this version of the program.

settings.json outlives the program that wrote it: a user who upgrades keeps
the file an older version saved, and anyone can edit it by hand. Everything
read from it therefore goes through normalize_settings(), which

- swaps a download link that is a retired default for the current default,
  while leaving links the user typed in themselves alone,
- replaces values of the wrong type or out of range with the default,
- fills in settings that did not exist when the file was written, and
- drops settings this version no longer uses.

Each change that the user would want to know about comes back as a note.

To change a default download link, edit the DEFAULT_*_URL constant in
config.py. Files saved from now on record the defaults they were saved with
("default_urls"), so they are upgraded without further work. Files saved
before that was recorded are recognised by RETIRED_DEFAULT_URLS below.
"""

import json
import os
from datetime import datetime

from .config import (DEFAULT_FAA_URL, DEFAULT_CCAR_URL, DEFAULT_OPENSKY_URL,
                     DEFAULT_NZCAA_URL, DEFAULT_CASA_URL)

# Bumped when the layout of settings.json changes in a way that needs more
# than the per-key checks below.
SETTINGS_VERSION = 1

# key -> (label shown to the user, current default)
DEFAULT_URLS = {
    "faa_url":     ("FAA",     DEFAULT_FAA_URL),
    "ccar_url":    ("CCAR",    DEFAULT_CCAR_URL),
    "opensky_url": ("OpenSky", DEFAULT_OPENSKY_URL),
    "nz_caa_url":  ("NZ CAA",  DEFAULT_NZCAA_URL),
    "casa_url":    ("CASA",    DEFAULT_CASA_URL),
}

# Defaults shipped by versions that did not yet record "default_urls". A
# saved link equal to one of these was never chosen by the user, so it
# follows the default. Nothing needs adding here for later changes.
RETIRED_DEFAULT_URLS = {
    # v2.0.0; the NZ CAA renamed the file in 2026.
    "nz_caa_url": (
        "https://www.aviation.govt.nz/assets/aircraft/aircraft-register/"
        "Aircraft-Register-for-website-.csv",
    ),
}

BOOL_DEFAULTS = {
    "download_faa": True,
    "download_ccar": True,
    "download_opensky": True,
    "download_nz_caa": True,
    "download_casa": True,
    "backup_vrs_db": True,
    "build_complete": False,
    "run_on_start": False,
    "skip_faa": False,
    "skip_ccar": False,
    "skip_opensky": False,
    "skip_nz_caa": False,
    "skip_casa": False,
    "skip_rules": False,
    "log_errors": True,
    "log_activity": False,
}

# Same default as the Settings dataclass, so an upgraded file and a fresh
# install agree. 0 means "always download".
AGE_KEYS = ("faa_max_age_days", "ccar_max_age_days", "opensky_max_age_days",
            "nz_caa_max_age_days", "casa_max_age_days")
AGE_DEFAULT = 30
AGE_MAX = 3650

RULES_ENTRY = "Rules.CSV (when present)"
MERGE_SOURCES = ("FAA", "CCAR", "CASA", "NZ CAA", "OpenSky")
DEFAULT_MERGE_ORDER = [RULES_ENTRY, "FAA", "CCAR", "CASA", "NZ CAA", "OpenSky"]

# Window placement is checked again where it is applied; here it only has to
# be the right type to be kept.
WINDOW_INT_KEYS = ("window_x", "window_y", "window_width", "window_height")


def _is_bool(v) -> bool:
    return isinstance(v, bool) or (type(v) is int and v in (0, 1))


def _is_int(v) -> bool:
    return isinstance(v, int) and not isinstance(v, bool)


def _is_url(v) -> bool:
    return (isinstance(v, str)
            and v.strip().lower().startswith(("http://", "https://"))
            and len(v.strip()) > len("https://"))


def _normalize_merge_order(saved):
    """Known sources once each, in the saved order, with any missing added."""
    if not isinstance(saved, list):
        return list(DEFAULT_MERGE_ORDER), saved is not None

    order, changed = [], False
    for item in saved:
        if isinstance(item, str) and item.startswith("Rules"):
            item = RULES_ENTRY
        if item in order or not (item == RULES_ENTRY or item in MERGE_SOURCES):
            changed = True
            continue
        order.append(item)

    # Sources added in later versions; earlier files simply lack them.
    if RULES_ENTRY not in order:
        order.insert(0, RULES_ENTRY)
    for source in MERGE_SOURCES:
        if source not in order:
            # Ahead of OpenSky, which fills gaps the registers leave.
            pos = order.index("OpenSky") if "OpenSky" in order else len(order)
            order.insert(pos, source)
    return order, changed


def normalize_settings(data):
    """Return (settings, notes): a clean settings dict and what was changed.

    The result holds every setting this version uses, except window placement
    when none was saved. Calling it again on its own output changes nothing.
    """
    notes = []
    if not isinstance(data, dict):
        data = {}
    out = {"settings_version": SETTINGS_VERSION}

    def bad(key, what):
        notes.append(f"settings.json: ignored {what} for {key}; using the default.")

    for key in ("work_dir", "vrs_dir"):
        v = data.get(key, "")
        if not isinstance(v, str):
            bad(key, "a value that is not a folder")
            v = ""
        out[key] = v

    for key, default in BOOL_DEFAULTS.items():
        v = data.get(key, default)
        if not _is_bool(v):
            bad(key, f"{v!r}")
            v = default
        out[key] = bool(v)

    saved_defaults = data.get("default_urls")
    if not isinstance(saved_defaults, dict):
        saved_defaults = {}
    for key, (label, current) in DEFAULT_URLS.items():
        v = data.get(key)
        retired = set(RETIRED_DEFAULT_URLS.get(key, ()))
        was_default = saved_defaults.get(key)
        if isinstance(was_default, str) and was_default != current:
            retired.add(was_default)

        if v is None:
            v = current
        elif not _is_url(v):
            notes.append(f"settings.json: the {label} download link was not a web "
                         f"address; reset to the default.")
            v = current
        elif v.strip() in retired:
            notes.append(f"The {label} download link has changed in this version "
                         f"and has been updated to {current}")
            v = current
        out[key] = v.strip()
    out["default_urls"] = {key: url for key, (_l, url) in DEFAULT_URLS.items()}

    for key in AGE_KEYS:
        v = data.get(key, AGE_DEFAULT)
        if not _is_int(v) or not 0 <= v <= AGE_MAX:
            bad(key, f"{v!r}")
            v = AGE_DEFAULT
        out[key] = v

    order, changed = _normalize_merge_order(data.get("merge_order"))
    if changed:
        notes.append("settings.json: the merge order had unknown or repeated "
                     "entries; they were removed.")
    out["merge_order"] = order

    failures = data.get("download_failures", {})
    out["download_failures"] = (
        {k: v for k, v in failures.items()
         if isinstance(k, str) and _is_int(v) and v > 0}
        if isinstance(failures, dict) else {})
    acked = data.get("download_alert_acked", [])
    out["download_alert_acked"] = (
        sorted({a for a in acked if isinstance(a, str)})
        if isinstance(acked, list) else [])

    for key in WINDOW_INT_KEYS:
        if _is_int(data.get(key)):
            out[key] = data[key]
    if _is_bool(data.get("window_zoomed")):
        out["window_zoomed"] = bool(data["window_zoomed"])
    ratio = data.get("pane_ratio")
    if isinstance(ratio, (int, float)) and not isinstance(ratio, bool):
        out["pane_ratio"] = float(ratio)

    return out, notes


def read_settings_file(path):
    """Read and normalize settings.json.

    Returns (settings, notes); settings is {} when there is no file. A file
    that is not valid JSON is renamed aside (so the next save does not
    destroy what the user may want to recover) and treated as absent.
    """
    if not os.path.exists(path):
        return {}, []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("not a JSON object")
    except OSError:
        return {}, []
    except ValueError as e:   # json.JSONDecodeError is a ValueError
        aside = f"{path}.bad-{datetime.now().strftime('%y%m%d-%H%M%S')}"
        try:
            os.replace(path, aside)
            where = f"it was renamed to {os.path.basename(aside)}"
        except OSError:
            where = "it could not be renamed and will be overwritten on exit"
        return {}, [f"settings.json could not be read ({e}); {where}. "
                    f"Starting from the default settings."]
    return normalize_settings(data)
