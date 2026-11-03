"""
Checks for VRS configuration mistakes that stop VRS reading the lookup cache,
or stop this program's updates reaching it.

VRS keeps Configuration.xml and PluginsConfiguration.txt in the same folder as
AircraftOnlineLookupCache.sqb (%LOCALAPPDATA%\\VirtualRadar on Windows,
~/.local/share/VirtualRadar under Mono), so everything is found from the VRS
Folder setting.
"""

import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import List

from .config import Settings

CACHE_NAME = "AircraftOnlineLookupCache.sqb"


@dataclass
class ConfigProblem:
    title: str      # One-line name of the problem
    detail: str     # What is wrong and what it breaks
    fix: str        # How to fix it

    def as_text(self) -> str:
        return f"{self.title}\n{self.detail}\nHow to fix: {self.fix}"


def _base_name(path: str) -> str:
    """File name from a path written on either Windows or Linux."""
    return re.split(r"[\\/]", path.strip())[-1]


def _read_basestation_settings(vrs_dir: str):
    """Return the <BaseStationSettings> element of Configuration.xml, or None."""
    path = os.path.join(vrs_dir, "Configuration.xml")
    try:
        root = ET.parse(path).getroot()
    except (OSError, ET.ParseError):
        return None
    return root.find("BaseStationSettings")


def check_vrs_config(settings: Settings) -> List[ConfigProblem]:
    """Return every problem found; an empty list means nothing to report."""
    vrs_dir = settings.vrs_dir
    if not vrs_dir:
        # Updating a copy in the working folder only, by choice
        return []

    problems: List[ConfigProblem] = []

    if not os.path.isdir(vrs_dir):
        problems.append(ConfigProblem(
            "VRS Folder does not exist",
            f"{vrs_dir} was not found, so the updated database cannot be "
            "copied back where VRS reads it.",
            "Browse to the folder that holds AircraftOnlineLookupCache.sqb. "
            "On Windows that is %LOCALAPPDATA%\\VirtualRadar; under Mono it is "
            "~/.local/share/VirtualRadar in the home folder of the user who runs VRS."))
        return problems

    cache_path = os.path.join(vrs_dir, CACHE_NAME)
    has_config = os.path.isfile(os.path.join(vrs_dir, "Configuration.xml"))
    has_cache = os.path.isfile(cache_path)

    if not has_config and not has_cache:
        problems.append(ConfigProblem(
            "VRS Folder is not a VRS data folder",
            f"{vrs_dir} has neither Configuration.xml nor {CACHE_NAME}, so VRS "
            "does not use this folder. Anything written here never reaches VRS. "
            "This usually means the program is running as a different user than "
            "VRS (for example with sudo), or the VRS install folder was picked "
            "instead of its data folder.",
            "Run this program as the same user that runs VRS, and set VRS Folder "
            "to that user's VirtualRadar data folder (where Configuration.xml is)."))
        return problems

    # Can this user write the cache back?
    target = cache_path if has_cache else vrs_dir
    if not os.access(target, os.W_OK):
        problems.append(ConfigProblem(
            "No permission to write to the VRS Folder",
            f"This user cannot write to {target}, so the updated database "
            "cannot be copied back for VRS to read.",
            "Run this program as the same user that runs VRS, or give this "
            "user write permission on that file and folder."))

    bs = _read_basestation_settings(vrs_dir)
    if bs is not None:
        db_file = (bs.findtext("DatabaseFileName") or "").strip()
        if db_file and _base_name(db_file).lower() == CACHE_NAME.lower():
            problems.append(ConfigProblem(
                "VRS is set to use the lookup cache as its BaseStation database",
                f"VRS's BaseStation database file is {db_file}. That file is "
                "the online lookup cache, not a BaseStation database, so VRS "
                "fails to open it (\"no such table DBHistory\", \"disk I/O "
                "error\" or a crash at start-up) and shows no aircraft details.",
                "In VRS, go to Tools > Options > Data Sources and clear the "
                "BaseStation database file name, or point it at a real "
                "BaseStation database such as BaseStation.sqb. Leave "
                f"{CACHE_NAME} where it is - VRS finds it on its own."))

        lookup_online = (bs.findtext("LookupAircraftDetailsOnline") or "").strip().lower()
        if lookup_online == "false" and not settings.build_complete:
            problems.append(ConfigProblem(
                "VRS online lookup is turned off",
                "VRS only adds aircraft to the lookup cache through its online "
                "lookup. With it off, aircraft VRS has not cached before are never "
                "added, and this program only refreshes aircraft already in the "
                "cache, so new aircraft never get details.",
                "In VRS, go to Tools > Options > Data Sources and tick "
                "\"Lookup aircraft details online\", or tick \"Build complete "
                "database\" here so this program adds every aircraft itself."))

    plugins_path = os.path.join(vrs_dir, "PluginsConfiguration.txt")
    try:
        with open(plugins_path, encoding="utf-8", errors="replace") as f:
            plugins_text = f.read()
    except OSError:
        plugins_text = ""
    if CACHE_NAME.lower() in plugins_text.lower():
        problems.append(ConfigProblem(
            "A VRS plugin is set to use the lookup cache as its database",
            "PluginsConfiguration.txt names the lookup cache as a plugin's "
            "database - usually the Database Writer. The plugin expects a "
            "BaseStation database, so it fails on the cache.",
            "In VRS, go to Tools > Plugins > Database Writer > Options and set "
            "\"Database filename\" to a separate BaseStation database (Create "
            "Database makes one), or untick Enabled if you do not need it."))

    return problems


def format_problems(problems: List[ConfigProblem]) -> str:
    """All problems as numbered paragraphs, for a dialog or the console."""
    return "\n\n".join(f"{i}. {p.as_text()}" for i, p in enumerate(problems, 1))
