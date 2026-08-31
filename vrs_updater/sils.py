"""
Sils.csv load / save / validation.

Shared by the silhouette lookup and the GUI Sils editor so both agree on the
on-disk format.

File layout (4 columns):

    0   FAA Manufacturer  - comma-separated list of manufacturer aliases, or "*"
    1   FAA Model         - comma-separated list of model aliases, or "*"
    2   Remap             - type code to use instead of column 3, when set
    3   Type              - ICAO type designator

Lookup is a linear scan that returns the FIRST row whose manufacturer and
model both match, so row order is significant. A row with an empty
manufacturer or model can never match - the scan skips it - and the file
ships with thousands of such rows.
"""

import csv
import os
import shutil
from typing import Dict, List, Optional, Tuple

HEADER = ['FAA Manufacturer', 'FAA Model', 'Remap', 'Type']
TOTAL_COLS = 4


class SilEntry:
    """One row of Sils.csv, with the comma-separated cells split into lists."""
    __slots__ = ('manufacturers', 'models', 'remap', 'type_code')

    def __init__(self, manufacturers=None, models=None, remap="", type_code=""):
        self.manufacturers: List[str] = list(manufacturers or [])
        self.models: List[str] = list(models or [])
        self.remap = remap
        self.type_code = type_code

    def copy(self) -> 'SilEntry':
        return SilEntry(self.manufacturers, self.models, self.remap, self.type_code)

    @property
    def usable(self) -> bool:
        """A row with no manufacturer or no model is skipped by the lookup."""
        return bool(self.manufacturers) and bool(self.models)

    @property
    def resolved(self) -> str:
        """The code the lookup returns: Remap when set, otherwise Type."""
        return self.remap or self.type_code

    def describe_manufacturers(self) -> str:
        return ", ".join(self.manufacturers)

    def describe_models(self) -> str:
        return ", ".join(self.models)


def split_aliases(cell: str) -> List[str]:
    """Split a comma-separated alias cell into a clean list."""
    return [part.strip() for part in (cell or "").split(",") if part.strip()]


def join_aliases(values: List[str]) -> str:
    """Join aliases the way the file does - comma separated, no spaces."""
    return ",".join(v.strip() for v in values if v.strip())


# ---------------------------------------------------------------------------
# Load / save
# ---------------------------------------------------------------------------

def load_sils(sils_path: str, quiet: bool = False) -> List[SilEntry]:
    """Load every row of Sils.csv, in file order, including unusable ones.

    Unusable rows are kept so that saving preserves them where they sit.
    """
    entries: List[SilEntry] = []
    if not os.path.exists(sils_path):
        return entries

    with open(sils_path, 'r', encoding='utf-8-sig', errors='replace', newline='') as f:
        for line_num, row in enumerate(csv.reader(f)):
            if line_num == 0 and row and row[0].strip().lower().startswith('faa manufacturer'):
                continue  # header
            if not any(v.strip() for v in row):
                continue  # blank line
            if len(row) < TOTAL_COLS:
                row = row + [''] * (TOTAL_COLS - len(row))

            entries.append(SilEntry(split_aliases(row[0]), split_aliases(row[1]),
                                    row[2].strip(), row[3].strip()))

    if not quiet:
        usable = sum(1 for e in entries if e.usable)
        print("  Loaded %d silhouette rows from Sils.csv (%d usable)."
              % (len(entries), usable))
    return entries


def entry_to_row(entry: SilEntry) -> List[str]:
    """Serialize one entry to its 4-column CSV row."""
    return [join_aliases(entry.manufacturers), join_aliases(entry.models),
            _clean(entry.remap), _clean(entry.type_code)]


def _clean(val: str) -> str:
    return (val or '').replace('\r', ' ').replace('\n', ' ').strip()


def save_sils(sils_path: str, entries: List[SilEntry], backup: bool = True) -> None:
    """Write Sils.csv atomically, keeping a .bak of the previous file.

    The BOM and CRLF line endings of the original are preserved so the file
    stays interchangeable with the VB.NET version and opens cleanly in Excel.
    """
    directory = os.path.dirname(os.path.abspath(sils_path))
    os.makedirs(directory, exist_ok=True)
    tmp_path = sils_path + '.tmp'

    with open(tmp_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f, lineterminator='\r\n')
        writer.writerow(HEADER)
        for entry in entries:
            writer.writerow(entry_to_row(entry))

    if backup and os.path.exists(sils_path):
        try:
            shutil.copy2(sils_path, sils_path + '.bak')
        except Exception:
            pass

    os.replace(tmp_path, sils_path)


# ---------------------------------------------------------------------------
# Matching - mirrors SilhouetteLookup._lookup_in_data
# ---------------------------------------------------------------------------

def entry_matches(entry: SilEntry, manufacturer: str, model: str) -> bool:
    """True if this row would match the given manufacturer / model pair.

    Matching is case-insensitive, and "*" in either column is a wildcard.
    """
    if not entry.usable:
        return False
    mfr_lower = (manufacturer or "").lower()
    model_lower = (model or "").lower()

    if not any(m == "*" or m.lower() == mfr_lower for m in entry.manufacturers):
        return False
    return any(m == "*" or m.lower() == model_lower for m in entry.models)


def find_match(entries: List[SilEntry], manufacturer: str,
               model: str) -> Optional[int]:
    """Index of the first row matching the pair, or None. First match wins."""
    for i, entry in enumerate(entries):
        if entry_matches(entry, manufacturer, model):
            return i
    return None


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_entry(entry: SilEntry) -> List[Tuple[str, str]]:
    """Check one row. Returns a list of (severity, message)."""
    issues: List[Tuple[str, str]] = []

    if not entry.manufacturers:
        issues.append(("error", "No manufacturer - the lookup skips this row."))
    if not entry.models:
        issues.append(("error", "No model - the lookup skips this row."))
    if not entry.resolved:
        issues.append(("warning",
                       "No Type and no Remap - this row stops the built-in mapping "
                       "from being assigned to these aircraft (it does not erase a "
                       "code already stored in VRS). Set a code if that is not what "
                       "you meant."))

    if entry.manufacturers == ["*"] and entry.models == ["*"]:
        issues.append(("warning",
                       "Wildcard on both columns - this row matches every "
                       "aircraft and hides everything below it."))

    for label, values in (("manufacturer", entry.manufacturers),
                          ("model", entry.models)):
        seen = set()
        for v in values:
            key = v.lower()
            if key in seen:
                issues.append(("warning", "Duplicate %s alias '%s'." % (label, v)))
            seen.add(key)

    for alias in entry.manufacturers:
        if "," in alias:
            issues.append(("error",
                           "Manufacturer alias '%s' contains a comma - commas "
                           "separate aliases and cannot appear inside one." % alias))
    for alias in entry.models:
        if "," in alias:
            issues.append(("error",
                           "Model alias '%s' contains a comma - commas separate "
                           "aliases and cannot appear inside one." % alias))

    code = entry.resolved
    if code and (len(code) > 4 or " " in code):
        issues.append(("warning",
                       "'%s' is unusual for an ICAO type designator "
                       "(expected up to 4 characters, no spaces)." % code))

    return issues


def find_shadowed(entries: List[SilEntry]) -> Dict[int, int]:
    """Find rows that can never be reached because an earlier row wins.

    The lookup returns the first matching row, so a row is unreachable when
    every manufacturer/model pair it covers is already claimed above it.
    Returns {shadowed_index: shadowing_index}.
    """
    shadowed: Dict[int, int] = {}
    claimed: Dict[Tuple[str, str], int] = {}
    wildcard_mfr: Dict[str, int] = {}   # model -> row that matches it for any mfr
    wildcard_model: Dict[str, int] = {}  # manufacturer -> row matching any model

    for i, entry in enumerate(entries):
        if not entry.usable:
            continue

        pairs = [(m.lower(), d.lower())
                 for m in entry.manufacturers for d in entry.models]

        blockers = set()
        for mfr, model in pairs:
            if (mfr, model) in claimed:
                blockers.add(claimed[(mfr, model)])
            elif model in wildcard_mfr:
                blockers.add(wildcard_mfr[model])
            elif mfr in wildcard_model:
                blockers.add(wildcard_model[mfr])
            else:
                blockers = set()   # at least one pair is still unclaimed
                break

        if blockers:
            shadowed[i] = min(blockers)

        for mfr, model in pairs:
            claimed.setdefault((mfr, model), i)
            if mfr == "*":
                wildcard_mfr.setdefault(model, i)
            if model == "*":
                wildcard_model.setdefault(mfr, i)

    return shadowed


def validate_all(entries: List[SilEntry],
                 skip_unusable: bool = True) -> List[Tuple[int, str, str]]:
    """Validate every row plus cross-row shadowing.

    Unusable rows are skipped by default - the shipped file contains
    thousands of them and reporting each one buries the real problems.
    Returns a list of (index, severity, message).
    """
    results: List[Tuple[int, str, str]] = []
    for i, entry in enumerate(entries):
        if skip_unusable and not entry.usable:
            continue
        for severity, msg in validate_entry(entry):
            results.append((i, severity, msg))
    for j, i in sorted(find_shadowed(entries).items()):
        results.append((j, "warning",
                        "Unreachable - row %d already matches everything this "
                        "row covers." % (i + 1)))
    return results
