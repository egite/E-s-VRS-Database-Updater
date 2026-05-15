"""
Utility functions for text cleaning, downloading, and progress reporting.
Replaces the VB.NET helper functions with efficient Python equivalents.
"""

import os
import re
import sys
import time
import zipfile
import shutil
from pathlib import Path
from urllib.request import urlretrieve
from urllib.error import URLError

# Try to use requests if available (better progress + error handling)
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False


# ---------------------------------------------------------------------------
# Text cleaning  (replaces Remove_All_Caps, Remove_Trailing_Spaces, etc.)
# ---------------------------------------------------------------------------

# Pre-compiled regex patterns for performance (called 800K+ times)
_RE_LLC = re.compile(r'\bLlc\b')
_RE_LLP = re.compile(r'\bLlp\b')


def title_case(s: str) -> str:
    """Convert ALL-CAPS text to Title Case, preserving common abbreviations.

    This replaces the VB.NET Remove_All_Caps function but uses Python's
    built-in str.title() with post-processing for edge cases.
    """
    if not s or len(s) <= 1:
        return s
    result = s.lower().title()
    # Fix common abbreviations that should stay uppercase
    for abbr in ("Llc", "Llp", "Ii", "Iii", "Iv", "Rv7a"):
        result = result.replace(f" {abbr} ", f" {abbr.upper()} ")
        if result.endswith(f" {abbr}"):
            result = result[: -len(abbr)] + abbr.upper()
    # Fix common words that should stay lowercase
    for word in (" Of ", " And ", " The ", " De ", " La "):
        result = result.replace(word, word.lower())
    # LLC/LLP at end
    result = _RE_LLC.sub('LLC', result)
    result = _RE_LLP.sub('LLP', result)
    return result


def clean_operator_name(name: str, registration: str = "",
                        detail_callback=None) -> str:
    """Clean and normalize operator/owner name.

    Handles the various VB.NET operator-name fixups:
    - Fix LLC/LLP casing (including "SomethingLlc" without space)
    - Fix "Of" -> "of"
    - Fix "Mc" prefix (McDonald -> McDonald, not Mcdonald)
    - Fix "X & X" and "X and X" patterns
    - Uppercase registration number if found within the operator name
    - Uppercase registration-without-N if found within the operator name

    If detail_callback is provided, it's called with detail messages
    matching the VB.NET TextBox4 output.
    """
    if not name:
        return name

    # Fix LLC at end without space (e.g. "SomethingLlc" -> "SomethingLLC")
    if name.upper().endswith("LLC"):
        name = name[:-3] + "LLC"
    name = name.replace(" Llc", " LLC")
    name = name.replace(" Llp", " LLP")
    name = name.replace(" Of ", " of ")
    name = name.replace(" Rv7a ", " RV7A ")

    # If operator starts with the registration, uppercase it
    if registration and name.lower().startswith(registration.lower()):
        name = registration.upper() + name[len(registration):]

    # Fix full registration embedded in operator name (case-insensitive)
    registration_in_operator = False
    if registration and len(name) >= len(registration):
        reg_upper = registration.upper()
        name_upper = name.upper()
        # Search for the registration anywhere (not just at start)
        start = 0
        if name_upper.startswith(reg_upper):
            start = len(reg_upper)  # Already handled above, skip
        idx = name_upper.find(reg_upper, start)
        if idx >= 0:
            name = name[:idx] + reg_upper + name[idx + len(reg_upper):]
            registration_in_operator = True
            if detail_callback:
                detail_callback(f"Operator has registration:  {name}.")

    # Fix registration-without-N embedded in operator name (VRS.vb lines 221-232)
    if registration and not registration_in_operator and len(registration) > 1:
        reg_no_n = registration[1:]  # Strip leading "N"
        if len(name) >= len(reg_no_n):
            name_upper = name.upper()
            reg_no_n_upper = reg_no_n.upper()
            idx = name_upper.find(reg_no_n_upper)
            if idx >= 0:
                name = name[:idx] + reg_no_n_upper + name[idx + len(reg_no_n):]
                if detail_callback:
                    detail_callback(f"Operator has registration w/o N:  {name}.")

    # Fix "X & x" pattern: capitalize second letter (VRS.vb lines 233-240)
    if len(name) >= 6 and name[1:4] == " & " and name[4] == " " and name[5] == " ":
        name = name[:2] + name[2].upper() + name[3:]
        if detail_callback:
            detail_callback(f"Operator format X&x:  {name}.")
    # Fix "X&x" pattern (VRS.vb lines 241-248)
    elif len(name) >= 3 and name[1] == "&" and len(name) > 3 and name[3] == " ":
        name = name[:2] + name[2].upper() + name[3:]
        if detail_callback:
            detail_callback(f"Operator format X&x:  {name}.")

    # Fix "x and x" / "x And x" pattern: capitalize after "and" (VRS.vb lines 249-256)
    if (len(name) >= 8 and name[1] == " " and
            name[2:6].lower() == "and " and name[7] == " "):
        name = name[:5] + name[5].upper() + name[6:]
        if detail_callback:
            detail_callback(f"Operator format X and x:  {name}.")

    # Fix "Mc" prefix -> "McX" (e.g., "Mcdonald" -> "McDonald")
    if len(name) > 3 and name[:2] == "Mc" and name[2] != " " and name[2].islower():
        name = "Mc" + name[2].upper() + name[3:]
        if detail_callback:
            detail_callback(f"Operator format Mcxx...{name}.")

    return name


def remove_non_ascii(s: str) -> str:
    """Remove non-ASCII characters, replace & with 'and'.

    Replaces RemoveNonASCII from VB.NET.
    """
    result = []
    for ch in s:
        code = ord(ch)
        if ch == '&':
            result.append('and')
        elif 31 < code < 127:
            result.append(ch)
    return ''.join(result)


def sql_escape(s: str) -> str:
    """Escape single quotes for SQLite string literals."""
    if s is None:
        return ""
    return s.replace("'", "''")


# ---------------------------------------------------------------------------
# Download helpers
# ---------------------------------------------------------------------------

class ProgressReporter:
    """Simple console progress reporter."""

    def __init__(self, task_name: str = ""):
        self.task_name = task_name
        self._start_time = time.time()
        self._last_pct = -1

    def _write(self, text: str):
        """Write to stdout if available (pythonw sets it to None)."""
        if sys.stdout is not None:
            sys.stdout.write(text)
            sys.stdout.flush()

    def update(self, current: int, total: int):
        if total <= 0:
            return
        pct = int(100 * current / total)
        if pct == self._last_pct:
            return
        self._last_pct = pct
        elapsed = time.time() - self._start_time
        if pct > 0:
            eta = elapsed / (pct / 100) * (1 - pct / 100)
            if eta > 60:
                eta_str = f"{eta / 60:.1f} min"
            else:
                eta_str = f"{eta:.0f} sec"
        else:
            eta_str = "..."
        bar = "#" * (pct // 2) + "-" * (50 - pct // 2)
        self._write(f"\r  {self.task_name} [{bar}] {pct}%  ETA: {eta_str}   ")
        if pct >= 100:
            self._write("\n")

    def done(self):
        elapsed = time.time() - self._start_time
        if elapsed > 60:
            t_str = f"{elapsed / 60:.1f} min"
        else:
            t_str = f"{elapsed:.1f} sec"
        self._write(f"\r  {self.task_name} ... Done ({t_str}).\n")


def download_file(url: str, dest: str, label: str = "Downloading") -> bool:
    """Download a file with progress indication.

    Uses `requests` if available, falls back to urllib.
    Returns True on success.
    """
    print(f"  {label}: {url}")
    # Remove existing file
    if os.path.exists(dest):
        os.remove(dest)

    if HAS_REQUESTS:
        try:
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) "
                              "Chrome/131.0.0.0 Safari/537.36"
            }
            resp = requests.get(url, stream=True, timeout=120, headers=headers)
            resp.raise_for_status()
            total = int(resp.headers.get("content-length", 0))
            downloaded = 0
            prog = ProgressReporter(label)
            with open(dest, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total:
                        prog.update(downloaded, total)
            prog.done()
            return True
        except Exception as e:
            print(f"  ERROR downloading: {e}")
            return False
    else:
        try:
            def _reporthook(block_num, block_size, total_size):
                if total_size > 0 and sys.stdout is not None:
                    downloaded = block_num * block_size
                    pct = min(100, int(100 * downloaded / total_size))
                    sys.stdout.write(f"\r  {label} ... {pct}%  ")
                    sys.stdout.flush()
            urlretrieve(url, dest, reporthook=_reporthook)
            print(f"\r  {label} ... Done.")
            return True
        except Exception as e:
            print(f"  ERROR downloading: {e}")
            return False


def extract_zip(zip_path: str, dest_dir: str, label: str = "Extracting"):
    """Extract a ZIP file to dest_dir, overwriting existing files."""
    print(f"  {label}...")
    if os.path.exists(dest_dir):
        shutil.rmtree(dest_dir, ignore_errors=True)
    os.makedirs(dest_dir, exist_ok=True)
    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(dest_dir)
    print(f"  {label}...Done.")


def safe_delete(path: str):
    """Delete a file or directory if it exists, ignoring errors."""
    try:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.exists(path):
            os.remove(path)
    except OSError:
        pass
