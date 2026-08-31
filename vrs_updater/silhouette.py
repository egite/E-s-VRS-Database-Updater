"""
Aircraft silhouette / ICAO type code mapping.

Maps manufacturer + model combinations to ICAO type designators used for
silhouette display in VRS.

This replaces the VB.NET Silhouette.vb and Sils.vb modules.
The lookup data comes from Sils.csv (preferred) or falls back to the
hardcoded arrays in Sils.vb.
"""

import os
from typing import Optional, Tuple, List

from .sils import SilEntry, load_sils, join_aliases


class SilhouetteLookup:
    """Loads Sils.csv data and provides type-code lookups."""

    def __init__(self):
        self.manufacturers: List[str] = []   # FAA_Manufacturer_Data
        self.models: List[str] = []          # FAA_Model_Data
        self.remaps: List[str] = []          # Remap_Data
        self.types: List[str] = []           # Type_Data
        self.entries: List[SilEntry] = []    # the Sils.csv rows, in file order
        self.builtin_start: int = 0          # index where built-in rows begin
        self.last_match_index: int = -1      # row that satisfied the last lookup
        self._cache: dict = {}               # (model, mfr, emitter) -> result

    def load_csv(self, csv_path: str) -> bool:
        """Load silhouette data: Sils.csv layered over the built-in table.

        Sils.csv is an overlay, not a replacement. Its rows are scanned first,
        so a row in the file overrides the built-in mapping for the
        manufacturer/model pairs it covers, and anything the file does not
        cover still resolves from the built-in table. A file row with no Type
        and no Remap suppresses the built-in mapping instead of replacing it.

        Only usable built-in rows are appended - a row with no manufacturer or
        no model can never match, so carrying thousands of them would just
        make every lookup slower.
        """
        self.manufacturers.clear()
        self.models.clear()
        self.remaps.clear()
        self.types.clear()
        self.entries = []
        self._cache = {}

        # ---- overlay: Sils.csv, scanned first so it wins -----------------
        found_csv = os.path.exists(csv_path)
        if found_csv:
            # Parsed by sils.py so the Sils editor and this lookup cannot
            # disagree about the file format.
            self.entries = load_sils(csv_path, quiet=True)

        builtin = self._layer(self.entries)
        if found_csv:
            usable = sum(1 for e in self.entries if e.usable)
            print(f"  Loaded {usable} silhouette mappings from Sils.csv, "
                  f"layered over {builtin} built-in mappings.")
            return True

        print(f"  Sils.csv not found; using {builtin} built-in silhouette mappings.")
        return False

    def load_entries(self, entries) -> None:
        """Load an in-memory overlay (used by the Sils editor's Test Lookup)."""
        self.manufacturers.clear()
        self.models.clear()
        self.remaps.clear()
        self.types.clear()
        self.entries = list(entries)
        self._cache = {}
        self._layer(self.entries)

    def _layer(self, entries) -> int:
        """Stack `entries` over the built-in table. Returns the built-in count.

        Only usable built-in rows are appended - a row with no manufacturer or
        no model can never match, so carrying thousands of them would just make
        every lookup slower.
        """
        def add(mfr, mdl, rmp, typ):
            self.manufacturers.append(mfr)
            self.models.append(mdl)
            self.remaps.append(rmp)
            self.types.append(typ)

        for entry in entries:
            add(join_aliases(entry.manufacturers), join_aliases(entry.models),
                entry.remap, entry.type_code)

        self.builtin_start = len(self.types)

        from .sils_data import SILS_DATA
        for mfr, mdl, rmp, typ in SILS_DATA:
            if mfr.strip() and mdl.strip():
                add(mfr, mdl, rmp, typ)

        return len(self.types) - self.builtin_start

    def determine_silhouette(self, icao_type_input: str, manufacturer: str,
                              emitter_type: str = "") -> Tuple[str, bool]:
        """Cached wrapper around the resolver.

        The table scan is linear and the merge calls this once or twice per
        aircraft, but a registry has far fewer distinct manufacturer/model
        pairs than records, so memoizing collapses the repeats.
        """
        key = (icao_type_input, manufacturer, emitter_type)
        hit = self._cache.get(key)
        if hit is not None:
            self.last_match_index = hit[1]
            return hit[0]
        result = self._determine_silhouette(icao_type_input, manufacturer,
                                            emitter_type)
        self._cache[key] = (result, self.last_match_index)
        return result

    def _determine_silhouette(self, icao_type_input: str, manufacturer: str,
                              emitter_type: str = "") -> Tuple[str, bool]:
        """Determine the ICAO type code for silhouette display.

        This is a direct port of the VB.NET Determine_Silhouette function,
        preserving all the manufacturer-specific normalization logic.

        Args:
            icao_type_input: The raw ICAO type code or model string
            manufacturer: The aircraft manufacturer name
            emitter_type: ADS-B emitter type (for surface vehicles, etc.)

        Returns:
            Tuple of (resolved_type_code, found) where found indicates a match.
        """
        if not icao_type_input:
            # Handle emitter-type-based fallback (surface vehicles)
            if emitter_type == "Surface vehicle - emergency":
                return "FIRE", True
            elif emitter_type == "Surface - service":
                return "TUG", True
            return "", False

        if icao_type_input in ("B--B", "----"):
            return "", False

        icao_text_orig = icao_type_input
        icao_text = icao_type_input.replace("-", "").replace(" ", "")
        found = False
        mfr = manufacturer.replace(",", "") if manufacturer else ""
        mfr_lower = mfr.lower()

        first2 = icao_text[:2]
        first3 = icao_text[:3]
        first4 = icao_text[:4]
        first5 = icao_text[:5]

        # ---- Manufacturer-specific normalization ----
        # VB.NET string comparison is case-insensitive by default,
        # so all manufacturer checks use mfr_lower.

        if mfr_lower in ("airbus industrie", "airbus sas") or mfr_lower.startswith("airbus"):
            if first4 in ("A330", "A350"):
                icao_text_orig = first3 + (first5[4] if len(first5) > 4 else "")
            elif first4 in ("A319", "A318", "A320", "A321"):
                icao_text_orig = first4

        elif mfr_lower.startswith("boeing"):
            if icao_text == "78710":
                icao_text = "B789"
                found = True
            elif icao_text_orig == "777F" or first4 == "777F":
                icao_text = "B77L"
                found = True
            elif first2 == "MD":
                mfr = "Mcdonnell Douglas"
                mfr_lower = mfr.lower()
            elif icao_text == "E3TF":
                found = True
            elif icao_text == "E6":
                found = True
            elif icao_text in ("A75N1(PT17)", "A75"):
                icao_text = "ST75"
                found = True
            elif icao_text in ("K35R", "C135"):
                found = True
            elif icao_text[0] != "B":
                icao_text = "B" + first2 + (first4[3] if len(first4) > 3 else "")

        elif mfr_lower.startswith("cessna") or mfr_lower in ("textron aviation inc", "textron aviation inc."):
            mfr = "Cessna"
            mfr_lower = "cessna"
            if first4 == "T182":
                icao_text = "C182"; found = True
            elif first4 == "T206":
                icao_text = "C210"; found = True
            elif first4 == "T210":
                icao_text = "C210"; found = True
            elif icao_text == "TU206A":
                icao_text = "C206"; found = True
            elif icao_text == "530":
                icao_text = "E530"; found = False
            elif icao_text == "P172D":
                icao_text = "C172"; found = True
            elif icao_text == "G36":
                found = False
            else:
                icao_text = "C" + first3

        elif mfr_lower == "eiriavion oy":
            icao_text = "AS25"

        elif mfr_lower == "glasflugel":
            icao_text = "AS25"; found = True

        elif mfr_lower.startswith("pilatus") or mfr_lower == "pilatus aircraft ltd":
            mfr = "Pilatus"
            mfr_lower = "pilatus"
            icao_text_orig = icao_text_orig.split("/")[0]

        elif mfr_lower == "piper":
            icao_text = first4

        elif mfr_lower == "mooney aircraft corp.":
            icao_text = icao_text.replace("-", "").split()[0]

        elif icao_text_orig == "BALL":
            icao_text = "Ball"; found = True

        elif not mfr:
            if icao_text == "BE3D":
                icao_text = "BE33"; found = True

        # ---- CSV/array lookup ----
        if not found:
            result = self._lookup_in_data(icao_text_orig, icao_text, mfr)
            if result:
                icao_text = result
                found = True

        # Boeing always counts as found
        if mfr_lower.startswith("boeing"):
            found = True

        return icao_text, found

    def _lookup_in_data(self, orig: str, cleaned: str, manufacturer: str) -> Optional[str]:
        """Search the Sils.csv data for a matching type code.

        Linear scan matching Silhouette.vb lines 119-144 exactly:
        iterate all entries, check manufacturer match (or wildcard *),
        then check model match (or wildcard *).

        VB.NET string comparison is case-insensitive by default,
        so we use .lower() for manufacturer and model matching.
        """
        mfr_lower = manufacturer.lower()
        orig_lower = orig.lower()
        self.last_match_index = -1
        for i in range(len(self.types)):
            mfr_entry = self.manufacturers[i]
            if mfr_entry:
                for mfr_part in mfr_entry.split(","):
                    mfr_part = mfr_part.strip()
                    if mfr_lower == mfr_part.lower() or mfr_part == "*":
                        model_entry = self.models[i]
                        if model_entry:
                            for model_part in model_entry.split(","):
                                model_part = model_part.strip()
                                if orig_lower == model_part.lower() or model_part == "*":
                                    result = self.remaps[i] if self.remaps[i] else self.types[i]
                                    self.last_match_index = i
                                    return result
        return None


# Module-level convenience for the shared lookup instance
_lookup = SilhouetteLookup()


def load_silhouettes(sils_csv_path: str) -> SilhouetteLookup:
    """Load the silhouette lookup data. Returns the lookup instance."""
    _lookup.load_csv(sils_csv_path)
    return _lookup


def determine_silhouette(model: str, manufacturer: str,
                          emitter_type: str = "") -> Tuple[str, bool]:
    """Module-level convenience for silhouette lookup."""
    return _lookup.determine_silhouette(model, manufacturer, emitter_type)
