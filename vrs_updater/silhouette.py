"""
Aircraft silhouette / ICAO type code mapping.

Maps manufacturer + model combinations to ICAO type designators used for
silhouette display in VRS.

This replaces the VB.NET Silhouette.vb and Sils.vb modules.
The lookup data comes from Sils.csv (preferred) or falls back to the
hardcoded arrays in Sils.vb.
"""

import csv
import os
from typing import Optional, Tuple, List


class SilhouetteLookup:
    """Loads Sils.csv data and provides type-code lookups."""

    def __init__(self):
        self.manufacturers: List[str] = []   # FAA_Manufacturer_Data
        self.models: List[str] = []          # FAA_Model_Data
        self.remaps: List[str] = []          # Remap_Data
        self.types: List[str] = []           # Type_Data

    def load_csv(self, csv_path: str) -> bool:
        """Load silhouette data from Sils.csv, falling back to built-in data.

        Mirrors VRS.vb behaviour: use Sils.csv if present, otherwise use the
        hardcoded mapping from Sils.vb (shipped as sils_data.py).
        """
        self.manufacturers.clear()
        self.models.clear()
        self.remaps.clear()
        self.types.clear()

        if os.path.exists(csv_path):
            with open(csv_path, 'r', encoding='utf-8', errors='replace') as f:
                reader = csv.reader(f)
                for row in reader:
                    if len(row) < 4:
                        continue
                    self.manufacturers.append(row[0])
                    self.models.append(row[1])
                    self.remaps.append(row[2])
                    self.types.append(row[3])
            print(f"  Loaded {len(self.types)} silhouette mappings from Sils.csv.")
            return True

        # Fall back to built-in data (equivalent to VB.NET Call Sils())
        from .sils_data import SILS_DATA
        for mfr, mdl, rmp, typ in SILS_DATA:
            self.manufacturers.append(mfr)
            self.models.append(mdl)
            self.remaps.append(rmp)
            self.types.append(typ)
        print(f"  Sils.csv not found; using {len(self.types)} built-in silhouette mappings.")

    def determine_silhouette(self, icao_type_input: str, manufacturer: str,
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
