"""
VRS database merge module.

Merges FAA, CCAR, and OpenSky intermediate databases into the VRS
AircraftOnlineLookupCache.sqb database.

Efficiency improvements over VB.NET version:
- Uses parameterized queries throughout (no string-built SQL)
- Pre-loads rules into structured data (not re-parsed per record)
- Uses dict-based ICAO lookups instead of per-record SQL queries
- Wraps updates in transactions for 10-50x speedup
- Manufacturer normalization via lookup table instead of if/else chains
"""

import glob
import os
import sqlite3
import shutil
from datetime import datetime, timezone
from typing import Dict, List, Optional, Tuple

from .utils import title_case, clean_operator_name, ProgressReporter
from .silhouette import SilhouetteLookup, load_silhouettes
from .config import Settings

# Detail log callback - GUI patches this to route to the Details tab
# Default: no-op (CLI mode doesn't need per-record detail spam)
def _default_detail_log(msg: str):
    pass

detail_log = _default_detail_log

# Phase callback - GUI patches this to update the phase indicator
def _default_phase_callback(phase: str):
    pass

phase_callback = _default_phase_callback


# Manufacturer normalization map (replaces repeated if/elseif chains in VB.NET)
MANUFACTURER_NORMALIZE = {
    "Cessna": "Cessna",
    "Boeing": "Boeing",
    "The Boeing": "Boeing",
    "Piper": "Piper",
    "Texas Engineering and Manufacturing Co. Inc.": "Temco",
    "Champion": "Champion",
    "Schweizer": "Schweizer",
    "Aeronca": "Aeronca",
    "Aero Commander": "Aero Commander",
    "Beech": "Beech",
    "Airbus": "Airbus",
    "Bell": "Bell",
}


def _normalize_manufacturer(mfr: str) -> str:
    """Normalize manufacturer name for silhouette lookup.

    Uses case-insensitive prefix matching against the normalization map.
    Replaces the long if/elseif chain in VB.NET.
    """
    if not mfr:
        return mfr
    mfr_lower = mfr.lower()
    # Check exact matches first, then prefix matches (case-insensitive)
    for prefix, normalized in MANUFACTURER_NORMALIZE.items():
        if mfr_lower == prefix.lower():
            return normalized
    for prefix, normalized in MANUFACTURER_NORMALIZE.items():
        if mfr_lower.startswith(prefix.lower()):
            return normalized
    return mfr


# ---------------------------------------------------------------------------
# Rules engine
# ---------------------------------------------------------------------------

class Rule:
    """A single rule from Rules.csv."""
    __slots__ = ('match_fields', 'change_fields', 'msg_field', 'msg_text')

    FIELD_NAMES = ['ICAO', 'Registration', 'Country', 'Manufacturer',
                   'Model', 'ModelIcao', 'Operator', 'OperatorICAO']

    def __init__(self, match_fields: Dict[str, str], change_fields: Dict[str, str],
                 msg_field: str = "", msg_text: str = ""):
        self.match_fields = match_fields    # field_name -> required value (or "!value" for negation)
        self.change_fields = change_fields  # field_name -> new value
        self.msg_field = msg_field
        self.msg_text = msg_text


def load_rules(rules_path: str) -> List[Rule]:
    """Load rules from Rules.csv.

    The CSV has columns:
      Rule, ICAO, Registration, Country, Manufacturer, Model, ModelIcao, Operator, OperatorICAO,
      (gap cols), ICAO, Registration, Country, Manufacturer, Model, ModelIcao, Operator, OperatorICAO,
      (gap), Message, From dB
    """
    rules = []
    if not os.path.exists(rules_path):
        return rules

    with open(rules_path, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.readlines()

    for line_num, line in enumerate(lines):
        if line_num == 0:
            continue  # Skip header

        fields = line.rstrip('\n').split(',')
        if len(fields) < 22:
            continue

        # Match fields (columns 1-8)
        match = {}
        for i, name in enumerate(Rule.FIELD_NAMES):
            val = fields[i + 1].strip() if i + 1 < len(fields) else ""
            if val:
                match[name] = val

        if not match:
            continue

        # Change fields (columns 12-19)
        change = {}
        for i, name in enumerate(Rule.FIELD_NAMES):
            val = fields[i + 12].strip() if i + 12 < len(fields) else ""
            if val:
                change[name] = val

        msg_field = fields[20].strip() if len(fields) > 20 else ""
        msg_text = fields[21].strip() if len(fields) > 21 else ""

        rules.append(Rule(match, change, msg_field, msg_text))

    print(f"  Loaded {len(rules)} rules from Rules.csv.")
    return rules


def apply_rules(rules: List[Rule], record: Dict[str, str]) -> Optional[Rule]:
    """Check if any rule matches the record and return the matched Rule.

    Returns the first matching Rule, or None if no rule matched.
    VB.NET string comparison is case-insensitive by default, so we
    use .lower() for all field matching.
    """
    for rule in rules:
        positive_count = 0
        positive_matched = 0
        negation_triggered = False

        for field_name, rule_val in rule.match_fields.items():
            field_val = record.get(field_name) or ""
            if rule_val.startswith("!"):
                # Negation: rule matches if field equals the negated value
                if field_val.lower() == rule_val[1:].lower():
                    negation_triggered = True
                    break
            else:
                positive_count += 1
                if field_val.lower() == rule_val.lower():
                    positive_matched += 1

        if negation_triggered:
            continue
        if positive_count > 0 and positive_matched == positive_count:
            return rule

    return None


# Mapping from Rule FIELD_NAMES (mixed case) to DB/existing-dict column names
_RULE_TO_DB = {
    'ICAO': 'Icao', 'Registration': 'Registration', 'Country': 'Country',
    'Manufacturer': 'Manufacturer', 'Model': 'Model', 'ModelIcao': 'ModelIcao',
    'Operator': 'Operator', 'OperatorICAO': 'OperatorIcao',
}


def _apply_rules_pass(rules: List[Rule], existing: Dict[str, dict],
                      dirty: set, utc_now: str):
    """Apply Rules.csv as a final pass over all VRS records.

    Runs after FAA, CCAR, and OpenSky merges so that rules always have the
    last word regardless of which data source populated the record.
    """
    if not rules:
        return

    print("  Applying rules...")
    prog = ProgressReporter("Rules")
    total = len(existing)
    rules_applied = 0
    count = 0

    for icao, ex in existing.items():
        count += 1
        prog.update(count, total)

        # Build a record using Rule field names
        record = {rk: (ex.get(dk) or "") for rk, dk in _RULE_TO_DB.items()}

        matched_rule = apply_rules(rules, record)

        if matched_rule:
            rules_applied += 1
            # Map changed Rule fields back to DB column names
            for rule_field, new_val in matched_rule.change_fields.items():
                db_col = _RULE_TO_DB.get(rule_field, rule_field)
                ex[db_col] = new_val

            # Build log message matching VB.NET format:
            #   msg_text + field_value + ". " + Registration + " -- From rules."
            reg = ex.get('Registration', '')
            change_msg = matched_rule.msg_text
            if matched_rule.msg_field:
                # Map msg_field names from Rules.csv to record keys
                msg_field_map = {
                    'ICAO': 'ICAO', 'Registration': 'Registration',
                    'Country': 'Country', 'Manufacturer': 'Manufacturer',
                    'Model': 'Model', 'ModelIcao': 'ModelIcao',
                    'Operator': 'Operator', 'Operator ICAO': 'OperatorICAO',
                }
                rk = msg_field_map.get(matched_rule.msg_field, '')
                if rk:
                    change_msg += record.get(rk, '') + "."
            if change_msg:
                detail_log(f"{change_msg} {reg}")

            dirty.add(icao)

    prog.done()

    if rules_applied > 0:
        detail_log(f"Rules: {rules_applied:,} applied.")
        print(f"  Rules: {rules_applied:,} applied.")


# ---------------------------------------------------------------------------
# Main merge
# ---------------------------------------------------------------------------

def _utc_now_str() -> str:
    """UTC timestamp in VRS format."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S") + ".0000000Z"


def _load_existing_vrs(conn: sqlite3.Connection) -> Dict[str, dict]:
    """Pre-load all existing VRS records into a dict keyed by ICAO.

    This is a major efficiency win: the VB.NET version does a SELECT per record,
    which means ~600K individual queries. Loading everything into a dict gives
    O(1) lookups.
    """
    print("  Pre-loading existing VRS database into memory...")
    existing = {}
    try:
        cur = conn.execute("SELECT * FROM AircraftDetail")
        cols = [desc[0] for desc in cur.description]
        for row in cur:
            rec = dict(zip(cols, row))
            icao = rec.get("Icao", "")
            if icao:
                existing[icao] = rec
    except Exception:
        pass  # Table might not exist yet
    print(f"  Loaded {len(existing):,} existing records.")
    return existing


def update_vrs(settings: Settings):
    """Main merge: FAA + CCAR + OpenSky -> VRS AircraftOnlineLookupCache.sqb.

    This is the Python port of the VB.NET Update_VRS() subroutine, with
    significant efficiency improvements.
    """
    faa_db = settings.faa_db_path
    vrs_db = settings.vrs_db_path
    ccar_db = settings.ccar_db_path
    nz_caa_db = settings.nz_caa_db_path
    casa_db = settings.casa_db_path
    opensky_db = settings.opensky_db_path

    build_complete = settings.build_complete

    # Validate FAA DB exists (unless skipped)
    if not settings.skip_faa:
        if not os.path.exists(faa_db):
            print("  ERROR: FAADatabase.sqb not found. Run with --download-faa first.")
            return

        # Save a dated snapshot of the FAA database for future PII recovery
        faa_snapshot = os.path.join(settings.work_dir,
                                    f"FAADatabase {datetime.now().strftime('%d%m%Y')}.sqb")
        if not os.path.exists(faa_snapshot):
            print(f"  Saving FAA snapshot as {os.path.basename(faa_snapshot)}...")
            shutil.copy2(faa_db, faa_snapshot)

    # Copy VRS database from VRS install location to working dir
    if settings.vrs_dir:
        vrs_source = os.path.join(settings.vrs_dir, "AircraftOnlineLookupCache.sqb")
        if os.path.exists(vrs_source):
            print("  Copying VRS database to working directory...")
            shutil.copy2(vrs_source, vrs_db)
            if settings.backup_vrs_db:
                backup_name = f"AircraftOnlineLookupCache-{datetime.now().strftime('%y%m%d-%H%M')}.sqb"
                shutil.copy2(vrs_source, os.path.join(settings.work_dir, backup_name))
                print(f"  Backup saved as {backup_name}")
        elif not os.path.exists(vrs_db):
            print("  ERROR: VRS AircraftOnlineLookupCache.sqb not found.")
            return

    if not os.path.exists(vrs_db):
        print("  ERROR: AircraftOnlineLookupCache.sqb not found in working directory.")
        return

    # Load silhouette data
    sil_lookup = load_silhouettes(settings.sils_path)

    # Load rules
    rules = load_rules(settings.rules_path)

    # Open VRS database
    vrs_conn = sqlite3.connect(vrs_db)
    vrs_conn.execute("PRAGMA journal_mode=WAL")
    vrs_conn.execute("PRAGMA synchronous=NORMAL")

    # Pre-load existing records for O(1) lookups
    existing = _load_existing_vrs(vrs_conn)

    # Snapshot original keys so we can split final writes into UPDATEs vs INSERTs
    original_icaos = set(existing.keys())
    # ICAOs touched by any merge — only these get written at the end
    dirty: set = set()

    utc_now = _utc_now_str()

    # Run merges in reverse priority order (highest priority runs last and wins)
    for source in reversed(settings.merge_order):
        if source == "OpenSky":
            phase_callback("Processing: OpenSky")
            if settings.skip_opensky:
                print("  OpenSky merge skipped.")
            elif os.path.exists(opensky_db):
                _merge_opensky(opensky_db, existing, dirty, sil_lookup, build_complete, utc_now)
            else:
                print("  OpenSky database not found, skipping.")
        elif source == "CCAR":
            phase_callback("Processing: CCAR")
            if settings.skip_ccar:
                print("  CCAR merge skipped.")
            elif os.path.exists(ccar_db):
                _merge_ccar(ccar_db, existing, dirty, sil_lookup, build_complete, utc_now)
            else:
                print("  CCAR database not found, skipping.")
        elif source == "NZ CAA":
            phase_callback("Processing: NZ CAA")
            if settings.skip_nz_caa:
                print("  NZ CAA merge skipped.")
            elif os.path.exists(nz_caa_db):
                _merge_nz_caa(nz_caa_db, existing, dirty, sil_lookup, build_complete, utc_now)
            else:
                print("  NZ CAA database not found, skipping.")
        elif source == "CASA":
            phase_callback("Processing: CASA")
            if settings.skip_casa:
                print("  CASA merge skipped.")
            elif os.path.exists(casa_db):
                _merge_casa(casa_db, existing, dirty, sil_lookup, utc_now)
            else:
                print("  CASA database not found, skipping.")
        elif source == "FAA":
            phase_callback("Processing: FAA")
            if settings.skip_faa:
                print("  FAA merge skipped.")
            else:
                _merge_faa(faa_db, existing, dirty, sil_lookup, build_complete, utc_now,
                           work_dir=settings.work_dir)
        elif source.startswith("Rules"):
            phase_callback("Processing: Rules")
            if settings.skip_rules:
                print("  Rules pass skipped.")
            else:
                _apply_rules_pass(rules, existing, dirty, utc_now)

    # Single bulk write of every touched record at the end of the run.
    phase_callback("Processing: Writing to disk")
    _final_flush(vrs_conn, existing, original_icaos, dirty, utc_now)

    vrs_conn.close()

    # Copy updated DB back to VRS location
    if settings.vrs_dir:
        vrs_dest = os.path.join(settings.vrs_dir, "AircraftOnlineLookupCache.sqb")
        if os.path.exists(os.path.dirname(vrs_dest)):
            print("  Copying updated database back to VRS location...")
            shutil.copy2(vrs_db, vrs_dest)

    print("  VRS database update complete.")


def _resolve_faa_operator(owner: str, registration: str, manu_kit: str,
                          faa_type: str, ref_lookup: dict,
                          ex: Optional[dict], old_faa_lookup: Dict[str, str],
                          icao: str, today_str: str) -> str:
    """Resolve the operator field for an FAA record, handling PII redaction.

    Port of VRS.vb lines 199-241.  When the FAA database shows no owner
    (PII redacted), the logic preserves existing VRS operator info and
    uses indicator prefixes:
      --PII Removed--         owner unknown, nothing to fall back on
      --Old PII (date): xxx   FAA blanked the owner; we kept the old VRS value
      --Old FAA (date): xxx   recovered from an older FAA database snapshot
      --PII Kit (date): xxx   kit plane whose builder name we still know
    """
    if owner:
        # FAA has an owner — use it regardless
        return owner

    # --- FAA owner is blank (PII redacted) ---

    # 1. Check existing VRS record
    if ex:
        existing_op = ex.get("Operator", "") or ""
        if existing_op:
            if existing_op.startswith(("--Old PII", "--Old FAA", "--PII Kit")):
                # Already tagged from a prior run — preserve the original date
                return existing_op
            if not existing_op.startswith("--PII"):
                tagged = f"--Old PII ({today_str}): {existing_op}"
                detail_log(f"PII info removed.  Using VRS db Operator:  {tagged}, {registration}.")
                return tagged

    # 2. Check older FAA database snapshot
    old_owner = old_faa_lookup.get(icao, "")
    if old_owner:
        tagged = f"--Old FAA ({today_str}): {old_owner}"
        detail_log(f"PII info removed.  Using old FAA db Operator:  {tagged}, {registration}.")
        return tagged

    # 3. Kit plane — use manufacturer as fallback
    if manu_kit:
        ref = ref_lookup.get(faa_type)
        kit_mfr = ref["Manufacturer"] if ref and ref.get("Manufacturer") else ""
        if kit_mfr:
            tagged = f"--PII Kit ({today_str}): {kit_mfr}"
            detail_log(f"PII info removed from kit plane.  Using manufacturer:  {kit_mfr}, {registration}.")
            return tagged
        else:
            detail_log(f"PII info removed from kit plane and manufacturer null:  {registration}.")
            return "--PII Removed--"

    # 4. Nothing to fall back on
    detail_log(f"PII information removed:  {registration}.")
    return "--PII Removed--"


def _load_old_faa_owners(work_dir: str, current_faa_db: str) -> Dict[str, str]:
    """Scan work_dir for older FAADatabase*.sqb files and load owner names.

    Returns a dict of ICAO -> Owner_Name from the most recent older snapshot.
    This lets us recover owner info for aircraft whose PII was redacted in
    the current FAA release.
    """
    pattern = os.path.join(work_dir, "FAADatabase*.sqb")
    current_name = os.path.basename(current_faa_db)
    today_snapshot = f"FAADatabase {datetime.now().strftime('%d%m%Y')}.sqb"
    candidates = []
    for path in glob.glob(pattern):
        basename = os.path.basename(path)
        if basename != current_name and basename != today_snapshot:
            candidates.append(path)

    if not candidates:
        return {}

    # Use the most recently modified older snapshot
    candidates.sort(key=os.path.getmtime, reverse=True)
    old_db = candidates[0]
    print(f"  Loading old FAA owner data from {os.path.basename(old_db)}...")

    lookup = {}
    try:
        conn = sqlite3.connect(old_db)
        for row in conn.execute("SELECT ICAO, Owner_Name FROM Master WHERE Owner_Name IS NOT NULL AND Owner_Name != ''"):
            lookup[row[0]] = row[1]
        conn.close()
    except Exception:
        pass

    print(f"  Loaded {len(lookup):,} owner records from old FAA database.")
    return lookup


def _merge_faa(faa_db: str, existing: Dict[str, dict],
               dirty: set, sil: SilhouetteLookup,
               build_complete: bool, utc_now: str,
               work_dir: str = ""):
    """Merge FAA data into VRS."""
    print("  Updating VRS database (FAA)...")

    faa_conn = sqlite3.connect(faa_db)
    faa_conn.row_factory = sqlite3.Row

    # Load aircraft reference table into dict for O(1) lookup
    ref_lookup = {}
    try:
        for row in faa_conn.execute("SELECT * FROM Aircraft_Reference"):
            ref_lookup[row["Type"]] = dict(row)
    except Exception:
        pass

    # Load old FAA database for PII recovery fallback
    old_faa_lookup = _load_old_faa_owners(work_dir, faa_db) if work_dir else {}

    total = faa_conn.execute("SELECT COUNT(*) FROM Master").fetchone()[0]
    prog = ProgressReporter("FAA merge")

    count = 0
    pii_old_pii = 0
    pii_old_faa = 0
    pii_kit = 0
    pii_removed = 0
    num_updates = 0
    num_inserts = 0
    today_str = datetime.now().strftime("%d-%b-%Y")

    for row in faa_conn.execute("SELECT * FROM Master"):
        count += 1
        prog.update(count, total)

        icao = row["ICAO"]
        registration = row["Registration"]
        owner = row["Owner_Name"] or ""
        year_built = row["Year"] or ""
        faa_type = row["Type"]
        manu_kit = row["Manufacturer_Kit"]
        type_kit = row["Type_Kit"]

        ex = existing.get(icao)

        if ex is None and not build_complete:
            # Not in VRS and not building complete — skip
            continue

        # Determine manufacturer and model
        manufacturer = ""
        model = ""
        if manu_kit:
            manufacturer = manu_kit
            model = type_kit or ""
        else:
            ref = ref_lookup.get(faa_type)
            if ref:
                manufacturer = ref["Manufacturer"] or ""
                model = ref["Model"] or ""

        # Preserve existing model if FAA has none (VRS.vb lines 287-292)
        if not model and ex and ex.get("Model"):
            model = ex["Model"]

        # Resolve operator with PII handling, then clean the result
        raw_operator = _resolve_faa_operator(
            owner, registration, manu_kit, faa_type, ref_lookup, ex,
            old_faa_lookup, icao, today_str)

        # Track PII case counts
        if raw_operator.startswith("--Old PII"):
            pii_old_pii += 1
        elif raw_operator.startswith("--Old FAA"):
            pii_old_faa += 1
        elif raw_operator.startswith("--PII Kit"):
            pii_kit += 1
        elif raw_operator == "--PII Removed--":
            pii_removed += 1

        operator = clean_operator_name(raw_operator, registration, detail_callback=detail_log)

        # Determine silhouette
        model_icao, found = sil.determine_silhouette(model, manufacturer)
        if not found:
            model_icao = ""

        # Preserve existing created time
        created_time = utc_now
        if ex and ex.get("CreatedUtc"):
            created_time = ex["CreatedUtc"]

        country = "United States"

        if ex is not None:
            num_updates += 1
            existing[icao] = {**ex, 'Registration': registration, 'Country': country,
                              'Manufacturer': manufacturer, 'Model': model,
                              'ModelIcao': model_icao, 'Operator': operator,
                              'YearBuilt': year_built, 'CreatedUtc': created_time}
        else:
            num_inserts += 1
            existing[icao] = {'Icao': icao, 'Registration': registration,
                              'Country': country, 'Manufacturer': manufacturer,
                              'Model': model, 'ModelIcao': model_icao,
                              'Operator': operator, 'YearBuilt': year_built,
                              'OperatorIcao': '', 'CreatedUtc': created_time}
        dirty.add(icao)

    faa_conn.close()
    prog.done()

    # Summary
    detail_log(f"FAA: {count:,} processed, {num_updates:,} updated, {num_inserts:,} new.")
    print(f"  FAA merge complete. {count:,} records processed ({num_updates:,} updated, {num_inserts:,} new).")

    pii_total = pii_old_pii + pii_old_faa + pii_kit + pii_removed
    if pii_total > 0:
        detail_log(f"PII summary:  {pii_total:,} redacted owners — "
                   f"{pii_old_pii:,} kept from VRS, "
                   f"{pii_old_faa:,} recovered from old FAA, "
                   f"{pii_kit:,} kit manufacturers, "
                   f"{pii_removed:,} no info available.")
        print(f"  PII: {pii_old_pii:,} kept from VRS, {pii_old_faa:,} from old FAA, "
              f"{pii_kit:,} kit mfrs, {pii_removed:,} removed ({pii_total:,} total)")

    detail_log("FAA database processing completed.")


def _merge_ccar(ccar_db: str, existing: Dict[str, dict],
                dirty: set, sil: SilhouetteLookup,
                build_complete: bool, utc_now: str):
    """Merge CCAR data into VRS."""
    print("  Updating VRS database (CCAR)...")

    ccar_conn = sqlite3.connect(ccar_db)
    ccar_conn.row_factory = sqlite3.Row

    total = ccar_conn.execute("SELECT COUNT(*) FROM Master").fetchone()[0]
    prog = ProgressReporter("CCAR merge")

    count = 0
    num_updates = 0
    num_inserts = 0

    for row in ccar_conn.execute("SELECT * FROM Master"):
        count += 1
        prog.update(count, total)

        icao = row["ICAO"]
        registration = row["Registration"]
        owner = row["Owner_Name"] or ""
        year_built = row["Year"] or ""
        manufacturer = row["Manufacturer"] or ""
        model = row["CCAR_Type"] or ""

        ex = existing.get(icao)

        if ex is None and not build_complete:
            continue

        # For CCAR, prefer existing data if available (fill gaps only)
        if ex:
            if not manufacturer and ex.get("Manufacturer"):
                manufacturer = ex["Manufacturer"]
            if not model and ex.get("Model"):
                model = ex["Model"]
            if not year_built and ex.get("YearBuilt"):
                year_built = ex["YearBuilt"]

        # Clean operator
        operator = clean_operator_name(owner, registration, detail_callback=detail_log)

        # Normalize manufacturer for silhouette lookup
        mfr_normalized = _normalize_manufacturer(manufacturer)

        # Determine silhouette (VRS.vb lines 618-630)
        # Try first word of normalized manufacturer, then full normalized name
        mfr_first_word = mfr_normalized.split()[0] if mfr_normalized else ""
        model_icao, found = sil.determine_silhouette(model, mfr_first_word)
        if not found:
            model_icao, found = sil.determine_silhouette(model, mfr_normalized)
        if not found:
            model_icao = ""

        # Use existing model_icao if we don't have one and existing does
        if not model_icao and ex and ex.get("ModelIcao"):
            model_icao = ex["ModelIcao"]

        created_time = utc_now
        if ex and ex.get("CreatedUtc"):
            created_time = ex["CreatedUtc"]

        country = "Canada"
        operator_icao = ex.get("OperatorIcao", "") if ex else ""

        if ex is not None:
            num_updates += 1
        else:
            num_inserts += 1

        existing[icao] = {'Icao': icao, 'Registration': registration,
                          'Country': country, 'Manufacturer': manufacturer,
                          'Model': model, 'ModelIcao': model_icao,
                          'Operator': operator, 'YearBuilt': year_built,
                          'OperatorIcao': operator_icao,
                          'CreatedUtc': created_time}
        dirty.add(icao)

    ccar_conn.close()
    prog.done()
    detail_log(f"CCAR: {count:,} processed, {num_updates:,} updated, {num_inserts:,} new.")
    detail_log("CCAR database processing completed.")
    print(f"  CCAR merge complete. {count:,} records processed ({num_updates:,} updated, {num_inserts:,} new).")


def _merge_nz_caa(nz_caa_db: str, existing: Dict[str, dict],
                  dirty: set, sil: SilhouetteLookup,
                  build_complete: bool, utc_now: str):
    """Merge NZ CAA data into VRS."""
    print("  Updating VRS database (NZ CAA)...")

    nz_conn = sqlite3.connect(nz_caa_db)
    nz_conn.row_factory = sqlite3.Row

    total = nz_conn.execute("SELECT COUNT(*) FROM Aircraft").fetchone()[0]
    prog = ProgressReporter("NZ CAA merge")

    count = 0
    num_updates = 0
    num_inserts = 0

    for row in nz_conn.execute("SELECT * FROM Aircraft"):
        count += 1
        prog.update(count, total)

        icao = row["ICAO"]
        registration = row["Registration"] or ""
        manufacturer = row["Manufacturer"] or ""
        model = row["Model"] or ""
        owner = row["Owner"] or ""
        serial = row["Serial"] or ""
        year_built = row["YearBuilt"] or ""

        ex = existing.get(icao)

        if ex is None and not build_complete:
            continue

        # Fill gaps from existing data
        if ex:
            if not manufacturer and ex.get("Manufacturer"):
                manufacturer = ex["Manufacturer"]
            if not model and ex.get("Model"):
                model = ex["Model"]
            if not year_built and ex.get("YearBuilt"):
                year_built = ex["YearBuilt"]
            if not owner and ex.get("Operator"):
                owner = ex["Operator"]

        # Normalize manufacturer for silhouette lookup
        mfr_normalized = _normalize_manufacturer(manufacturer)

        # Determine silhouette
        mfr_first_word = mfr_normalized.split()[0] if mfr_normalized else ""
        model_icao, found = sil.determine_silhouette(model, mfr_first_word)
        if not found:
            model_icao, found = sil.determine_silhouette(model, mfr_normalized)
        if not found:
            model_icao = ""

        # Use existing model_icao if we don't have one
        if not model_icao and ex and ex.get("ModelIcao"):
            model_icao = ex["ModelIcao"]

        created_time = utc_now
        if ex and ex.get("CreatedUtc"):
            created_time = ex["CreatedUtc"]

        country = "New Zealand"
        operator_icao = ex.get("OperatorIcao", "") if ex else ""

        if ex is not None:
            num_updates += 1
        else:
            num_inserts += 1

        existing[icao] = {'Icao': icao, 'Registration': registration,
                          'Country': country, 'Manufacturer': manufacturer,
                          'Model': model, 'ModelIcao': model_icao,
                          'Operator': owner, 'YearBuilt': year_built,
                          'OperatorIcao': operator_icao,
                          'CreatedUtc': created_time}
        dirty.add(icao)

    nz_conn.close()
    prog.done()
    detail_log(f"NZ CAA: {count:,} processed, {num_updates:,} updated, {num_inserts:,} new.")
    detail_log("NZ CAA database processing completed.")
    print(f"  NZ CAA merge complete. {count:,} records processed ({num_updates:,} updated, {num_inserts:,} new).")


def _merge_casa(casa_db: str, existing: Dict[str, dict],
                dirty: set, sil: SilhouetteLookup,
                utc_now: str):
    """Merge CASA data into VRS using registration-based matching.

    CASA has no ICAO hex column, so we build a reverse lookup from existing
    VRS records (registration -> ICAO) and match by VH- registration.
    This means CASA can only UPDATE existing records, never insert new ones.
    """
    print("  Updating VRS database (CASA)...")

    # Build reverse lookup: registration -> ICAO hex for Australian VH- aircraft
    reg_to_icao = {}
    for icao, rec in existing.items():
        reg = rec.get("Registration", "") or ""
        if reg.upper().startswith("VH-"):
            reg_to_icao[reg.upper()] = icao

    print(f"  Found {len(reg_to_icao):,} VH- registrations in existing VRS data.")

    casa_conn = sqlite3.connect(casa_db)
    casa_conn.row_factory = sqlite3.Row

    total = casa_conn.execute("SELECT COUNT(*) FROM Aircraft").fetchone()[0]
    prog = ProgressReporter("CASA merge")

    count = 0
    num_updates = 0
    num_skipped = 0

    for row in casa_conn.execute("SELECT * FROM Aircraft"):
        count += 1
        prog.update(count, total)

        registration = (row["Registration"] or "").strip()
        if not registration:
            continue

        # Look up ICAO hex by registration
        icao = reg_to_icao.get(registration.upper())
        if not icao:
            num_skipped += 1
            continue

        manufacturer = row["Manufacturer"] or ""
        model = row["Model"] or ""
        owner = row["Owner"] or ""
        serial = row["Serial"] or ""
        year_built = row["YearBuilt"] or ""
        icao_type = row["ICAOTypeDesig"] or ""

        ex = existing.get(icao)
        if ex is None:
            # Should not happen since we built reg_to_icao from existing, but guard
            num_skipped += 1
            continue

        # Fill gaps from existing data (don't overwrite good data with blanks)
        if not manufacturer and ex.get("Manufacturer"):
            manufacturer = ex["Manufacturer"]
        if not model and ex.get("Model"):
            model = ex["Model"]
        if not year_built and ex.get("YearBuilt"):
            year_built = ex["YearBuilt"]
        if not owner and ex.get("Operator"):
            owner = ex["Operator"]

        # Use CASA's ICAO type designator directly as ModelIcao (valuable data)
        model_icao = icao_type
        if not model_icao:
            # Fall back to existing ModelIcao or silhouette lookup
            if ex.get("ModelIcao"):
                model_icao = ex["ModelIcao"]
            elif model:
                mfr_normalized = _normalize_manufacturer(manufacturer)
                mfr_first_word = mfr_normalized.split()[0] if mfr_normalized else ""
                model_icao, found = sil.determine_silhouette(model, mfr_first_word)
                if not found:
                    model_icao, found = sil.determine_silhouette(model, mfr_normalized)
                if not found:
                    model_icao = ""

        created_time = utc_now
        if ex.get("CreatedUtc"):
            created_time = ex["CreatedUtc"]

        country = "Australia"
        operator_icao = ex.get("OperatorIcao", "") or ""

        num_updates += 1

        existing[icao] = {'Icao': icao, 'Registration': registration,
                          'Country': country, 'Manufacturer': manufacturer,
                          'Model': model, 'ModelIcao': model_icao,
                          'Operator': owner, 'YearBuilt': year_built,
                          'OperatorIcao': operator_icao,
                          'CreatedUtc': created_time}
        dirty.add(icao)

    casa_conn.close()
    prog.done()
    detail_log(f"CASA: {count:,} processed, {num_updates:,} updated, {num_skipped:,} skipped (no matching ICAO).")
    detail_log("CASA database processing completed.")
    print(f"  CASA merge complete. {count:,} records processed ({num_updates:,} updated, {num_skipped:,} no match).")


def _merge_opensky(opensky_db: str, existing: Dict[str, dict],
                   dirty: set, sil: SilhouetteLookup,
                   build_complete: bool, utc_now: str):
    """Merge OpenSky data into VRS."""
    print("  Updating VRS database (OpenSky)...")

    os_conn = sqlite3.connect(opensky_db)
    os_conn.row_factory = sqlite3.Row

    total = os_conn.execute("SELECT COUNT(*) FROM Aircraft").fetchone()[0]
    prog = ProgressReporter("OpenSky merge")

    count = 0
    num_updates = 0
    num_inserts = 0

    for row in os_conn.execute("SELECT * FROM Aircraft"):
        count += 1
        prog.update(count, total)

        icao = row["ICAO"]
        registration = row["Registration"] or ""
        manufacturer = row["Manufacturer"] or ""
        model = row["Model"] or ""
        model_icao = row["ModelIcao"] or ""
        operator = row["Operator"] or ""
        operator_icao = row["OperatorIcao"] or ""
        serial = row["Serial"] or ""
        year_built = row["YearBuilt"] or ""

        ex = existing.get(icao)

        if ex is None and not build_complete:
            continue

        # For OpenSky: fill gaps from existing, don't overwrite good data
        # Year: prefer existing VRS year over OpenSky (VRS.vb lines 762-771)
        if ex:
            ex_year = str(ex.get("YearBuilt", "") or "")
            if ex_year:
                year_built = ex_year
                if ex_year[:4].isdigit() and int(ex_year[:4]) < 1900:
                    year_built = ""
            # else: keep OpenSky year_built as-is

            # Prefer existing operator if available
            if not operator and ex.get("Operator"):
                operator = ex["Operator"]

            # Prefer existing manufacturer/model if available
            if not manufacturer and ex.get("Manufacturer"):
                manufacturer = ex["Manufacturer"]
            if not model and ex.get("Model"):
                model = ex["Model"]

            # Prefer existing serial
            if not serial and ex.get("Serial"):
                serial = ex["Serial"]

        # Clean operator
        if operator:
            operator = clean_operator_name(operator, registration, detail_callback=detail_log)

        # Silhouette logic (matches VRS.vb lines 862-885):
        # If existing VRS record has a non-empty ModelIcao, keep it.
        # Otherwise, if OpenSky's ModelIcao is also empty, run lookup.
        if ex:
            ex_model_icao = ex.get("ModelIcao", "")
            if ex_model_icao and ex_model_icao != "NULL":
                # Existing VRS record has a silhouette - keep it
                model_icao = ex_model_icao
            elif not model_icao and model:
                # Both existing and OpenSky are empty - try lookup
                mfr_normalized = _normalize_manufacturer(manufacturer)
                model_icao, found = sil.determine_silhouette(model, mfr_normalized)
                if not found:
                    model_icao, found = sil.determine_silhouette(model, manufacturer)
                if not found:
                    model_icao = ""
        else:
            # No existing record - try lookup if OpenSky didn't provide one
            if not model_icao and model:
                mfr_normalized = _normalize_manufacturer(manufacturer)
                model_icao, found = sil.determine_silhouette(model, mfr_normalized)
                if not found:
                    model_icao, found = sil.determine_silhouette(model, manufacturer)
                if not found:
                    model_icao = ""

        # Preserve existing operator_icao
        if ex and not operator_icao and ex.get("OperatorIcao"):
            operator_icao = ex["OperatorIcao"]

        created_time = utc_now
        if ex and ex.get("CreatedUtc"):
            created_time = ex["CreatedUtc"]

        if ex is not None:
            num_updates += 1
        else:
            num_inserts += 1

        # OpenSky has no Country field, but other merges in this run may have
        # set one — preserve whatever's currently in memory.
        country = ex.get('Country', '') if ex else ''

        existing[icao] = {'Icao': icao, 'Registration': registration,
                          'Country': country,
                          'Manufacturer': manufacturer, 'Model': model,
                          'ModelIcao': model_icao, 'Operator': operator,
                          'OperatorIcao': operator_icao, 'Serial': serial,
                          'YearBuilt': year_built, 'CreatedUtc': created_time}
        dirty.add(icao)

    os_conn.close()
    prog.done()
    detail_log(f"OpenSky: {count:,} processed, {num_updates:,} updated, {num_inserts:,} new.")
    detail_log("OpenSky database processing completed.")
    print(f"  OpenSky merge complete. {count:,} records processed ({num_updates:,} updated, {num_inserts:,} new).")


def _final_flush(vrs_conn: sqlite3.Connection,
                 existing: Dict[str, dict],
                 original_icaos: set,
                 dirty: set,
                 utc_now: str):
    """Single bulk write at the end of all merges.

    All per-source merges only mutate the in-memory `existing` dict and add the
    touched ICAO to `dirty`. This function then partitions `dirty` by whether
    the row was originally in the DB (UPDATE) or new (INSERT) and writes
    everything in one transaction. Untouched rows are left alone.
    """
    if not dirty:
        print("")
        print("  No records were modified — nothing to write.")
        return

    print("")
    print("  Building write batches from in-memory cache...")
    updates = []
    inserts = []
    for icao in dirty:
        ex = existing[icao]
        if icao in original_icaos:
            updates.append((
                ex.get('Registration', '') or '',
                ex.get('Country', '') or '',
                ex.get('Manufacturer', '') or '',
                ex.get('Model', '') or '',
                ex.get('ModelIcao', '') or '',
                ex.get('Operator', '') or '',
                ex.get('YearBuilt', '') or '',
                ex.get('CreatedUtc', utc_now) or utc_now,
                utc_now,
                ex.get('OperatorIcao', '') or '',
                icao,
            ))
        else:
            inserts.append((
                icao,
                ex.get('Registration', '') or '',
                ex.get('Country', '') or '',
                ex.get('Manufacturer', '') or '',
                ex.get('Model', '') or '',
                ex.get('ModelIcao', '') or '',
                ex.get('YearBuilt', '') or '',
                ex.get('Operator', '') or '',
                ex.get('OperatorIcao', '') or '',
                ex.get('CreatedUtc', utc_now) or utc_now,
                utc_now,
            ))

    print(f"  Writing {len(updates):,} updates and {len(inserts):,} inserts "
          f"to disk in a single transaction...")
    flush_start = datetime.now()

    with vrs_conn:
        if updates:
            # Country = '' means "no source set a country, keep what's in the DB"
            vrs_conn.executemany("""
                UPDATE AircraftDetail SET
                    Registration = ?, Country = CASE WHEN ?='' THEN Country ELSE ? END,
                    Manufacturer = ?, Model = ?, ModelIcao = ?,
                    Operator = ?, YearBuilt = ?,
                    CreatedUtc = ?, UpdatedUtc = ?,
                    OperatorIcao = ?
                WHERE Icao = ?
            """, [(u[0], u[1], u[1], u[2], u[3], u[4], u[5], u[6], u[7], u[8], u[9], u[10]) for u in updates])

        if inserts:
            vrs_conn.executemany("""
                INSERT OR IGNORE INTO AircraftDetail
                    (Icao, Registration, Country, Manufacturer, Model, ModelIcao,
                     YearBuilt, Operator, OperatorIcao, CreatedUtc, UpdatedUtc)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)
            """, inserts)

    elapsed = (datetime.now() - flush_start).total_seconds()
    total_rows = len(updates) + len(inserts)
    rate = total_rows / elapsed if elapsed > 0 else 0
    print(f"  Disk write complete: {total_rows:,} rows in {elapsed:.2f}s "
          f"({rate:,.0f} rows/sec).")
