"""
Rules.csv load / save / validation.

Shared by the merge pass (vrs_merge) and the GUI rules editor so both agree
on the on-disk format.

File layout (23 columns, unchanged from the VB.NET original):

    0        Rule number (cosmetic - renumbered on save)
    1-8      Match fields   (ICAO .. OperatorICAO)
    9-11     Visual gap
    12-19    Change fields  (ICAO .. OperatorICAO)
    20       Visual gap
    21       Message text
    22       "From dB" - name of the field whose value is appended to the message
"""

import csv
import os
import shutil
from typing import Dict, List, Optional, Tuple

FIELD_NAMES = ['ICAO', 'Registration', 'Country', 'Manufacturer',
               'Model', 'ModelIcao', 'Operator', 'OperatorICAO']

MATCH_START = 1
CHANGE_START = 12
MSG_TEXT_COL = 21
MSG_FIELD_COL = 22
TOTAL_COLS = 23

HEADER = (['Rule'] + FIELD_NAMES + ['', '', ''] + FIELD_NAMES
          + ['', 'Message', 'From dB'])

# Mapping from rule field names (mixed case) to VRS DB column names
RULE_TO_DB = {
    'ICAO': 'Icao', 'Registration': 'Registration', 'Country': 'Country',
    'Manufacturer': 'Manufacturer', 'Model': 'Model', 'ModelIcao': 'ModelIcao',
    'Operator': 'Operator', 'OperatorICAO': 'OperatorIcao',
}

# Accepted values for the "From dB" column
MSG_FIELD_CHOICES = ['', 'ICAO', 'Registration', 'Country', 'Manufacturer',
                     'Model', 'ModelIcao', 'Operator', 'Operator ICAO']

# "From dB" column value -> rule field name
MSG_FIELD_MAP = {
    'ICAO': 'ICAO', 'Registration': 'Registration', 'Country': 'Country',
    'Manufacturer': 'Manufacturer', 'Model': 'Model', 'ModelIcao': 'ModelIcao',
    'Operator': 'Operator', 'Operator ICAO': 'OperatorICAO',
}


class Rule:
    """A single rule from Rules.csv."""
    __slots__ = ('match_fields', 'change_fields', 'msg_field', 'msg_text')

    FIELD_NAMES = FIELD_NAMES  # kept for backwards compatibility

    def __init__(self, match_fields: Optional[Dict[str, str]] = None,
                 change_fields: Optional[Dict[str, str]] = None,
                 msg_field: str = "", msg_text: str = ""):
        self.match_fields = dict(match_fields or {})    # field -> value ("!value" negates)
        self.change_fields = dict(change_fields or {})  # field -> new value
        self.msg_field = msg_field
        self.msg_text = msg_text

    def copy(self) -> 'Rule':
        return Rule(self.match_fields, self.change_fields,
                    self.msg_field, self.msg_text)

    # -- human-readable summaries for the editor ------------------------
    def describe_match(self) -> str:
        parts = []
        for name in FIELD_NAMES:
            val = self.match_fields.get(name)
            if not val:
                continue
            if val.startswith('!'):
                parts.append("%s ≠ %s" % (name, val[1:]))
            else:
                parts.append("%s = %s" % (name, val))
        return "  AND  ".join(parts)

    def describe_change(self) -> str:
        parts = ["%s → %s" % (name, self.change_fields[name])
                 for name in FIELD_NAMES if self.change_fields.get(name)]
        return ",  ".join(parts)


# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------

def rule_matches(rule: Rule, record: Dict[str, str]) -> bool:
    """True if `record` satisfies `rule`'s match fields.

    Comparison is case-insensitive (VB.NET string comparison default).
    A "!value" match field is a negation: the rule is rejected when the
    field equals that value. A rule made only of negations never matches.
    """
    positive_count = 0
    positive_matched = 0

    for field_name, rule_val in rule.match_fields.items():
        field_val = record.get(field_name) or ""
        if rule_val.startswith("!"):
            if field_val.lower() == rule_val[1:].lower():
                return False
        else:
            positive_count += 1
            if field_val.lower() == rule_val.lower():
                positive_matched += 1

    return positive_count > 0 and positive_matched == positive_count


def apply_rules(rules: List[Rule], record: Dict[str, str]) -> Optional[Rule]:
    """Return the first rule matching `record`, or None."""
    for rule in rules:
        if rule_matches(rule, record):
            return rule
    return None


# ---------------------------------------------------------------------------
# Load / save
# ---------------------------------------------------------------------------

def load_rules(rules_path: str, quiet: bool = False) -> List[Rule]:
    """Load rules from Rules.csv.

    Uses the csv module so quoted values containing commas survive a
    round trip. Short rows are padded rather than dropped, so a file
    re-saved by Excel (which trims trailing empty columns) still loads.
    """
    rules: List[Rule] = []
    if not os.path.exists(rules_path):
        return rules

    with open(rules_path, 'r', encoding='utf-8-sig', errors='replace', newline='') as f:
        for line_num, fields in enumerate(csv.reader(f)):
            if line_num == 0 and fields and fields[0].strip().lower() == 'rule':
                continue  # header
            if not any(v.strip() for v in fields):
                continue  # blank line

            if len(fields) < TOTAL_COLS:
                fields = fields + [''] * (TOTAL_COLS - len(fields))

            match = {}
            for i, name in enumerate(FIELD_NAMES):
                val = fields[MATCH_START + i].strip()
                if val:
                    match[name] = val
            if not match:
                continue

            change = {}
            for i, name in enumerate(FIELD_NAMES):
                val = fields[CHANGE_START + i].strip()
                if val:
                    change[name] = val

            msg_text = fields[MSG_TEXT_COL].strip()
            msg_field = fields[MSG_FIELD_COL].strip()

            rules.append(Rule(match, change, msg_field, msg_text))

    if not quiet:
        print("  Loaded %d rules from Rules.csv." % len(rules))
    return rules


def _clean(val: str) -> str:
    """Strip line breaks - a value spanning lines would corrupt the file."""
    return (val or '').replace('\r', ' ').replace('\n', ' ').strip()


def rule_to_row(rule: Rule, number: int) -> List[str]:
    """Serialize one rule to its 23-column CSV row."""
    row = [''] * TOTAL_COLS
    row[0] = str(number)
    for i, name in enumerate(FIELD_NAMES):
        row[MATCH_START + i] = _clean(rule.match_fields.get(name, ''))
        row[CHANGE_START + i] = _clean(rule.change_fields.get(name, ''))
    row[MSG_TEXT_COL] = _clean(rule.msg_text)
    row[MSG_FIELD_COL] = _clean(rule.msg_field)
    return row


def save_rules(rules_path: str, rules: List[Rule], backup: bool = True) -> None:
    """Write rules to Rules.csv atomically, keeping a .bak of the previous file.

    Rules are renumbered from 1 in list order. CRLF line endings and the
    original header are preserved so the file stays interchangeable with
    the VB.NET version.
    """
    directory = os.path.dirname(os.path.abspath(rules_path))
    os.makedirs(directory, exist_ok=True)
    tmp_path = rules_path + '.tmp'

    with open(tmp_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, lineterminator='\r\n')
        writer.writerow(HEADER)
        for i, rule in enumerate(rules, start=1):
            writer.writerow(rule_to_row(rule, i))

    if backup and os.path.exists(rules_path):
        try:
            shutil.copy2(rules_path, rules_path + '.bak')
        except Exception:
            pass

    os.replace(tmp_path, rules_path)


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_rule(rule: Rule) -> List[Tuple[str, str]]:
    """Check a single rule. Returns a list of (severity, message).

    Severity is "error" (the rule cannot work) or "warning".
    """
    issues: List[Tuple[str, str]] = []

    positives = {k: v for k, v in rule.match_fields.items() if not v.startswith('!')}
    negations = {k: v for k, v in rule.match_fields.items() if v.startswith('!')}

    if not rule.match_fields:
        issues.append(("error", "No match fields - the rule can never fire."))
    elif not positives:
        issues.append(("error",
                       "Only negated match fields - a rule needs at least one "
                       "positive match to fire."))

    if not rule.change_fields:
        issues.append(("warning", "No change fields - the rule matches but changes nothing."))

    for name, val in positives.items():
        if rule.change_fields.get(name, '').lower() == val.lower():
            issues.append(("warning", "%s is set to the same value it matches on." % name))

    icao = positives.get('ICAO', '')
    if icao and (len(icao) != 6 or any(c not in '0123456789abcdefABCDEF' for c in icao)):
        issues.append(("warning", "ICAO '%s' is not 6 hex characters." % icao))
    if rule.change_fields.get('ICAO', ''):
        issues.append(("warning", "Changing ICAO rewrites the record's key field."))

    for source, label in ((rule.match_fields, "match"), (rule.change_fields, "change")):
        op = source.get('OperatorICAO', '')
        op = op[1:] if op.startswith('!') else op
        if op and len(op) != 3:
            issues.append(("warning",
                           "OperatorICAO '%s' in %s is %d characters, expected 3."
                           % (op, label, len(op))))

    for name, val in negations.items():
        if len(val) == 1:
            issues.append(("error", "%s is just '!' with no value." % name))

    return issues


def find_shadowed(rules: List[Rule]) -> Dict[int, int]:
    """Find rules that can never be reached because an earlier rule always wins.

    Rules are first-match-wins, so rule j is unreachable when some earlier
    rule i has no negations and its positive match fields are a subset of
    j's, with identical values. Returns {shadowed_index: shadowing_index}.
    """
    shadowed: Dict[int, int] = {}

    def positives(rule: Rule) -> Dict[str, str]:
        return {k: v.lower() for k, v in rule.match_fields.items()
                if not v.startswith('!')}

    for j in range(len(rules)):
        pj = positives(rules[j])
        if not pj:
            continue
        for i in range(j):
            if any(v.startswith('!') for v in rules[i].match_fields.values()):
                continue  # negations make the earlier rule conditional - can't be sure
            pi = positives(rules[i])
            if not pi:
                continue
            if all(pj.get(k) == v for k, v in pi.items()):
                shadowed[j] = i
                break

    return shadowed


def validate_all(rules: List[Rule]) -> List[Tuple[int, str, str]]:
    """Validate every rule plus cross-rule shadowing.

    Returns a list of (rule_index, severity, message).
    """
    results: List[Tuple[int, str, str]] = []
    for i, rule in enumerate(rules):
        for severity, msg in validate_rule(rule):
            results.append((i, severity, msg))
    for j, i in sorted(find_shadowed(rules).items()):
        results.append((j, "warning",
                        "Unreachable - rule %d matches everything this rule does "
                        "and runs first." % (i + 1)))
    return results
