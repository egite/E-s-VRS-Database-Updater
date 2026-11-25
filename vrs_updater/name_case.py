"""
Proper casing for owner, operator and manufacturer names.

Registries publish names in capitals ("SKY HIGH AVN LLC", "UMB BANK NA TR").
Plain title case turns every acronym into a word ("Umb Bank Na"), so each
word is checked on its own and kept in capitals when it looks like an
acronym:

- it is in ALWAYS_UPPER (LLC, DBA, NASA, US, ...);
- it mixes letters and digits (N123AB, RV7A, 3M), unless it is an ordinal
  such as 1st or 22nd;
- it has two or three letters and is not a word or name in COMMON_WORDS
  ("PHI", "UMB", "TJ");
- it has no vowels and is not in COMMON_WORDS ("TVPX", "PLLC", "BNSF").

Abbreviations that are not acronyms (Co, Jr, St, Mfg, Mgmt) are listed in
KEEP_TITLE and stay in title case even though the rules above would catch
them.

Names of people ("SMITH DAN V", "CUI JIN", "ERB PAUL R") are full of short
names that are not English words, so person=True turns the last two rules
off. Only PERSON_SUFFIXES (III, MD, DDS) and two-letter initials with no
vowels ("WRIGHT TJ") stay in capitals.

The lists were tuned against the owner names in the FAA registry: to add a
word, put it in lower case in the matching set below.
"""

import re

# ---------------------------------------------------------------------------
# Word lists (all lower case)
# ---------------------------------------------------------------------------

# Always written in capitals, even when they look like a word ("US") or are
# longer than three letters and contain vowels ("NASA").
ALWAYS_UPPER = frozenset("""
    llc llp lllp pllc lc lp na pc dba
    us usa uk
    nasa noaa usda usaf usfs aphis fepp cavu
    faa ems atp
""".split())

# Abbreviations that are written in title case even though they are not
# words or have no vowels.
KEEP_TITLE = frozenset("""
    co cos corp inc ltd pty cie
    jr sr dr mr mrs mt st ste wm tr
    mc mac
    svc svcs mfg mgmt sch flt blvd ctr cntr dept bldg hwy pkwy
    natl intl assn assoc bros hldgs twp mtn govt trvl trns ave ent div srvs sys
""".split())

# In person=True mode, the only short words kept in capitals: generational
# numerals, degrees, and the FAA's "no middle name" marker.
PERSON_SUFFIXES = frozenset("""
    ii iii iv vi vii viii ix
    md dds dmd dvm phd cpa nmn
""".split())

# Lower-cased when they are neither the first nor the last word
# ("Bank of Utah", "Ponce de Leon", "Liore et Olivier").
LOWER_IN_MIDDLE = frozenset(("of", "and", "the", "de", "la", "und", "et"))

# Words with a fixed mixed-case spelling.
SPECIAL_FORMS = {"gmbh": "GmbH"}

# Two- and three-letter words and names that are not acronyms. Anything of
# that length that is missing from here is written in capitals.
COMMON_WORDS = frozenset("""
    a i
    ad ag ah am an as at aw ax be by do eh go ha he hi ho if in is it la lo
    ma me my no of oh ok on op or ow ox oz pa pi re so to uh um up we ya ye
    yo

    ace act add ado aft age ago aha aid ail aim air ale all amp and ant any
    ape apt arc are ark arm art ash ask asp ate auk awe awl axe aye bad bag
    ban bar bat bay bed bee beg bet bib bid big bin bio bit boa bob bog boo
    bop bow box boy bra bud bug bum bun bus but buy bye she cab cad cam can cap
    car cat caw chi cob cod cog con coo cop cot cow coy cry cub cud cue cup
    cur cut dab dad dam day den dew did die dig dim din dip doc doe dog don
    dot dry dub dud due dug duo dye ear eat ebb eel egg ego eke elf elk elm
    emu end eon era ere erg eve ewe eye fab fad fan far fat fax fed fee fen
    few fez fib fig fin fir fit fix flu fly foe fog for fox fro fry fun fur
    gab gag gal gap gas gay gee gel gem get gig gin gnu gob god got gum gun
    gut guy gym had hag ham has hat haw hay hem hen her hew hex hey hid him
    hip his hit hob hoe hog hop hot how hub hue hug hum hut ice icy ill imp
    ink inn ion ire irk its ivy jab jag jam jar jaw jay jet jig job jog jot
    joy jug jut keg ken key kid kin kit lab lad lag lap law lax lay lea led
    leg let lid lie lip lit lob log loo lot low lug lye mad man map mar mat
    maw max may men met mid mix mob mod mom moo mop mow mud mug mum nab nag
    nap nay neb net new nib nil nip nit nod nor not now nub nun nut oak oar
    oat odd ode off oft ohm oil old one opt orb ore our out ova owe owl own
    pad pal pan pap par pat paw pay pea peg pen pep per pet pew pie pig pin
    pit ply pod pop pot pow pro pry pub pug pun pup pus put rad rag ram ran
    rap rat raw ray red rib rid rig rim rip rob rod roe rot row rub rug rum
    run rut rye sad sag sap sat saw sax say sea sec see set sew sex shy sin
    sip sir sis sit six ski sky sly sob sod son sop sow soy spa spy sty sub
    sue sum sun sup tab tad tag tan tap tar tax tea tee ten the thy tic tie
    tin tip toe tog ton too top tot tow toy try tub tug tun tux two urn use
    van vat vet vex via vie vim vow wad wag war was wax way web wed wee wet
    who why wig win wit woe wok won woo wow wry yak yam yap yaw yea yen yes
    yet yew yin yon you yum zag zap zed zen zig zip zit zoo

    aer aero eco geo neo ops uni tri gen med lux neu bon tex rio fe paz
    biz blu com evo lil mag mas oso rev tec uno vee

    al bo cy di ed jo ty
    abe ada adi ali amy ann ari art ava avi bea ben bev bob boe cal dan deb
    dee del dom don dot eli ely eva eve fay flo gil gus guy hal ian ida ike ira
    isa jan jax jay jed jen jim joe jon kai kay ken kim kip kit lea lee len
    leo les liz lou luc lyn mae max meg mel mia mo moe nan ned nel noe ole
    pam pat peg rae ray reg rex rod ron roy sal sam sid sol sue tad ted tim
    tod tom toy uwe val vic wes zac zak
    ahn ash bae bai cao cho cox dai day dix doe dow dye ely eng fan fox
    fry gay han hay hoy hsu kan kay kee lam lau law lim lin liu loy lum luo
    moy nix orr ott pak pei poe rao roe rue ryu seo shi sim sun tam tan teo
    tye vos wan wei woo xie yan yao yee yim yip yoo yun zhu
    du hu ko le li lu ly ng su vu wu xu yi yu

    da das de del den der des di dos du el la las le les los san van von
""".split())

_VOWELS = frozenset("aeiouy")

# A chunk between spaces is split into pieces at these characters, which are
# kept. Each piece is cased on its own ("TRI-STATE" -> "Tri-State").
_RE_SEPARATORS = re.compile(r"([-/.,()+\"])")
_RE_ORDINAL = re.compile(r"^\d+(st|nd|rd|th|s)$")
_RE_LETTER_RUN = re.compile(r"[^\W\d_]+")


# ---------------------------------------------------------------------------
# Casing
# ---------------------------------------------------------------------------

def _case_letters(piece: str, person: bool) -> str:
    """Case a piece with no separators in it."""
    low = piece.lower()
    if not any(ch.isalpha() for ch in low):
        return piece
    if any(ch.isdigit() for ch in low):
        if _RE_ORDINAL.match(low):
            return low
        # N100GZ, SR22T, AV8R in capitals; KENTEM100 and M2TIGER have
        # letter runs long enough to be words: Kentem100, M2Tiger.
        return _RE_LETTER_RUN.sub(
            lambda m: (m.group().capitalize()
                       if len(m.group()) > 3
                       and any(ch in _VOWELS for ch in m.group())
                       else m.group().upper()),
            low)
    if not low.isalpha():
        return low[:1].upper() + low[1:]
    if low in SPECIAL_FORMS:
        return SPECIAL_FORMS[low]
    if low in ALWAYS_UPPER:
        return piece.upper()
    if low in KEEP_TITLE or low in COMMON_WORDS:
        return low.capitalize()
    if len(low) == 1:
        return low.upper()
    has_vowel = any(ch in _VOWELS for ch in low)
    if person:
        if low in PERSON_SUFFIXES or (len(low) == 2 and not has_vowel):
            return piece.upper()
        return low.capitalize()
    if len(low) <= 3 or not has_vowel:
        return piece.upper()
    return low.capitalize()


def _case_apostrophe(piece: str, person: bool) -> str:
    """Case a piece that may contain apostrophes ("O'BRIEN", "JOHN'S")."""
    if "'" not in piece:
        return _case_letters(piece, person)
    head, _, tail = piece.partition("'")
    if not head:                      # leading apostrophe: "'N"
        return "'" + _case_apostrophe(tail, person)
    if len(tail) <= 2:                # JOHN'S -> John's, CAN'T -> Can't
        return _case_letters(head, person) + "'" + tail.lower()
    return _case_letters(head, person) + "'" + _case_apostrophe(tail, person)


def _case_chunk(chunk: str, person: bool) -> str:
    """Case one space-separated chunk of a name."""
    if "&" in chunk:
        parts = chunk.split("&")
        # "AT&T", "R&D", "B&B": short letter groups joined by & are initials.
        if any(parts) and all(p.isalpha() and len(p) <= 3 for p in parts if p):
            return chunk.upper()
        return "&".join(_case_chunk(p, person) for p in parts)
    pieces = _RE_SEPARATORS.split(chunk)
    return "".join(p if (i % 2) else _case_apostrophe(p, person)
                   for i, p in enumerate(pieces))


def proper_case(s: str, person: bool = False) -> str:
    """Return a registry name in proper case with acronyms in capitals.

    person=True is for the name of a person, where short words that are
    not English words are names rather than acronyms.

    Names that already contain lower-case letters keep any word that is
    written entirely in capitals, since the source has marked it as an
    acronym ("Lufthansa AG").
    """
    if not s or len(s) <= 1:
        return s
    mixed = any(ch.islower() for ch in s)
    words = s.split(" ")
    last = len(words) - 1
    out = []
    for i, word in enumerate(words):
        if mixed and len(word) > 1 and word.isupper():
            out.append(word)
            continue
        cased = _case_chunk(word, person)
        if 0 < i < last and cased.lower() in LOWER_IN_MIDDLE:
            cased = cased.lower()
        out.append(cased)
    return " ".join(out)
