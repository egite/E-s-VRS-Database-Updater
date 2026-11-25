# v2.5.0

Operator and manufacturer names keep their acronyms in capitals.

**Downloads:** `VRS_Database_Updater_x64.exe` and `Rules.csv`.

Upgrading is just replacing the executable. Your working folder, `Rules.csv`
and `settings.json` carry over unchanged.

---

## Acronyms in names stay in capitals

The registers publish names in capitals, and the program used to turn every
word into title case, so "UMB BANK NA" became "Umb Bank Na". Only LLC, LLP
and a few Roman numerals were put back in capitals. Now each word is checked
on its own and stays in capitals when it looks like an acronym:

- **Two or three letters that are not a word or a name:** UMB Bank NA, PHI
  Aviation, CSC Delaware Trust Co. Short words and names such as Sky, Fly, Jet,
  Lee and Fox stay in title case.
- **No vowels:** TVPX, PLLC, BNSF.
- **Letters mixed with digits:** N100GZ, SR22T, RV7A. Ordinals are written
  1st and 426th rather than 1St and 426Th.
- **Known acronyms:** LLC, DBA, US, USA, NASA, USDA, APHIS and others.

Abbreviations that aren't acronyms, such as Co, Jr, St, Mfg and Mgmt, stay in
title case.

**FAA owners who are people are cased as names.** Short names that aren't
English words, such as Erb, Cyr or Jin, would otherwise look like acronyms, so
for owners the FAA registers as individuals only suffixes (II, III, MD, DDS),
two-letter initials and DBA stay in capitals.

**Other fixes:**
- **Apostrophes:** Stueve's rather than Stueve'S, O'Neal.
- **Ampersands:** AT&T rather than At&T.
- **Spellings:** GmbH, Co-Op, and "et" or "und" in lower case inside a name
  (Liore et Olivier).
- **Names that already arrive in mixed case** keep the words they write in
  capitals (Lufthansa AG).

Because of this, the first run of v2.5.0 updates the operator of many aircraft
in VRS, almost always just its capitalization. `Rules.csv` matching ignores
case, so your rules apply as before.

## Fixes

**The exit countdown stays on its own line.** After a **Run on Start** run, the
countdown to the program closing crept up onto the separator line above it
with every tick. It now stays on its own line and is shown in red.
