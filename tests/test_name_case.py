"""Tests for vrs_updater.name_case. Run with:  python -m unittest discover tests"""

import unittest

from vrs_updater.name_case import proper_case
from vrs_updater.utils import title_case


class OrganisationNames(unittest.TestCase):
    """Real owner names from the FAA registry."""

    def check(self, raw, expected):
        self.assertEqual(proper_case(raw), expected)

    def test_short_non_words_are_acronyms(self):
        self.check("UMB BANK NA TRUSTEE", "UMB Bank NA Trustee")
        self.check("PHI AVIATION LLC", "PHI Aviation LLC")
        self.check("CSC DELAWARE TRUST CO TRUSTEE", "CSC Delaware Trust Co Trustee")
        self.check("TJ WATERSENSE LLC", "TJ Watersense LLC")

    def test_short_words_and_names_are_not(self):
        self.check("SKY HIGH AIR LLC", "Sky High Air LLC")
        self.check("FLY JET FUN INC", "Fly Jet Fun Inc")
        self.check("LEE FOX AND COX LTD", "Lee Fox and Cox Ltd")
        self.check("DEL RIO AVIATION INC", "Del Rio Aviation Inc")

    def test_words_without_vowels_are_acronyms(self):
        self.check("TVPX AIRCRAFT SOLUTIONS INC TRUSTEE",
                   "TVPX Aircraft Solutions Inc Trustee")
        self.check("DRONESCAPE PLLC", "Dronescape PLLC")

    def test_abbreviations_stay_title_case(self):
        self.check("ACME MFG CO", "Acme Mfg Co")
        self.check("ST LOUIS FLT SVC MGMT", "St Louis Flt Svc Mgmt")

    def test_known_long_acronyms(self):
        self.check("USDA FOREST SERVICE FEPP", "USDA Forest Service FEPP")
        self.check("NASA LANGLEY RESEARCH CENTER", "NASA Langley Research Center")
        self.check("US ARMY", "US Army")
        self.check("DIRECTIONAL CAPITAL LLC DBA", "Directional Capital LLC DBA")

    def test_roman_numerals(self):
        self.check("BATELEUR EAGLE HOLDINGS II LLC", "Bateleur Eagle Holdings II LLC")
        self.check("MARK IV AVIATION LLC", "Mark IV Aviation LLC")

    def test_letters_and_digits(self):
        self.check("N100GZ INC TRUSTEE", "N100GZ Inc Trustee")
        self.check("CIRRUS SR22T 2190 INC", "Cirrus SR22T 2190 Inc")
        self.check("KENTEM100 LLC", "Kentem100 LLC")
        self.check("1ST AVENUE LLC", "1st Avenue LLC")
        self.check("426TH AERO SQUADRON LLC", "426th Aero Squadron LLC")

    def test_punctuation(self):
        self.check("STUEVE'S DAIRY", "Stueve's Dairy")
        self.check("O'NEAL AIR", "O'Neal Air")
        self.check("AT&T SERVICES INC", "AT&T Services Inc")
        self.check("D&R LLC", "D&R LLC")
        self.check("TRI-STATE CO-OP", "Tri-State Co-Op")
        self.check("COULSON AVIATION (USA) INC", "Coulson Aviation (USA) Inc")

    def test_small_words_lower_case_in_the_middle_only(self):
        self.check("BANK OF UTAH", "Bank of Utah")
        self.check("THE BOEING COMPANY", "The Boeing Company")
        self.check("LIORE ET OLIVIER", "Liore et Olivier")

    def test_special_forms(self):
        self.check("APPARATEBAU NABERN GMBH", "Apparatebau Nabern GmbH")

    def test_mixed_case_source_keeps_its_capitals(self):
        self.check("Lufthansa AG", "Lufthansa AG")
        self.check("Air New Zealand", "Air New Zealand")

    def test_empty_and_short(self):
        self.check("", "")
        self.check("A", "A")


class PersonNames(unittest.TestCase):

    def check(self, raw, expected):
        self.assertEqual(proper_case(raw, person=True), expected)

    def test_short_names_are_not_acronyms(self):
        self.check("SMITH DAN V", "Smith Dan V")
        self.check("ERB PAUL R", "Erb Paul R")
        self.check("CUI JIN", "Cui Jin")
        self.check("NG NAM H", "Ng Nam H")

    def test_suffixes_and_initials(self):
        self.check("ASKEW STELL BLAKE III", "Askew Stell Blake III")
        self.check("VANDERMOLEN LOUIS A MD", "Vandermolen Louis A MD")
        self.check("WRIGHT TJ", "Wright TJ")
        self.check("COUTCHES ROBERT HERCULES DBA", "Coutches Robert Hercules DBA")


class TitleCaseWrapper(unittest.TestCase):

    def test_title_case_uses_proper_case(self):
        self.assertEqual(title_case("UMB BANK NA"), "UMB Bank NA")
        self.assertEqual(title_case("CUI JIN", person=True), "Cui Jin")


if __name__ == "__main__":
    unittest.main()
