"""
Test the segment layouts.
"""

import json
import re

from efoli import EdifactFormat

from migmose.mig.segmentlayout import Datenelement, Datenelementgruppe, Segmentlayouts
from migmose.parsing import parse_segmentlayouts
from unittests import path_to_test_fv2310, path_to_test_fv2510, path_to_test_fv2610

ordchg = path_to_test_fv2310 / "ORDCHGMIG-informatorischeLesefassung1.1_99991231_20231001.docx"
contrl = path_to_test_fv2610 / "CONTRLMIG-informatorischeLesefassung2.0b_99991231_20221001.docx"
contrl_with_letters = path_to_test_fv2510 / "CONTRLMIG-informatorischeLesefassung2.0b_20251210_20221001.docx"


class TestSegmentlayout:
    """
    A class with pytest unit tests for segment layouts.
    """

    def test_one_layout_per_segment(self):
        segmentlayouts = parse_segmentlayouts(ordchg).segmentlayouts
        assert [(layout.zaehler, layout.nr, layout.bezeichnung) for layout in segmentlayouts] == [
            ("0010", "00001", "UNH"),
            ("0020", "00002", "BGM"),
            ("0030", "00003", "DTM"),
            ("0100", "00004", "RFF"),
            ("0100", "00005", "RFF"),
            ("0100", "00006", "RFF"),
            ("0160", "00007", "NAD"),
            ("0260", "00008", "CTA"),
            ("0270", "00009", "COM"),
            ("0160", "00010", "NAD"),
            ("2380", "00011", "UNS"),
            ("2450", "00012", "UNT"),
        ]

    def test_data_elements_and_groups(self):
        nad = parse_segmentlayouts(ordchg).segmentlayouts[6]
        qualifier, identifikation = nad.elemente
        assert qualifier == Datenelement(
            bezeichnung="3035",
            name="Beteiligter, Qualifier",
            standard_status="M",
            standard_format="an..3",
            bdew_status="M",
            bdew_format="an..3",
            anwendung="\tMS\tDokumenten-/Nachrichtenaussteller bzw. -\nabsender",
        )
        assert isinstance(identifikation, Datenelementgruppe)
        assert (identifikation.bezeichnung, identifikation.name) == ("C082", "Identifikation des Beteiligten")
        assert (identifikation.standard_status, identifikation.bdew_status) == ("C", "R")
        assert [datenelement.bezeichnung for datenelement in identifikation.datenelemente] == ["3039", "1131", "3055"]
        assert identifikation.datenelemente[1].bdew_status == "N"
        assert identifikation.datenelemente[1].bdew_format == ""

    def test_data_element_after_a_group_belongs_to_the_segment(self):
        unh, bgm = parse_segmentlayouts(ordchg).segmentlayouts[:2]
        assert [(type(element), element.bezeichnung) for element in unh.elemente] == [
            (Datenelement, "0062"),
            (Datenelementgruppe, "S009"),
        ]
        assert [(type(element), element.bezeichnung) for element in bgm.elemente] == [
            (Datenelementgruppe, "C002"),
            (Datenelementgruppe, "C106"),
            (Datenelement, "1225"),
        ]

    def test_segment_numbers_with_a_letter(self):
        segmentlayouts = parse_segmentlayouts(contrl_with_letters).segmentlayouts
        assert [layout.nr for layout in segmentlayouts] == ["C0001", "C0002", "C0003", "C0004", "C0005", "C0006"]

    def test_row_that_continues_anwendung(self):
        """the codes of a data element go on in rows that fill nothing but 'Anwendung / Bemerkung'"""
        uci = parse_segmentlayouts(contrl).segmentlayouts[1]
        syntaxfehler = uci.elemente[4]
        assert isinstance(syntaxfehler, Datenelement)
        assert syntaxfehler.bezeichnung == "0085"
        codes = re.findall(r"(?m)^\t(\w+)\t", syntaxfehler.anwendung)
        assert codes == ["2", "7", "12", "13", "16", "20", "21", "23", "25", "26", "28", "29", "32"]

    def test_segmentlayouts(self, snapshot):
        assert parse_segmentlayouts(ordchg).model_dump() == snapshot

    def test_to_json(self, tmp_path):
        segmentlayouts = parse_segmentlayouts(contrl)
        segmentlayouts.to_json(EdifactFormat.CONTRL, tmp_path)
        with open(tmp_path / "segmentlayouts.json", encoding="utf-8") as file:
            content = json.load(file)
        assert Segmentlayouts.model_validate(content) == segmentlayouts
