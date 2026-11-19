"""
contains classes for the segment layouts of a MIG: the data elements and data element groups of each segment
"""

import json
from pathlib import Path
from typing import Any

from efoli import EdifactFormat
from loguru import logger
from pydantic import BaseModel


class Datenelement(BaseModel):
    """
    class for a data element in the layout of a segment, e.g. (UTILMD, NAD):
    {
        "bezeichnung": "3039",
        "name": "Beteiligter, Identifikation",
        "standard_status": "M",
        "standard_format": "an..35",
        "bdew_status": "M",
        "bdew_format": "an..35",
        "anwendung": "MP-ID"
    }
    The texts are those of the document, with its line breaks and tabs.
    """

    bezeichnung: str
    name: str
    standard_status: str
    standard_format: str
    bdew_status: str
    bdew_format: str
    anwendung: str  #: the column 'Anwendung / Bemerkung', which also lists the codes of the data element


class Datenelementgruppe(BaseModel):
    """
    class for a data element group (composite) in the layout of a segment, e.g. (UTILMD, NAD):
    {
        "bezeichnung": "C082",
        "name": "Identifikation des Beteiligten",
        "standard_status": "C",
        "bdew_status": "R",
        "anwendung": "",
        "datenelemente": [...]
    }
    """

    bezeichnung: str
    name: str
    standard_status: str
    bdew_status: str
    anwendung: str
    datenelemente: list[Datenelement]


class Segmentlayout(BaseModel):
    """
    class for the layout of one segment. zaehler, nr and bezeichnung are those of the segment in the
    Nachrichtenstruktur, so that nr connects a layout with its NachrichtenstrukturZeile.
    """

    zaehler: str
    nr: str
    bezeichnung: str
    elemente: list[Datenelement | Datenelementgruppe]  #: in the order of the segment


class Segmentlayouts(BaseModel):
    """
    class for the segment layouts of a MIG, in the order of the document
    """

    segmentlayouts: list[Segmentlayout]

    def to_json(self, message_type: EdifactFormat, output_dir: Path) -> dict[str, Any]:
        """
        writes the Segmentlayouts as json
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        file_path = output_dir.joinpath("segmentlayouts.json")
        structured_json = self.model_dump()
        with open(file_path, "w", encoding="utf-8") as json_file:
            json.dump(structured_json, json_file, indent=4)
        logger.info(f"Wrote Segmentlayouts for {message_type} to {file_path}")
        return structured_json
