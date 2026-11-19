"""
Test the command line interface.
"""

from pathlib import Path

from click.testing import CliRunner

from migmose.__main__ import main
from unittests import path_to_test_edi_energy_mirror_repo


def _run(output_dir: Path, *file_types: str) -> set[str]:
    """runs migmose for the CONTRL of FV2610 and returns the names of the files it writes"""
    arguments = [
        "-eemp",
        str(path_to_test_edi_energy_mirror_repo),
        "-o",
        str(output_dir),
        "-fv",
        "FV2610",
        "-mf",
        "CONTRL",
    ]
    for file_type in file_types:
        arguments += ["-ft", file_type]
    result = CliRunner().invoke(main, arguments)
    assert result.exit_code == 0, result.output
    return {path.name for path in (output_dir / "FV2610" / "CONTRL").iterdir()}


class TestMain:
    """
    A class with pytest unit tests for the command line interface.
    """

    def test_default_file_types_are_those_of_the_nachrichtenstruktur(self, tmp_path):
        assert _run(tmp_path) == {
            "nachrichtenstruktur.csv",
            "nested_nachrichtenstruktur.json",
            "reduced_nested_nachrichtenstruktur.json",
            "CONTRL2.0b.tree",
        }

    def test_sgh_json(self, tmp_path):
        assert _run(tmp_path, "sgh_json") == {"CONTRL.sgh.json"}

    def test_segmentlayout_json(self, tmp_path):
        assert _run(tmp_path, "segmentlayout_json") == {"segmentlayouts.json"}

    def test_segmentlayout_json_next_to_the_nachrichtenstruktur(self, tmp_path):
        assert _run(tmp_path, "csv", "segmentlayout_json") == {"nachrichtenstruktur.csv", "segmentlayouts.json"}
