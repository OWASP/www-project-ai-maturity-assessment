#!/usr/bin/env python3
"""Validate the answer scale and scoring formulas in released AIMA toolkits."""

import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

TOOLKITS = (
    Path("releases/V1/OWASP-AIMA-Toolkit_V1.xlsx"),
    Path("releases/V1/OWASP-AIMA-Toolkit_V1.0.1.xlsx"),
    Path("source/v1/AIMA_Assessment_Toolkit.xlsx"),
    Path("source/latest/AIMA_Assessment_Toolkit.xlsx"),
)
SCORE_ROWS = (
    7,
    14,
    21,
    29,
    36,
    43,
    51,
    58,
    65,
    73,
    80,
    87,
    95,
    102,
    109,
    117,
    124,
    131,
    139,
    146,
    153,
    161,
    168,
    175,
)
ANSWER_RANGES = " ".join(f"E{row}:E{row + 5}" for row in SCORE_ROWS)
ANSWER_LIST = '"To evaluate,0,1,2,3"'


def worksheet_xml(archive: ZipFile, sheet_name: str) -> ET.Element:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relation_targets = {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }

    for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
        if sheet.attrib["name"] == sheet_name:
            relation_id = sheet.attrib[f"{{{REL_NS}}}id"]
            target = relation_targets[relation_id].lstrip("/")
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            return ET.fromstring(archive.read(target))

    raise AssertionError(f"missing worksheet: {sheet_name}")


def cell_formulas(sheet: ET.Element) -> dict[str, str]:
    return {
        cell.attrib["r"]: formula.text or ""
        for cell in sheet.findall(f".//{{{MAIN_NS}}}c")
        if (formula := cell.find(f"{{{MAIN_NS}}}f")) is not None
    }


def validate_toolkit(path: Path) -> None:
    with ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_names = [
            sheet.attrib["name"]
            for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet")
        ]
        assert sheet_names == [
            "Introduction",
            "Assessment Questions",
            "Results",
            "Graphical visualization",
        ], f"{path}: unexpected sheets: {sheet_names}"

        questions = worksheet_xml(archive, "Assessment Questions")
        validations = questions.findall(f".//{{{MAIN_NS}}}dataValidation")
        assert len(validations) == 1, f"{path}: expected one answer validation rule"
        validation = validations[0]
        formula = validation.find(f"{{{MAIN_NS}}}formula1")
        assert formula is not None and formula.text == ANSWER_LIST, (
            f"{path}: answer list must be {ANSWER_LIST}, got "
            f"{None if formula is None else formula.text!r}"
        )
        assert validation.attrib["sqref"] == ANSWER_RANGES, (
            f"{path}: answer validation does not cover every question"
        )

        question_formulas = cell_formulas(questions)
        for row in SCORE_ROWS:
            expected = f"SUM(E{row}:E{row + 5})/6"
            assert question_formulas.get(f"F{row}") == expected, (
                f"{path}: F{row} must be {expected}"
            )

        results = cell_formulas(worksheet_xml(archive, "Results"))
        for result_row, score_row in enumerate(SCORE_ROWS, start=4):
            expected = f"'Assessment Questions'!F{score_row}"
            assert results.get(f"D{result_row}") == expected, (
                f"{path}: D{result_row} must be {expected}"
            )

        maximum_score = sum([3] * 6) / 6
        assert maximum_score == 3


def main() -> None:
    for toolkit in TOOLKITS:
        validate_toolkit(toolkit)
        print(f"Validated {toolkit}")


if __name__ == "__main__":
    main()
