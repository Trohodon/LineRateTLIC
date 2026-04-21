from __future__ import annotations

from dataclasses import dataclass
from pathlib import PurePosixPath
import re
import zipfile
from xml.etree import ElementTree as ET


MAIN_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")


@dataclass(frozen=True)
class XlsxCell:
    ref: str
    value: str | None
    formula: str | None = None


def _column_index(cell_ref: str) -> int:
    match = CELL_REF_RE.fullmatch(cell_ref)
    if match is None:
        raise ValueError(f"Invalid cell reference '{cell_ref}'.")

    letters = match.group(1)
    index = 0
    for letter in letters:
        index = index * 26 + (ord(letter) - ord("A") + 1)
    return index - 1


def _load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: list[str] = []

    for si in root.findall("main:si", MAIN_NS):
        text = "".join(node.text or "" for node in si.iterfind(".//main:t", MAIN_NS))
        values.append(text)

    return values


def _sheet_targets(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    rel_map = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall("rel:Relationship", REL_NS)
    }

    targets: dict[str, str] = {}
    for sheet in workbook_root.find("main:sheets", MAIN_NS) or []:
        name = sheet.attrib["name"]
        rel_id = sheet.attrib[f"{{{OFFICE_REL}}}id"]
        target = rel_map[rel_id].lstrip("/")
        targets[name] = target if target.startswith("xl/") else str(PurePosixPath("xl").joinpath(PurePosixPath(target)))

    return targets


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> str | None:
    cell_type = cell.attrib.get("t")
    value_node = cell.find("main:v", MAIN_NS)
    inline_node = cell.find("main:is", MAIN_NS)

    if cell_type == "s" and value_node is not None:
        return shared_strings[int(value_node.text)]

    if cell_type == "inlineStr" and inline_node is not None:
        return "".join(node.text or "" for node in inline_node.iterfind(".//main:t", MAIN_NS))

    if value_node is not None:
        return value_node.text

    return None


def read_sheet_cells(filepath: str, sheet_name: str) -> list[list[XlsxCell]]:
    with zipfile.ZipFile(filepath) as archive:
        shared_strings = _load_shared_strings(archive)
        targets = _sheet_targets(archive)

        if sheet_name not in targets:
            raise KeyError(f"Sheet '{sheet_name}' was not found in '{filepath}'.")

        sheet_root = ET.fromstring(archive.read(targets[sheet_name]))
        sheet_data = sheet_root.find("main:sheetData", MAIN_NS)
        if sheet_data is None:
            return []

        rows: list[list[XlsxCell]] = []
        for row in sheet_data.findall("main:row", MAIN_NS):
            cells: list[XlsxCell] = []
            for cell in row.findall("main:c", MAIN_NS):
                formula_node = cell.find("main:f", MAIN_NS)
                cells.append(
                    XlsxCell(
                        ref=cell.attrib["r"],
                        value=_cell_value(cell, shared_strings),
                        formula=formula_node.text if formula_node is not None else None,
                    )
                )
            rows.append(cells)

        return rows


def read_sheet_table(filepath: str, sheet_name: str) -> list[list[str | None]]:
    rows = read_sheet_cells(filepath, sheet_name)
    table: list[list[str | None]] = []

    for row_cells in rows:
        if not row_cells:
            table.append([])
            continue

        row_size = max(_column_index(cell.ref) for cell in row_cells) + 1
        row_values: list[str | None] = [None] * row_size
        for cell in row_cells:
            row_values[_column_index(cell.ref)] = cell.value
        table.append(row_values)

    return table


def list_sheet_names(filepath: str) -> list[str]:
    with zipfile.ZipFile(filepath) as archive:
        return list(_sheet_targets(archive).keys())
