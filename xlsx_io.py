"""Small dependency-free XLSX adapter for flat, auditable tables.

The demo deliberately avoids a mandatory Excel dependency.  The adapter reads
the first (or named) worksheet and writes simple workbooks with styled headers.
It is not intended to replace a full spreadsheet engine.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Sequence
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _column_index(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference.upper())
    if not letters:
        raise ValueError(f"invalid XLSX cell reference: {reference}")
    result = 0
    for letter in letters.group(0):
        result = result * 26 + ord(letter) - ord("A") + 1
    return result - 1


def _column_name(index: int) -> str:
    result = ""
    value = index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _shared_strings(archive: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: List[str] = []
    for item in root.findall(f"{{{MAIN_NS}}}si"):
        strings.append("".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t")))
    return strings


def _worksheet_path(archive: zipfile.ZipFile, sheet_name: str | None) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in relationships.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    sheets = workbook.find(f"{{{MAIN_NS}}}sheets")
    if sheets is None or not list(sheets):
        raise ValueError("XLSX workbook contains no worksheets")
    selected = None
    for sheet in sheets:
        if sheet_name is None or sheet.attrib.get("name") == sheet_name:
            selected = sheet
            break
    if selected is None:
        raise ValueError(f"worksheet not found: {sheet_name}")
    relation_id = selected.attrib[f"{{{DOC_REL_NS}}}id"]
    target = targets[relation_id].replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return f"xl/{target}"


def read_xlsx_rows(path: Path, sheet_name: str | None = None) -> List[Dict[str, str]]:
    """Read a flat worksheet as dictionaries using its first row as headers."""
    with zipfile.ZipFile(path) as archive:
        shared = _shared_strings(archive)
        worksheet = ET.fromstring(archive.read(_worksheet_path(archive, sheet_name)))
    parsed_rows: List[List[str]] = []
    sheet_data = worksheet.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        return []
    for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
        values: List[str] = []
        for cell in row.findall(f"{{{MAIN_NS}}}c"):
            index = _column_index(cell.attrib.get("r", "A1"))
            while len(values) <= index:
                values.append("")
            kind = cell.attrib.get("t")
            value_node = cell.find(f"{{{MAIN_NS}}}v")
            if kind == "inlineStr":
                inline = cell.find(f"{{{MAIN_NS}}}is")
                value = "" if inline is None else "".join(
                    node.text or "" for node in inline.iter(f"{{{MAIN_NS}}}t")
                )
            elif value_node is None:
                value = ""
            elif kind == "s":
                value = shared[int(value_node.text or "0")]
            elif kind == "b":
                value = "true" if value_node.text == "1" else "false"
            else:
                value = value_node.text or ""
            values[index] = value
        parsed_rows.append(values)
    if not parsed_rows:
        return []
    headers = [header.strip() for header in parsed_rows[0]]
    return [
        {header: (row[index] if index < len(row) else "") for index, header in enumerate(headers) if header}
        for row in parsed_rows[1:]
        if any(cell != "" for cell in row)
    ]


def _cell_xml(reference: str, value: object, header: bool = False) -> str:
    style = ' s="1"' if header else ""
    if value is None:
        return f'<c r="{reference}"{style}/>'
    if isinstance(value, bool):
        return f'<c r="{reference}" t="b"{style}><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{reference}"{style}><v>{value}</v></c>'
    text = escape(str(value))
    preserve = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<c r="{reference}" t="inlineStr"{style}><is><t{preserve}>{text}</t></is></c>'


def _sheet_xml(headers: Sequence[str], rows: Iterable[Mapping[str, object]]) -> str:
    materialized = list(rows)
    widths = []
    for header in headers:
        width = max([len(str(header)), *(len(str(row.get(header, ""))) for row in materialized)])
        widths.append(min(max(width + 2, 10), 42))
    columns = "".join(
        f'<col min="{index + 1}" max="{index + 1}" width="{width}" customWidth="1"/>'
        for index, width in enumerate(widths)
    )
    xml_rows = []
    all_rows: List[Sequence[object]] = [list(headers)]
    all_rows.extend([[row.get(header, "") for header in headers] for row in materialized])
    for row_index, row in enumerate(all_rows, start=1):
        cells = "".join(
            _cell_xml(f"{_column_name(column_index)}{row_index}", value, row_index == 1)
            for column_index, value in enumerate(row)
        )
        xml_rows.append(f'<row r="{row_index}">{cells}</row>')
    last_cell = f"{_column_name(max(0, len(headers) - 1))}{max(1, len(all_rows))}"
    auto_filter = f'<autoFilter ref="A1:{last_cell}"/>' if headers else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{MAIN_NS}">'
        f'<dimension ref="A1:{last_cell}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
        'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<cols>{columns}</cols><sheetData>{"".join(xml_rows)}</sheetData>{auto_filter}'
        '</worksheet>'
    )


def write_xlsx(path: Path, sheets: Mapping[str, Sequence[Mapping[str, object]]]) -> None:
    """Write a simple multi-sheet XLSX workbook with filters and frozen headers."""
    if not sheets:
        raise ValueError("at least one worksheet is required")
    path.parent.mkdir(parents=True, exist_ok=True)
    names = []
    for raw_name in sheets:
        name = re.sub(r"[\\/*?:\[\]]", "_", raw_name)[:31] or "Sheet"
        if name in names:
            raise ValueError(f"duplicate XLSX worksheet name: {name}")
        names.append(name)
    content_overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, len(names) + 1)
    )
    workbook_sheets = "".join(
        f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, name in enumerate(names, start=1)
    )
    workbook_rels = "".join(
        f'<Relationship Id="rId{index}" Type="{DOC_REL_NS}/worksheet" '
        f'Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, len(names) + 1)
    )
    workbook_rels += (
        f'<Relationship Id="rId{len(names) + 1}" Type="{DOC_REL_NS}/styles" Target="styles.xml"/>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{MAIN_NS}"><fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>'
        '</fonts><fills count="3"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/>'
        '<bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            f'{content_overrides}<Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>',
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rId1" '
            f'Type="{DOC_REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{MAIN_NS}" xmlns:r="{DOC_REL_NS}"><sheets>{workbook_sheets}</sheets></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PKG_REL_NS}">{workbook_rels}</Relationships>',
        )
        archive.writestr("xl/styles.xml", styles)
        for index, rows in enumerate(sheets.values(), start=1):
            headers = list(rows[0].keys()) if rows else ["status"]
            materialized = rows if rows else [{"status": "no rows"}]
            archive.writestr(f"xl/worksheets/sheet{index}.xml", _sheet_xml(headers, materialized))
