from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

from .schema import DATE_FIELDS, FIELD_ORDER

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN_NS, "r": DOC_REL_NS, "p": PKG_REL_NS}


def _column_name(reference):
    return reference[: re.search(r"\d", reference).start()]


def _string_from_node(node):
    return "".join(text.text or "" for text in node.findall(".//m:t", NS))


def _excel_date(value):
    try:
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date().isoformat()
    except (TypeError, ValueError):
        return str(value)


def read_workbook(source, with_headers=False):
    """Return the first worksheet as a list of ordered dictionaries."""
    raw = source.read_bytes() if isinstance(source, (str, Path)) else source
    if not isinstance(raw, bytes):
        raw = raw.read()

    with zipfile.ZipFile(BytesIO(raw)) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            strings = [_string_from_node(item) for item in shared.findall("m:si", NS)]

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relations.findall("p:Relationship", NS)
        }
        first_sheet = workbook.find("m:sheets/m:sheet", NS)
        if first_sheet is None:
            raise ValueError("The workbook does not contain a worksheet.")
        target = targets[first_sheet.attrib[f"{{{DOC_REL_NS}}}id"]].lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        root = ET.fromstring(archive.read(target))

        rows = []
        for row in root.findall(".//m:sheetData/m:row", NS):
            values = {}
            for cell in row.findall("m:c", NS):
                cell_type = cell.attrib.get("t")
                if cell_type == "inlineStr":
                    value = _string_from_node(cell)
                else:
                    node = cell.find("m:v", NS)
                    value = "" if node is None else node.text or ""
                    if cell_type == "s" and value:
                        value = strings[int(value)]
                values[_column_name(cell.attrib["r"])] = value
            rows.append(values)

    if not rows:
        return []
    ordered_columns = sorted(rows[0], key=_column_index)
    headers = [rows[0][column] for column in ordered_columns]
    records = []
    for row in rows[1:]:
        record = {header: row.get(column, "") for column, header in zip(ordered_columns, headers)}
        if not any(str(value).strip() for value in record.values()):
            continue
        for field in DATE_FIELDS:
            value = record.get(field, "")
            if value and str(value).replace(".", "", 1).isdigit():
                record[field] = _excel_date(value)
        records.append(record)
    return (records, headers) if with_headers else records


def _column_index(column):
    value = 0
    for character in column:
        value = value * 26 + ord(character) - ord("A") + 1
    return value


def _column_letter(index):
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _xml_text(value):
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def write_workbook(records, sheet_name="Orig_smartsheet"):
    """Create a compact XLSX workbook using inline strings and standard XML."""
    rows = []
    for row_number, values in enumerate([FIELD_ORDER] + records, start=1):
        if isinstance(values, list):
            fields = values
        else:
            fields = [values.get(field, "") for field in FIELD_ORDER]
        cells = []
        for column_number, value in enumerate(fields, start=1):
            reference = f"{_column_letter(column_number)}{row_number}"
            if value is None:
                value = ""
            style = " s=\"1\"" if row_number == 1 else ""
            cells.append(
                f'<c r="{reference}"{style} t="inlineStr"><is><t xml:space="preserve">'
                f"{_xml_text(value)}</t></is></c>"
            )
        rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')

    last_cell = f"{_column_letter(len(FIELD_ORDER))}{len(rows)}"
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{MAIN_NS}">'
        f'<dimension ref="A1:{last_cell}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" '
        'topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        f'<sheetData>{"".join(rows)}</sheetData>'
        f'<autoFilter ref="A1:{_column_letter(len(FIELD_ORDER))}1"/>'
        '</worksheet>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rId1" '
        f'Type="{DOC_REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<workbook xmlns="{MAIN_NS}" xmlns:r="{DOC_REL_NS}"><sheets>'
        f'<sheet name="{sheet_name}" sheetId="1" r:id="rId1"/>'
        '</sheets></workbook>'
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<Relationships xmlns="{PKG_REL_NS}"><Relationship Id="rId1" '
        f'Type="{DOC_REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/></Relationships>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<styleSheet xmlns="{MAIN_NS}"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font>'
        '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos"/></font></fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/>'
        '</fill><fill><patternFill patternType="solid"><fgColor rgb="FF147D72"/></patternFill></fill></fills>'
        '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>'
        '<cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
        '</styleSheet>'
    )
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/worksheets/sheet1.xml", sheet)
        archive.writestr("xl/styles.xml", styles)
    return output.getvalue()

