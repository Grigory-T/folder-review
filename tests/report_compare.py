"""Compare the existing report view, explicitly projecting out the new state extension."""

import re
import xml.etree.ElementTree as ET
import zipfile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def normalized_workbook(path):
    """Compare every ZIP member, cell type, style, formula and layout.

    The only omissions are access-time values, scan start/end timestamps,
    and workbook creation/modification timestamps. Independent scans may
    legitimately produce different values for these fields.
    """
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        headers = list(
            next(
                workbook["folder_review"].iter_rows(
                    min_row=6, max_row=6, values_only=True
                )
            )
        )
        access_column = get_column_letter(headers.index("Accessed at") + 1)
        status = list(workbook["scan_status"].iter_rows(values_only=True))
        volatile_status_rows = {
            index
            for index, row in enumerate(status, 1)
            if row and row[0] in ("Started at", "Completed at")
        }
        status_values = {}
        for row in status[1:]:
            if not row or row[0] is None:
                break
            if not row[0].startswith("Comparison state "):
                status_values[row[0]] = row[1]
        rows = sum(1 for _ in workbook["folder_review"].iter_rows(min_row=7))
    finally:
        workbook.close()
    result = {}
    has_state = "_folder_review_state" in workbook.sheetnames
    with zipfile.ZipFile(path) as archive:
        assert archive.testzip() is None
        for name in archive.namelist():
            if has_state and name == "xl/worksheets/sheet4.xml":
                continue
            data = archive.read(name)
            if name in (
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "docProps/core.xml",
            ):
                tree = ET.fromstring(data)
                if name == "docProps/core.xml":
                    for element in tree:
                        if element.tag.rsplit("}", 1)[-1] in ("created", "modified"):
                            element.text = "VOLATILE"
                else:
                    for row in tree.findall("s:sheetData/s:row", NS):
                        number = int(row.attrib["r"])
                        for cell in row.findall("s:c", NS):
                            if (
                                name.endswith("sheet1.xml")
                                and number >= 7
                                and cell.attrib["r"] == f"{access_column}{number}"
                            ) or (
                                name.endswith("sheet2.xml")
                                and number in volatile_status_rows
                                and cell.attrib["r"] == f"B{number}"
                            ):
                                for value in cell:
                                    value.text = "VOLATILE"
                data = ET.tostring(tree)
            if name.endswith((".xml", ".rels")):
                tree = ET.fromstring(data)
                if has_state:
                    if name == "xl/workbook.xml":
                        sheets = tree.find("s:sheets", NS)
                        for sheet in list(sheets):
                            if sheet.get("name") == "_folder_review_state":
                                sheets.remove(sheet)
                    elif name == "xl/_rels/workbook.xml.rels":
                        for relation in list(tree):
                            if relation.get("Target") == "worksheets/sheet4.xml":
                                tree.remove(relation)
                        for relation in tree:
                            number = int(relation.get("Id")[3:])
                            if number > 4:
                                relation.set("Id", f"rId{number - 1}")
                    elif name == "[Content_Types].xml":
                        for part in list(tree):
                            if part.get("PartName") == "/xl/worksheets/sheet4.xml":
                                tree.remove(part)
                    elif name == "docProps/app.xml":
                        for vector in tree.iter(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}vector"
                        ):
                            for part in list(vector):
                                if part.text == "_folder_review_state":
                                    vector.remove(part)
                                    vector.set("size", str(len(vector)))
                        for value in tree.iter(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}i4"
                        ):
                            if value.text == "4":
                                value.text = "3"
                    elif name == "xl/worksheets/sheet2.xml":
                        sheet_data = tree.find("s:sheetData", NS)
                        removed = []
                        for row in list(sheet_data):
                            texts = [
                                element.text or ""
                                for element in row.iter("{" + NS["s"] + "}t")
                            ]
                            if any(t.startswith("Comparison state ") for t in texts):
                                removed.append(int(row.get("r")))
                                sheet_data.remove(row)
                        for row in sheet_data:
                            number = int(row.get("r"))
                            shift = sum(r < number for r in removed)
                            row.set("r", str(number - shift))
                            for cell in row.findall("s:c", NS):
                                cell.set(
                                    "r",
                                    re.sub(
                                        r"[0-9]+$", str(number - shift), cell.get("r")
                                    ),
                                )
                        dimension = tree.find("s:dimension", NS)
                        if dimension is not None and removed:
                            dimension.set(
                                "ref",
                                re.sub(
                                    r"[0-9]+$",
                                    lambda m, removed=removed: str(
                                        int(m[0]) - len(removed)
                                    ),
                                    dimension.get("ref"),
                                ),
                            )
                data = ET.tostring(tree)
            result[name] = data
    return result, status_values, rows
