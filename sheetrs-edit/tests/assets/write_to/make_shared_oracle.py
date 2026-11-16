"""Makes `shared_beyond.xlsx` and `shared_beyond.json`: what LibreOffice 26.8 shows for the copies of a shared formula
that would read outside the sheet (past column XFD), which `tests/shared_beyond.rs` compares `SheetRS` with.

THE RULE (`CONTRIBUTING.md`, *Parity and fidelity*): the reference is the workbook SAVED AND OPENED AGAIN. For each cell
of the fixture the oracle records, informationally, what LibreOffice shows when it opens the fixture (`first_load`:
`#REF!` and a text such as `=#REF!2+1`, which is its memory and not reproducible) and, as THE EXPECTED RESULT, what it
shows after saving the file as an XLSX and opening that one (`reread`: the value, as text) together with the formula text
it wrote in the XML of that file (`saved`: `#REF!+1`, a bare `#REF!`, in the notation of the file).

An ODS has no shared formulas (each cell holds its own text), so the case is XLSX only.

    /opt/libreoffice26.8/program/python make_shared_oracle.py     (LibreOffice listening: scripts/libreoffice-server.sh)
"""
import json
import os
import re
import sys
import tempfile
import time
import zipfile
from xml.sax.saxutils import escape

import uno
from com.sun.star.beans import PropertyValue

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the zip dates of a fixture in the repository are constants)

HERE = os.path.dirname(os.path.abspath(__file__))
NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
# The last three columns of the sheet: XFB, XFC and XFD (0-based 16381..16383).
COLUMNS = (16381, 16382, 16383)
# (row, formula of the first cell). Each group is the three cells XFB:XFD of a row, sharing the formula of XFB.
GROUPS = [
    (1, "XFC2+1"),
    (3, "SUM(XFB1:XFC1)"),
    (4, "Data!XFC1"),
    (5, "COUNT(XFC1)"),
    (6, "COUNTA(XFC1)"),
    (7, "IFERROR(XFC1,77)"),
    (8, "SUM(XFC1:XFD1)"),
    (9, "XFC1:XFD1"),
    (10, "ISERROR(XFC1)"),
]
LETTERS = {16381: "XFB", 16382: "XFC", 16383: "XFD"}


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def cached(value):
    """The `<v>` of a cached result given as the text LibreOffice shows: a number, a boolean or an error."""
    if value is None:
        return "<v>1</v>", ""
    if value in ("TRUE", "FALSE"):
        return f"<v>{int(value == 'TRUE')}</v>", ' t="b"'
    if value.startswith("#"):
        return f"<v>{value}</v>", ' t="e"'
    return f"<v>{value}</v>", ""


def calc_sheet(results=None):
    """The sheet `Calc`: each group of GROUPS is three cells that share a formula. `results` are the values the cells cache
    (`None`: a placeholder, which is the first pass: the second caches what LibreOffice calculates, so that the file holds
    right results and what it shows of a cell it does not calculate is the true value)."""
    rows = ""
    for si, (row, text) in enumerate(GROUPS):
        cells = ""
        for i, col in enumerate(COLUMNS):
            at = f"{LETTERS[col]}{row}"
            v, t = cached(None if results is None else results[at])
            if i == 0:
                cells += f'<c r="{at}"{t}><f t="shared" ref="XFB{row}:XFD{row}" si="{si}">{escape(text)}</f>{v}</c>'
            else:
                cells += f'<c r="{at}"{t}><f t="shared" si="{si}"/>{v}</c>'
        rows += f'<row r="{row}">{cells}</row>'
    return f'<?xml version="1.0" encoding="UTF-8"?><worksheet {NS}><sheetData>{rows}</sheetData></worksheet>'
    return f'<?xml version="1.0" encoding="UTF-8"?><worksheet {NS}><sheetData>{rows}</sheetData></worksheet>'


def make_fixture(path, results=None):
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    sheets = {
        "xl/worksheets/sheet1.xml": f'<?xml version="1.0" encoding="UTF-8"?><worksheet {NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
        "xl/worksheets/sheet2.xml": calc_sheet(results),
    }
    parts = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        "_rels/.rels": f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<?xml version="1.0" encoding="UTF-8"?><workbook {NS} xmlns:r="{rel}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Calc" sheetId="2" r:id="rId2"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{rel}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
        **sheets,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(parts):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))  # constant: the fixture is deterministic
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, parts[name])


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(30):
        try:
            ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        except Exception:
            time.sleep(1)
    raise SystemExit("no LibreOffice listening on port 2002: scripts/libreoffice-server.sh")


def read(doc):
    sheet = doc.Sheets.getByName("Calc")
    out = {}
    for row, _ in GROUPS:
        for col in COLUMNS:
            cell = sheet.getCellByPosition(col, row - 1)
            out[f"{LETTERS[col]}{row}"] = {"formula": cell.Formula, "value": cell.String}
    return out


def saved_formulas(path):
    text = zipfile.ZipFile(path).read("xl/worksheets/sheet2.xml").decode()
    return {m.group(1): m.group(2) for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*><f[^>]*>([^<]*)</f>', text)}


def main():
    fixture = os.path.join(HERE, "shared_beyond.xlsx")
    desktop = connect()
    load = lambda p: desktop.loadComponentFromURL("file://" + p, "_blank", 0, (pv("Hidden", True),))
    # First pass: what LibreOffice calculates for the cells; second pass: the file caches those results.
    make_fixture(fixture)
    doc = load(fixture)
    doc.calculateAll()
    calculated = {cell: v["value"] for cell, v in read(doc).items()}
    doc.close(True)
    make_fixture(fixture, calculated)
    fixture_determinism.normalize(fixture)
    doc = load(fixture)
    first_load = read(doc)
    with tempfile.TemporaryDirectory() as folder:
        own = os.path.join(folder, "own.xlsx")
        doc.storeToURL("file://" + own, (pv("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(True)
        saved = saved_formulas(own)
        doc = load(own)
        reread = read(doc)
        doc.close(True)
    oracle = {"cells": {}}
    for cell in reread:
        oracle["cells"][cell] = {
            "first_load": first_load[cell],  # informational: the memory, never compared
            "reread": reread[cell]["value"],  # THE EXPECTED VALUE
            "saved": saved[cell],  # THE EXPECTED TEXT, as the file holds it
        }
    with open(os.path.join(HERE, "shared_beyond.json"), "w") as f:
        json.dump(oracle, f, indent=1, sort_keys=True)
        f.write("\n")
    print("wrote shared_beyond.xlsx and shared_beyond.json:", len(oracle["cells"]), "cells")


main()
