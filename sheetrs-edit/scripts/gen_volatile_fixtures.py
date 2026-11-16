#!/usr/bin/env python3
"""The workbooks of the tests of `now` and `seed` (`NOW()`, `TODAY()` and `RAND()` fixed by a
calculation): one sheet `V` that holds the three volatile functions and a few cells to edit, with no
cached result (a formula that is calculated is calculated by the engine), as an XLSX and as an ODS.

    V!A1 =NOW()   V!B1 =TODAY()   V!C1 =RAND()
    V!A2 5        V!B2 =A2*2
    V!A3 7

They are written with the standard library only (no application recalculates them: what the tests
read is what the engine writes) and are the same bytes on every run.

    python3 sheetrs-edit/scripts/gen_volatile_fixtures.py
"""
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fixture_determinism  # noqa: E402

OUT = ROOT / "sheetrs-edit" / "tests" / "assets" / "volatile"
DATE = (2024, 1, 1, 0, 0, 0)


def write(path, entries, stored_first=None):
    with zipfile.ZipFile(path, "w") as z:
        for name, data in entries:
            info = zipfile.ZipInfo(name, DATE)
            info.compress_type = (
                zipfile.ZIP_STORED if name == stored_first else zipfile.ZIP_DEFLATED
            )
            z.writestr(info, data)


def xlsx(path):
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        '<row r="1"><c r="A1"><f>NOW()</f></c><c r="B1"><f>TODAY()</f></c><c r="C1"><f>RAND()</f></c></row>'
        '<row r="2"><c r="A2"><v>5</v></c><c r="B2"><f>A2*2</f></c></row>'
        '<row r="3"><c r="A3"><v>7</v></c></row>'
        "</sheetData></worksheet>"
    )
    write(
        path,
        [
            (
                "[Content_Types].xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                "</Types>",
            ),
            (
                "_rels/.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                "</Relationships>",
            ),
            (
                "xl/workbook.xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="V" sheetId="1" r:id="rId1"/></sheets></workbook>',
            ),
            (
                "xl/_rels/workbook.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                "</Relationships>",
            ),
            ("xl/worksheets/sheet1.xml", sheet),
        ],
    )


def ods(path):
    ns = (
        'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
        'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"'
    )
    content = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<office:document-content {ns} office:version="1.3"><office:body><office:spreadsheet>'
        '<table:table table:name="V">'
        "<table:table-row>"
        '<table:table-cell table:formula="of:=NOW()"/>'
        '<table:table-cell table:formula="of:=TODAY()"/>'
        '<table:table-cell table:formula="of:=RAND()"/>'
        "</table:table-row>"
        "<table:table-row>"
        '<table:table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>'
        '<table:table-cell table:formula="of:=[.A2]*2"/>'
        "</table:table-row>"
        "<table:table-row>"
        '<table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell>'
        "</table:table-row>"
        "</table:table></office:spreadsheet></office:body></office:document-content>"
    )
    manifest = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
        '<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        "</manifest:manifest>"
    )
    write(
        path,
        [
            ("mimetype", "application/vnd.oasis.opendocument.spreadsheet"),
            ("META-INF/manifest.xml", manifest),
            ("content.xml", content),
        ],
        stored_first="mimetype",
    )


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    for make, name in ((xlsx, "volatile.xlsx"), (ods, "volatile.ods")):
        make(OUT / name)
        fixture_determinism.normalize(OUT / name)
        print("wrote", OUT / name)
