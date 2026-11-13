"""Makes `sheetrs-io/tests/assets/dates.{xlsx,ods}`, `dates_1904.{xlsx,ods}` and `dates_1900.ods`: numbers with
date and time formats in the three date systems of LibreOffice (1899-12-30, the default; 1904-01-01; and
the 1900-01-01 of StarCalc 1.0, an ODS only), so the oracle records what LibreOffice itself shows for each: the serials around the
1900 leap-year quirk (0, 1, 59, 60, 61), negative ones, the end of the supported years, a rounding up to
midnight and fractions of a second, in a date format, a date and time format, a time format, an elapsed
time format and the formats an XLSX numbers (`mm-dd-yy`, `d-mmm-yy`, `h:mm AM/PM`).

The XLSX files are written by hand (a minimal package); the ODS files are what LibreOffice saves from them.
It also writes `other_null_date.ods`, an ODS whose null date is none of the three (1950-06-15), which
LibreOffice can store through its API and SheetRS refuses with an error: it has no oracle.
Run it with a LibreOffice listening on a socket (see `make_oracle.py`):

    /opt/libreoffice26.8/program/python make_dates.py
"""
import os, time, zipfile
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.util import Date

import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)

# Column by column: the format code, written as the XLSX writes it (a custom one) or a built-in id.
FORMATS = [
    ("yyyy-mm-dd", None),
    ("yyyy-mm-dd hh:mm:ss", None),
    ("hh:mm:ss", None),
    ("[hh]:mm:ss", None),
    ("yyyy-mm-dd hh:mm:ss.000", None),
    ("General", 0),
    ("dd/mm/yyyy", None),
    ("d-mmm-yy", 15),
    ("mm-dd-yy", 14),
    ("h:mm AM/PM", 18),
    ("mm:ss", None),
    ("hh:mm:ss.00", None),
]

SERIALS = [
    0, 1, 2, 59, 60, 61, 366, 367, 25569, 36526, 44927.5, 45366, 45366.4375, 45366.999999,
    45366.99999999, 45366.5000058, 45366.0000115, 0.5, 0.75, 0.999999, 1.5, 0.000011574,
    -1, -0.25, -365.5, -693593, -693594, 2958465, 2958466, 3e6, 11967899, 11967900, 12000000,
]


def xlsx(path, date1904):
    ids = {}
    xfs = []
    numfmts = []
    next_id = 164
    for text, builtin in FORMATS:
        if builtin is not None:
            fid = builtin
        else:
            fid = next_id
            next_id += 1
            numfmts.append(f'<numFmt numFmtId="{fid}" formatCode="{text}"/>')
        xfs.append(f'<xf numFmtId="{fid}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<numFmts count="{len(numfmts)}">{"".join(numfmts)}</numFmts>'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs></styleSheet>'
    )
    rows = ""
    for r, serial in enumerate(SERIALS, 1):
        cells = "".join(
            f'<c r="{chr(65 + c)}{r}" s="{c}"><v>{serial!r}</v></c>' for c in range(len(FORMATS))
        )
        rows += f'<row r="{r}">{cells}</row>'
    pr = '<workbookPr date1904="1"/>' if date1904 else ""
    parts = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{pr}<sheets><sheet name="Dates" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": styles,
        "xl/worksheets/sheet1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{rows}</sheetData></worksheet>',
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


xlsx(os.path.join(OUT, "dates.xlsx"), False)
xlsx(os.path.join(OUT, "dates_1904.xlsx"), True)
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
made = ["dates.xlsx", "dates_1904.xlsx"]
for name, ods, null_date in (
    ("dates.xlsx", "dates.ods", None),
    ("dates_1904.xlsx", "dates_1904.ods", None),
    ("dates.xlsx", "dates_1900.ods", Date(1, 1, 1900)),
    # None of the three: LibreOffice stores it (its API), SheetRS refuses it with an error, so it has no oracle.
    ("dates.xlsx", "other_null_date.ods", Date(15, 6, 1950)),
):
    doc = desktop.loadComponentFromURL("file://" + os.path.join(OUT, name), "_blank", 0, (prop("Hidden", True),))
    if null_date is not None:
        doc.NullDate = null_date
    doc.storeToURL("file://" + os.path.join(OUT, ods), (prop("FilterName", "calc8"),))
    doc.close(True)
    made.append(ods)
for name in made:
    fixture_determinism.normalize(os.path.join(OUT, name))
print(", ".join(made))
