"""Makes `sheetrs-io/tests/assets/errors.{xlsx,ods}`: formulas that hold an error written in the formula
(a bare `#REF!`, `#N/A`, `#DIV/0!`) and calls to a function that does not exist (`casa()`), in macro-free
files, so the oracle records what LibreOffice itself shows for each.

The XLSX is written by hand (a minimal package, formulas with their cached values as Excel writes them);
the ODS is what LibreOffice saves from that XLSX, so its formulas are in the notation LibreOffice
writes. Run it with a LibreOffice listening on a socket (see `make_oracle.py`):

    /opt/libreoffice26.8/program/python make_errors.py
"""
import os, time, zipfile
import uno
from com.sun.star.beans import PropertyValue

import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)

# (formula as the XLSX stores it, cached error value)
FORMULAS = [
    ("#REF!+5", "#REF!"),
    ("SUM(#REF!,1)", "#REF!"),
    ("#REF!", "#REF!"),
    ("IF(TRUE,#REF!,1)", "#REF!"),
    ("#N/A+1", "#N/A"),
    ("#DIV/0!+1", "#DIV/0!"),
    ("casa()", "#NAME?"),
    ("CASA()", "#NAME?"),
    ("_xlfn.casa()", "#NAME?"),
    ("_xlfn.CASA()", "#NAME?"),
    ("casa(1,2)+1", "#NAME?"),
    ("SUM(casa(),1)", "#NAME?"),
]


def xlsx(path):
    rows = "".join(
        f'<row r="{i}"><c r="A{i}" t="e"><f>{f.replace("&", "&amp;")}</f><v>{v}</v></c></row>'
        for i, (f, v) in enumerate(FORMULAS, 1)
    )
    parts = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Errors" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{rows}</sheetData></worksheet>',
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


xlsx(os.path.join(OUT, "errors.xlsx"))
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = desktop.loadComponentFromURL(
    "file://" + os.path.join(OUT, "errors.xlsx"), "_blank", 0, (prop("Hidden", True),)
)
doc.storeToURL("file://" + os.path.join(OUT, "errors.ods"), (prop("FilterName", "calc8"),))
doc.close(True)
for name in ("errors.xlsx", "errors.ods"):
    fixture_determinism.normalize(os.path.join(OUT, name))
print("errors.xlsx, errors.ods")
