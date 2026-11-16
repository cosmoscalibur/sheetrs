"""Makes `external.{xlsx,ods}` of `tests/from_cache.rs`: `Calc!A1` reads another workbook, as LibreOffice 26.8 writes it
(`[1]Sheet1!A1` and the `externalLink` part of an XLSX; `['file:///...'#$Sheet1.A1]` and the table that caches the
cells it read in an ODS), and the cells of `make_fixtures.py` around it.

Run it with a LibreOffice listening on a socket (`scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python make_external.py

The other workbook is written in a temporary folder and its location replaced by `/external/` in the output, so that
the fixture does not depend on where it was made.
"""
import os
import sys
import tempfile
import time
import zipfile

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "..", "..", "scripts"))
import fixture_determinism  # the same bytes on every run, see its docstring

FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8"}
# The results LibreOffice cached for B1, C1 and E1 are replaced by numbers the formulas do not give (see
# `make_fixtures.py`), so that a cell that was calculated is told from one that kept its cached result.
CACHED = (("43", "1000"), ("10", "999"), ("11", "1000"))


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def new():
    return desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))


def rewrite(path, folder):
    """The package with the folder of the other workbook replaced and the results of `CACHED` changed."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith((".xml", ".rels")):
                text = data.decode().replace(f"file://{folder}/", "file:///external/").replace(folder, "/external")
                if item.filename in ("xl/worksheets/sheet1.xml", "content.xml"):
                    for old, new in CACHED:
                        text = text.replace(f"<v>{old}</v>", f"<v>{new}</v>")
                        text = text.replace(f'office:value="{old}"', f'office:value="{new}"')
                        text = text.replace(f"<text:p>{old}</text:p>", f"<text:p>{new}</text:p>")
                data = text.encode()
            zout.writestr(item, data)
    os.replace(tmp, path)


with tempfile.TemporaryDirectory() as folder:
    other = new()
    other.Sheets.getByIndex(0).Name = "Sheet1"
    other.Sheets.getByIndex(0).getCellByPosition(0, 0).setValue(42)
    other.storeToURL(f"file://{folder}/other.xlsx", (pv("FilterName", FILTERS["xlsx"]),))
    other.close(True)
    for extension in ("xlsx", "ods"):
        doc = new()
        sheet = doc.Sheets.getByIndex(0)
        sheet.Name = "Calc"
        sheet.getCellByPosition(3, 0).setValue(5)  # D1
        sheet.getCellByPosition(0, 0).setFormula(f"='file://{folder}/other.xlsx'#$Sheet1.A1")  # A1
        sheet.getCellByPosition(1, 0).setFormula("=A1+1")  # B1
        sheet.getCellByPosition(2, 0).setFormula("=D1*2")  # C1
        sheet.getCellByPosition(4, 0).setFormula("=C1+1")  # E1
        doc.calculateAll()
        out = os.path.join(HERE, f"external.{extension}")
        doc.storeToURL("file://" + out, (pv("FilterName", FILTERS[extension]),))
        doc.close(True)
        rewrite(out, folder)
        fixture_determinism.normalize(out)
