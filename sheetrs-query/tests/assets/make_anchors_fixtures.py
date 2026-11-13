"""Makes `anchors.{xlsx,ods}`, the workbooks of the searches with `^` and `$` (`libreoffice_find.rs`):
a sheet `T` with one text a row, among them texts of several lines, with a line break at the start or
at the end, and the headers of the issue that asked for the anchors in `find`. The oracle
(`make_find_oracle.py`) records the cells that `findAll` of LibreOffice finds in them.

Run it with a LibreOffice listening on a socket (`scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python make_anchors_fixtures.py
"""
import os, sys, time
import uno
from com.sun.star.beans import PropertyValue

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

HERE = os.path.dirname(os.path.abspath(__file__))
TEXTS = [
    "a", "ab", "ba", "b", "A", "aa", "a b", "Code", "code", "Code x", "Resumen Contadia",
    "x\na", "a\nx", "a\nx\nb", "x\na\nb", "ab\ncd", "a\n", "\na", "a\n\n", "a ", " a",
]
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
sheet = doc.Sheets.getByIndex(0)
sheet.Name = "T"
for i, text in enumerate(TEXTS):
    sheet.getCellByPosition(0, i).String = text
doc.storeToURL("file://" + os.path.join(HERE, "anchors.xlsx"), (pv("FilterName", "Calc MS Excel 2007 XML"),))
doc.storeToURL("file://" + os.path.join(HERE, "anchors.ods"), (pv("FilterName", "calc8"),))
for ext in ("xlsx", "ods"):
    fixture_determinism.normalize(os.path.join(HERE, "anchors." + ext))
doc.close(True)
