"""Which formulas of `visibility.{xlsx,ods}` (`tests/assets/cached`) LibreOffice 26.8 calculates when rows of `Data` are hidden or shown, by UNO and with no
save: each argument row is toggled and the `SUBTOTAL`s of `Calc!A1:A13` are read (a cell that shows `S` kept the result the file caches).

    /opt/libreoffice26.8/program/python probe_hide_subtotal.py DIR 2,5     # DIR is `tests/assets/cached`; LibreOffice listening on port 2002

OBSERVED: in an ODS a `SUBTOTAL` that reads a single cell of a row that flipped is calculated; in an XLSX it keeps its result.
"""
import sys, time, uno
from com.sun.star.beans import PropertyValue

def pv(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v; return p
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
BASE = sys.argv[1]
rows = [int(x) for x in sys.argv[2].split(',')]
for ext in ("xlsx", "ods"):
    doc = desktop.loadComponentFromURL(f"file://{BASE}/visibility.{ext}", "_blank", 0, (pv("Hidden", True),))
    data = doc.Sheets.getByName("Data")
    calc = doc.Sheets.getByName("Calc")
    for r in rows:
        data.Rows.getByIndex(r - 1).IsVisible = not data.Rows.getByIndex(r - 1).IsVisible
    out = []
    for i in range(13):
        c = calc.getCellByPosition(0, i)
        out.append(c.getString())
    print(ext, "toggle", rows, out, [calc.getCellByPosition(1, i).getString() for i in range(3)])
    doc.close(True)
