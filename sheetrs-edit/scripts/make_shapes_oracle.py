#!/usr/bin/env python3
"""Makes `sheetrs-io/tests/assets/oracle/shapes_delete.json`: where LibreOffice puts the two shapes of
`sheetrs-edit/tests/assets/structure/book.{xlsx,ods}` (sheet `T`) after deleting each band of rows
and of columns, through UNO (`Rows.removeByIndex`, `Columns.removeByIndex`).

    scripts/libreoffice-server.sh
    /opt/libreoffice26.8/program/python sheetrs-edit/scripts/make_shapes_oracle.py

Each entry: the format, the axis, the first index and the count of the band, and the shapes left,
as [anchor row, anchor column, x, y, width, height] in 1/100 mm, sorted by x then y.
"""
import json, os, subprocess, sys, time
import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..", "..")
OUT = os.path.join(ROOT, "sheetrs-io", "tests", "assets", "oracle", "shapes_delete.json")


def props(**kw):
    out = []
    for k, v in kw.items():
        p = PropertyValue()
        p.Name, p.Value = k, v
        out.append(p)
    return tuple(out)


def connect():
    for _ in range(60):
        try:
            local = uno.getComponentContext()
            resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        except Exception:
            time.sleep(1)
    raise SystemExit("LibreOffice is not listening on 2002: scripts/libreoffice-server.sh")


def shapes(doc):
    sheet = doc.Sheets.getByName("T")
    found = [
        [a.Anchor.CellAddress.Row, a.Anchor.CellAddress.Column, a.Position.X, a.Position.Y, a.Size.Width, a.Size.Height]
        for a in (sheet.DrawPage.getByIndex(i) for i in range(sheet.DrawPage.Count))
    ]
    return sorted(found, key=lambda s: (s[2], s[3]))


def main():
    desktop = connect()
    cases = []
    for extension in ("xlsx", "ods"):
        book = os.path.abspath(os.path.join(ROOT, "sheetrs-edit", "tests", "assets", "structure", f"book.{extension}"))
        for axis, firsts in (("rows", range(12)), ("columns", range(7))):
            for count in range(1, 6):
                for first in firsts:
                    try:
                        doc = desktop.loadComponentFromURL("file://" + book, "_blank", 0, props(Hidden=True))
                        sheet = doc.Sheets.getByName("T")
                        (sheet.Rows if axis == "rows" else sheet.Columns).removeByIndex(first, count)
                        left = shapes(doc)
                        doc.close(True)
                    except Exception:
                        # A deletion that LibreOffice cannot do (it dies on a few): no entry.
                        subprocess.run([os.path.join(ROOT, "scripts", "libreoffice-server.sh")], capture_output=True)
                        desktop = connect()
                        continue
                    cases.append({"format": extension, "axis": axis, "first": first, "count": count, "shapes": left})
    with open(OUT, "w") as f:
        json.dump(cases, f, indent=1)
        f.write("\n")
    print(len(cases), "cases ->", OUT)


if __name__ == "__main__":
    main()
