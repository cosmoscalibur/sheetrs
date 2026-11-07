"""Writes `<file>.oracle.json` for each fixture: what LibreOffice itself says of every cell when it
opens that file (its value, the text it shows, the format code, the formula in its own notation, the
error code).

Run it with a LibreOffice listening on a socket, in the folder that holds the fixtures:

    soffice --headless --accept="socket,host=localhost,port=2002;urp;" &
    python3 make_oracle.py

The tests of `sheetrs-query` and `sheetrs-io` compare SheetRS with these files: the oracle is the
behaviour of LibreOffice on that very file, not the other format of the same workbook. The reference
version is recorded in `version` (24.2.7 until 26.8 can be installed here).
"""
import json, os, sys, time
import uno
from com.sun.star.beans import PropertyValue

FIXTURES = ["cells", "formulas", "formats"]
HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.dirname(HERE)

local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


config = ctx.ServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
access = config.createInstanceWithArguments(
    "com.sun.star.configuration.ConfigurationAccess", (prop("nodepath", "/org.openoffice.Setup/Product"),)
)
version = access.getByName("ooSetupVersionAboutBox")

for fixture in FIXTURES:
    for ext in ("xlsx", "ods"):
        path = os.path.join(ASSETS, f"{fixture}.{ext}")
        doc = desktop.loadComponentFromURL("file://" + path, "_blank", 0, (prop("Hidden", True),))
        formats = doc.NumberFormats
        out = {"version": version, "file": f"{fixture}.{ext}", "cells": []}
        for sheet in doc.Sheets:
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(False)
            last_row, last_col = cursor.RangeAddress.EndRow, cursor.RangeAddress.EndColumn
            for row in range(last_row + 1):
                for col in range(last_col + 1):
                    cell = sheet.getCellByPosition(col, row)
                    kind = cell.Type.value
                    if kind == "EMPTY":
                        continue
                    entry = {
                        "sheet": sheet.Name,
                        "row": row,
                        "col": col,
                        "type": kind,
                        "string": cell.String,
                        "value": cell.Value,
                        "error": cell.getError(),
                        "format": formats.getByKey(cell.NumberFormat).FormatString,
                    }
                    if kind == "FORMULA":
                        entry["formula"] = cell.FormulaLocal
                        entry["result"] = cell.FormulaResultType2
                    out["cells"].append(entry)
        doc.close(True)
        with open(os.path.join(HERE, f"{fixture}.{ext}.oracle.json"), "w") as f:
            json.dump(out, f, indent=1, ensure_ascii=False)
        print(fixture, ext, len(out["cells"]), "cells,", version)
