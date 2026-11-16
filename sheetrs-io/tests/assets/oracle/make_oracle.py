"""Writes `<file>.oracle.json` for each fixture: what LibreOffice itself says of every cell when it
opens that file (its value, the text it shows, the format code, the formula in its own notation, the
error code).

Run it with a LibreOffice listening on a socket, in the folder that holds the fixtures:

    soffice --headless --accept="socket,host=localhost,port=2002;urp;" &
    /opt/libreoffice26.8/program/python make_oracle.py [fixture ...]

The tests of `sheetrs-query` and `sheetrs-io` compare SheetRS with these files: the oracle is the
behaviour of LibreOffice on that very file. The reference
version is recorded in `version` (26.8.2).

The fixtures `locale_<tag>` (made by `make_locale.py`) are recorded by a LibreOffice running with that
regional setting: start it with a profile whose `ooSetupSystemLocale` is the tag, on another port,
and name the port and the fixture:

    LO_PORT=2003 /opt/libreoffice26.8/program/python make_oracle.py locale_es-CO

(`scripts/libreoffice-server.sh` with `LO_PORT` and `LO_PROFILE` starts it; the profile needs
`<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>es-CO</value></prop></item>`
in its `registrymodifications.xcu`.)
"""
import json, os, sys, time
import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.dirname(HERE)
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
# (name of the oracle, folder of the files relative to the root of the repository[, extensions])
FIXTURES = [
    ("cells", "sheetrs-io/tests/assets"),
    ("inherit", "sheetrs-io/tests/assets"),
    ("inherit_more", "sheetrs-io/tests/assets"),
    ("inherit_text", "sheetrs-io/tests/assets"),
    ("formulas", "sheetrs-io/tests/assets"),
    ("formats", "sheetrs-io/tests/assets"),
    ("book", "sheetrs-edit/tests/assets/structure"),
    ("input", "sheetrs-edit/tests/assets/replace"),
    ("errors", "sheetrs-io/tests/assets"),
    ("lint_book", "tests"),
    ("dates", "sheetrs-io/tests/assets"),
    ("dates_1904", "sheetrs-io/tests/assets"),
    # An XLSX cannot say the 1900-01-01 null date (StarCalc 1.0): this one is an ODS only.
    ("dates_1900", "sheetrs-io/tests/assets", ("ods",)),
    # What a LibreOffice with the es-CO regional setting shows (see the docstring).
    ("locale_es-CO", "sheetrs-io/tests/assets"),
]

local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve(f"uno:socket,host=localhost,port={os.environ.get('LO_PORT', '2002')};urp;StarOffice.ComponentContext")
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

# Names on the command line limit the run to those fixtures (the others keep their file: what depends
# on the clock, `NOW()`, would change in every regeneration).
only = sys.argv[1:]
for fixture, folder, *extensions in FIXTURES:
    if only and fixture not in only:
        continue
    for ext in extensions[0] if extensions else ("xlsx", "ods"):
        path = os.path.join(ROOT, folder, f"{fixture}.{ext}")
        doc = desktop.loadComponentFromURL("file://" + path, "_blank", 0, (prop("Hidden", True),))
        formats = doc.NumberFormats
        out = {
            "version": version,
            "file": f"{fixture}.{ext}",
            "sheets": [sheet.Name for sheet in doc.Sheets],
            "cells": [],
            "notes": [],
        }
        for sheet in doc.Sheets:
            for i in range(sheet.Annotations.Count):
                note = sheet.Annotations.getByIndex(i)
                at = note.Position
                # `String` of a note of an XLSX is empty until its shape has been asked for.
                text = note.AnnotationShape.String
                out["notes"].append(
                    {"sheet": sheet.Name, "row": at.Row, "col": at.Column, "text": text, "author": note.Author}
                )
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
                        # The folder of the document is where LibreOffice resolves the files an external
                        # reference names; it is written as `<folder>` so the file does not depend on
                        # where the repository is.
                        entry["formula"] = cell.FormulaLocal.replace(
                            "file://" + os.path.dirname(path) + "/", "file://<folder>/"
                        )
                        entry["result"] = cell.FormulaResultType2
                    out["cells"].append(entry)
        doc.close(True)
        with open(os.path.join(HERE, f"{fixture}.{ext}.oracle.json"), "w") as f:
            json.dump(out, f, indent=1, ensure_ascii=False)
        print(fixture, ext, len(out["cells"]), "cells,", version)
