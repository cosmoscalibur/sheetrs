"""Which formulas LibreOffice 26.8 calculates again, and which keep the result the file caches, when an edit is made and the
file is saved and opened again (the default load): the evidence of the rules of `sheetrs_compiler::Cache`
(`sheetrs-compiler/docs/FIDELITY.md`, *What a deletion reaches*).

The workbook is `tests/assets/cached/dirty.xlsx` or `.ods` (`make_fixtures.py`): every result it caches is the text `S`, and a
formula that shows `S` after the edit kept it, one that shows anything else was calculated. Column A of `Calc` holds one
formula for each label and column B reads it (what depends on a calculated result is calculated too). The table has a row per
formula and a column per edit, `K` kept and `R` calculated again. Run it with the Python of LibreOffice and a server
(`scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python probe_dirty.py xlsx|ods [fixture]

With the fixtures `classes` and `classes_volatile` (one formula for each class of function, `make_fixtures.py`) the edits are the ones that do not delete anything
(hiding and showing rows, writing a value) and the table says, for each function, whether `LibreOffice` calculates it when it opens the file
(`no edit`) and when each edit is made.
"""
import os
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8"}
# (the label of the edit, the operation)
EDITS = [
    ("no edit", None),
    ("rows Data!3", ("rows", "Data", 3, 3)),
    ("rows Data!2:3", ("rows", "Data", 2, 3)),
    ("rows Data!6", ("rows", "Data", 6, 6)),
    ("rows Data!7:8", ("rows", "Data", 7, 8)),
    ("rows Data!1", ("rows", "Data", 1, 1)),
    ("rows Data!4", ("rows", "Data", 4, 4)),
    ("cols Data!B", ("cols", "Data", "B", "B")),
    ("cols Data!C", ("cols", "Data", "C", "C")),
    ("cols Data!D", ("cols", "Data", "D", "D")),
    ("cols Data!A", ("cols", "Data", "A", "A")),
    ("sheet Other", ("sheet", "Other")),
    ("sheet Data", ("sheet", "Data")),
    ("name n_cell1", ("name", "n_cell1")),
    ("rows Calc!2", ("rows", "Calc", 2, 2)),
    ("cols Calc!D", ("cols", "Calc", "D", "D")),
]
CLASSES_EDITS = [
    ("no edit", None),
    ("hide Data!2", ("hide", "Data", 2)),
    ("hide Data!3", ("hide", "Data", 3)),
    ("hide Data!6", ("hide", "Data", 6)),
    ("write Data!A1", ("set", "Data", "A1", 5)),
    ("write Data!B1", ("set", "Data", "B1", "y")),
    ("rows Data!3", ("rows", "Data", 3, 3)),
    ("cols Data!B", ("cols", "Data", "B", "B")),
    ("rows Calc!2", ("rows", "Calc", 2, 2)),
]


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


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


def column(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def edit(doc, op):
    if op is None:
        return
    sheets = doc.Sheets
    if op[0] == "rows":
        sheets.getByName(op[1]).Rows.removeByIndex(op[2] - 1, op[3] - op[2] + 1)
    elif op[0] == "cols":
        sheets.getByName(op[1]).Columns.removeByIndex(column(op[2]), column(op[3]) - column(op[2]) + 1)
    elif op[0] == "hide":
        sheets.getByName(op[1]).Rows.getByIndex(op[2] - 1).IsVisible = False
    elif op[0] == "set":
        cell = sheets.getByName(op[1]).getCellRangeByName(op[2])
        cell.setString(op[3]) if isinstance(op[3], str) else cell.setValue(op[3])
    elif op[0] == "sheet":
        sheets.removeByName(op[1])
    else:
        doc.NamedRanges.removeByName(op[1])


def main(extension, fixture="dirty"):
    global EDITS
    if fixture in ("classes", "classes_volatile"):
        EDITS = CLASSES_EDITS
    desktop = connect()
    source = os.path.join(HERE, "..", "cached", f"{fixture}.{extension}")

    def load(path):
        return desktop.loadComponentFromURL("file://" + os.path.abspath(path), "_blank", 0, (pv("Hidden", True),))

    # What each row of Calc is, from the formulas of the file.
    doc = load(source)
    calc = doc.Sheets.getByName("Calc")
    formulas = []
    for row in range(80):
        text = calc.getCellByPosition(0, row).getFormula()
        if text:
            formulas.append(text)
    doc.close(True)
    table = []
    for label, op in EDITS:
        doc = load(source)
        edit(doc, op)
        with tempfile.TemporaryDirectory() as folder:
            own = os.path.join(folder, f"saved.{extension}")
            doc.storeToURL("file://" + own, (pv("FilterName", FILTERS[extension]),))
            doc.close(True)
            doc = load(own)
            calc = doc.Sheets.getByName("Calc")
            # The formula at the row (the deletions of `Calc` move the rows: the table is of what is left).
            table.append([calc.getCellByPosition(0, row).getString() == "S" for row in range(len(formulas))])
            doc.close(True)
    print(f"{extension}: K = the formula kept the result the file caches, R = it was calculated again")
    for i, (label, _) in enumerate(EDITS):
        print(f"{i:2} {label}")
    print(" " * 44 + " ".join(f"{i:2}" for i in range(len(EDITS))))
    for row, text in enumerate(formulas):
        print(f"{text[:43]:43} " + " ".join(f"{'K' if col[row] else 'R':>2}" for col in table))


main(sys.argv[1] if len(sys.argv) > 1 else "xlsx", *(sys.argv[2:3]))
