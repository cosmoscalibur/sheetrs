#!/usr/bin/env python3
"""Generate tests/assets/native.ods: the oracle formulas as a spreadsheet *created in
LibreOffice*, not imported from XLSX.

tests/assets/semantics.xlsx is loaded in LibreOffice only to read its formulas (in
LibreOffice's own names) and data; they are typed into a new document with ODF's default
settings (case-sensitive comparisons, regular expressions...), plus the functions whose
LibreOffice name means something else than the Excel one (CEILING, FLOOR), recalculated in
full and saved as ODS. See docs/FIDELITY.md ("Document state and namespace").

Run with LibreOffice's Python:  SOFFICE=soffice26 .../program/python scripts/gen_native_fixture.py
"""
import os
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

SOFFICE = os.environ.get("SOFFICE", "soffice")
HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "..", "tests", "assets")
PORT = 2013

# Native LibreOffice names (the file is an ODS: these are not Excel's functions of the same name).
NATIVE = []
for fn in ("CEILING", "FLOOR"):
    for args in ("-2.1;2", "2.1;-2", "-2.1;-2", "2.1;2", "0;2", "4.3;0", "-2.1;2;1", "2.1;-2;1", "-2.1;-2;1", "2.1;2;1", "-4.3;0.5", "4.3;0.5;1", "1;0;1", "-1;0"):
        NATIVE.append(f"{fn}({args})")
for fn in ("CEILING.XCL", "FLOOR.XCL", "CEILING.MATH", "FLOOR.MATH", "CEILING.PRECISE", "FLOOR.PRECISE", "ISO.CEILING"):
    for args in ("-2.1;2", "2.1;-2", "-2.1;-2", "2.1;2", "0;2", "4.3;0", "-4.3;0.5"):
        NATIVE.append(f"{fn}({args})")
for fn in ("CEILING", "FLOOR", "CEILING.XCL", "FLOOR.XCL"):
    for args in ("2.1", "-2.1", "0;0", "-4.3;0"):
        NATIVE.append(f"{fn}({args})")
# INDIRECT reads the reference text in the document's syntax (Calc A1 here, Excel A1 in an
# XLSX import), and ADDRESS writes the sheet in it.
NATIVE += [
    'INDIRECT("Data.A3")', 'INDIRECT("$Data.A3")', 'INDIRECT("Data!A3")', 'INDIRECT("A3")',
    'SUM(INDIRECT("Data.A1:A3"))', 'SUM(INDIRECT("Data.A1:Data.A3"))', 'SUM(INDIRECT("$Data.$A$1:$A$3"))',
    'INDIRECT("Calc.B1")', 'ADDRESS(1;1;1;1;"Data")', 'ADDRESS(2;3;4;1;"Data")', 'ADDRESS(1;1;1;1;"My sheet")',
]
for fn in ("CEILING.MATH", "FLOOR.MATH"):
    for args in ("-2.1;2;1", "2.1;-2;1", "-2.1;-2;1", "-2.1;2;0", "2.1;2;1"):
        NATIVE.append(f"{fn}({args})")


def pv(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v; return p


def main():
    profile = tempfile.mkdtemp(prefix="lo_native_")
    proc = subprocess.Popen([SOFFICE, "--headless", "--invisible", "--norestore", f"-env:UserInstallation=file://{profile}",
                             f"--accept=socket,host=localhost,port={PORT};urp;"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    ctx = None
    for _ in range(90):
        try:
            r = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            ctx = r.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext"); break
        except Exception:
            time.sleep(1)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    src = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(os.path.join(ASSETS, "semantics.xlsx"))), "_blank", 0, (pv("Hidden", True),))
    calc = src.Sheets.getByName("Calc")
    cur = calc.createCursor(); cur.gotoEndOfUsedArea(False)
    n = cur.RangeAddress.EndRow + 1
    formulas = [calc.getCellByPosition(0, i).getFormula() for i in range(n)]
    labels = [calc.getCellByPosition(1, i).getString() for i in range(n)]
    # Every sheet of the source but `Calc` is copied (data and hidden rows), in the source order: the
    # formulas of `Calc` read `Data` and, over several sheets, `Wide` (`Data:Wide!A1:A3`); a formula
    # typed into a document without the sheet it names is written back garbled.
    sheets = []
    for name in src.Sheets.ElementNames:
        if name == "Calc":
            continue
        sheet = src.Sheets.getByName(name)
        cur2 = sheet.createCursor(); cur2.gotoEndOfUsedArea(False)
        rows, cols = cur2.RangeAddress.EndRow + 1, cur2.RangeAddress.EndColumn + 1
        values = sheet.getCellRangeByPosition(0, 0, cols - 1, rows - 1).getFormulaArray()
        hidden = [not sheet.Rows.getByIndex(i).IsVisible for i in range(rows)]
        sheets.append((name, rows, cols, values, hidden))

    doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
    for k, (name, *_rest) in enumerate(sheets):
        doc.Sheets.insertNewByName(name, k)
    doc.Sheets.insertNewByName("Calc", len(sheets))
    for name in list(doc.Sheets.ElementNames):
        if name not in [x[0] for x in sheets] + ["Calc"]:
            doc.Sheets.removeByName(name)
    for name, rows, cols, values, hidden in sheets:
        d = doc.Sheets.getByName(name)
        d.getCellRangeByPosition(0, 0, cols - 1, rows - 1).setFormulaArray(values)
        for i, h in enumerate(hidden):
            if h:
                d.Rows.getByIndex(i).IsVisible = False
    c = doc.Sheets.getByName("Calc")
    for i, (f, label) in enumerate(zip(formulas, labels)):
        c.getCellByPosition(0, i).setFormula(f)
        c.getCellByPosition(1, i).setString(label)
    for j, f in enumerate(NATIVE):
        c.getCellByPosition(0, n + j).setFormula("=" + f)
        c.getCellByPosition(1, n + j).setString(f)
    doc.calculateAll()
    out = os.path.join(ASSETS, "native.ods")
    doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(out)), (pv("FilterName", "calc8"),))
    fixture_determinism.normalize(out)
    doc.close(True); src.close(True)
    try:
        desktop.terminate()
    except Exception:
        pass
    proc.wait(timeout=60)
    print(n, "formulas +", len(NATIVE), "native ->", out)


main()
