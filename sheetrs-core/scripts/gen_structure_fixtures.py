#!/usr/bin/env python3
"""Writes sheetrs-core/tests/assets/structure/book.{xlsx,ods}: the workbook that the tests of rows,
columns, sheets and names deleted (`sheetrs-core/tests/structure_*`) start from, built by LibreOffice.

Sheets `U` and `V`, after `T`, hold numbers in A1:B10 for the 3D references of `F`.
Sheet `T` holds numbers in A1:B10 and formulas in C1:C10 over its own cells (alternating patterns); sheet `F` holds the
formulas that read `T` (one reference form per row of the table in `F_FORMULAS`); the names are
`NameOne`, `NameTwo`, `NameThree` and `Gone`; `T!A7:B8` is merged; `T` has notes on B3, B5, B6 and B7,
manual page breaks before rows 3, 6 and 10, a shape anchored to C10 and another anchored to D2 that
resizes with its cells.

Usage (LibreOffice 26.8 and its python3-uno):
    python3 sheetrs-core/scripts/gen_structure_fixtures.py OUT_DIR
"""
import os, subprocess, sys, time
import uno
from com.sun.star.awt import Size
from com.sun.star.beans import PropertyValue

F_FORMULAS = [
    "=$T.A2", "=$T.A4", "=$T.A5", "=$T.A6", "=SUM($T.A2:A3)", "=SUM($T.A2:A4)", "=SUM($T.A4:A5)",
    "=SUM($T.A4:A8)", "=SUM($T.A2:A8)", "=SUM($T.A5:A6)", "=SUM($T.A:A)", "=SUM($T.4:5)",
    "=SUM($T.3:6)", "=SUM($T.1:3)", "=$T.$A$6", "=SUM($T.A6:B7)", "=NameOne+1", "=NameTwo",
    "=SUM($T.B2:D2)", "=SUM($T.C1:C10)", "=SUM($T.B:C)",
    # 3D references over `T`, `U` and `V`, in that order.
    "=SUM($T.A2:$V.A4)", "=SUM($T.A1:$V.A1)", "=SUM($T.A4:$V.A5)", "=SUM($T.A6:$V.A8)",
    "=SUM($T.B2:$V.D2)", "=SUM($U.A2:$V.A6)", "=SUM($T.A3:$U.A9)",
    "=SUM($T.A4:$V.B4)", "=SUM($T.B1:$V.B10)",
]


def pv(**kw):
    out = []
    for k, v in kw.items():
        p = PropertyValue()
        p.Name, p.Value = k, v
        out.append(p)
    return tuple(out)


def main(out_dir):
    port = 2007
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore",
                             f"--accept=socket,host=localhost,port={port};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        for _ in range(90):
            try:
                local = uno.getComponentContext()
                r = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
                ctx = r.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, pv(Hidden=True))
        doc.Sheets.getByIndex(0).Name = "T"
        for index, name in enumerate(("U", "V", "F"), 1):
            doc.Sheets.insertNewByName(name, index)
        t, f = doc.Sheets.getByName("T"), doc.Sheets.getByName("F")
        for name, scale in (("U", 100), ("V", 1000)):
            for r in range(1, 11):
                doc.Sheets.getByName(name).getCellRangeByName(f"A{r}").setValue(scale + r)
                doc.Sheets.getByName(name).getCellRangeByName(f"B{r}").setValue(scale + r * 10)
        for r in range(1, 11):
            t.getCellRangeByName(f"A{r}").setValue(r)
            t.getCellRangeByName(f"B{r}").setValue(r * 10)
            # Neighboring cells with one relative pattern are a formula group once LibreOffice
            # imports the file, and it moves the ranges of a group on its own terms; alternating
            # patterns keep each cell a formula of its own.
            if r % 2:
                t.getCellRangeByName(f"C{r}").setFormula(f"=SUM(A{max(r - 2, 1)}:A{r})")
            else:
                t.getCellRangeByName(f"C{r}").setFormula(f"=A{r}*B{r}")
        t.getCellRangeByName("D5").setFormula("=A5*B5")
        t.getCellRangeByName("A7:B8").merge(True)
        # Notes, manual page breaks (before rows 3, 6 and 10) and a shape anchored to cell C10.
        for row in (3, 5, 6, 7):
            t.Annotations.insertNew(t.getCellByPosition(1, row - 1).CellAddress, f"note {row}")
        for index in (2, 5, 9):
            t.Rows.getByIndex(index).IsStartOfNewPage = True
        shape = doc.createInstance("com.sun.star.drawing.RectangleShape")
        t.DrawPage.add(shape)
        shape.Size = Size(2000, 1000)
        shape.Position = t.getCellRangeByName("C10").Position
        shape.Anchor = t.getCellByPosition(2, 9)
        # A second shape that resizes with its cells, from B2 to C4: an ODS keeps its end cell.
        sized = doc.createInstance("com.sun.star.drawing.RectangleShape")
        t.DrawPage.add(sized)
        sized.Position = t.getCellRangeByName("D2").Position
        sized.Size = Size(2200, 1200)
        sized.Anchor = t.getCellByPosition(3, 1)
        sized.ResizeWithCell = True
        for i, formula in enumerate(F_FORMULAS, 1):
            f.getCellRangeByName(f"A{i}").setFormula(formula)
        base = f.getCellRangeByName("A1").CellAddress
        for name, content in [("NameOne", "$T.$A$6"), ("NameTwo", "$T.$A$1:$A$10"), ("NameThree", "$F.$B$1"), ("Gone", "$T.$A$4")]:
            doc.NamedRanges.addNewByName(name, content, base, 0)
        doc.calculateAll()
        os.makedirs(out_dir, exist_ok=True)
        for ext, flt in (("xlsx", "Calc MS Excel 2007 XML"), ("ods", "calc8")):
            doc.storeToURL("file://" + os.path.abspath(os.path.join(out_dir, f"book.{ext}")), pv(FilterName=flt))
        doc.close(True)
        try:
            desktop.terminate()
        except Exception:
            pass
    finally:
        proc.wait(timeout=60)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
