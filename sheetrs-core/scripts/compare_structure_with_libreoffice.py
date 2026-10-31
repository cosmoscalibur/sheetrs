#!/usr/bin/env python3
"""Compares `sheetedit` with LibreOffice 26.8 on one operation at a time.

For each operation (rows and columns deleted, a sheet removed) it makes the operation in
LibreOffice through UNO on the workbook, and with `sheetedit` on the same file; LibreOffice opens
both results and the formulas and values of sheet `F` (column A), and the notes, manual page
breaks and shapes of sheet `T`, are compared. A reference that
goes is `#REF!` in both, whatever its spelling. Removing a sheet differs in an XLSX: LibreOffice
reads the bare `#REF!` that it writes itself as `#NAME?` (docs/FIDELITY.md of sheetrs-compiler).

Usage (LibreOffice 26.8, python3-uno and a built `sheetedit`):
    cargo build -p sheetrs-edit
    python3 sheetrs-core/scripts/compare_structure_with_libreoffice.py sheetrs-core/tests/assets/structure/book.ods
    python3 sheetrs-core/scripts/compare_structure_with_libreoffice.py sheetrs-core/tests/assets/structure/book.xlsx
"""
import os, re, subprocess, sys, tempfile, time
import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
SHEETEDIT = os.environ.get("SHEETEDIT", os.path.join(HERE, "..", "..", "target", "debug", "sheetedit"))
PORT = 2008
ROWS = 30  # formulas in column A of sheet F

# (label, sheetedit arguments, the same operation made in LibreOffice, and whether `sheetedit` must
# stop: LibreOffice places a shape by its size and the height of the rows when a deletion removes
# its cell together with the rows before it, and the file holds neither)
OPERATIONS = [
    ("delete rows T!4:5", ["--delete-rows", "T!4:5"],
     lambda d: d.Sheets.getByName("T").Rows.removeByIndex(3, 2)),
    ("delete rows T!2 and T!7:8", ["--delete-rows", "T!2", "T!7:8"],
     lambda d: (d.Sheets.getByName("T").Rows.removeByIndex(6, 2), d.Sheets.getByName("T").Rows.removeByIndex(1, 1))),
    ("delete columns T!B", ["--delete-cols", "T!B"],
     lambda d: d.Sheets.getByName("T").Columns.removeByIndex(1, 1)),
    ("remove sheet T", ["--remove-sheets", "T"], lambda d: d.Sheets.removeByName("T")),
    # Shapes: one anchored to C10, one anchored to D2 that ends in D4; the anchor goes, or the rows
    # between its two cells.
    ("delete rows T!10 (the anchor of a shape)", ["--delete-rows", "T!10"],
     lambda d: d.Sheets.getByName("T").Rows.removeByIndex(9, 1)),
    ("delete rows T!9:10 (the anchor of a shape with the row before it)", ["--delete-rows", "T!9:10"],
     lambda d: d.Sheets.getByName("T").Rows.removeByIndex(8, 2), True),
    ("delete rows T!10:12", ["--delete-rows", "T!10:12"],
     lambda d: d.Sheets.getByName("T").Rows.removeByIndex(9, 3)),
    ("delete rows T!1:3 (the anchor of a shape with the rows before it)", ["--delete-rows", "T!1:3"],
     lambda d: d.Sheets.getByName("T").Rows.removeByIndex(0, 3), True),
    ("delete rows T!3 (inside a shape)", ["--delete-rows", "T!3"],
     lambda d: d.Sheets.getByName("T").Rows.removeByIndex(2, 1)),
    ("delete columns T!C (the anchor column of a shape)", ["--delete-cols", "T!C"],
     lambda d: d.Sheets.getByName("T").Columns.removeByIndex(2, 1)),
    ("delete columns T!B:C (the anchor column with the one before)", ["--delete-cols", "T!B:C"],
     lambda d: d.Sheets.getByName("T").Columns.removeByIndex(1, 2), True),
    ("delete columns T!E", ["--delete-cols", "T!E"],
     lambda d: d.Sheets.getByName("T").Columns.removeByIndex(4, 1)),
    ("delete rows T!2 and columns T!C", ["--delete-rows", "T!2", "--delete-cols", "T!C"],
     lambda d: (d.Sheets.getByName("T").Columns.removeByIndex(2, 1), d.Sheets.getByName("T").Rows.removeByIndex(1, 1))),
    # The edited sheet is in the middle or at an end of the span of the 3D references.
    ("delete rows U!2:3", ["--delete-rows", "U!2:3"],
     lambda d: d.Sheets.getByName("U").Rows.removeByIndex(1, 2)),
    ("delete rows U!1", ["--delete-rows", "U!1"],
     lambda d: d.Sheets.getByName("U").Rows.removeByIndex(0, 1)),
    ("delete rows V!4:6", ["--delete-rows", "V!4:6"],
     lambda d: d.Sheets.getByName("V").Rows.removeByIndex(3, 3)),
    ("delete rows F!1:3", ["--delete-rows", "F!1:3"],
     lambda d: d.Sheets.getByName("F").Rows.removeByIndex(0, 3)),
    ("delete columns U!A", ["--delete-cols", "U!A"],
     lambda d: d.Sheets.getByName("U").Columns.removeByIndex(0, 1)),
    ("remove sheet U", ["--remove-sheets", "U"], lambda d: d.Sheets.removeByName("U")),
    ("remove sheets T and V", ["--remove-sheets", "T", "V"],
     lambda d: (d.Sheets.removeByName("T"), d.Sheets.removeByName("V"))),
]


def props(**kw):
    out = []
    for k, v in kw.items():
        p = PropertyValue()
        p.Name, p.Value = k, v
        out.append(p)
    return tuple(out)


def same_reference(formula):
    """The formula without the sheet that the second end of a range repeats (`$U.A3:U.A9` and
    `$U.A3:A9` are one reference), and a range of one cell as that cell."""
    formula = re.sub(r":\$?\w+\.", ":", formula)
    return re.sub(r"(\$?[A-Z]+\$?\d+):\1\b", r"\1", formula)


def column_a(doc, sheet):
    doc.calculateAll()
    cells = doc.Sheets.getByName(sheet)
    return [(c.getFormula(), c.getString()) for c in (cells.getCellRangeByName(f"A{i}") for i in range(1, ROWS + 1))]


def layout(doc, sheet="T"):
    """Where the notes, the manual page breaks and the shapes of `sheet` are."""
    cells = doc.Sheets.getByName(sheet)
    notes = sorted((n.Position.Row, n.Position.Column, n.String)
                   for n in (cells.Annotations.getByIndex(i) for i in range(cells.Annotations.Count)))
    breaks = [i for i in range(40) if cells.Rows.getByIndex(i).IsStartOfNewPage]
    shapes = sorted((a.Anchor.CellAddress.Row, a.Anchor.CellAddress.Column, a.Position.X, a.Position.Y,
                     a.Size.Width, a.Size.Height)
                    for a in (cells.DrawPage.getByIndex(i) for i in range(cells.DrawPage.Count)))
    return {"notes": notes, "page breaks (rows)": breaks, "shapes (anchor row, column, x, y, width, height)": shapes}


def main(book):
    extension = os.path.splitext(book)[1]
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore",
                             f"--accept=socket,host=localhost,port={PORT};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    ok = True
    try:
        for _ in range(90):
            try:
                local = uno.getComponentContext()
                resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        load = lambda path: desktop.loadComponentFromURL("file://" + os.path.abspath(path), "_blank", 0, props(Hidden=True))
        for label, args, native, *stops in OPERATIONS:
            out = tempfile.mktemp(suffix=extension)
            run = subprocess.run([SHEETEDIT, book, *args, "--recalc-on-open", "-o", out], capture_output=True, text=True)
            if run.returncode:
                if stops and "which the file does not hold" in run.stderr:
                    print("STOP " + label)
                    continue
                print("FAILED", label, run.stderr.strip()[-200:])
                ok = False
                continue
            if stops:
                print("DIFF " + label + ": sheetedit did not stop")
                ok = False
                continue
            expected_doc = load(book)
            native(expected_doc)
            expected = column_a(expected_doc, "F")
            expected_layout = layout(expected_doc) if not label.startswith("remove sheet") else {}
            expected_doc.close(True)
            got_doc = load(out)
            got = column_a(got_doc, "F")
            got_layout = layout(got_doc) if not label.startswith("remove sheet") else {}
            got_doc.close(True)
            os.remove(out)
            # LibreOffice reads the bare `#REF!` of an XLSX as an unknown name, also in the files it
            # writes itself: `#NAME?` where it had `#REF!` is the known difference.
            known = extension == ".xlsx" and label.startswith("remove sheet")
            differ = [(i + 1, e, g) for i, (e, g) in enumerate(zip(expected, got))
                      if (e[1] != g[1] and not (known and (e[1], g[1]) == ("#REF!", "#NAME?")))
                      or (same_reference(e[0]) != same_reference(g[0]) and "REF!" not in (e[0] + g[0]).upper())]
            moved = {k: (v, got_layout[k]) for k, v in expected_layout.items() if v != got_layout[k]}
            print(("OK   " if not differ and not moved else "DIFF ") + label)
            for row in differ:
                print("   F.A%d LibreOffice=%r sheetedit=%r" % row)
            for what, (e, g) in moved.items():
                print("   %s: LibreOffice=%r sheetedit=%r" % (what, e, g))
            ok = ok and not differ and not moved
        desktop.terminate()
    finally:
        proc.wait(timeout=60)
    print("ALL SAME" if ok else "DIFFERENCES")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
