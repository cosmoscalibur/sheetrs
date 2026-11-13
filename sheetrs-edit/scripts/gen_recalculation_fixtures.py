#!/usr/bin/env python3
"""Writes `sheetrs-edit/tests/assets/structure/expected_<name>.{xlsx,ods}`: the workbook that
LibreOffice 26.8 leaves when it does a sequence of operations on `structure/book.{xlsx,ods}` itself,
ending in a conversion of formulas to values. It calculates as it goes, so the conversion reads the
results the deletions before it left (a formula over a row that went reads one row less), which is
what the editor reproduces by calculating the results again with the engine before it converts
(`sheetrs-edit/tests/recalculation.rs`).

    scripts/libreoffice-server.sh        # port 2002
    /opt/libreoffice26.8/program/python sheetrs-edit/scripts/gen_recalculation_fixtures.py [NAME ...]
"""
import os
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.normpath(os.path.join(HERE, "..", "tests", "assets", "structure"))
FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8"}

# Each operation addresses the workbook as the ones before it left it. The test has the same list.
SEQUENCES = {
    # No change before the conversion: what the file caches is what is converted, an error (F!A18)
    # included, which LibreOffice leaves as an empty text.
    "value_only": [("formula_to_value", "F")],
    "delete_rows_then_value": [("delete_rows", "T", 2, 1), ("formula_to_value", "F")],
    "delete_everything_then_value": [
        ("delete_rows", "T", 2, 1),
        ("delete_columns", "U", 1, 1),
        ("delete_sheet", "V"),
        ("delete_name", "Gone"),
        ("formula_to_value", "F"),
    ],
    "value_in_the_middle": [
        ("delete_rows", "T", 2, 1),
        ("formula_to_value", "F", "A1:A8"),
        ("delete_rows", "T", 1, 1),
    ],
}


def pv(**kw):
    out = []
    for k, v in kw.items():
        p = PropertyValue()
        p.Name, p.Value = k, v
        out.append(p)
    return tuple(out)


local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
ctx = None
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def delete_rows(doc, sheet, first, count):
    doc.Sheets.getByName(sheet).Rows.removeByIndex(first - 1, count)


def delete_columns(doc, sheet, first, count):
    doc.Sheets.getByName(sheet).Columns.removeByIndex(first, count)


def delete_sheet(doc, name):
    doc.Sheets.removeByName(name)


def delete_name(doc, name):
    doc.NamedRanges.removeByName(name)


def formula_to_value(doc, sheet, area=None):
    """`Data > Calculate > Formula to Value` (`.uno:ConvertFormulaToValue`) on `area` or the sheet."""
    # What the application shows after a deletion is the results calculated again. A document that
    # was loaded from a file and driven through UNO without a window does not interpret the formulas
    # it adjusted until something reads them, and the conversion would take the results the file
    # held: `calculateAll` is the recalculation the screen does as it repaints.
    doc.calculateAll()
    sh = doc.Sheets.getByName(sheet)
    controller = doc.CurrentController
    controller.setActiveSheet(sh)
    controller.select(sh.getCellRangeByName(area) if area else sh)
    dispatcher = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    dispatcher.executeDispatch(controller.Frame, ".uno:ConvertFormulaToValue", "", 0, ())


DISPATCH = {
    "delete_rows": delete_rows, "delete_columns": delete_columns, "delete_sheet": delete_sheet,
    "delete_name": delete_name, "formula_to_value": formula_to_value,
}


def main(only):
    for name, operations in SEQUENCES.items():
        if only and name not in only:
            continue
        for ext, filter_name in FILTERS.items():
            doc = desktop.loadComponentFromURL(
                "file://%s/book.%s" % (OUT, ext), "_blank", 0, pv(Hidden=True))
            for operation in operations:
                DISPATCH[operation[0]](doc, *operation[1:])
            path = "%s/expected_%s.%s" % (OUT, name, ext)
            doc.storeToURL("file://" + path, pv(FilterName=filter_name))
            fixture_determinism.normalize(path)
            doc.close(True)


if __name__ == "__main__":
    main(set(sys.argv[1:]))
