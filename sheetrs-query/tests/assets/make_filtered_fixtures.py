"""Makes `filtered.{xlsx,ods}` and `filtered_range.{xlsx,ods}`, the workbooks of the filtered-row
tests (`Find::filtered`, `SearchFiltered`): LibreOffice builds a sheet `T` with a key column and a
note column, filters A1:B6 on `key = x` (rows 3 and 5 are hidden by the filter; row 1, the header,
is filtered as well, as LibreOffice shows it), and has row 10 hidden by hand, outside the filter.

- `filtered.*`: the filter is the autofilter of an anonymous database range (a filter of the sheet).
  LibreOffice writes it as `<autoFilter>` of the worksheet in the XLSX and reads it back as a filter.
- `filtered_range.*`: the filter belongs to a named database range. An XLSX keeps it in a table part
  and LibreOffice does not read that one back as a filter; an ODS does.

Run it with a LibreOffice listening on a socket (`scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python make_filtered_fixtures.py
"""
import os, time
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.sheet import TableFilterField2
from com.sun.star.sheet.FilterOperator2 import EQUAL
from com.sun.star.table import CellRangeAddress

import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

HERE = os.path.dirname(os.path.abspath(__file__))
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def criterion():
    f = TableFilterField2()
    f.Field, f.IsNumeric, f.StringValue, f.Operator = 0, False, "x", EQUAL
    return f


def build(stem, named):
    doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
    sheet = doc.Sheets.getByIndex(0)
    sheet.Name = "T"
    sheet.getCellByPosition(0, 0).String = "key"
    sheet.getCellByPosition(1, 0).String = "note"
    for i, (key, note) in enumerate(
        [("x", "hit a"), ("y", "hit b"), ("x", "hit c"), ("y", "hit d"), ("x", "hit e")]
    ):
        sheet.getCellByPosition(0, i + 1).String = key
        sheet.getCellByPosition(1, i + 1).String = note
    sheet.getCellByPosition(1, 8).String = "hit outside visible"
    sheet.getCellByPosition(1, 9).String = "hit outside hidden"
    sheet.Rows.getByIndex(9).IsVisible = False
    if named:
        area = CellRangeAddress()
        area.Sheet, area.StartColumn, area.StartRow, area.EndColumn, area.EndRow = 0, 0, 0, 1, 5
        doc.DatabaseRanges.addNewByName("f", area)
        db = doc.DatabaseRanges.getByName("f")
        db.AutoFilter = True
        db.getFilterDescriptor().setFilterFields2((criterion(),))
        db.refresh()
    else:
        controller = doc.getCurrentController()
        rng = sheet.getCellRangeByName("A1:B6")
        controller.select(rng)
        dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
        dispatcher.executeDispatch(controller.getFrame(), ".uno:DataFilterAutoFilter", "", 0, ())
        time.sleep(1)
        descriptor = rng.createFilterDescriptor(True)
        descriptor.setFilterFields2((criterion(),))
        rng.filter(descriptor)
    doc.storeToURL("file://" + os.path.join(HERE, stem + ".xlsx"), (pv("FilterName", "Calc MS Excel 2007 XML"),))
    doc.storeToURL("file://" + os.path.join(HERE, stem + ".ods"), (pv("FilterName", "calc8"),))
    for ext in ("xlsx", "ods"):
        fixture_determinism.normalize(os.path.join(HERE, stem + "." + ext))
    doc.close(True)


build("filtered", False)
build("filtered_range", True)
