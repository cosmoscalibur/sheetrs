"""Does the conversion of formulas to values of LibreOffice 26.8 (*Data > Calculate > Formula to Value*, `.uno:ConvertFormulaToValue`) force a calculation of
what it converts, or does it write the value the cell holds? The evidence of `sheetrs-edit/src/plan.rs` (a conversion is one step in every mode).

    /opt/libreoffice26.8/program/python probe_convert.py DIR     # DIR is `tests/assets/cached`; LibreOffice listening on port 2002

`book.{xlsx,ods}` cache `from-excel` in `Calc!A7` where the formula gives `t`, `140` in `A6` where `rate*2` gives 14 and `99` in `A8` where `ROW()` gives 8. The file
is opened (macro security as installed), the range `Calc!A1:A9` is converted by the dispatcher, and the values are read: a cell that shows what the file caches was
converted as it was; one that shows what the formula gives was calculated, by the load (an ODS number in a cell with no format) or by the conversion.
"""
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def pv(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v; return p


local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
for ext in ("xlsx", "ods"):
    doc = desktop.loadComponentFromURL(f"file://{sys.argv[1]}/book.{ext}", "_blank", 0, ())
    calc = doc.Sheets.getByName("Calc")
    before = [calc.getCellByPosition(0, i).getString() for i in range(9)]
    formulas_before = [calc.getCellByPosition(0, i).getFormula() for i in range(9)]
    controller = doc.getCurrentController()
    controller.setActiveSheet(calc)
    controller.select(calc.getCellRangeByName("A1:A9"))
    dispatcher.executeDispatch(controller.getFrame(), ".uno:ConvertFormulaToValue", "", 0, ())
    time.sleep(0.5)
    after = [calc.getCellByPosition(0, i).getString() for i in range(9)]
    formulas_after = [calc.getCellByPosition(0, i).getFormula() for i in range(9)]
    print(ext, "values at load   ", before)
    print(ext, "values converted ", after)
    print(ext, "formulas gone    ", [f == "" or not f.startswith("=") for f in formulas_after])
    doc.close(True)
