"""What `CELL("width"|"format"|"color"|"prefix"|"protect"|"parentheses")` of LibreOffice 26.8 depends on (a probe, not a test): the
evidence of why the engine does not answer them (`sheetrs-engine/src/functions/info.rs`, `notes::CELL_FORMAT`) and why a workbook that
holds one is not calculated by the default mode (`docs/approved-differences.md`, *Proposed*).

The same formulas, over the same cell, give another result when only the column width, the number format, the alignment, the
protection or the colour of negative numbers of the cell changes: the information is in the formatting of the cell, which neither the
model of the engine nor the cell the readers hand over (`Cell`: value, formula, number format) holds for the alignment, the protection
and the width.

    /opt/libreoffice26.8/program/python probe_cell_format.py       # LibreOffice listening on port 2002
"""
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table.CellHoriJustify import LEFT, RIGHT

local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
hidden = PropertyValue()
hidden.Name, hidden.Value = "Hidden", True
doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (hidden,))
sheet = doc.Sheets.getByIndex(0)
sheet.getCellRangeByName("A1").setValue(-5)
sheet.getCellRangeByName("B1").setString("text")
types = ("width", "format", "color", "prefix", "protect", "parentheses")
for k, t in enumerate(types):
    sheet.getCellByPosition(3 + k, 0).setFormula(f'=CELL("{t}";$A$1)')
sheet.getCellByPosition(9, 0).setFormula('=CELL("prefix";$B$1)')


def show(label):
    doc.calculateAll()
    values = [sheet.getCellByPosition(3 + k, 0).getString() for k in range(len(types))]
    print(f"{label:34}" + "  ".join(f"{t}={v}" for t, v in zip(types, values)) + f"  prefix(text)={sheet.getCellByPosition(9, 0).getString()}")


show("default")
sheet.Columns.getByIndex(0).Width = 6000
show("column A wider")
fmt = doc.NumberFormats
key = fmt.queryKey("0.00", doc.CharLocale, False)
if key == -1:
    key = fmt.addNew("0.00", doc.CharLocale)
sheet.getCellRangeByName("A1").NumberFormat = key
show("format 0.00")
key = fmt.addNew("0.0;[RED](0.0)", doc.CharLocale)
sheet.getCellRangeByName("A1").NumberFormat = key
show("format 0.0;[RED](0.0)")
sheet.getCellRangeByName("B1").HoriJustify = RIGHT
show("text aligned right")
sheet.getCellRangeByName("B1").HoriJustify = LEFT
show("text aligned left")
prot = sheet.getCellRangeByName("A1").CellProtection
prot.IsLocked = False
sheet.getCellRangeByName("A1").CellProtection = prot
show("A1 not locked")
doc.close(True)
