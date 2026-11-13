"""Makes `sheetrs-io/tests/assets/inherit_more.{xlsx,ods}`: probes of the rules by which LibreOffice
gives a formula cell the number format of its result when the formula chooses a cell with `IF`,
`CHOOSE` or `INDIRECT` through an argument that has to be worked out (a function of constants and
cells) or through a defined name (see `sheetrs-query/src/inherit.rs`).

A1:A3 hold a date, a currency and a percentage; the defined names point at them (`Dt`, `Cur`,
`Pct`, `Rng`) or hold a constant (`Idx`, `Flag`, `Ref`). Each formula is written twice: as it is,
and inside `IF(TRUE,<f>,RAND())`, which makes the cell volatile (LibreOffice interprets volatile
cells when it opens an XLSX; the others keep the format the file says). The styles of the formula
cells are taken out afterwards, as in `make_inherit.py`.

    /opt/libreoffice26.8/program/python make_inherit_more.py
"""
import os, re, shutil, tempfile, time, zipfile
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale
from com.sun.star.table import CellAddress

import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)

OPERANDS = [(45000, "YYYY-MM-DD"), (10, "[$$-409]#,##0.00"), (0.5, "0%")]
FORMULAS = [
    # The condition or the index is a function of constants and cells.
    "IF(ISNUMBER($A$1),$A$1,$A$2)", "IF(ISTEXT($A$1),$A$1,$A$2)", "IF(LEN(\"ab\")>1,$A$1,$A$2)",
    "IF(ABS(-1)=1,$A$1,$A$2)", "IF(AND(ISNUMBER($A$1),$A$3>0),$A$1,$A$2)", "IF(ROW($A$2)=2,$A$1,$A$2)",
    "IF(ISBLANK($A$9),$A$3,$A$1)", "IF(SUM($A$1:$A$3)>1,$A$2,$A$1)", "IF(NOT(ISNUMBER($A$1)),$A$1,$A$3)",
    "CHOOSE(ROUND(1.4,0),$A$1,$A$2)", "CHOOSE(MOD(5,2)+1,$A$1,$A$2)", "CHOOSE(COUNT($A$1:$A$3),$A$1,$A$2,$A$3)",
    "CHOOSE(MAX(1,2),$A$1,$A$2)", "CHOOSE(IF(TRUE,2,1),$A$1,$A$2)", "CHOOSE(INT(2.9),$A$1,$A$2,$A$3)",
    "INDIRECT(\"A\"&ROW($A$3))", "INDIRECT(CONCATENATE(\"A\",1))", "INDIRECT(ADDRESS(1,1))",
    "INDIRECT(UPPER(\"a1\"))", "INDIRECT(LEFT(\"A2x\",2))", "INDIRECT(\"A\"&LEN(\"ab\"))",
    "INDIRECT(\"A\"&COLUMN($B$1))", "INDIRECT(\"A\"&SUM(1,1))", "INDIRECT(\"A\"&COUNT($A$1:$A$3))",
    # A defined name stands for a cell, a range or a constant.
    "Dt", "Dt+1", "Cur", "Cur*2", "Pct", "Dt+Cur", "IF(TRUE,Dt,1)", "IF(Flag,$A$1,$A$2)",
    "IF(Idx=2,Dt,Cur)", "CHOOSE(Idx,$A$1,$A$2,$A$3)", "INDIRECT(Ref)", "INDIRECT(\"Dt\")",
    "SUM(Rng)", "MIN(Rng)", "MAX(Rng)", "AVERAGE(Rng)", "INDEX(Rng,2)", "OFFSET(Dt,1,0)",
    "OFFSET(Dt,Idx-1,0)", "SUM(Dt,Cur)", "Idx", "Flag", "Idx+Dt", "Dt&\"x\"", "IF(Flag,Dt,Cur)+1",
    "IFERROR(Dt,0)", "ABS(Dt)", "Dt-Dt",
]
NAMES = [("Dt", "$Inherit.$A$1"), ("Cur", "$Inherit.$A$2"), ("Pct", "$Inherit.$A$3"),
         ("Rng", "$Inherit.$A$1:$A$3"), ("Idx", "2"), ("Flag", "TRUE()"), ("Ref", "\"A2\"")]


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (prop("Hidden", True),))
sheet = doc.Sheets.getByIndex(0)
sheet.Name = "Inherit"
formats = doc.NumberFormats
en = Locale("en", "US", "")


def key(code):
    k = formats.queryKey(code, en, False)
    return k if k != -1 else formats.addNew(code, en)


for i, (value, code) in enumerate(OPERANDS):
    cell = sheet.getCellByPosition(0, i)
    cell.Value = value
    cell.NumberFormat = key(code)
origin = CellAddress()
origin.Sheet, origin.Column, origin.Row = 0, 0, 0
for name, content in NAMES:
    doc.NamedRanges.addNewByName(name, content, origin, 0)
for j, body in enumerate(FORMULAS):
    sheet.getCellByPosition(1, 20 + j).Formula = "=" + body.replace(",", ";")
    sheet.getCellByPosition(2, 20 + j).Formula = "=IF(TRUE;" + body.replace(",", ";") + ";RAND())"
doc.calculateAll()

tmp = tempfile.mkdtemp()
for ext, filt in (("xlsx", "Calc MS Excel 2007 XML"), ("ods", "calc8")):
    doc.storeToURL("file://" + os.path.join(tmp, f"inherit_more.{ext}"), (prop("FilterName", filt),))
doc.close(True)


def rewrite(path, member, edit):
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename == member:
                data = edit(data.decode("utf-8")).encode("utf-8")
            z.writestr(info, data)


def xlsx_edit(xml):
    return re.sub(r'(<c r="[A-Z]+\d+")( s="\d+")?([^>]*>)(<f)', r"\1\3\4", xml)


def ods_edit(xml):
    xml = re.sub(r'(<table:table-cell )([^>]*?)table:style-name="[^"]*" ?([^>]*table:formula=)', r"\1\2\3", xml)
    xml = re.sub(r'(<table:table-cell )([^>]*table:formula=[^>]*?) table:style-name="[^"]*"', r"\1\2", xml)
    return re.sub(r' table:default-cell-style-name="[^"]*"', "", xml)


rewrite(os.path.join(tmp, "inherit_more.xlsx"), "xl/worksheets/sheet1.xml", xlsx_edit)
rewrite(os.path.join(tmp, "inherit_more.ods"), "content.xml", ods_edit)
for ext in ("xlsx", "ods"):
    shutil.copy(os.path.join(tmp, f"inherit_more.{ext}"), os.path.join(OUT, f"inherit_more.{ext}"))
    fixture_determinism.normalize(os.path.join(OUT, f"inherit_more.{ext}"))
print("wrote", OUT)
