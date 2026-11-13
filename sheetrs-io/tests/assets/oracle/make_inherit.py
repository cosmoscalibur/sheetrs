"""Makes `sheetrs-io/tests/assets/inherit.{xlsx,ods}`: the probes of the rules by which
LibreOffice gives a formula cell the number format of its result (see `sheetrs-query/src/inherit.rs`).

Column A holds one operand per row (a date, a time, a percentage, a currency, ...); the formulas of
the row combine it with operators and functions, each in two forms: as written, and inside
`IF(TRUE,<f>,RAND())`, which makes the cell volatile (LibreOffice interprets volatile cells when it
opens an XLSX, the others keep the format the file says).

LibreOffice writes the files; the styles of the formula cells are then taken out (the cell keeps
the default style, `General`), so that what the file says is nothing and what LibreOffice shows is
what its interpreter infers. Run it with a LibreOffice listening on a socket (see
`scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python make_inherit.py
"""
import os, re, shutil, tempfile, time, zipfile
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale

import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)

OPERANDS = [  # (value, format code)
    (45000, "YYYY-MM-DD"),
    (0.25, "HH:MM:SS"),
    (0.5, "0%"),
    (10, "[$$-409]#,##0.00"),
    (1234.5, "#,##0.00"),
    (None, "BOOLEAN"),
    (1234.5, "0.0E+0"),
    (5, "General"),
    (45000.5, "YYYY-MM-DD HH:MM"),
    (0.75, "[HH]:MM:SS"),
    (10, "[$EUR-407] #,##0.00"),
    (0.125, "0.00%"),
]
# `{a}` is the operand of the row; the cells A1:A4 are a date, a time, a percentage and a currency.
FORMS = [
    "{a}", "{a}+$A$1", "$A$1+{a}", "{a}+$A$2", "$A$2+{a}", "{a}+$A$3", "$A$3+{a}", "{a}+$A$4",
    "$A$4+{a}", "{a}-$A$1", "$A$1-{a}", "{a}-$A$2", "{a}+1", "1+{a}", "{a}+{a}", "{a}*2",
    "2*{a}", "{a}*$A$4", "{a}/2", "2/{a}", "{a}/$A$4", "{a}^2", "-{a}", "{a}%", "{a}&\"x\"",
    "{a}={a}", "{a}>1", "NOT({a})", "SUM({a},1)", "SUM(1,{a})", "SUM({a},{a})", "SUM({a}:{a})",
    "MIN({a},1)", "MAX({a},1)", "AVERAGE({a},1)", "PRODUCT({a},1)", "ABS({a})", "ROUND({a},0)",
    "IF(TRUE,{a},1)", "IF(FALSE,1,{a})", "CHOOSE(1,{a},1)", "INDEX({a}:{a},1)",
    "IFERROR({a},0)", "ISNUMBER({a})", "OFFSET({a},0,0)", "INDIRECT(\"A{n}\")",
]
GLOBALS = ["DATE(2024,1,2)", "TODAY()", "NOW()", "TIME(1,2,3)", "TRUE", "RAND()>2", "DATE(2024,1,2)+1",
           "SUM(A1:A4)", "MIN(A1:A4)", "MAX(A1:A4)", "AVERAGE(A1:A4)", "A1-A9", "A9-A1", "A2+A10",
           "A1+A3", "A3+A3", "A3*A3", "A4*A4", "A4*A11", "A11+A4", "A4+A11"]


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def col(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


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
    if value is None:
        cell.Formula = "=TRUE()"
    else:
        cell.Value = value
    cell.NumberFormat = key(code)


def put(c, r, text):
    sheet.getCellByPosition(c, r).Formula = "=" + text.replace(",", ";")


for i in range(len(OPERANDS)):
    for j, form in enumerate(FORMS):
        body = form.format(a=f"$A${i + 1}", n=i + 1)
        put(1 + 2 * j, i, body)
        put(2 + 2 * j, i, f"IF(TRUE,{body},RAND())")
base = len(OPERANDS) + 1
for j, form in enumerate(GLOBALS):
    put(1 + 2 * j, base, form)
    put(2 + 2 * j, base, f"IF(TRUE,{form},RAND())")
doc.calculateAll()

tmp = tempfile.mkdtemp()
for ext, filt in (("xlsx", "Calc MS Excel 2007 XML"), ("ods", "calc8")):
    path = os.path.join(tmp, f"inherit.{ext}")
    doc.storeToURL("file://" + path, (prop("FilterName", filt),))
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
    # A formula cell takes the default style (index 0): `<c r=".." s="N" ..><f>` -> no `s`.
    return re.sub(r'(<c r="[A-Z]+\d+")( s="\d+")?([^>]*>)(<f)', r"\1\3\4", xml)


def ods_edit(xml):
    xml = re.sub(r'(<table:table-cell )([^>]*?)table:style-name="[^"]*" ?([^>]*table:formula=)', r"\1\2\3", xml)
    xml = re.sub(r'(<table:table-cell )([^>]*table:formula=[^>]*?) table:style-name="[^"]*"', r"\1\2", xml)
    xml = re.sub(r' table:default-cell-style-name="[^"]*"', "", xml)
    return xml


rewrite(os.path.join(tmp, "inherit.xlsx"), "xl/worksheets/sheet1.xml", xlsx_edit)
rewrite(os.path.join(tmp, "inherit.ods"), "content.xml", ods_edit)
for ext in ("xlsx", "ods"):
    shutil.copy(os.path.join(tmp, f"inherit.{ext}"), os.path.join(OUT, f"inherit.{ext}"))
    fixture_determinism.normalize(os.path.join(OUT, f"inherit.{ext}"), freeze_volatile=True)
print("wrote", OUT)
