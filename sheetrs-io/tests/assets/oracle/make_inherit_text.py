"""Makes `sheetrs-io/tests/assets/inherit_text.{xlsx,ods}`: probes of the format LibreOffice gives a
formula cell whose operands are cells that hold text, are formatted as text (`@`) or are empty (see
`sheetrs-query/src/inherit.rs`, the aggregates).

A1:A14 hold, in this order: text formatted `@`, a numeric-looking text (`007`) formatted `@`, a number
formatted `@`, text in `General`, a number in `#,##0.00`, a date, an empty cell formatted `0.00%`, a
currency, a percentage, a number in `General`, a date in a text-formatted-looking text (`2023-03-15`) formatted `@`,
a time, a text in a date format and an empty cell in a currency format. Each formula is written twice: as it is, and inside `IF(TRUE,<f>,RAND())`, which makes the
cell volatile (LibreOffice interprets volatile cells when it opens an XLSX; the others keep the format
the file says). The styles of the formula cells are taken out afterwards, as in `make_inherit.py`.

    /opt/libreoffice26.8/program/python make_inherit_text.py
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

TEXT = "@"
# (kind, content, format code): kind "s" text, "n" number, "e" empty (only the format)
OPERANDS = [
    ("s", "abc", TEXT), ("s", "007", TEXT), ("n", 5, TEXT), ("s", "x", None), ("n", 10, "#,##0.00"),
    ("n", 45000, "YYYY-MM-DD"), ("e", None, "0.00%"), ("n", 10, "[$$-409]#,##0.00"), ("n", 0.5, "0%"),
    ("n", 7, None), ("s", "2023-03-15", TEXT), ("n", 0.75, "HH:MM:SS AM/PM"),
]
CELLS = [f"$A${i}" for i in range(1, 13)]
FORMULAS = []
for c in CELLS:
    # A reference to an empty cell that has a format is left out: the reader does not report it.
    FORMULAS += ([] if c == "$A$7" else [c]) + [c + "&\"x\"", c + "+1", c + "*2", "-" + c]
FUNCS = ["SUM", "MIN", "MAX", "AVERAGE", "PRODUCT", "SUMSQ", "COUNT", "COUNTA", "MEDIAN", "AVERAGEA", "MAXA", "MINA"]
ARGS = [
    "$A$1", "$A$2", "$A$3", "$A$4", "$A$7", "$A$11", "$A$1,$A$2", "$A$1,$A$5", "$A$5,$A$1", "$A$4,$A$6",
    "$A$7,$A$6", "$A$1,$A$7", "$A$7,$A$1", "$A$1:$A$2", "$A$1:$A$4", "$A$1:$A$6", "$A$3:$A$5", "$A$1:$A$3",
    "$A$7:$A$10", "$A$1:$A$12", "1,$A$1", "$A$1,1", "$A$3,$A$5", "$A$5,$A$3",
    "$A$3,$A$6", "$A$11,$A$6", "$A$6,$A$11", "$A$7,$A$7", "1,$A$6", "$A$6,1", "$A$10,$A$6", "$A$6,$A$10",
    "$A$5,$A$6", "$A$6,$A$5", "$A$8,$A$6", "$A$6,$A$8", "$A$9,$A$6", "$A$12,$A$6", "$A$6,$A$12", "$A$10,$A$8",
    "$A$13", "$A$14", "$A$13,$A$8", "$A$8,$A$13",
    "$A$8,$A$10", "$A$5,$A$8", "$A$9,$A$8", "$A$8,$A$9", "$A$6,$A$7,$A$8", "$A$7,$A$8,$A$6", "$A$4,$A$1,$A$6",
]
for f in FUNCS:
    for a in ARGS:
        if f == "SUMSQ" and ("$A$6" in a or "$A$12" in a):
            continue  # the square of a date or a time is not a date any more
        FORMULAS.append(f"{f}({a})")
for f in ("IF($A$1=\"abc\",$A$1,$A$5)", "IF($A$4=\"x\",$A$2,$A$5)", "IF(TRUE,$A$1,$A$5)", "CHOOSE(1,$A$1,$A$5)",
          "INDEX($A$1:$A$3,1)", "INDEX($A$1:$A$3,3)", "OFFSET($A$1,1,0)", "INDIRECT(\"A1\")", "INDIRECT(\"A3\")",
          "IFERROR($A$1,0)", "IFERROR($A$3,0)", "SUM(IF(TRUE,$A$1,0))"):
    FORMULAS.append(f)


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


for i, (kind, value, code) in enumerate(OPERANDS):
    cell = sheet.getCellByPosition(0, i)
    if kind == "s":
        cell.String = value
    elif kind == "n":
        cell.Value = value
    if code:
        cell.NumberFormat = key(code)
for j, body in enumerate(FORMULAS):
    sheet.getCellByPosition(1 + 2 * j, 20).Formula = "=" + body.replace(",", ";")
    sheet.getCellByPosition(2 + 2 * j, 20).Formula = "=IF(TRUE;" + body.replace(",", ";") + ";RAND())"
doc.calculateAll()

tmp = tempfile.mkdtemp()
for ext, filt in (("xlsx", "Calc MS Excel 2007 XML"), ("ods", "calc8")):
    doc.storeToURL("file://" + os.path.join(tmp, f"inherit_text.{ext}"), (prop("FilterName", filt),))
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


rewrite(os.path.join(tmp, "inherit_text.xlsx"), "xl/worksheets/sheet1.xml", xlsx_edit)
rewrite(os.path.join(tmp, "inherit_text.ods"), "content.xml", ods_edit)
for ext in ("xlsx", "ods"):
    shutil.copy(os.path.join(tmp, f"inherit_text.{ext}"), os.path.join(OUT, f"inherit_text.{ext}"))
    fixture_determinism.normalize(os.path.join(OUT, f"inherit_text.{ext}"))
print("wrote", OUT)
