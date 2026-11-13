"""Makes `tests/lint_book.{xlsx,ods}`: one workbook with at least one violation of every linter rule a
file can trigger, plus clean sheets, so that the tests of the linter, the statistics, the readers and
the editor have an integrated book of the kind that the old `minimal_test` was, without its cache errors.

LibreOffice 26.8 writes both files from the same document in memory, after `calculateAll()`, so every
cached value is what LibreOffice itself computes (the oracle `lint_book.<ext>.oracle.json` records it).
Each file is judged by what LibreOffice says of that file.
A few facts a document in memory cannot express are written into the saved archives afterwards:

* XLSX: the sheet `Vault` is made `veryHidden` (LibreOffice has no such state), and a `xl/vbaProject.bin`
  part is added (VBA1101); the ODS gets a Basic library through UNO instead.
* Both: the last-modified date of the metadata is set to 2020-01-15 (FILE1002 is a matter of age).

Run it with a LibreOffice listening on a socket (see `scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python tests/make_lint_book.py

Where each rule is triggered (default configuration of the linter unless said otherwise):

| Rule | Sheet and cells |
| ---- | --------------- |
| ERR101 | name `Gone` (its sheet was deleted) |
| ERR102 | `Errors!A1` (`=1/0`), `A2` (`=NA()`), `A5` (`=EXP(1000)`) and the end of the chain `A4` |
| ERR103 | `Errors!A3` and `A4`, formulas that read an error cell |
| CALC201 | `Calc!B2`, `B3` (a decimal literal in the formula) |
| CALC202 | `Loops!A1:A2` (a loop) |
| CALC203 | `Calc!B4` (`++`), `B5` (`+-`) |
| CALC204 | `Calc!B6` (`VLOOKUP` without the match flag), `B7` (`MATCH` without it) |
| CALC205 | `Calc!B8` (`SUM` of two overlapping ranges) |
| REF301 | name `TaxRate` (used by no formula) |
| REF302 | sheets `Sheet2` and `Sheet 2` |
| REF303 | sheet `Empty` |
| REF304 | `Blanks` (a formatted cell far past the data) |
| REF305 | `Blanks` (blank rows and columns before the data and between blocks) |
| REF306 | sheets `Unlinked` and `Empty` (no link with any other sheet) |
| REF307 | `Refs!B1` (`A:A`), `B2` (`2:2`) |
| REF309 | `Refs!B3` (reads the empty `D50`) |
| REF310 | `Refs!B4` (`SUM(A2:A5)` with numbers still in `A6:A8`) |
| REF311 | name `Local`, defined for the book and for the sheet `Calc` |
| INT401 | `Seq!D6` (a constant in a column of formulas) |
| INT402 | `Seq!D11` (a hole in the column) |
| INT403 | `Seq!D16` (a formula with another pattern) |
| CPX502 | `Complexity!A30:C30`, `A32:B33` (merged cells) |
| CPX503 | `Complexity!F1:F10` (six conditional formats) |
| CPX504 | `Complexity!B1` (IF nested six deep) |
| CPX505 | `Complexity!B2` (seven nested parentheses) |
| CPX506 | `Complexity!B3` (nine operators) |
| CPX507 | `Complexity!B4` (reads three sheets), `Summary!B2` |
| CPX508 | `Complexity!B5` (eleven references) |
| CPX509 | `Complexity!B6` (more than 255 characters); `B7` is not (its prefix is not counted) |
| VUL601 | `Vulnerable!G1:G3` (the same formula in adjacent cells) |
| VUL602 | `Vulnerable!B1:B5` (`NOW`, `RAND`, `OFFSET`, `INDIRECT`, `CELL`) |
| VUL603 | `Vulnerable!C1:C2` (`=""` and `LEN()=0`) |
| VUL604 | `Vulnerable!D1:D3` (`VLOOKUP`, `HLOOKUP`, `LOOKUP`) |
| VUL605 | `Vulnerable!H1:H3`, `J1:L1` (matrix formulas) |
| VUL606 | `Vulnerable!E1:E8` (deprecated functions) |
| DATA701 | sheets `Sheet2`, `Sheet 2` (default names) |
| DATA702 | `Data!B2:B4` (numbers kept as text) |
| DATA703 | `Data!C2`, `C4`, `C5` (date formats other than `mm/dd/yyyy`) |
| DATA704 | `Data!D2` (text of 260 characters) |
| DATA705 | `Data!E2:E3` (spaces before and after) |
| DATA706 | `Data!F2:F3` (formulas that read text that looks like a number) |
| EXT801 | `Links!B1:B3` (formulas that read another workbook) |
| EXT802 | `Links!A1:A2` (the same URL), `A4` |
| HID901 | sheets `Hidden` and `Vault` (`veryHidden` in the XLSX only) |
| HID902 | `Hide!B:D`, `F` (columns), `3:7` (rows) |
| FILE1002 | the metadata date (2020-01-15) |
| VBA1101 | `xl/vbaProject.bin` (XLSX) and a `Basic/` library (ODS) |
| CPX501, FILE1001 | with a configuration (`max_sheets`, `max_file_size_mb`): no file is big enough |
| FILE1003 | not in this book: the 1904 date system changes the serial of every date; a unit test of the rule and `sheetrs-io/tests/date_system_tests.rs` cover it |

A sheet with no violation is `Summary` apart from `CPX507`, and `Inputs`, which is clean.
"""
import os, re, shutil, tempfile, time, zipfile
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale
from com.sun.star.table import CellAddress
from com.sun.star.awt import Rectangle

import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

HERE = os.path.dirname(os.path.abspath(__file__))


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(30):
        try:
            return resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        except Exception:
            time.sleep(1)
    raise SystemExit("no LibreOffice listening on port 2002")


ctx = connect()
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def new_doc():
    return desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (prop("Hidden", True),))


# The workbook that the external links of `Links` read (its cached values go into the book). It sits in
# the folder where the books are saved, so that LibreOffice writes its location relative to the book
# (`prices.ods`), as it does when a person saves next to the file they link to.
# A folder with a fixed name (not `mkdtemp()`): LibreOffice writes the path of the external workbook into
# the ODS, and a random one would change the file on every run.
EXTERNAL = os.path.join(tempfile.gettempdir(), "sheetrs-lint-book")
shutil.rmtree(EXTERNAL, ignore_errors=True)
os.makedirs(EXTERNAL)
ext = new_doc()
ext.Sheets.getByIndex(0).Name = "Prices"
ext.Sheets.getByIndex(0).getCellRangeByName("B1").Value = 10
ext.Sheets.getByIndex(0).getCellRangeByName("B2").Value = 20
ext.Sheets.getByIndex(0).getCellRangeByName("B3").Value = 30
ext.storeToURL("file://" + EXTERNAL + "/prices.ods", (prop("FilterName", "calc8"),))
ext.close(True)

doc = new_doc()
en = Locale("en", "US", "")
formats = doc.NumberFormats


def fmt_key(code):
    k = formats.queryKey(code, en, False)
    return k if k != -1 else formats.addNew(code, en)


def add_sheet(name):
    sheets = doc.Sheets
    sheets.insertNewByName(name, sheets.Count)
    return sheets.getByName(name)


first = doc.Sheets.getByIndex(0)
first.Name = "Summary"
SHEETS = ["Inputs", "Calc", "Loops", "Errors", "Refs", "Blanks", "Empty", "Unlinked", "Seq", "Complexity",
          "Vulnerable", "Data", "Sheet2", "Sheet 2", "Hide", "Links", "Hidden", "Vault", "Escapes"]
S = {"Summary": first}
for n in SHEETS:
    S[n] = add_sheet(n)


def cell(sheet, a1):
    return S[sheet].getCellRangeByName(a1)


def text(sheet, a1, value):
    cell(sheet, a1).String = value


def num(sheet, a1, value):
    cell(sheet, a1).Value = value


def f(sheet, a1, formula):
    cell(sheet, a1).Formula = formula


# ---- Inputs: clean data with a note and a validation ----------------------------------------------
for j, h in enumerate(["Item", "Qty", "Price"]):
    S["Inputs"].getCellByPosition(j, 0).String = h
items = [("Apples", 10, 1.5), ("Pears", 7, 2.25), ("Plums", 12, 0.8), ("Kiwis", 5, 3.1), ("Figs", 9, 4.0)]
for i, (a, q, p) in enumerate(items, start=1):
    S["Inputs"].getCellByPosition(0, i).String = a
    S["Inputs"].getCellByPosition(1, i).Value = q
    S["Inputs"].getCellByPosition(2, i).Value = p
text("Inputs", "D1", "Rate")
num("Inputs", "D2", 0.19)
S["Inputs"].Annotations.insertNew(CellAddress(1, 0, 0), "Quantity in units")
v = cell("Inputs", "E2:E6").Validation
v.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
v.ShowErrorMessage = True
v.setFormula1('"open";"closed"')
cell("Inputs", "E2:E6").Validation = v

# ---- Names -----------------------------------------------------------------------------------------
doc.NamedRanges.addNewByName("Rate", "$Inputs.$D$2", CellAddress(0, 0, 0), 0)
doc.NamedRanges.addNewByName("TaxRate", "$Inputs.$D$2", CellAddress(0, 0, 0), 0)
doc.NamedRanges.addNewByName("Local", "$Inputs.$B$2", CellAddress(0, 0, 0), 0)
S["Calc"].NamedRanges.addNewByName("Local", "$Inputs.$B$3", CellAddress(0, 0, 0), 0)
doc.Sheets.insertNewByName("Gone sheet", doc.Sheets.Count)
doc.NamedRanges.addNewByName("Gone", "$'Gone sheet'.$A$1", CellAddress(0, 0, 0), 0)

# ---- Calc: CALC201 to CALC205 ----------------------------------------------------------------------
text("Calc", "A1", "Formula problems")
f("Calc", "B2", "=$Inputs.B2*1.19")
f("Calc", "B3", "=$Inputs.C2*3.14")
f("Calc", "B4", "=$Inputs.B2++$Inputs.B3")
f("Calc", "B5", "=$Inputs.B2+-$Inputs.B3")
f("Calc", "B6", "=VLOOKUP($Inputs.A2;$Inputs.A2:C6;2)")
f("Calc", "B7", "=MATCH($Inputs.B2;$Inputs.B2:B6)")
f("Calc", "B8", "=SUM($Inputs.B2:B6;$Inputs.B4:B6)")
f("Calc", "B9", "=Local+Rate")

# ---- Loops: CALC202 --------------------------------------------------------------------------------
f("Loops", "A1", "=A2")
f("Loops", "A2", "=A1+1")

# ---- Errors: ERR102, ERR103 ------------------------------------------------------------------------
f("Errors", "A1", "=1/0")
f("Errors", "A2", "=NA()")
f("Errors", "A3", "=A1+1")
f("Errors", "A4", "=A3*2")
f("Errors", "A5", "=EXP(1000)")
text("Errors", "B1", "ok")

# ---- Refs: REF307, REF309, REF310 ------------------------------------------------------------------
for i in range(1, 9):
    num("Refs", f"A{i}", i * 10)
f("Refs", "B1", "=SUM($Inputs.B:B)")
f("Refs", "B2", "=SUM($Inputs.2:2)")
f("Refs", "B3", "=D50*2")
f("Refs", "B4", "=SUM(A2:A5)")

# ---- Blanks: REF304, REF305 ------------------------------------------------------------------------
for i, row in enumerate([(1, 2), (3, 4)]):
    for j, v_ in enumerate(row):
        num("Blanks", f"{'DE'[j]}{4 + i}", v_)
num("Blanks", "D11", 5)
num("Blanks", "E11", 6)
cell("Blanks", "H16").CellBackColor = 0xFFFF00

# ---- Unlinked: REF306 ------------------------------------------------------------------------------
text("Unlinked", "A1", "Scratch")
num("Unlinked", "A2", 1)

# ---- Seq: INT401, INT402, INT403 -------------------------------------------------------------------
for r in range(2, 20):
    num("Seq", f"B{r}", r)
    num("Seq", f"C{r}", r * 2)
for r in range(2, 20):
    if r == 6:
        num("Seq", f"D{r}", 99)       # INT401
    elif r == 11:
        continue                       # INT402: one absent cell between formulas
    elif r == 16:
        f("Seq", f"D{r}", f"=B{r}*C{r}")  # INT403
    else:
        f("Seq", f"D{r}", f"=B{r}+C{r}")

# ---- Complexity: CPX502 to CPX509 ------------------------------------------------------------------
for r in range(1, 14):
    num("Complexity", f"A{r}", r)
f("Complexity", "B1", "=IF(A1>1;1;IF(A1>2;2;IF(A1>3;3;IF(A1>4;4;IF(A1>5;5;IF(A1>6;6;0))))))")
f("Complexity", "B2", "=((((((A2))))))+0")
f("Complexity", "B3", "=A1+A2+A3+A4+A5+A6+A7+A8+A9+A10")
f("Complexity", "B4", "=$Inputs.B2+$Calc.B2+$Seq.B2")
f("Complexity", "B5", "=SUM(A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11)")
f("Complexity", "B6", '="' + "a" * 130 + '"&"' + "b" * 130 + '"')
# 253 characters as the user sees it, but the file writes `_xlfn.CONCAT` (XLSX) or `COM.MICROSOFT.CONCAT`
# (ODS) and the text is longer than 255: CPX509 measures what the user sees, so it is not reported.
f("Complexity", "B7", '=CONCAT("' + "a" * 120 + '";"' + "b" * 120 + '")')
text("Complexity", "A30", "Merged banner")
cell("Complexity", "A30:C30").merge(True)
cell("Complexity", "A32:B33").merge(True)
cf = cell("Complexity", "F1:F10").ConditionalFormat
for k in range(6):
    cf.addNew((prop("Operator", uno.Enum("com.sun.star.sheet.ConditionOperator", "GREATER")),
               prop("Formula1", str(k + 1)), prop("StyleName", "Good" if k % 2 else "Bad")))
cell("Complexity", "F1:F10").ConditionalFormat = cf
for r in range(1, 11):
    num("Complexity", f"F{r}", r)

# ---- Vulnerable: VUL601 to VUL606 ------------------------------------------------------------------
for r in range(1, 6):
    num("Vulnerable", f"A{r}", r)
f("Vulnerable", "B1", "=NOW()")
f("Vulnerable", "B2", "=RAND()")
f("Vulnerable", "B3", "=OFFSET(A1;1;0)")
f("Vulnerable", "B4", '=INDIRECT("A1")')
f("Vulnerable", "B5", '=CELL("address";A1)')
f("Vulnerable", "C1", '=A1=""')
f("Vulnerable", "C2", "=LEN(A2)=0")
f("Vulnerable", "D1", "=VLOOKUP(A1;A1:A5;1;0)")
f("Vulnerable", "D2", "=HLOOKUP(A1;A1:A5;1;0)")
f("Vulnerable", "D3", "=LOOKUP(A1;A1:A5;A1:A5)")
f("Vulnerable", "E1", '=CONCATENATE(A1;"x")')
f("Vulnerable", "E2", "=STDEV(A1:A5)")
f("Vulnerable", "E3", "=RANK(A1;A1:A5)")
f("Vulnerable", "E4", "=MODE(A1:A5)")
f("Vulnerable", "E5", "=PERCENTILE(A1:A5;0.5)")
f("Vulnerable", "E6", "=FORECAST(6;A1:A5;A1:A5)")
f("Vulnerable", "E7", "=CEILING(A1;1)")
f("Vulnerable", "E8", "=FLOOR(A1;1)")
for r in (1, 2, 3):
    f("Vulnerable", f"G{r}", "=$Inputs.$B$2*$Inputs.$C$2")
cell("Vulnerable", "H1:H3").setArrayFormula("=A1:A3*2")
cell("Vulnerable", "J1:L1").setArrayFormula("=TRANSPOSE(A1:A3)")

# ---- Data: DAT702 to DAT706 ------------------------------------------------------------------------
for a1, v_ in (("B2", "123"), ("B3", "007"), ("B4", "45.5")):
    text("Data", a1, v_)
for a1, code, serial in (("C2", "DD/MM/YYYY", 45000), ("C3", "MM/DD/YYYY", 45001),
                         ("C4", "YYYY-MM-DD", 45002), ("C5", "DD-MMM-YY", 45003)):
    num("Data", a1, serial)
    cell("Data", a1).NumberFormat = fmt_key(code)
text("Data", "D2", "long text " * 26)
text("Data", "E2", "North ")
text("Data", "E3", " South")
f("Data", "F2", "=B2+1")
f("Data", "F3", "=SUM(B2;B3)")

# ---- Sheet2, Sheet 2: DATA701, REF302 --------------------------------------------------------------
text("Sheet2", "A1", "one")
f("Sheet2", "A2", "=$Inputs.B2")
text("Sheet 2", "A1", "two")
f("Sheet 2", "A2", "=$Inputs.B3")

# ---- Hide: HID902 ----------------------------------------------------------------------------------
for r in range(1, 9):
    for c in range(0, 7):
        S["Hide"].getCellByPosition(c, r - 1).Value = r * 10 + c
f("Hide", "H1", "=SUM(A1:G8)")
cols = S["Hide"].Columns
for c in (1, 2, 3, 5):
    cols.getByIndex(c).IsVisible = False
rows = S["Hide"].Rows
for r in range(2, 7):
    rows.getByIndex(r).IsVisible = False

# ---- Links: EXT801, EXT802 -------------------------------------------------------------------------
for r in (1, 2, 3):
    f("Links", f"B{r}", f"='file://{EXTERNAL}/prices.ods'#$Prices.B{r}")
text("Links", "A1", "https://example.com/data")
text("Links", "A2", "https://example.com/data")
text("Links", "A4", "ftp://files.example.com/archive")

# ---- Escapes: text with quotes, `<`, `>` and `&` in formulas (XML escapes in the file) -------------
f("Escapes", "A1", '="Say ""hi"" <b>ok</b> & more"')
f("Escapes", "A2", '=IF($Inputs.B2>5;"big > 5";"small < 5")')
f("Escapes", "A3", '=IF(A1<>"";"has & text";"none")')
f("Escapes", "A4", "=$'Sheet 2'.A1")

# ---- Hidden sheets, HID901 -------------------------------------------------------------------------
text("Hidden", "A1", "kept out of sight")
num("Hidden", "A2", 1)
text("Vault", "A1", "never listed")
num("Vault", "A2", 2)
S["Hidden"].IsVisible = False
S["Vault"].IsVisible = False

# ---- Summary: clean, apart from CPX507 and the links of the other sheets ---------------------------
text("Summary", "A1", "Summary")
f("Summary", "B2", "=$Inputs.B2+$Calc.B2+$Data.F2")
f("Summary", "B3", "=SUM($Inputs.B2:B6)*Rate")
f("Summary", "B4", "=$Sheet2.A2+$'Sheet 2'.A2")
f("Summary", "B5", "=$Loops.A1+$Errors.A1")
f("Summary", "B6", "=$Refs.B4+$Blanks.D4+$Seq.D2+$Complexity.B1+$Vulnerable.A1")
f("Summary", "B7", "=$Hide.H1+$Links.B1+$Hidden.A2+$Vault.A2")
f("Summary", "B9", "=$Escapes.A3")
rng = S["Summary"].getCellRangeByName("A9:B12")
S["Summary"].Charts.addNewByName("Quantities", Rectangle(8000, 500, 9000, 6000),
                                 (S["Inputs"].getCellRangeByName("A1:B6").RangeAddress,), True, True)

doc.Sheets.removeByName("Gone sheet")
doc.calculateAll()
# A document made through UNO has no creation date; the readers give `created` when the file says it.
from com.sun.star.util import DateTime
created = DateTime()
created.Year, created.Month, created.Day, created.Hours, created.Minutes, created.Seconds = 2019, 6, 1, 9, 30, 0
doc.DocumentProperties.CreationDate = created

tmp = EXTERNAL
# A Basic library for the ODS (VBA1101).
try:
    lib = doc.BasicLibraries
    if not lib.hasByName("Standard"):
        lib.createLibrary("Standard")
    lib.getByName("Standard").insertByName("Module1", "Sub Main\nEnd Sub\n")
except Exception as error:  # pragma: no cover
    print("no Basic library:", error)
paths = {}
for ext_, filt in (("xlsx", "Calc MS Excel 2007 XML"), ("ods", "calc8")):
    paths[ext_] = os.path.join(tmp, f"lint_book.{ext_}")
    doc.storeToURL("file://" + paths[ext_], (prop("FilterName", filt),))
doc.close(True)


def rewrite(path, edits, add=()):
    """Rewrites the archive: `edits` maps a member to a function of its text, `add` lists new members."""
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename in edits:
                data = edits[info.filename](data.decode("utf-8")).encode("utf-8")
            z.writestr(info, data)
        for name, data in add:
            z.writestr(name, data)


OLD = "2020-01-15T10:00:00"


def xlsx_workbook(xml):
    return re.sub(r'(<sheet name="Vault"[^>]*?)state="hidden"', r'\1state="veryHidden"', xml)


def xlsx_core(xml):
    xml = re.sub(r"(<dcterms:modified[^>]*>)[^<]*", r"\g<1>" + OLD + "Z", xml)
    return xml


def xlsx_types(xml):
    if "bin" in xml:
        return xml
    return xml.replace("<Default ", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)


def xlsx_rels(xml):
    return xml.replace(
        "</Relationships>",
        '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
        'Target="vbaProject.bin"/></Relationships>')


rewrite(paths["xlsx"], {"xl/workbook.xml": xlsx_workbook, "docProps/core.xml": xlsx_core,
                        "[Content_Types].xml": xlsx_types, "xl/_rels/workbook.xml.rels": xlsx_rels},
        add=[("xl/vbaProject.bin", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 504)])


def ods_meta(xml):
    return re.sub(r"(<dc:date>)[^<]*", r"\g<1>" + OLD + ".000000000", xml)


rewrite(paths["ods"], {"meta.xml": ods_meta})
for ext_ in ("xlsx", "ods"):
    shutil.copy(paths[ext_], os.path.join(HERE, f"lint_book.{ext_}"))
    # Everything that depends on the moment of the run becomes a constant, except the modification date
    # set above, which FILE1002 is about (`fixture_determinism.DELIBERATE_MODIFIED`).
    assert fixture_determinism.DELIBERATE_MODIFIED[f"tests/lint_book.{ext_}"] == OLD
    fixture_determinism.normalize(os.path.join(HERE, f"lint_book.{ext_}"), freeze_volatile=True)

print("wrote", HERE)
