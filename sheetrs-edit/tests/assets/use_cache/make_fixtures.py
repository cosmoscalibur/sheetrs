#!/usr/bin/env python3
"""The workbooks of `tests/use_cache.rs`: formulas that cache a result the engine would not give, so that a
formula that keeps its cache is told from one that was calculated again.

    python3 make_fixtures.py     # needs openpyxl and `soffice` (LibreOffice) on the PATH

Writes next to this file, for each workbook, an XLSX and the ODS LibreOffice makes of it (it keeps the results the
XLSX caches: it does not recalculate a file on load by default):

  book      Data (A1:A6 = 10..60, B1 = "x", C1 = 7, the name `rate` = Data!$C$1) and Calc, with the formulas the
            tests reach in different ways
  indirect  `book` and a formula with INDIRECT: it can read any cell
  repro     issue 126: `=C1` and `=IF(B1<>"","x",0)` in Calc, which Excel caches as 0 and "x"
  array     an array formula (`Calc!E1:F3`) and eight formulas that nothing reaches
  volatile  `TODAY()`, `NOW()` and `RAND()` in Calc!A1:A3, which no edit of the tests reaches, with a result
            that is not the one they give
  rules     what a deletion reaches without changing the text of the formula: a whole row of `Data`, `ROW()`,
            `SHEET()`, a defined name, `ISFORMULA` and `FORMULATEXT` of a cell; the sheet `Other` is the one the tests delete
  empty     three formulas that cache an empty result (<v></v>, office:value="") and the forty fillers
  quarter   40 formulas `=SUM(Data!A{k}:A{k+40})` of Calc, over `Data!A1:A80` (all 11): a quarter of them is the
            limit above which the cache is not used
  runs      300 formulas over `Data` that read the cells a replacement of 65 separate cells, and one of a run of
            70 adjacent ones, rewrites (the rectangles `Touched` hands the compiler)

Each is written by `build` with the results a formula "cached by Excel" has (a number that the formula does not
give, so that a formula that kept it is told from one that was calculated again), and every workbook is made
byte-for-byte reproducible by `fixture_determinism.normalize`.
"""
import os
import re
import subprocess
import sys
import tempfile
import zipfile

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "..", "..", "scripts"))
import fixture_determinism  # the same bytes on every run, see its docstring

# (cell, formula, cached value): a text is cached as `t="str"`.
CALC = [
    ("A1", "=SUM(Data!A1:A6)", 210),
    ("A2", "=Data!A2*2", 40),
    ("A3", "=A2+1", 41),
    ("A4", "=Data!A6*3", 180),
    ("A5", "=COUNT(Data!A:A)", 6),
    ("A6", "=rate*2", 140),
    ("A7", '=IF(Data!B1<>"","t",0)', "from-excel"),
    ("A8", "=ROW()", 99),
    ("A9", "=Data!A5+Data!A4", 90),
    ("B1", "=A1+A5", 216),
    ("B2", "=SUBTOTAL(109,Data!A1:A6)", 777),
    # Forty formulas that no test reaches: the walk is used where a share of the formulas is stale only.
    *[(f"C{k}", f"=Data!C1+{k}", 7 + k) for k in range(1, 41)],
]
INDIRECT = [("A10", '=INDIRECT("Data!A"&COUNTA(Data!B1))', 999)]
REPRO = [("B1", "=C1", 0), ("B2", '=IF(B1<>"","x",0)', "x")]
FILLER = [(f"C{k}", f"=Data!C1+{k}", 1000 + k) for k in range(1, 41)]
# Volatile functions that no edit of the tests reaches. Nothing depends on them: LibreOffice calculates them again
# when it opens the file, and a cell that read one would cache a different number on every run.
VOLATILE = [("A1", "=TODAY()", 1), ("A2", "=NOW()", 1), ("A3", "=RAND()", 0.5), *FILLER]
# The formulas a deletion reaches although their text does not change; the results are not the ones they give.
RULES = [
    ("A1", "=SUM(Data!1:1)", 555),
    ("A2", "=ROW()", 99),
    ("A3", "=rate*2", 140),
    ("A4", "=SHEET()", 99),
    ("A5", "=Data!A2*2", 777),
    ("A6", "=COUNT(Data!A:A)", 888),
    ("A7", "=_xlfn.ISFORMULA(Data!A2)", 777),
    ("A8", "=_xlfn.FORMULATEXT(Data!A2)", 778),
    *FILLER,
]
# Formulas whose cached result is empty (`<v></v>` of an XLSX, `office:value=""` of an ODS: what a program that
# writes no results leaves), written as a number first and blanked by `blank`.
EMPTY = [(f"A{k}", f"=Data!A{k + 1}*2", 123456) for k in range(1, 4)] + FILLER
QUARTER = [(f"A{k}", f"=SUM(Data!A{k}:A{k + 40})", 5000 + k) for k in range(1, 41)]
# Replacements: column A is rewritten at the odd rows 1..129 (65 separate cells), row 1 from B to BS (70 adjacent
# ones) and B41. A formula for each of those cells, formulas that read a cell between them, and some that nothing
# reaches (a quarter of the formulas is the limit).
RUNS = (
    [(f"A{n}", f"=Data!A{r}*2", 2000 + r) for n, r in enumerate(range(1, 130, 2), start=1)]
    + [("A66", "=Data!A2*2", 3002), ("A67", "=Data!A100*2", 3100), ("A68", "=Data!A131*2", 3131)]
    + [(f"A{69 + i}", f"=Data!{openpyxl.utils.get_column_letter(c)}1*2", 4000 + i) for i, c in enumerate(range(2, 72))]
    + [("A139", "=Data!B41*2", 4041), ("A140", "=Data!C20*2", 4020), ("A141", "=Data!BR30*2", 4030)]
    + [(f"A{142 + k}", f"=Data!A200+{k}", 7000 + k) for k in range(160)]
)


def book_data(data):
    for i, v in enumerate([10, 20, 30, 40, 50, 60], start=1):
        data.cell(row=i, column=1, value=v)
    data["B1"] = "x"
    data["C1"] = 7


def rules_data(data):
    book_data(data)
    data["D3"] = 5  # a column of its own: the tests delete it, and no formula reads it


def eleven(rows):
    def populate(data):
        for i in range(1, rows + 1):
            data.cell(row=i, column=1, value=11)

    return populate


def runs_data(data):
    for r in range(1, 131):
        data.cell(row=r, column=1, value=11)
    for c in range(2, 72):
        data.cell(row=1, column=c, value=11)
    for row, column in ((20, 3), (30, 70), (41, 2)):  # C20 and BR30 are between the replaced cells; B41 is one
        data.cell(row=row, column=column, value=11)


def build(path, formulas, populate=book_data, others=()):
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Data"
    populate(data)
    wb.defined_names["rate"] = DefinedName("rate", attr_text="Data!$C$1")
    sheet = wb.create_sheet("Calc")
    for cell, formula, _ in formulas:
        sheet[cell] = formula
    for name in others:  # after Calc: it stays `xl/worksheets/sheet2.xml`
        wb.create_sheet(name)["A1"] = 1
    wb.save(path)
    cached = {cell: value for cell, _, value in formulas}
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet2.xml":
                xml = data.decode()
                for cell, value in cached.items():
                    pattern = re.compile(r'<c r="%s"([^>]*)><f>(.*?)</f>(?:<v\s*/>|<v></v>)</c>' % cell)
                    kind = ' t="str"' if isinstance(value, str) else ""
                    xml, n = pattern.subn(
                        lambda m: '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (cell, kind, m.group(2), value), xml)
                    assert n == 1, cell
                data = xml.encode()
            if item.filename == "xl/workbook.xml":
                # openpyxl asks the application to recalculate on load: these results are meant to be read.
                data = data.decode().replace(' fullCalcOnLoad="1"', "").encode()
            zout.writestr(item, data)
    os.replace(tmp, path)


def to_ods(path):
    with tempfile.TemporaryDirectory() as profile:
        subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "ods",
                        "--outdir", HERE, path], check=True, capture_output=True)


def blank(name):
    """Empties the results that `build` cached as 123456, in the XLSX and in the ODS."""
    for extension, entry, edit in (
        ("xlsx", "xl/worksheets/sheet2.xml", lambda x: x.replace("<v>123456</v>", "<v></v>")),
        ("ods", "content.xml", lambda x: re.sub(r'office:value="123456"([^>]*)><text:p>123456</text:p>',
                                                r'office:value=""\1><text:p/>', x)),
    ):
        path = os.path.join(HERE, f"{name}.{extension}")
        tmp = path + ".tmp"
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == entry:
                    text = data.decode()
                    changed = edit(text)
                    assert changed != text, (name, extension)
                    data = changed.encode()
                zout.writestr(item, data)
        os.replace(tmp, path)


def finish(name, freeze_volatile=False):
    for extension in ("xlsx", "ods"):
        fixture_determinism.normalize(os.path.join(HERE, f"{name}.{extension}"), freeze_volatile=freeze_volatile)


def array_book():
    """`Calc!E1:F3` is the array formula `=Data!A1:B3` and `H1:H8` eight formulas that nothing reaches, in both
    formats as LibreOffice writes them: a result for each member cell of the block."""
    from openpyxl.worksheet.formula import ArrayFormula

    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Data"
    for i in range(1, 4):
        data.cell(row=i, column=1, value=10 * i)
        data.cell(row=i, column=2, value=i)
    calc = wb.create_sheet("Calc")
    calc["E1"] = ArrayFormula("E1:F3", "=Data!A1:B3")
    for k in range(1, 9):
        calc[f"H{k}"] = f"=Data!A1+{k}"
    with tempfile.TemporaryDirectory() as folder, tempfile.TemporaryDirectory() as profile:
        source = os.path.join(folder, "array.xlsx")
        wb.save(source)

        def convert(target, path, outdir):
            subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to",
                            target, "--outdir", outdir, path], check=True, capture_output=True)

        convert("ods", source, HERE)
        # The XLSX is the one LibreOffice writes from the ODS, with the members of the block.
        os.makedirs(os.path.join(folder, "out"))
        convert("xlsx", os.path.join(HERE, "array.ods"), os.path.join(folder, "out"))
        os.replace(os.path.join(folder, "out", "array.xlsx"), os.path.join(HERE, "array.xlsx"))
    finish("array")


def main():
    books = (
        ("book", CALC, book_data, (), False),
        ("indirect", CALC + INDIRECT, book_data, (), False),
        ("repro", REPRO, book_data, (), False),
        ("volatile", VOLATILE, book_data, (), True),
        ("rules", RULES, rules_data, ("Other",), False),
        ("empty", EMPTY, book_data, (), False),
        ("quarter", QUARTER, eleven(80), (), False),
        ("runs", RUNS, runs_data, (), False),
    )
    for name, formulas, populate, others, freeze in books:
        path = os.path.join(HERE, f"{name}.xlsx")
        build(path, formulas, populate, others)
        to_ods(path)
        if name == "empty":
            blank(name)
        finish(name, freeze)
    array_book()


if __name__ == "__main__":
    main()
