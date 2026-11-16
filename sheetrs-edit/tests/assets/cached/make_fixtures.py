#!/usr/bin/env python3
"""The workbooks of `tests/calculation_mode.rs`: formulas that cache a result the engine would not give, so that a
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
  refs      every function of a family over a reference that an edit removes (whole column, whole row, range, cell)
  dirty     what `LibreOffice` recalculates (and keeps) when a deletion of rows, columns, a sheet or a name reaches a
            formula (`tests/calculation_mode.rs` and `write_to_libreoffice.rs`): formulas of every kind over `Data`
            (a cell, a range, whole columns and rows, a defined name that points to a cell, a range, a whole row, a name
            that uses another one), `ROW`/`COLUMN`, a formula that holds an error `#REF!`, formulas of `Data` and
            `Other`, and 120 formulas nothing reaches (no `CELL` or `FORMULATEXT`: `LibreOffice` calculates those when it loads any file) (column I of `Calc`). Every result cached is the text `S`, not a number,
            so that what LibreOffice keeps is told from what it calculates in an ODS too (it calculates the numbers of
            an ODS on load)
  classes   one formula of `Calc` (column A, column B reads it) for each class of function that `LibreOffice` may calculate when it
            opens a file or when something is edited, every result cached as the text `S`: the position ones (`ROW`, `COLUMN`),
            the whole-column and whole-row reads, `HYPERLINK`, the types of `CELL` and `INFO` that the engine answers and that give the same twice (`CELL("width"|"format"|"color"|"prefix"|"protect"|"parentheses")` depend on the column width, the format, the alignment and the protection of the cell, which the engine's model does not hold), `FORMULATEXT`
            and `ISFORMULA`, `SHEET`, `SUBTOTAL`, and formulas of no class as the control (the volatile functions, `INDIRECT`,
            `OFFSET`, `CELL("filename")` and the `INFO` that depend on the machine are in `probe_dirty.py classes` of the oracle's
            folder: they give another result each time they are calculated, so no file can be compared with them)
  visibility  `Data` (values in rows 1-5, 7-9 and 15) with rows 3-4, 6 and 10-14 already hidden (6 and 10-14 are empty; an ODS writes 10-13 as one
            repeated row), and in `Calc` a SUBTOTAL over each row of `Data`, over rows 9-12, 13-14, 16-17 and 10-15, and a group of three SUMs, every result cached as the
            text `S`: what hiding or showing rows calculates when the rows are hidden already, or visible (`probe_hide_groups.py`)
  limit     `book` and a formula the engine does not calculate (`TEXT` with the code of another language, which it
            has no names for): where a run that has to calculate marks the workbook to be recalculated
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
LIMIT = CALC + [("B3", '=TEXT(45367,"[$-40C]dddd")', "lundi")]
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
    ("A4", "=_xlfn.SHEET()", 99),
    ("A5", "=Data!A2*2", 777),
    ("A6", "=COUNT(Data!A:A)", 888),
    ("A7", "=_xlfn.ISFORMULA(Data!A2)", 777),
    ("A8", "=_xlfn.FORMULATEXT(Data!A2)", 778),
    *FILLER,
]
# Formulas whose cached result is empty (`<v></v>` of an XLSX, `office:value=""` of an ODS: what a program that
# writes no results leaves), written as a number first and blanked by `blank`.
EMPTY = [(f"A{k}", f"=Data!A{k + 1}*2", 123456) for k in range(1, 4)] + FILLER
# What each function does with a reference that goes: the functions of the family, over whole columns, whole
# rows, a range, a range of two columns and one cell of `Data`, and some others (results are not the ones they
# give).
REFS = []
for _fn in ("COUNT", "SUM", "COUNTA", "AVERAGE", "MAX", "MIN", "COUNTBLANK"):
    for _ref in ("Data!A:A", "Data!1:1", "Data!A1:A6", "Data!A2:B3", "Data!A3"):
        REFS.append((f"A{len(REFS) + 1}", f"={_fn}({_ref})", 777))
for _text in ('COUNTIF(Data!A:A,">1")', 'COUNTIF(Data!A1:A6,">1")', 'SUMIF(Data!A1:A6,">1")', "Data!A3+1", "Data!A3",
              "ROWS(Data!A1:A6)", "ISREF(Data!A3)", "ISERROR(Data!A3)", "IFERROR(Data!A3,7)", "SUM(Data!A1:A3)",
              "INDEX(Data!A1:A6,2)", "SUMPRODUCT(Data!A1:A6)", "COUNT(Data!A3,5)", "SUM(Data!A3,5)"):
    REFS.append((f"A{len(REFS) + 1}", "=" + _text, 777))
# An error written in the formulas (a `#REF!` typed or read from a file, which LibreOffice reads as the error it is),
# next to the references an edit removes (`Data!A3` on rows, `Other!A1` on a sheet): what a deletion of the run does to
# each. Rows 21 to 26 mix the two kinds.
PREREF = [(f"A{n}", "=" + f, 777) for n, f in enumerate([
    "COUNT(#REF!)", "COUNTA(#REF!)", "SUM(#REF!)", "MAX(#REF!)", "IFERROR(#REF!,0)", "ISERROR(#REF!)",
    "COUNT(#REF!,5)", "SUM(#REF!,5)", "COUNTBLANK(#REF!)",
    "COUNT(Data!#REF!)", "COUNTA(Data!#REF!)", "SUM(Data!#REF!)", "IFERROR(Data!#REF!,0)", "ISERROR(Data!#REF!)",
    "#REF!+1", "SUM(Data!A1:A6)", "COUNT(Data!A1:A6)", "IF(ISERROR(#REF!),1,2)", "COUNT(Data!A3:A4)",
    "IFERROR(#REF!,0)+COUNT(Data!A3)",
    "IFERROR(Data!A3,7)", "IFERROR(Other!A1,7)", "COUNT(Data!A3)", "COUNT(Other!A1)",
    "SUM(Data!A3,Other!A1)", "IFERROR(Data!A3+Other!A1,7)",
], start=1)]
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


# The `dirty` workbook: one formula of `Calc` for each of these (column A; column B reads it, to tell what depends on a
# stale result), every result cached as the text `S`.
DIRTY_NAMES = {
    "n_cell3": "Data!$A$3", "n_cell5": "Data!$A$5", "n_cell1": "Data!$A$1", "n_rng16": "Data!$A$1:$A$6",
    "n_rng12": "Data!$A$1:$A$2", "n_rng35": "Data!$A$3:$A$5", "n_row3": "Data!$3:$3", "n_colA": "Data!$A:$A",
    "n_c1": "Data!$C$1", "n_b1": "Data!$B$1", "n_other": "Other!$A$1", "n_const": "5", "n_nest": "n_cell5+1",
    "n_rng24": "Data!$A$2:$A$4", "n_d3": "Data!$D$3", "n_nestr": "SUM(n_rng16)", "n_nestm": "n_cell1+n_cell5",
}
DIRTY = (
    [f"={n}*2" if n != "n_b1" else f"=LEN({n})" for n in DIRTY_NAMES]
    + [f"=SUM({n})" for n in ("n_rng16", "n_rng12", "n_rng35", "n_colA", "n_row3", "n_rng24")]
    + ["=Data!A1", "=Data!A3", "=Data!A5", "=Data!A6", "=Data!C1", "=LEN(Data!B1)", "=Data!D3",
       "=SUM(Data!A1:A6)", "=SUM(Data!A2:A3)", "=SUM(Data!A3:A4)", "=SUM(Data!A3:A3)", "=SUM(Data!A1:A2)",
       "=SUM(Data!A4:A6)", "=SUM(Data!A1:C1)", "=SUM(Data!A:A)", "=SUM(Data!1:1)", "=SUM(Data!3:3)",
       "=COUNT(Data!A:A)", "=Other!A1", "=SUM(Other!A1:A2)", "=Data!A1+Other!A1", "=Data!A3+Other!A1",
       "=IFERROR(Other!A1,7)", "=COUNT(Other!A1)", "=E1", "=ROW()", "=ROW(Data!A3)", "=COLUMN(Data!B1)",
       "=ROWS(Data!A1:A6)", "=_xlfn.SHEET(Other!A1)", "=COUNT(#REF!)", "=COUNT(Data!#REF!)", "=SUM(#REF!)+1",
       "=SUM(E1:E5)", "=ROW(E3)", "=Data!E5", "=Data!F1", "=Other!B1", "=ROWS(Data!A4:A6)",
       "=COLUMNS(Data!B1:C1)", "=ROW(Data!A5)", "=COLUMN(Data!C1)", "=INDEX(Data!A1:A6,2)", "=SUMPRODUCT(Data!A1:A6)"]
)


def dirty_book(path):
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Data"
    book_data(data)
    data["D3"] = 5
    for name, text in DIRTY_NAMES.items():
        wb.defined_names[name] = DefinedName(name, attr_text=text)
    calc = wb.create_sheet("Calc")
    for k, formula in enumerate(DIRTY, start=1):
        calc[f"A{k}"] = formula
        calc[f"B{k}"] = f'=A{k}&"!"'
    # Constants and formulas to the right of the formulas, which a deletion of the column between moves.
    for k, v in enumerate((5, 6, 7, 8), start=1):
        calc[f"E{k}"] = v
    calc["F1"], calc["F2"], calc["F3"] = "=COLUMN()", "=ROW()", "=A3"
    # Formulas nothing reaches, outside the area the oracle compares (A1:F80): the walk is used where a share of the
    # formulas is stale only.
    for k in range(120):
        calc[f"I{k + 1}"] = f"=$E$4+{k}"
    data["E1"], data["E3"], data["E5"], data["E6"], data["F1"] = "=A1*2", "=A3*2", "=A5*2", "=A6*2", "=C1+1"
    other = wb.create_sheet("Other")
    other["A1"], other["A2"], other["B1"] = 1, 2, "=A1+1"
    wb.save(path)
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            xml = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                text, n = re.subn(
                    r'<c r="([A-Z]+\d+)"([^>]*)><f>(.*?)</f>(?:<v\s*/>|<v></v>)</c>',
                    lambda m: '<c r="%s"%s t="str"><f>%s</f><v>S</v></c>' % (m.group(1), m.group(2), m.group(3)),
                    xml.decode())
                assert n, item.filename
                xml = text.encode()
            if item.filename == "xl/workbook.xml":
                xml = xml.decode().replace(' fullCalcOnLoad="1"', "").encode()
            zout.writestr(item, xml)
    os.replace(tmp, path)


# The `classes` workbook: a formula of `Calc` for each of these, every result cached as the text `S`.
CLASSES = [
    "ROW()", "COLUMN()", "ROW(Data!A3)", "COLUMN(Data!B1)", "ROWS(Data!A1:A6)", "COLUMNS(Data!A1:C1)",
    "COUNT(Data!A:A)", "SUM(Data!A:A)", "COUNTA(Data!1:1)", "MATCH(30,Data!A:A,0)", "VLOOKUP(10,Data!A:C,3,0)",
    'SUMIF(Data!A:A,">1")', "SUMPRODUCT(Data!A:A)",
    'HYPERLINK("http://x.y","t")', "HYPERLINK(Data!B1)", 'HYPERLINK("http://x.y")',
    'CELL("row",Data!A5)', 'CELL("col",Data!A5)', 'CELL("address",Data!A5)', 'CELL("contents",Data!A5)',
    'CELL("type",Data!A5)', 'CELL("row")',
    'INFO("system")', 'INFO("recalc")', 'INFO("numfile")',
    "_xlfn.FORMULATEXT(Data!E5)", "_xlfn.FORMULATEXT(Data!A1)", "_xlfn.ISFORMULA(Data!E5)", "_xlfn.ISFORMULA(Data!A1)",
    "_xlfn.SHEET()", "_xlfn.SHEET(Data!A1)", "_xlfn.SHEETS()", "_xlfn.SHEETS(Data!A1)", "AREAS(Data!A1:A2)",
    "SUBTOTAL(109,Data!A1:A6)",
    "ISBLANK(Data!A1)", "N(Data!A1)", "T(Data!B1)", "TYPE(Data!A1)", "ISREF(Data!A1)", "ISNUMBER(Data!A1)",
    "DATE(2020,1,1)", "1+1", "SUM(Data!A1:A6)", 'TEXT(Data!A1,"0.00")', "Data!A2",
]


# What gives another result each time it is calculated or depends on the machine: `classes_volatile`, for `probe_dirty.py` only.
CLASSES_VOLATILE = [
    'CELL("filename")', 'CELL("filename",Data!A5)', 'INFO("osversion")', 'INFO("release")', 'INFO("directory")', 'INFO("origin")',
    "NOW()", "TODAY()", "RAND()", "RANDBETWEEN(1,5)",
    'INDIRECT("Data!A1")', "OFFSET(Data!A1,1,0)", "OFFSET(Data!A1,1,0,2,1)", "SUM(Data!A1:A6)", "1+1",
]


def classes_book(path, formulas=None):
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Data"
    book_data(data)
    data["E5"] = "=A5*2"
    calc = wb.create_sheet("Calc")
    for k, formula in enumerate(formulas or CLASSES, start=1):
        calc[f"A{k}"] = "=" + formula
        calc[f"B{k}"] = f'=A{k}&"!"'
    wb.save(path)
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            xml = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                text, n = re.subn(
                    r'<c r="([A-Z]+\d+)"([^>]*)><f>(.*?)</f>(?:<v\s*/>|<v></v>)</c>',
                    lambda m: '<c r="%s"%s t="str"><f>%s</f><v>S</v></c>' % (m.group(1), m.group(2), m.group(3)),
                    xml.decode())
                assert n, item.filename
                xml = text.encode()
            if item.filename == "xl/workbook.xml":
                xml = xml.decode().replace(' fullCalcOnLoad="1"', "").encode()
            zout.writestr(item, xml)
    os.replace(tmp, path)


def visibility_book(path):
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Data"
    for r in (*range(1, 6), 7, 8, 9, 15):  # row 6 is empty (and hidden), and so are 10-14; 15 is the last with a value
        data.cell(row=r, column=1, value=r * 10)
    for r in (3, 4, 6, 10, 11, 12, 13, 14):
        data.row_dimensions[r].hidden = True
    calc = wb.create_sheet("Calc")
    for k in range(1, 9):
        calc[f"A{k}"] = f"=SUBTOTAL(109,Data!A{k})"
    calc["A9"], calc["A10"], calc["A11"] = "=SUBTOTAL(109,Data!A9:A12)", "=SUBTOTAL(109,Data!A13:A14)", "=SUBTOTAL(109,Data!A10)"
    calc["A12"], calc["A13"] = "=SUBTOTAL(109,Data!A16:A17)", "=SUBTOTAL(109,Data!A10:A15)"
    for k in range(1, 4):
        calc[f"B{k}"] = f"=SUM(Data!A{k}:A{k + 2})"
    wb.save(path)
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            xml = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet2.xml":
                text, n = re.subn(
                    r'<c r="([A-Z]+\d+)"([^>]*)><f>(.*?)</f>(?:<v\s*/>|<v></v>)</c>',
                    lambda m: '<c r="%s"%s t="str"><f>%s</f><v>S</v></c>' % (m.group(1), m.group(2), m.group(3)),
                    xml.decode())
                assert n == 16, n
                xml = text.encode()
            if item.filename == "xl/workbook.xml":
                xml = xml.decode().replace(' fullCalcOnLoad="1"', "").encode()
            zout.writestr(item, xml)
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
        ("limit", LIMIT, book_data, (), False),
        ("repro", REPRO, book_data, (), False),
        ("volatile", VOLATILE, book_data, (), True),
        ("rules", RULES, rules_data, ("Other",), False),
        ("refs", REFS, book_data, (), False),
        ("preref", PREREF, book_data, ("Other",), False),
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
    path = os.path.join(HERE, "dirty.xlsx")
    dirty_book(path)
    to_ods(path)
    finish("dirty")
    path = os.path.join(HERE, "visibility.xlsx")
    visibility_book(path)
    to_ods(path)
    finish("visibility")
    path = os.path.join(HERE, "classes.xlsx")
    classes_book(path)
    to_ods(path)
    finish("classes")
    path = os.path.join(HERE, "classes_volatile.xlsx")
    classes_book(path, CLASSES_VOLATILE)
    to_ods(path)
    finish("classes_volatile")
    array_book()


if __name__ == "__main__":
    main()
