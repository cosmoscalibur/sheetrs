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
"""
import os
import re
import subprocess
import tempfile
import zipfile

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))

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


def build(path, formulas):
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Data"
    for i, v in enumerate([10, 20, 30, 40, 50, 60], start=1):
        data.cell(row=i, column=1, value=v)
    data["B1"] = "x"
    data["C1"] = 7
    wb.defined_names["rate"] = DefinedName("rate", attr_text="Data!$C$1")
    sheet = wb.create_sheet("Calc")
    for cell, formula, _ in formulas:
        sheet[cell] = formula
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


def main():
    for name, formulas in (("book", CALC), ("indirect", CALC + INDIRECT), ("repro", REPRO)):
        path = os.path.join(HERE, f"{name}.xlsx")
        build(path, formulas)
        to_ods(path)
    array_book()


if __name__ == "__main__":
    main()
