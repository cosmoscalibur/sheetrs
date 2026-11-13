#!/usr/bin/env python3
"""Writes sheetrs-edit/tests/assets/book/: `book.{xlsx,ods}`, the integrated workbook that the
editor tests start from (it replaces `tests/minimal_test.*`), its oracle `book.<ext>.oracle.json`,
and for each case of `CASES` the workbook `LibreOffice` 26.8 leaves after doing the same operation
itself, `expected_<case>.{xlsx,ods}`, plus `cases.json`.

The book is a small report, made so that every kind of content the editor has to carry or rewrite is
in it, each file judged on its own:

  Summary   formulas over the other sheets, a defined name, a text formula, an external reference
  Data      A1:E1 merged title; table A2:E12 (text, numbers, a currency, formulas), totals in row 13,
            a single-cell and a multi-cell array formula, a hidden column (F) and a hidden row (8), a
            conditional format, a validation, a hyperlink, a note, a chart, a local name (`Local`)
  Rates     numbers and a formula that reads a name
  Text      repeated strings, rich text (part in bold), numbers typed as text, dates, formulas on text
  Hidden    a hidden sheet with formulas over `Data` and `Summary`
  Empty     no cells

Names: `Total` (Data!D13), `UnitPrice` (Data!C3:C12), `Local` (only in `Data`).

Run it with a LibreOffice listening on a socket (26.8 until another is the reference):

    scripts/libreoffice-server.sh        # port 2002
    /opt/libreoffice26.8/program/python sheetrs-edit/scripts/gen_book_fixtures.py [CASE ...]

With case names it makes only those again (the book and the oracle are kept).

"""
import json
import os
import shutil
import sys
import time

import uno
from com.sun.star.awt import Rectangle

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale
from com.sun.star.table import CellRangeAddress

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.normpath(os.path.join(HERE, "..", "tests", "assets", "book"))
os.makedirs(OUT, exist_ok=True)

local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8"}


def pv(**kw):
    out = []
    for k, v in kw.items():
        p = PropertyValue()
        p.Name, p.Value = k, v
        out.append(p)
    return tuple(out)


def enum(kind, name):
    return uno.Enum("com.sun.star.sheet." + kind, name)


def new():
    return desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, pv(Hidden=True))


def load(path):
    return desktop.loadComponentFromURL("file://" + path, "_blank", 0, pv(Hidden=True))


def save(doc, stem, ext):
    doc.storeToURL("file://%s/%s.%s" % (OUT, stem, ext), pv(FilterName=FILTERS[ext]))
    fixture_determinism.normalize("%s/%s.%s" % (OUT, stem, ext))


def build():
    doc = new()
    sheets = doc.Sheets
    summary = sheets.getByIndex(0)
    summary.Name = "Summary"
    for index, name in enumerate(("Data", "Rates", "Text", "Hidden", "Empty"), 1):
        sheets.insertNewByName(name, index)
    data, rates, text, hidden = (sheets.getByName(n) for n in ("Data", "Rates", "Text", "Hidden"))
    formats = doc.NumberFormats
    en_us = Locale("en", "US", "")

    def fmt(cell, code):
        key = formats.queryKey(code, en_us, False)
        cell.NumberFormat = key if key != -1 else formats.addNew(code, en_us)

    # ---- Data
    data.getCellRangeByName("A1").String = "Quarterly sales"
    data.getCellRangeByName("A1:E1").merge(True)
    for col, head in enumerate(("Region", "Units", "Price", "Revenue", "Level")):
        data.getCellByPosition(col, 1).String = head
    regions = ["North", "South", "East", "West", "North", "South", "East", "West", "Centre", "Centre"]
    for i, region in enumerate(regions):
        row = 2 + i  # zero based: rows 3..12
        r = row + 1
        data.getCellByPosition(0, row).String = region
        data.getCellByPosition(1, row).Value = 10 * (i + 1)
        data.getCellByPosition(2, row).Value = 12.5 + i
        fmt(data.getCellByPosition(2, row), "[$$-409]#,##0.00")
        data.getCellByPosition(3, row).Formula = "=B%d*C%d" % (r, r)
        fmt(data.getCellByPosition(3, row), "#,##0.00")
        # Neighbors with one relative pattern make a formula group on import; the alternating
        # patterns keep each cell a formula of its own where it matters, the rest is a group.
        data.getCellByPosition(4, row).Formula = '=IF(D%d>1000;"high";"low")' % r
    data.getCellRangeByName("A13").String = "Total"
    data.getCellRangeByName("B13").Formula = "=SUM(B3:B12)"
    data.getCellRangeByName("D13").Formula = "=SUM(D3:D12)"
    fmt(data.getCellRangeByName("D13"), "#,##0.00")
    # Array formulas: one cell and a range.
    data.getCellRangeByName("H1").setArrayFormula("=SUM(B3:B12*C3:C12)")
    data.getCellRangeByName("H3:H5").setArrayFormula("=B3:B5*2")
    data.getCellRangeByName("G3").Formula = "=H3+1"
    # Hidden column F and hidden row 8.
    data.getCellRangeByName("F2").String = "hidden column"
    data.Columns.getByIndex(5).IsVisible = False
    data.Rows.getByIndex(7).IsVisible = False
    # Conditional format on the revenue and a validation on the units.
    cf = data.getCellRangeByName("D3:D12").ConditionalFormat
    cf.addNew(pv(Operator=enum("ConditionOperator", "GREATER"), Formula1="1000", StyleName="Good"))
    data.getCellRangeByName("D3:D12").ConditionalFormat = cf
    rng = data.getCellRangeByName("B3:B12")
    validation = rng.Validation
    validation.Type = enum("ValidationType", "WHOLE")
    validation.setOperator(enum("ConditionOperator", "GREATER_EQUAL"))
    validation.setFormula1("0")
    validation.ShowErrorMessage = True
    rng.Validation = validation
    # A hyperlink, a note and a chart.
    link = data.getCellRangeByName("A15")
    field = doc.createInstance("com.sun.star.text.TextField.URL")
    field.URL = "https://example.com/report"
    field.Representation = "Report site"
    link.insertTextContent(link.createTextCursor(), field, False)
    data.Annotations.insertNew(data.getCellRangeByName("B3").CellAddress, "Check the first units")
    area = data.getCellRangeByName("A2:B7").RangeAddress
    rect = Rectangle()
    rect.X, rect.Y, rect.Width, rect.Height = 500, 9000, 9000, 6000
    data.Charts.addNewByName("Units", rect, (area,), True, True)

    # ---- Rates (a formula that reads a name)
    for i, v in enumerate((0.05, 0.07, 0.1)):
        rates.getCellByPosition(0, i).Value = v
        fmt(rates.getCellByPosition(0, i), "0.0%")
    rates.getCellRangeByName("B1").Formula = "=INDEX(UnitPrice;2)"
    rates.getCellRangeByName("C1").Formula = "=SUM(UnitPrice)"

    # ---- Text
    for i, s in enumerate(("Hello World hello", "repeat", "repeat", "repeat", "123", "Hello")):
        text.getCellByPosition(0, i).String = s
    rich = text.getCellRangeByName("B1")
    rich.String = "Bold and plain"
    cursor = rich.createTextCursor()
    cursor.goRight(4, True)
    cursor.CharWeight = 150.0
    text.getCellRangeByName("B2").Value = 45000
    fmt(text.getCellRangeByName("B2"), "yyyy-mm-dd")
    text.getCellRangeByName("B3").Value = 0.256
    fmt(text.getCellRangeByName("B3"), "0.0%")
    text.getCellRangeByName("C1").Formula = '=A1&" "&A6'
    text.getCellRangeByName("C2").Formula = '=LEN(A1)'
    text.getCellRangeByName("C3").Formula = '=UPPER("abc")'

    # ---- Hidden
    hidden.getCellRangeByName("A1").Formula = "=SUM($Data.B3:B12)"
    hidden.getCellRangeByName("A2").Formula = "=$Summary.A1+1"
    hidden.IsVisible = False

    # ---- Names
    def address(sheet, col, row):
        return sheet.getCellByPosition(col, row).CellAddress

    doc.NamedRanges.addNewByName("Total", "$Data.$D$13", address(data, 0, 0), 0)
    doc.NamedRanges.addNewByName("UnitPrice", "$Data.$C$3:$C$12", address(data, 0, 0), 0)
    data.NamedRanges.addNewByName("Local", "$Data.$B$3", address(data, 0, 0), 0)

    # ---- Summary
    ext = os.path.join(OUT, "ext.xlsx")
    summary.getCellRangeByName("A1").Formula = "=SUM($Data.D3:D12)"
    summary.getCellRangeByName("A2").Formula = "=Total*2"
    summary.getCellRangeByName("A3").Formula = "=AVERAGE(UnitPrice)"
    summary.getCellRangeByName("A4").Formula = '=VLOOKUP("North";$Data.A3:D12;4;0)'
    summary.getCellRangeByName("A5").Formula = "=$Hidden.A1+1"
    summary.getCellRangeByName("A6").Formula = '=COUNTIF($Data.E3:E12;"high")'
    summary.getCellRangeByName("A7").Formula = '="Total: "&TEXT(Total;"#,##0")'
    summary.getCellRangeByName("A8").Formula = "='file://%s'#$Sheet1.B3" % ext
    summary.getCellRangeByName("B1").Formula = "=$Rates.B1+$Text.C2"
    summary.getCellRangeByName("C1").String = "net"
    return doc


def make_external():
    doc = new()
    sheet = doc.Sheets.getByIndex(0)
    sheet.Name = "Sheet1"
    sheet.getCellRangeByName("B3").Value = 42
    doc.storeToURL("file://%s/ext.xlsx" % OUT, pv(FilterName=FILTERS["xlsx"]))
    fixture_determinism.normalize("%s/ext.xlsx" % OUT)
    doc.close(True)


def oracle(doc, ext):
    """What LibreOffice says of the saved file: the sheets and every formula cell."""
    sheets = []
    formulas = []
    for sheet in doc.Sheets.ElementNames:
        sh = doc.Sheets.getByName(sheet)
        sheets.append({"name": sheet, "visible": bool(sh.IsVisible)})
        cursor = sh.createCursor()
        cursor.gotoEndOfUsedArea(False)
        for r in range(cursor.RangeAddress.EndRow + 1):
            for c in range(cursor.RangeAddress.EndColumn + 1):
                cell = sh.getCellByPosition(c, r)
                if cell.Type.value == "FORMULA":
                    formulas.append({
                        "sheet": sheet, "row": r, "col": c, "formula": cell.Formula.replace(OUT, "<folder>"),
                        "array": bool(cell.ArrayFormula),
                        "string": cell.String,
                    })
    return {"sheets": sheets, "formulas": formulas, "names": sorted(doc.NamedRanges.ElementNames)}


# ---- The operations, done by LibreOffice itself.
def delete_rows(doc, sheet, first, count):
    doc.Sheets.getByName(sheet).Rows.removeByIndex(first - 1, count)


def delete_columns(doc, sheet, first, count):
    doc.Sheets.getByName(sheet).Columns.removeByIndex(first, count)


def delete_sheet(doc, name):
    doc.Sheets.removeByName(name)


def delete_name(doc, name):
    doc.NamedRanges.removeByName(name)


def formula_to_value(doc, sheet, area=None):
    """`Data > Calculate > Formula to Value` (`.uno:ConvertFormulaToValue`) on `area` or the sheet."""
    sh = doc.Sheets.getByName(sheet)
    controller = doc.CurrentController
    controller.setActiveSheet(sh)
    controller.select(sh.getCellRangeByName(area) if area else sh)
    dispatcher = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    dispatcher.executeDispatch(controller.Frame, ".uno:ConvertFormulaToValue", "", 0, ())


def replace_all(doc, sheet, text, with_, **opts):
    sh = doc.Sheets.getByName(sheet) if sheet else doc.Sheets
    descriptor = (sh if sheet else doc.Sheets.getByIndex(0)).createReplaceDescriptor()
    descriptor.SearchString = text
    descriptor.ReplaceString = with_
    descriptor.SearchCaseSensitive = opts.get("match_case", False)
    descriptor.SearchWords = opts.get("entire", False)
    descriptor.SearchRegularExpression = opts.get("regex", False)
    descriptor.SearchType = {"formula": 0, "value": 1}.get(opts.get("in", "formula"), 0)
    if sheet:
        sh.replaceAll(descriptor)
    else:
        for name in doc.Sheets.ElementNames:
            doc.Sheets.getByName(name).replaceAll(descriptor)


CASES = [
    dict(name="delete_rows", ops=[("delete_rows", "Data", 10, 2)], edit={"delete_rows": ["Data!10:11"]}),
    dict(name="delete_rows_over_the_chart", ops=[("delete_rows", "Data", 15, 16)], edit={"delete_rows": ["Data!15:30"]}),
    dict(name="delete_rows_hidden_row", ops=[("delete_rows", "Data", 8, 1)], edit={"delete_rows": ["Data!8"]}),
    dict(name="delete_columns", ops=[("delete_columns", "Data", 1, 1)], edit={"delete_columns": ["Data!B"]}),
    dict(name="delete_columns_hidden", ops=[("delete_columns", "Data", 5, 2)], edit={"delete_columns": ["Data!F:G"]}),
    dict(name="delete_sheet_data", ops=[("delete_sheet", "Data")], edit={"delete_sheets": ["Data"]}),
    dict(name="delete_sheet_hidden", ops=[("delete_sheet", "Hidden")], edit={"delete_sheets": ["Hidden"]}),
    dict(name="delete_name", ops=[("delete_name", "Total")], edit={"delete_names": ["Total"]}),
    dict(name="formula_to_value_sheet", ops=[("formula_to_value", "Data")], edit={"formula_to_value": ["Data"]}),
    dict(name="formula_to_value_columns_rows", ops=[("formula_to_value", "Data", "D1:E1048576"), ("formula_to_value", "Data", "A5:H8")],
         edit={"formula_to_value": ["Data!D:E", "Data!5:8"]}),
    dict(name="formula_to_value_every_sheet",
         ops=[("formula_to_value", n) for n in ("Summary", "Data", "Rates", "Text", "Hidden", "Empty")],
         edit={"formula_to_value": ["Summary", "Data", "Rates", "Text", "Hidden", "Empty"]}),
    dict(name="plan_delete_and_formula_to_value",
         ops=[("delete_sheet", "Summary"), ("delete_sheet", "Empty"), ("formula_to_value", "Data", "D1:E1048576")],
         edit={"delete_sheets": ["Summary", "Empty"], "formula_to_value": ["Data!D:E"]}),
    dict(name="replace_value", ops=[("replace_all", None, "repeat", "again", dict(**{"in": "value"}))],
         edit={"replace": {"text": "repeat", "with_": "again"}}),
    dict(name="replace_rich_text", ops=[("replace_all", None, "plain", "simple", dict(**{"in": "value"}))],
         edit={"replace": {"text": "plain", "with_": "simple"}}),
    dict(name="replace_name_in_formula", ops=[("replace_all", None, "Total", "UnitPrice", dict(**{"in": "formula"}))],
         edit={"replace": {"text": "Total", "with_": "UnitPrice", "search_in": "formula"}}),
    dict(name="replace_formula_sheet_ref", ops=[("replace_all", None, "Data", "Rates", dict(**{"in": "formula"}))],
         edit={"replace": {"text": "Data", "with_": "Rates", "search_in": "formula"}}),
    dict(name="replace_hidden_sheet", ops=[("replace_all", "Hidden", "$Data", "$Rates", dict(**{"in": "formula"}))],
         edit={"replace": {"text": "$Data", "with_": "$Rates", "search_in": "formula", "scope": [{"sheet": "Hidden"}]}}),
]

DISPATCH = {
    "delete_rows": delete_rows, "delete_columns": delete_columns, "delete_sheet": delete_sheet,
    "delete_name": delete_name, "formula_to_value": formula_to_value, "replace_all": replace_all,
}


def run(op, doc):
    name, args = op[0], list(op[1:])
    kwargs = args.pop() if args and isinstance(args[-1], dict) else {}
    DISPATCH[name](doc, *args, **kwargs)


def main(only):
    if not only:
        make_external()
        doc = build()
        for ext in FILTERS:
            save(doc, "book", ext)
        doc.close(True)
        for ext in FILTERS:
            doc = load("%s/book.%s" % (OUT, ext))
            with open("%s/book.%s.oracle.json" % (OUT, ext), "w") as fh:
                json.dump(oracle(doc, ext), fh, indent=1)
            doc.close(True)
    for case in CASES:
        if only and case["name"] not in only:
            continue
        for ext in FILTERS:
            doc = load("%s/book.%s" % (OUT, ext))
            for op in case["ops"]:
                run(op, doc)
            save(doc, "expected_" + case["name"], ext)
            doc.close(True)
    if not only:
        with open(OUT + "/cases.json", "w") as fh:
            json.dump([{"name": c["name"], **c["edit"]} for c in CASES], fh, indent=1)


if __name__ == "__main__":
    main(set(sys.argv[1:]))
