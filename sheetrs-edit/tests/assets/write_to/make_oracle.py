"""What LibreOffice 26.8 shows of a workbook after an edit, recorded for the tests that compare a write with it.

THE RULE (policy of the project, `CONTRIBUTING.md`, *Parity and fidelity*): the expected values of an edit are what
LibreOffice 26.8 shows when the workbook it left after the edit is SAVED AND OPENED AGAIN (re-read). The state of its
memory right after the edit is not a reference: the same workbook, saved and opened by someone else, shows other values
(it is not reproducible), which is a defect of LibreOffice and not a behavior to copy. The rule is for every operation
and both formats.

For each workbook of the list of fixtures, in each format, and for each edit of its list, LibreOffice 26.8 makes the
edit through UNO on the loaded file, saves the file (`storeToURL`) and loads it again. The record holds:

  reread_default        the cells (formula and value), the hidden rows and columns and the visibility of each sheet of
                        that file as LibreOffice shows them when it opens it (default load: the results the file
                        caches, plus what LibreOffice itself calculates on load), for an edit with no `calculateAll()`;
  reread                the same after the edit and `calculateAll()` (every formula calculated: the recalculation the
                        screen does as it repaints);
  cache_differences     where the same file, after a hard recalculation (`calculateAll`), shows another value than
                        `reread`;
  in_memory_differences INFORMATIONAL ONLY, never a reference: where the memory of LibreOffice right after the edit
                        (and `calculateAll`) differs from `reread`. They are the evidence that the memory is not stable.

An edit is a list of operations, `[name, ...arguments]`, applied through UNO. Run it with a LibreOffice listening on a
socket (`scripts/libreoffice-server.sh`):

    ASSETS=DIR /opt/libreoffice26.8/program/python make_oracle.py [FIXTURE ...]

`ASSETS` is the folder that holds `FIXTURE.xlsx` and `FIXTURE.ods`; with fixture names it makes only those again.
"""
import json
import os
import re
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.environ.get("ASSETS", ".")
FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8"}
ROWS, COLUMNS = 80, 6  # the area of each sheet that is compared: A1:F80

# (label, the operations). An operation is `[name, ...arguments]`:
#   delete_rows "S!3:5"      delete_columns "S!B:D"      delete_sheets "S"      delete_names "name"
#   formula_to_value "S!A:A"
#   replace {text, with, scope, regex, in}           set_value "S!A2" 25 (an input of a calculation: the value is
#   written and every formula calculated)
# `FIXTURES` pairs the name of each workbook with its list of edits.
FIXTURES = []


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(30):
        try:
            ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        except Exception:
            time.sleep(1)
    raise SystemExit("no LibreOffice listening on port 2002: scripts/libreoffice-server.sh")


desktop = connect()


def load(path):
    return desktop.loadComponentFromURL("file://" + os.path.abspath(path), "_blank", 0, (pv("Hidden", True),))


def a1(row, col):
    return f"{chr(ord('A') + col)}{row + 1}"


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def dump(doc):
    """The cells, hidden rows and columns and visibility of each sheet of `doc`, as the test compares them."""
    out = []
    for i in range(doc.Sheets.Count):
        sheet = doc.Sheets.getByIndex(i)
        cells = {}
        # One call for the area finds the cells that hold something (an empty string is no cell), so that the
        # other calls are for them only.
        filled = sheet.getCellRangeByPosition(0, 0, COLUMNS - 1, ROWS - 1).getFormulaArray()
        for r in range(ROWS):
            for c in range(COLUMNS):
                if filled[r][c] == "":
                    continue
                cell = sheet.getCellByPosition(c, r)
                if cell.Type.value == "EMPTY":
                    continue
                formula = cell.getFormula() if cell.Type.value == "FORMULA" else ""
                if cell.getError() != 0 or (cell.Type.value == "FORMULA" and cell.FormulaResultType2 == 4):
                    value = {"t": "e", "v": cell.getString()}
                elif cell.Type.value == "VALUE" or (cell.Type.value == "FORMULA" and cell.FormulaResultType2 == 1):
                    value = {"t": "n", "v": cell.getValue()}
                else:
                    value = {"t": "s", "v": cell.getString()}
                cells[a1(r, c)] = {"f": formula, "v": value}
        out.append({
            "name": sheet.Name,
            "visible": bool(sheet.IsVisible),
            "hidden_rows": [r for r in range(ROWS) if not sheet.Rows.getByIndex(r).IsVisible],
            "hidden_columns": [c for c in range(COLUMNS) if not sheet.Columns.getByIndex(c).IsVisible],
            "cells": cells,
        })
    return out


def rows_of(spec):
    sheet, rest = spec.rsplit("!", 1)
    first, _, last = rest.partition(":")
    return sheet, int(first) - 1, int(last or first) - 1


def columns_of(spec):
    sheet, rest = spec.rsplit("!", 1)
    first, _, last = rest.partition(":")
    return sheet, col_index(first), col_index(last or first)


def apply_lo(doc, operations, calculate_all=True):
    """The operations on the loaded document, in a fixed order: the replacements,
    the conversions, then the deletions of rows and of columns (from the last to the first), the names and
    the sheets. With `calculate_all` a conversion is of what the formulas give after `calculateAll()`;
    without it, of what the cells hold (the results the file caches plus what the edit made LibreOffice
    calculate)."""
    by_name = {}
    for name, *args in operations:
        by_name.setdefault(name, []).append(args[0])
    sheets = doc.Sheets
    for name, *args in operations:
        if name == "set_value":
            cell = sheets.getByName(args[0].split("!")[0]).getCellRangeByName(args[0].split("!")[1])
            if isinstance(args[1], str):
                cell.setString(args[1])
            else:
                cell.setValue(args[1])
    for spec in by_name.get("replace", []):
        sheet = sheets.getByName(spec["scope"].split("!")[0])
        target = sheet.getCellRangeByName(spec["scope"].split("!")[1])
        d = target.createReplaceDescriptor()
        d.SearchString, d.ReplaceString = spec["text"], spec["with"]
        d.SearchType = 0 if spec.get("in") == "formula" else 1
        d.SearchRegularExpression = bool(spec.get("regex"))
        d.SearchCaseSensitive = False
        d.SearchWords = False
        target.replaceAll(d)
    if by_name.get("formula_to_value") and calculate_all:
        # With `calculate_all` a conversion is of what the formulas give, not of what the file caches:
        # LibreOffice calculates them first. Its own load calculates only some of them, and an ODS and an
        # XLSX differ in which.
        doc.calculateAll()
    for spec in by_name.get("formula_to_value", []):
        # What a conversion reads is what the cells hold now: the results the file caches.
        sheet = sheets.getByName(spec.split("!")[0])
        _, first, last = columns_of(spec)
        target = sheet.getCellRangeByPosition(first, 0, last, ROWS - 1)
        target.setDataArray(target.getDataArray())
    for spec in sorted(by_name.get("delete_rows", []), key=lambda s: rows_of(s)[1], reverse=True):
        sheet, first, last = rows_of(spec)
        sheets.getByName(sheet).Rows.removeByIndex(first, last - first + 1)
    for spec in sorted(by_name.get("delete_columns", []), key=lambda s: columns_of(s)[1], reverse=True):
        sheet, first, last = columns_of(spec)
        sheets.getByName(sheet).Columns.removeByIndex(first, last - first + 1)
    for name in by_name.get("delete_names", []):
        doc.NamedRanges.removeByName(name)
    for name in by_name.get("delete_sheets", []):
        sheets.removeByName(name)



def differences(expected, opened):
    """Where `opened` is not `expected`: a list of `where: expected -> opened`, for formulas, values, hidden
    rows and columns, and sheets (the names of the arguments are those of the first use: `reread` against `written`)."""
    diffs = []
    names = {s["name"] for s in expected} | {s["name"] for s in opened}
    e = {s["name"]: s for s in expected}
    o = {s["name"]: s for s in opened}
    for name in sorted(names):
        if name not in e or name not in o:
            diffs.append(f"sheet {name}: {'only in expected' if name in e else 'only in opened'}")
            continue
        for key in ("visible", "hidden_rows", "hidden_columns"):
            if e[name][key] != o[name][key]:
                diffs.append(f"{name} {key}: {e[name][key]} -> {o[name][key]}")
        for cell in sorted(set(e[name]["cells"]) | set(o[name]["cells"]), key=lambda c: (int(c[1:]), c[0])):
            x, y = e[name]["cells"].get(cell), o[name]["cells"].get(cell)
            if x is None or y is None:
                diffs.append(f"{name}!{cell}: {x} -> {y}")
                continue
            if same_formula(x["f"]) != same_formula(y["f"]):
                diffs.append(f"{name}!{cell} formula: {x['f']} -> {y['f']}")
            if not same_value(x["v"], y["v"]):
                diffs.append(f"{name}!{cell} value: {x['v']['v']!r} -> {y['v']['v']!r}")
    return diffs


def same_formula(formula):
    """The formula without the sheet that the second end of a range repeats, and a range of one cell as that cell."""
    formula = re.sub(r":\$?\w+\.", ":", formula)
    return re.sub(r"(\$?[A-Z]+\$?\d+):\1\b", r"\1", formula)


def same_value(a, b):
    if a["t"] == "n" and b["t"] == "n":
        return abs(a["v"] - b["v"]) <= 1e-9 * max(1.0, abs(a["v"]))
    return a == b



def main(only):
    """`only`: the fixtures to make again (their entries replace the ones of the `oracle.json` that is there)."""
    path = os.path.join(HERE, "oracle.json")
    oracle = {"area": f"A1:{a1(ROWS - 1, COLUMNS - 1)}", "fixtures": {}}
    if only:
        with open(path) as f:
            oracle = json.load(f)
    for fixture, edits in FIXTURES:
        if only and fixture not in only:
            continue
        oracle["fixtures"][fixture] = {}
        for extension in ("xlsx", "ods"):
            source = os.path.join(ASSETS, f"{fixture}.{extension}")
            results = []
            for label, operations in edits:
                # No `calculateAll()`: the edit on the loaded file, saved and opened again; LibreOffice keeps what the
                # file caches and calculates what the edit reaches (and what it calculates on load).
                doc = load(source)
                apply_lo(doc, operations, calculate_all=False)
                with tempfile.TemporaryDirectory() as folder:
                    own = os.path.join(folder, f"default.{extension}")
                    doc.storeToURL("file://" + own, (pv("FilterName", FILTERS[extension]),))
                    doc.close(True)
                    doc = load(own)
                    reread_default = dump(doc)
                    doc.close(True)
                doc = load(source)
                apply_lo(doc, operations)
                doc.calculateAll()
                in_memory = dump(doc)
                with tempfile.TemporaryDirectory() as folder:
                    own = os.path.join(folder, f"own.{extension}")
                    doc.storeToURL("file://" + own, (pv("FilterName", FILTERS[extension]),))
                    doc.close(True)
                    doc = load(own)
                    reread = dump(doc)  # THE EXPECTED VALUES: the file opened as a user opens it
                    doc.calculateAll()
                    reread_hard = dump(doc)
                    doc.close(True)
                results.append({
                    "label": label,
                    "operations": operations,
                    "reread": reread,
                    "reread_default": reread_default,
                    "cache_differences": differences(reread, reread_hard),
                    "in_memory_differences": differences(reread, in_memory),
                })
                r = results[-1]
                print(f"{fixture}.{extension} {label}: {len(r['cache_differences'])} cells where the re-read "
                      f"changes when calculated, {len(r['in_memory_differences'])} (informational) in memory",
                      flush=True)
            oracle["fixtures"][fixture][extension] = {"edits": results}
    with open(path, "w") as f:
        json.dump(oracle, f, indent=1, sort_keys=True)
        f.write("\n")


main(set(sys.argv[1:]))
