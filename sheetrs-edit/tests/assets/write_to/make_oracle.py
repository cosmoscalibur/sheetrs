"""Makes `oracle.json`, what `tests/write_to_libreoffice.rs` compares the in-memory write (`BookEdit::write_to`,
the one Python and WebAssembly use) with: for each workbook of `tests/assets/use_cache/` (`book`, `rules`, `refs`), in each
format, and for each edit of `EDITS`,

  expected   the cells (formula and value), the hidden rows and columns and the visibility of each sheet after
             LibreOffice 26.8 itself made the edit through UNO on the loaded file and calculated every formula
             (`calculateAll`);
  opened     what LibreOffice shows of the file `sheetrs edit` writes for the same edit, with every formula
             calculated again (`calculateAll`), and
  differences  the cells, hidden rows and columns and sheets where `opened` is not `expected`, for the formulas
             and for the values. They are reported to the maintainer: nothing here is an approved difference;
  reread_differences  the same, between what LibreOffice shows of the file it saves itself after the edit
             (loaded again and calculated: `own_reread`) and `opened`: whether the file `sheetrs edit` writes
             reads back as the one LibreOffice writes (a reference that went is spelled `#REF!` in the files of
             both, and read differently from what the edit gave).

The edit is the same list of operations in the three places: `sheetrs edit` flags, UNO calls, and the methods of
`BookEdit` that the Rust test calls. Run it with a LibreOffice listening on a socket
(`scripts/libreoffice-server.sh`) and a built `sheetrs` (`cargo build -p sheetrs-cli`):

    /opt/libreoffice26.8/program/python make_oracle.py
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "..", "use_cache")
SHEETRS = os.environ.get("SHEETRS", os.path.join(HERE, "..", "..", "..", "..", "target", "debug", "sheetrs"))
FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8"}
ROWS, COLUMNS = 80, 6  # the area of each sheet that is compared: A1:F80

# (label, the operations). An operation is `[name, ...arguments]`:
#   delete_rows / hide_rows / show_rows "S!3:5"      delete_columns / hide_columns / show_columns "S!B:D"
#   delete_sheets / hide_sheets / show_sheets "S"    delete_names "name"       formula_to_value "S!A:A"
#   replace {text, with, scope, regex, in}
BOOK = [
    ("delete rows Data!3", [["delete_rows", "Data!3"]]),
    ("delete rows Data!2:3", [["delete_rows", "Data!2:3"]]),
    ("delete columns Data!B", [["delete_columns", "Data!B"]]),
    ("delete columns Data!D (no cell reads it)", [["delete_columns", "Data!D"]]),
    ("delete sheet Data", [["delete_sheets", "Data"]]),
    ("delete name rate", [["delete_names", "rate"]]),
    ("replace 20 by 25 in Data!A2", [["replace", {"text": "20", "with": "25", "scope": "Data!A2"}]]),
    ("replace ^x$ by y in Data!B1 (regex)",
     [["replace", {"text": "^x$", "with": "y", "scope": "Data!B1", "regex": True}]]),
    ("replace $ by ! in Data!B1 (regex, the end of the text)",
     [["replace", {"text": "$", "with": "!", "scope": "Data!B1", "regex": True}]]),
    ("replace ^ by > in Data!B1 (regex, the start: no match of no characters)",
     [["replace", {"text": "^", "with": ">", "scope": "Data!B1", "regex": True}]]),
    ("replace * by + in the formula of Calc!A2", [["replace", {"text": "*2", "with": "*3", "scope": "Calc!A2", "in": "formula"}]]),
    ("hide rows Data!2", [["hide_rows", "Data!2"]]),
    ("hide rows Data!2:3 and show Data!3", [["hide_rows", "Data!2:3"], ["show_rows", "Data!3"]]),
    ("hide columns Data!B", [["hide_columns", "Data!B"]]),
    ("hide columns Data!B:C and show Data!C", [["hide_columns", "Data!B:C"], ["show_columns", "Data!C"]]),
    ("hide sheet Data", [["hide_sheets", "Data"]]),
    ("formula to value Calc!A:A", [["formula_to_value", "Calc!A:A"]]),
    ("delete rows Data!3 and columns Data!B", [["delete_rows", "Data!3"], ["delete_columns", "Data!B"]]),
    ("replace and hide rows", [["replace", {"text": "20", "with": "25", "scope": "Data!A2"}], ["hide_rows", "Data!4"]]),
]
RULES = [
    ("delete columns Data!D", [["delete_columns", "Data!D"]]),
    ("delete rows Data!6", [["delete_rows", "Data!6"]]),
    ("delete sheet Other", [["delete_sheets", "Other"]]),
    ("delete name rate", [["delete_names", "rate"]]),
    ("replace Data!B1 x by y", [["replace", {"text": "x", "with": "y", "scope": "Data!B1"}]]),
    ("replace 20 by 25 in Data!A2", [["replace", {"text": "20", "with": "25", "scope": "Data!A2"}]]),
    ("replace *2 by *3 in the formula of Calc!A5", [["replace", {"text": "*2", "with": "*3", "scope": "Calc!A5", "in": "formula"}]]),
    ("hide rows Data!2:4", [["hide_rows", "Data!2:4"]]),
    ("hide columns Data!D", [["hide_columns", "Data!D"]]),
]
REFS = [
    ("delete sheet Data", [["delete_sheets", "Data"]]),
    ("delete rows Data!3", [["delete_rows", "Data!3"]]),
    ("delete rows Data!1:6", [["delete_rows", "Data!1:6"]]),
    ("delete rows Data!2:3", [["delete_rows", "Data!2:3"]]),
    ("delete columns Data!A", [["delete_columns", "Data!A"]]),
    ("delete columns Data!A:B", [["delete_columns", "Data!A:B"]]),
]
FIXTURES = [("book", BOOK), ("rules", RULES), ("refs", REFS)]


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
        for r in range(ROWS):
            for c in range(COLUMNS):
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


def apply_lo(doc, operations):
    """The operations on the loaded document, in the order of the planner: the replacements, the visibility,
    the conversions, then the deletions of rows and of columns (from the last to the first), the names and
    the sheets."""
    by_name = {}
    for name, *args in operations:
        by_name.setdefault(name, []).append(args[0])
    sheets = doc.Sheets
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
    for name, *args in operations:
        verb, _, kind = name.partition("_")
        if verb not in ("hide", "show"):
            continue
        visible = verb == "show"
        if kind == "rows":
            sheet, first, last = rows_of(args[0])
            for r in range(first, last + 1):
                sheets.getByName(sheet).Rows.getByIndex(r).IsVisible = visible
        elif kind == "columns":
            sheet, first, last = columns_of(args[0])
            for c in range(first, last + 1):
                sheets.getByName(sheet).Columns.getByIndex(c).IsVisible = visible
        elif kind == "sheets":
            sheets.getByName(args[0]).IsVisible = visible
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


def flags(operations):
    out = ["edit"]
    groups = {}
    replaces = []
    for name, *args in operations:
        if name == "replace":
            replaces.append(args[0])
        else:
            groups.setdefault(name, []).append(args[0])
    command = []
    for name, values in groups.items():
        command += ["--" + name.replace("_", "-"), *values]
    # Each replacement is a search of its own: they share the options, so every edit here has one.
    for spec in replaces:
        command += ["--replace", spec["text"], "--with", spec["with"], "--scope", spec["scope"]]
        if spec.get("regex"):
            command += ["--regex"]
        if spec.get("in"):
            command += ["--in", spec["in"]]
    return out + command


def differences(expected, opened):
    """Where `opened` is not `expected`: a list of `where: expected -> opened`, for formulas, values, hidden
    rows and columns, and sheets."""
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


def main():
    oracle = {"area": f"A1:{a1(ROWS - 1, COLUMNS - 1)}", "fixtures": {}}
    for fixture, edits in FIXTURES:
        oracle["fixtures"][fixture] = {}
        for extension in ("xlsx", "ods"):
            source = os.path.join(ASSETS, f"{fixture}.{extension}")
            results = []
            for label, operations in edits:
                doc = load(source)
                apply_lo(doc, operations)
                doc.calculateAll()
                expected = dump(doc)
                with tempfile.TemporaryDirectory() as folder:
                    own = os.path.join(folder, f"own.{extension}")
                    doc.storeToURL("file://" + own, (pv("FilterName", FILTERS[extension]),))
                    doc.close(True)
                    doc = load(own)
                    doc.calculateAll()
                    own_reread = dump(doc)
                    doc.close(True)
                    out = os.path.join(folder, f"out.{extension}")
                    done = subprocess.run([SHEETRS, *flags(operations), "-i", source, "-o", out],
                                          capture_output=True, text=True)
                    if done.returncode != 0:
                        sys.exit(f"{fixture}.{extension} {label}: {done.stderr}")
                    doc = load(out)
                    written = dump(doc)
                    doc.calculateAll()
                    opened = dump(doc)
                    doc.close(True)
                results.append({
                    "label": label,
                    "operations": operations,
                    "expected": expected,
                    "differences": differences(expected, opened),
                    "reread_differences": differences(own_reread, opened),
                    "written_differences": differences(expected, written),
                })
                print(f"{fixture}.{extension} {label}: {len(results[-1]['differences'])} differences after "
                      f"calculating, {len(results[-1]['written_differences'])} as written, "
                      f"{len(results[-1]['reread_differences'])} from the file LibreOffice saves", flush=True)
            oracle["fixtures"][fixture][extension] = {"edits": results}
    with open(os.path.join(HERE, "oracle.json"), "w") as f:
        json.dump(oracle, f, indent=1, sort_keys=True)
        f.write("\n")


main()
