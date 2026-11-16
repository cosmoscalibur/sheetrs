"""Makes `oracle.json`, what `tests/write_to_libreoffice.rs` compares the in-memory write (`BookEdit::write_to`,
the one Python and WebAssembly use) with.

THE RULE (policy of the project, `CONTRIBUTING.md`, *Parity and fidelity*): the expected values of an edit are what
LibreOffice 26.8 shows when the workbook it left after the edit is SAVED AND OPENED AGAIN (re-read). The state of its
memory right after the edit is not a reference: the same workbook, saved and opened by someone else, shows other values
(it is not reproducible), which is a defect of LibreOffice and not a behavior to copy. The rule is for every operation
and both formats.

For each workbook of `tests/assets/cached/` (`book`, `rules`, `refs`, `preref`, `dirty`), in each format, and for each edit of
`EDITS`, LibreOffice 26.8 makes the edit through UNO on the loaded file and calculates every formula (`calculateAll`: the
recalculation the screen does as it repaints, so that no result stays stale), saves the file (`storeToURL`) and loads it
again. The oracle holds:

  reread                the cells (formula and value), the hidden rows and columns and the visibility of each sheet of
                        that file as LibreOffice shows them when it opens it (default load: the results the file
                        caches, plus what LibreOffice itself calculates on load), after the edit and `calculateAll()`.
                        THE EXPECTED VALUES of the mode `calculate_all` (`sheetrs edit --calculate-all`);
  reread_default        the same, for an edit with no `calculateAll()`: what LibreOffice keeps of the results the file
                        caches and calculates because the edit reaches it. THE EXPECTED VALUES of the default mode;
  differences_default   where the file of `sheetrs edit` (default mode) is not `reread_default`;
  cache_differences     where the same file, after a hard recalculation (`calculateAll`), shows another value than
                        `reread` (a cross-check; `reread` is the official one);
  opened                what LibreOffice shows of the file `sheetrs edit` writes for the same edit, loaded with its
                        default load (`written`) and after `calculateAll` (`opened`);
  differences           where `written` is not `reread`, for the formulas and for the values: the file of `sheetrs edit`
                        against the expected one. Nothing here is an approved difference unless the test says so;
  hard_differences      the same, between `opened` (a hard recalculation of the file of `sheetrs edit`) and the file of
                        LibreOffice recalculated the same way: whether the two files are the same program for
                        LibreOffice;
  in_memory_differences INFORMATIONAL ONLY, never a reference: where the memory of LibreOffice right after the edit
                        (and `calculateAll`) differs from `reread`. They are the evidence that the memory is not stable
                        (a reference that went counts as `#REF!` there and as an error value or an unknown name after
                        saving and loading).

The edit is the same list of operations in the three places: `sheetrs edit` flags, UNO calls, and the methods of
`BookEdit` that the Rust test calls. Run it with a LibreOffice listening on a socket
(`scripts/libreoffice-server.sh`) and a built `sheetrs` (`cargo build -p sheetrs-cli`):

    /opt/libreoffice26.8/program/python make_oracle.py [FIXTURE ...]

With fixture names (`book`, `rules`, `refs`, `preref`, `dirty`, `classes`) it makes only those again.
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
ASSETS = os.path.join(HERE, "..", "cached")
SHEETRS = os.environ.get("SHEETRS", os.path.join(HERE, "..", "..", "..", "..", "target", "debug", "sheetrs"))
FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8"}
ROWS, COLUMNS = 80, 6  # the area of each sheet that is compared: A1:F80

# (label, the operations). An operation is `[name, ...arguments]`:
#   delete_rows / hide_rows / show_rows "S!3:5"      delete_columns / hide_columns / show_columns "S!B:D"
#   delete_sheets / hide_sheets / show_sheets "S"    delete_names "name"       formula_to_value "S!A:A"
#   replace {text, with, scope, regex, in}           set_value "S!A2" 25 (an input of a calculation: the value is
#   written and every formula calculated; the API has it, `sheetrs edit` has no flag for it, so the file is not compared)
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
    ("write 25 in Data!A2 (input of a calculation)", [["set_value", "Data!A2", 25]]),
    ("write 9 in Data!C1 (the name rate)", [["set_value", "Data!C1", 9]]),
    ("write y in Data!B1 (a text)", [["set_value", "Data!B1", "y"]]),
    ("write 55 in Data!A3 and delete rows Data!3", [["set_value", "Data!A3", 55], ["delete_rows", "Data!3"]]),
]
RULES = [
    ("delete columns Data!D", [["delete_columns", "Data!D"]]),
    ("delete rows Data!6", [["delete_rows", "Data!6"]]),
    ("delete rows Data!2 (the cell ISFORMULA and FORMULATEXT ask about)", [["delete_rows", "Data!2"]]),
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
    ("write 55 in Data!A3", [["set_value", "Data!A3", 55]]),
]
# An error written in the formulas next to references an edit removes (`Data!A3` on rows, `Other!A1` on a sheet).
PREREF = [
    ("delete rows Data!3", [["delete_rows", "Data!3"]]),
    ("delete columns Data!C", [["delete_columns", "Data!C"]]),
    ("delete sheet Other", [["delete_sheets", "Other"]]),
    ("delete sheet Data", [["delete_sheets", "Data"]]),
    ("delete rows Data!3 and sheet Other", [["delete_rows", "Data!3"], ["delete_sheets", "Other"]]),
    ("hide rows Data!2 (no deletion)", [["hide_rows", "Data!2"]]),
    ("replace x by y in Data!B1 (no deletion)", [["replace", {"text": "x", "with": "y", "scope": "Data!B1"}]]),
    ("write 5 in Data!A1 (no deletion)", [["set_value", "Data!A1", 5]]),
]
# What a deletion reaches in a workbook whose results are all stale (the text `S`: `LibreOffice` keeps it where nothing reaches
# the formula, and calculates it where something does): a number of the ODS would be calculated on load.
DIRTY = [
    ("delete rows Data!3", [["delete_rows", "Data!3"]]),
    ("delete rows Data!2:3", [["delete_rows", "Data!2:3"]]),
    ("delete rows Data!6 (the last of the range)", [["delete_rows", "Data!6"]]),
    ("delete rows Data!7:8 (below everything)", [["delete_rows", "Data!7:8"]]),
    ("delete rows Data!1", [["delete_rows", "Data!1"]]),
    ("delete rows Data!4", [["delete_rows", "Data!4"]]),
    ("delete rows Data!2 and Data!5", [["delete_rows", "Data!2"], ["delete_rows", "Data!5"]]),
    ("delete columns Data!B", [["delete_columns", "Data!B"]]),
    ("delete columns Data!C", [["delete_columns", "Data!C"]]),
    ("delete columns Data!D", [["delete_columns", "Data!D"]]),
    ("delete columns Data!A", [["delete_columns", "Data!A"]]),
    ("delete columns Data!E", [["delete_columns", "Data!E"]]),
    ("delete sheet Other", [["delete_sheets", "Other"]]),
    ("delete sheet Data", [["delete_sheets", "Data"]]),
    ("delete name n_cell1", [["delete_names", "n_cell1"]]),
    ("delete rows Calc!100 (below the formulas)", [["delete_rows", "Calc!100"]]),
    ("delete rows Calc!2 (the formulas below move)", [["delete_rows", "Calc!2"]]),
    ("delete columns Calc!C (empty)", [["delete_columns", "Calc!C"]]),
    ("delete columns Calc!D (the cells to its right move)", [["delete_columns", "Calc!D"]]),
]
# One formula for each class of function (`classes`, every result cached as the text `S`): what `LibreOffice` calculates when it opens
# the file (the first edit leaves every formula alone) and what each edit reaches.
CLASSES = [
    ("hide columns Data!H (no edit: no cell is in it)", [["hide_columns", "Data!H"]]),
    ("hide rows Data!2", [["hide_rows", "Data!2"]]),
    ("hide rows Data!3", [["hide_rows", "Data!3"]]),
    ("hide rows Data!6", [["hide_rows", "Data!6"]]),
    ("hide rows Data!2 and show Data!2", [["hide_rows", "Data!2"], ["show_rows", "Data!2"]]),
    ("write 5 in Data!A1", [["set_value", "Data!A1", 5]]),
    ("write y in Data!B1 (a text)", [["set_value", "Data!B1", "y"]]),
    ("delete rows Data!3", [["delete_rows", "Data!3"]]),
    ("delete rows Data!5 (the cell FORMULATEXT and ISFORMULA ask about)", [["delete_rows", "Data!5"]]),
    ("delete columns Data!B", [["delete_columns", "Data!B"]]),
    ("delete rows Calc!2", [["delete_rows", "Calc!2"]]),
]
# A formula that reads cells at run time (`INDIRECT`) next to formulas that cache results the engine would not give: `LibreOffice` calculates
# the first when it opens the file, and what reads it, and the others keep what the file caches.
INDIRECT = [
    ("hide columns Data!H (no edit: no cell is in it)", [["hide_columns", "Data!H"]]),
    ("delete rows Data!3", [["delete_rows", "Data!3"]]),
    ("delete columns Data!D", [["delete_columns", "Data!D"]]),
    ("write 25 in Data!A2", [["set_value", "Data!A2", 25]]),
    ("write y in Data!B1 (what the INDIRECT counts)", [["set_value", "Data!B1", "y"]]),
]
# `visibility`: rows 3-4, 6 and 10-14 of `Data` are hidden in the file (10-13 are one repeated row of an ODS; 6 and 10-14 are empty). LibreOffice
# dirties the formulas of the rows whose visibility CHANGES: hiding a hidden row or showing a visible one reaches nothing.
VISIBILITY = [
    ("hide rows Data!3 (already hidden)", [["hide_rows", "Data!3"]]),
    ("show rows Data!2 (already visible)", [["show_rows", "Data!2"]]),
    ("hide rows Data!2 (visible)", [["hide_rows", "Data!2"]]),
    ("show rows Data!3 (hidden)", [["show_rows", "Data!3"]]),
    ("hide rows Data!2:4 (two of them hidden already)", [["hide_rows", "Data!2:4"]]),
    ("show rows Data!3:5 (two of them hidden)", [["show_rows", "Data!3:5"]]),
    ("show rows Data!6 (hidden and empty, between values)", [["show_rows", "Data!6"]]),
    ("hide rows Data!16 (visible and empty, past the last value)", [["hide_rows", "Data!16"]]),
    ("hide rows Data!5:7", [["hide_rows", "Data!5:7"]]),
    ("show rows Data!10:15 (hidden empty rows and a visible value)", [["show_rows", "Data!10:15"]]),
    ("hide rows Data!11 (inside a repeated hidden row)", [["hide_rows", "Data!11"]]),
    ("show rows Data!12 (inside a repeated hidden row)", [["show_rows", "Data!12"]]),
    ("hide rows Data!9:12 (partly hidden)", [["hide_rows", "Data!9:12"]]),
    ("hide Data!2 and show Data!2 (in one plan)", [["hide_rows", "Data!2"], ["show_rows", "Data!2"]]),
    ("hide Data!3 and show Data!3 (hidden already)", [["hide_rows", "Data!3"], ["show_rows", "Data!3"]]),
    ("hide Data!2:3 and show Data!3:4", [["hide_rows", "Data!2:3"], ["show_rows", "Data!3:4"]]),
]
FIXTURES = [("book", BOOK), ("rules", RULES), ("refs", REFS), ("preref", PREREF), ("dirty", DIRTY), ("classes", CLASSES), ("indirect", INDIRECT), ("visibility", VISIBILITY)]


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
    """The operations on the loaded document, in the order of the planner: the replacements, the visibility,
    the conversions, then the deletions of rows and of columns (from the last to the first), the names and
    the sheets. With `calculate_all` a conversion is of what the formulas give after `calculateAll()` (the
    mode `calculate_all` of `sheetrs`); without it, of what the cells hold (the results the file caches
    plus what the edit made LibreOffice calculate: the default mode)."""
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


def flags(operations, calculate_all=True):
    """The flags of `sheetrs edit` for the edit, `None` when it has an operation that only the API has."""
    if any(name == "set_value" for name, *_ in operations):
        return None
    out = ["edit"] + (["--calculate-all"] if calculate_all else [])
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
                # THE DEFAULT MODE: the edit on the loaded file, no `calculateAll()`, saved and opened again:
                # LibreOffice keeps what the file caches and calculates what the edit reaches (and what it
                # calculates on load).
                doc = load(source)
                apply_lo(doc, operations, calculate_all=False)
                with tempfile.TemporaryDirectory() as folder:
                    own = os.path.join(folder, f"default.{extension}")
                    doc.storeToURL("file://" + own, (pv("FilterName", FILTERS[extension]),))
                    doc.close(True)
                    doc = load(own)
                    reread_default = dump(doc)
                    doc.close(True)
                    command = flags(operations, calculate_all=False)
                    if command is None:
                        written_default = None
                    else:
                        out = os.path.join(folder, f"default_out.{extension}")
                        done = subprocess.run([SHEETRS, *command, "-i", source, "-o", out],
                                              capture_output=True, text=True)
                        if done.returncode != 0:
                            sys.exit(f"{fixture}.{extension} {label}: {done.stderr}")
                        doc = load(out)
                        written_default = dump(doc)
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
                    out = os.path.join(folder, f"out.{extension}")
                    command = flags(operations)
                    if command is None:
                        written = opened = None
                    else:
                        done = subprocess.run([SHEETRS, *command, "-i", source, "-o", out],
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
                    "reread": reread,
                    "reread_default": reread_default,
                    "differences_default": [] if written_default is None
                    else differences(reread_default, written_default),
                    "cache_differences": differences(reread, reread_hard),
                    "differences": [] if written is None else differences(reread, written),
                    "hard_differences": [] if opened is None else differences(reread_hard, opened),
                    "in_memory_differences": differences(reread, in_memory),
                })
                r = results[-1]
                print(f"{fixture}.{extension} {label}: {len(r['differences'])} differences from the re-read, "
                      f"{len(r['hard_differences'])} after calculating both, {len(r['cache_differences'])} "
                      f"cells where the re-read changes when calculated, {len(r['in_memory_differences'])} "
                      f"(informational) in memory", flush=True)
            oracle["fixtures"][fixture][extension] = {"edits": results}
    with open(path, "w") as f:
        json.dump(oracle, f, indent=1, sort_keys=True)
        f.write("\n")


main(set(sys.argv[1:]))
