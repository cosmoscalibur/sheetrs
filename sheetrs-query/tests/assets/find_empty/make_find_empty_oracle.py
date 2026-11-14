"""Makes `input.{xlsx,ods}` and `oracle.json`, the fixtures of `tests/find_empty.rs`: a sheet `T` with
one text a row (the texts of `sheetrs-edit/tests/assets/replace_anchors/make_anchors_oracle.py`, of
one and of several lines) and, for each regular expression of the searches of that script and each
combination of `match_case` and `entire_cells`, the rows that LibreOffice itself reports with
`Sheet.findAll`, loaded from each format. The patterns that match the empty text (`^`, `$`, `x*`,
`a*$`, `(^|$)`) are the point: which cells does `findAll` report for them?

It also records the empty search string in every mode (it finds nothing).

Run it with a LibreOffice listening on a socket (`scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python make_find_empty_oracle.py
"""
import ast, json, os, re, sys, time
import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

REPLACE_ORACLE = os.path.join(
    HERE, "..", "..", "..", "..", "sheetrs-edit", "tests", "assets", "replace_anchors", "make_anchors_oracle.py"
)


def literal(name):
    """The value of the module-level list `name` of the script that makes the replace oracle."""
    tree = ast.parse(open(REPLACE_ORACLE, encoding="utf-8").read())
    for node in tree.body:
        if isinstance(node, ast.Assign) and node.targets[0].id == name:
            return ast.literal_eval(node.value)
    raise KeyError(name)


TEXTS = literal("TEXTS")
PATTERNS = []
for search, _with in literal("CASES"):
    if search not in PATTERNS:
        PATTERNS.append(search)
FLAGS = [(False, False), (True, False), (False, True), (True, True)]  # (match_case, entire_cells)

local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def open_doc(path):
    return desktop.loadComponentFromURL("file://" + path, "_blank", 0, (pv("Hidden", True),))


def rows_of(result):
    """The 0-based rows of `T.A1:A3,T.A5` (a column of a sheet)."""
    if result is None:
        return []
    rows = set()
    for part in filter(None, result.getRangeAddressesAsString().split(",")):
        rng = part.rsplit(".", 1)[1]
        a, _, b = rng.partition(":")
        first = int(re.sub(r"[^0-9]", "", a))
        last = int(re.sub(r"[^0-9]", "", b or a))
        rows.update(range(first - 1, last))
    return sorted(rows)


doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
sheet = doc.Sheets.getByIndex(0)
sheet.Name = "T"
for i, text in enumerate(TEXTS):
    sheet.getCellByPosition(0, i).String = text
doc.storeToURL("file://" + os.path.join(HERE, "input.xlsx"), (pv("FilterName", "Calc MS Excel 2007 XML"),))
doc.storeToURL("file://" + os.path.join(HERE, "input.ods"), (pv("FilterName", "calc8"),))
for ext in ("xlsx", "ods"):
    fixture_determinism.normalize(os.path.join(HERE, "input." + ext))
doc.close(True)

oracle = {"texts": TEXTS, "cases": [], "empty": []}
for ext in ("xlsx", "ods"):
    doc = open_doc(os.path.join(HERE, "input." + ext))
    sheet = doc.Sheets.getByName("T")
    loaded = [sheet.getCellByPosition(0, i).String for i in range(len(TEXTS))]
    assert loaded == TEXTS, (ext, loaded)
    for search in PATTERNS:
        for match_case, entire in FLAGS:
            d = sheet.createSearchDescriptor()
            d.SearchString = search
            d.SearchType = 1
            d.SearchRegularExpression = True
            d.SearchCaseSensitive = match_case
            d.SearchWords = entire
            d.setPropertyValue("SearchFormatted", False)
            oracle["cases"].append(
                {"format": ext, "find": search, "match_case": match_case, "entire_cells": entire,
                 "rows": rows_of(sheet.findAll(d))}
            )
    for mode in ("literal", "regex", "wildcards"):
        for match_case, entire in FLAGS:
            d = sheet.createSearchDescriptor()
            d.SearchString = ""
            d.SearchType = 1
            d.SearchRegularExpression = mode == "regex"
            d.SearchWildcard = mode == "wildcards"
            d.SearchCaseSensitive = match_case
            d.SearchWords = entire
            d.setPropertyValue("SearchFormatted", False)
            oracle["empty"].append(
                {"format": ext, "mode": mode, "match_case": match_case, "entire_cells": entire,
                 "rows": rows_of(sheet.findAll(d))}
            )
    doc.close(True)
with open(os.path.join(HERE, "oracle.json"), "w") as f:
    json.dump(oracle, f, indent=1, ensure_ascii=False)
print(len(oracle["cases"]), "searches,", len(oracle["empty"]), "empty searches")
