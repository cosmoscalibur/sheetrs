"""Writes `find_formatted.json`: what the Find & Replace of LibreOffice (`findAll`) finds, with and
without `SearchFormatted`, in the workbooks `sheetrs-edit/tests/assets/replace/input.*` (cells whose
shown text differs from their input line) and `sheetrs-io/tests/assets/inherit.*` (results of
formulas whose format they take from the result). `tests/libreoffice_find.rs` runs the same searches
through `sheetrs-query` and compares the cells.

Run it with a LibreOffice listening on a socket (`scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python make_find_oracle.py
"""
import json, os, re, time
import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
FILES = {
    "input": "sheetrs-edit/tests/assets/replace",
    "inherit": "sheetrs-io/tests/assets",
    "filtered": "sheetrs-query/tests/assets",
    "filtered_range": "sheetrs-query/tests/assets",
    "threaded": "sheetrs-query/tests/assets",
    "anchors": "sheetrs-query/tests/assets",
}
# (file, text, formatted, entire cells, regular expression[, filtered: `SearchFiltered`[, in notes]])
# `threaded.xlsx` has an Excel threaded comment (A1: the placeholder Excel keeps in the legacy note,
# with the first message after `Comment:`, and a reply that only the threaded part holds):
# LibreOffice reads the legacy note and ignores the threaded part, and so does the reader.
# `anchors.*` (`make_anchors_fixtures.py`) has texts of one and several lines for `^` and `$`.
# `filtered.*` is an autofilter on an anonymous database range (A1:B6, key = x) with a row hidden
# by hand below it; `filtered_range.*` is the same filter on a named database range, which an XLSX
# keeps in a table part (LibreOffice does not read that one back as filtered).
SEARCHES = [
    ("input", "1,234.50", True, False, False),
    ("input", "1,234.50", False, False, False),
    ("input", "1,234", True, False, False),
    ("input", "1234.5", True, False, False),
    ("input", "1234.5", False, False, False),
    ("input", "15 Mar 2023", True, True, False),
    ("input", "15 Mar", True, False, False),
    ("input", "Mar", False, False, False),
    ("input", "03/15/2023", False, False, False),
    ("input", "03/15/2023", True, False, False),
    ("input", "$10.00", True, True, False),
    ("input", "$10.00", False, True, False),
    ("input", "x5", True, False, False),
    ("input", "x5", False, False, False),
    ("input", "yes", True, True, False),
    ("input", "yes", False, True, False),
    ("input", "25.6%", True, True, False),
    ("input", "25.6", True, False, False),
    ("input", "0.256", True, False, False),
    ("input", "50.00%", True, True, False),
    ("input", "50%", False, True, False),
    ("input", "2023", True, False, False),
    ("input", "\\d+,\\d+", True, False, True),
    ("input", "2,469.00", True, True, False),
    ("input", "2469", True, True, False),
    ("input", "2469", False, True, False),
    ("inherit", "2023-03-15", True, True, False),
    ("inherit", "03/15/23", True, True, False),
    ("inherit", "$10.00", True, True, False),
    ("inherit", "$45,010.00", True, True, False),
    ("inherit", "12.50%", True, True, False),
    ("inherit", "06:00:00 AM", True, True, False),
    ("inherit", "45000.25", True, False, False),
    ("inherit", "45000.25", False, False, False),
    ("inherit", "TRUE", True, True, False),
    ("inherit", "03/15/23 06:00 AM", True, False, False),
    ("filtered", "hit", False, False, False, False),
    ("filtered", "hit", False, False, False, True),
    ("filtered", "hit b", False, False, False, False),
    ("filtered", "outside", False, False, False, False),
    ("filtered", "key", False, False, False, False),
    ("filtered", "key", False, False, False, True),
    ("filtered_range", "hit", False, False, False, False),
    ("filtered_range", "hit", False, False, False, True),
    ("threaded", "Real thread", False, False, False, False, True),
    ("threaded", "Threaded comment", False, False, False, False, True),
    ("threaded", "A reply", False, False, False, False, True),
    ("anchors", "^a", False, False, True),
    ("anchors", "^a", False, True, True),
    ("anchors", "a$", False, False, True),
    ("anchors", "a$", False, True, True),
    ("anchors", "^a$", False, False, True),
    ("anchors", "^a$", False, True, True),
    ("anchors", "^(a|b)$", False, False, True),
    ("anchors", "^(a|b)$", False, True, True),
    ("anchors", "^(Code|Resumen Contadia)$", False, False, True),
    ("anchors", "^(Code|Resumen Contadia)$", False, True, True),
    ("anchors", "(^a|b$)", False, False, True),
    ("anchors", "(^a|b$)", False, True, True),
    ("anchors", "a^b", False, False, True),
    ("anchors", "a^b", False, True, True),
    ("anchors", "a$b", False, False, True),
    ("anchors", "a$b", False, True, True),
    ("anchors", "^x$", False, False, True),
    ("anchors", "^x$", False, True, True),
    ("anchors", "^a.", False, False, True),
    ("anchors", "^a.", False, True, True),
    ("anchors", "a.$", False, False, True),
    ("anchors", "a.$", False, True, True),
    ("anchors", "^.", False, False, True),
    ("anchors", "^.", False, True, True),
    ("anchors", "^ab|cd$", False, False, True),
    ("anchors", "^ab|cd$", False, True, True),
    ("anchors", "(^|x)a", False, False, True),
    ("anchors", "(^|x)a", False, True, True),
    ("anchors", "a($|x)", False, False, True),
    ("anchors", "a($|x)", False, True, True),
    ("anchors", "^a|a$", False, False, True),
    ("anchors", "^a|a$", False, True, True),
    ("anchors", "^b", False, False, True),
    ("anchors", "^b", False, True, True),
    ("anchors", "b$", False, False, True),
    ("anchors", "b$", False, True, True),
    ("anchors", "^code", False, False, True),
    ("anchors", "^code", False, True, True),
    ("anchors", "^code x$", False, False, True),
    ("anchors", "^code x$", False, True, True),
    ("anchors", "x$", False, False, True),
    ("anchors", "x$", False, True, True),
    ("anchors", "^x", False, False, True),
    ("anchors", "^x", False, True, True),
    ("anchors", "a.\\s$", False, False, True),
    ("anchors", "a.\\s$", False, True, True),
    ("anchors", "a\\s$", False, False, True),
    ("anchors", "a\\s$", False, True, True),
    ("anchors", "(^a|a$)\\s", False, False, True),
    ("anchors", "(^a|a$)\\s", False, True, True),
    ("anchors", "^.*x", False, False, True),
    ("anchors", "^.*x", False, True, True),
    ("anchors", "^\\s", False, False, True),
    ("anchors", "^\\s", False, True, True),
    ("anchors", "\\s$", False, False, True),
    ("anchors", "\\s$", False, True, True),
    ("anchors", "a$|^b", False, False, True),
    ("anchors", "a$|^b", False, True, True),
    ("anchors", "(^a)+", False, False, True),
    ("anchors", "(^a)+", False, True, True),
]


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def cells_of(text):
    """`Sheet1.A1:A4,Sheet1.B2` as a sorted list of `Sheet1!A1`."""
    out = set()
    for part in filter(None, text.split(",")):
        sheet, rng = part.rsplit(".", 1)
        sheet = sheet.strip("$'")
        a, _, b = rng.partition(":")
        b = b or a

        def split(x):
            m = re.match(r"\$?([A-Z]+)\$?(\d+)$", x)
            col = 0
            for ch in m.group(1):
                col = col * 26 + ord(ch) - 64
            return int(m.group(2)), col

        (r0, c0), (r1, c1) = split(a), split(b)
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                letters, n = "", c
                while n:
                    n, rem = divmod(n - 1, 26)
                    letters = chr(65 + rem) + letters
                out.add(f"{sheet}!{letters}{r}")
    return sorted(out)


local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
result = []
for stem, folder in FILES.items():
    for ext in ("xlsx", "ods"):
        path = os.path.join(REPO, folder, f"{stem}.{ext}")
        if not os.path.exists(path):  # a fixture that exists in one format only (`threaded.xlsx`)
            continue
        doc = desktop.loadComponentFromURL("file://" + path, "_blank", 0, (prop("Hidden", True),))
        for file, text, formatted, entire, regex, *more in SEARCHES:
            filtered = bool(more and more[0])
            in_notes = bool(len(more) > 1 and more[1])
            if file != stem:
                continue
            found = []
            for sheet in doc.Sheets:
                d = sheet.createSearchDescriptor()
                d.SearchString = text
                d.SearchType = 2 if in_notes else 1
                d.SearchWords = entire
                d.SearchRegularExpression = regex
                d.SearchCaseSensitive = False
                d.setPropertyValue("SearchFormatted", formatted)
                d.setPropertyValue("SearchFiltered", filtered)
                res = sheet.findAll(d)
                if res is not None:
                    found += cells_of(res.getRangeAddressesAsString())
            result.append(
                {
                    "file": f"{stem}.{ext}",
                    "folder": folder,
                    "text": text,
                    "formatted": formatted,
                    "filtered": filtered,
                    "in_notes": in_notes,
                    "entire_cells": entire,
                    "regex": regex,
                    "found": sorted(found),
                }
            )
        doc.close(True)
with open(os.path.join(HERE, "find_formatted.json"), "w") as f:
    json.dump(result, f, indent=1, ensure_ascii=False)
print(len(result), "searches")
