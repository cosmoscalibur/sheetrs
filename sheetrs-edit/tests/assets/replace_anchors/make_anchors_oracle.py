"""Makes `input.{xlsx,ods}` and `oracle.json`, the fixtures of `tests/replace_anchors.rs`: a sheet `T`
with one text a row (texts of one and of several lines), and, for each search of `CASES` and each
combination of `match_case` and `entire_cells`, the text of each cell after LibreOffice itself ran
`Sheet.replaceAll` with a regular expression over the workbook, loaded from each format.

Run it with a LibreOffice listening on a socket (`scripts/libreoffice-server.sh`):

    /opt/libreoffice26.8/program/python make_anchors_oracle.py
"""
import json, os, sys, time
import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

TEXTS = [
    "a", "ab", "abc", "ba", "aa", "aaa", "A", "AB", "abab", "abcabc", "abc abc", "a b", "x", "xa", "ax",
    "xax", "bab", "xabc", "abcx", "aab", "b", "bb", "a1", "ab12",
    "a\nb", "a\n", "\na", "a\nb\nc", "ab\n", "ab\nab", "a\n\n", "a\u2028", "x\na",
    "a\nx", "abc\nabc", "b\na", "\n", "\n\n", "a\nb\n", "ab\n\n", "é", "éa", "aé",
    "\U0001F600", "\U0001F600a", "a\U0001F600", "a\U0001F600b",
]
# (search, replacement). The replacement reads `&`, `$0`..`$9` and `\t`, `\&`, `\$`, `\\`.
CASES = [
    ("^abc", "X"), ("abc$", "X"), ("^abc$", "X"), ("^a", "X"), ("a$", "X"), ("^a$", "X"), ("^b", "X"),
    ("b$", "X"), ("^ab", "X"), ("ab$", "X"), ("^", "X"), ("$", "X"), ("^$", "X"), ("a*$", "X"),
    ("^a*", "X"), ("a*", "X"), ("x*", "X"), ("a?$", "X"), ("^a?", "X"), ("a+$", "X"), ("^a+", "X"),
    ("(^|x)a", "X"), ("a($|x)", "X"), ("(a|b$)", "X"), ("(^a|b$)", "X"), ("^(a|b)$", "X"),
    ("^(a|b)", "X"), ("(a|b)$", "X"), ("a^b", "X"), ("a$b", "X"), ("^a|b$", "X"), ("a$|^b", "X"),
    ("^.", "X"), (".$", "X"), ("^.*", "X"), (".*$", "X"), ("^.+", "X"), (".+$", "X"), ("^.*$", "X"),
    ("^a.", "X"), ("a.$", "X"), ("^\\s", "X"), ("\\s$", "X"), ("a\\s$", "X"), ("^\\w+$", "X"),
    ("\\w+$", "X"), ("^\\w+", "X"), ("^a|a$", "X"), ("(^a)+", "X"), ("^^a", "X"), ("a$$", "X"),
    ("$a", "X"), ("b^", "X"), ("^a*b", "X"), ("a*b$", "X"), ("^(ab)*", "X"), ("(ab)*$", "X"),
    ("[ab]$", "X"), ("^[ab]", "X"), ("^[^a]", "X"), ("[^a]$", "X"), ("^\\d", "X"), ("\\d$", "X"),
    ("\\d+$", "X"), ("^a{2}", "X"), ("a{2}$", "X"), ("^a{0,1}", "X"), ("a{0,2}$", "X"),
    ("^(a|)", "X"), ("(|a)$", "X"), ("(^|$)", "X"), ("(^|b)", "X"), ("(b|$)", "X"), ("(^)", "X"),
    ("($)", "X"), ("^(?:a|b)", "X"), ("(?:a|b)$", "X"),
    ("^(a)", "<$1>"), ("(a)$", "<$1>"), ("^(a)(b)", "<$2$1>"), ("^a", "[&]"), ("a$", "[&]"),
    ("^a", "[$0]"), ("a$", "[$0]"), ("^", "[&]"), ("$", "[&]"), ("$", "[$0]"), ("^", "[$0]"),
    ("^(a*)", "<$1>"), ("(a*)$", "<$1>"), ("^(a|b)", "<$1>"), ("(^|x)(a)", "<$1|$2>"),
    ("$", "\\t"), ("^", "\\t"), ("$", "\\&"), ("$", "\\$"), ("$", "\\\\"), ("$", ""), ("^a", ""),
    ("a$", ""), ("^abc$", ""), ("$", "&&"), ("a$", "$$"),
    ("a|", "X"), ("|a", "X"), ("^|a", "X"), ("a|$", "X"), ("$|^", "X"), ("\\s*$", "X"), ("^\\s*", "X"),
    (".*", "X"), (".?", "X"), (".?$", "X"), ("a?", "X"), ("(ab)?", "X"), ("b*$", "X"), ("[ab]*", "X"),
    ("[ab]*$", "X"), ("x?$", "X"), ("(a|b)*$", "X"), ("^(a|b)*", "X"), ("a*?", "X"), ("a*?b", "X"),
    ("a*?$", "X"), ("a??", "X"), ("^a*?", "X"), ("\\W*$", "X"), (".", "X"), ("..", "X"), (".$", "[&]"),
    # A replacement is made in the text itself, and the next search goes on in what is left of it.
    ("^a", ""), ("^ab", ""), ("^a*", ""), ("^a|b", ""), ("^(ab)", "a"), ("^a", "a"), ("^b*a", "b"),
    ("(^|b)a", ""), ("^.", ""), ("^.", "$0$0"), ("^a", "aa"), ("^(a)", "$1$1"), ("^a|^b", ""),
    ("(^|a)b", ""), ("^\\w", ""), ("^\\w+", ""), ("^[ab]", ""), ("^a?", ""), ("$", "$0"),
    ("a|b", "X"), ("a", "X"), ("a", "&&"), ("(a)", "$1$1"),
]
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

oracle = {"texts": TEXTS, "cases": []}
for ext in ("xlsx", "ods"):
    doc = open_doc(os.path.join(HERE, "input." + ext))
    sheet = doc.Sheets.getByName("T")
    loaded = [sheet.getCellByPosition(0, i).String for i in range(len(TEXTS))]
    assert loaded == TEXTS, (ext, loaded)
    for search, with_ in CASES:
        for match_case, entire in FLAGS:
            for i, text in enumerate(TEXTS):
                sheet.getCellByPosition(0, i).String = text
            d = sheet.createReplaceDescriptor()
            d.SearchString, d.ReplaceString = search, with_
            d.SearchType = 1
            d.SearchRegularExpression = True
            d.SearchCaseSensitive = match_case
            d.SearchWords = entire
            d.setPropertyValue("SearchFormatted", False)
            sheet.replaceAll(d)
            after = []
            for i in range(len(TEXTS)):
                try:
                    after.append(sheet.getCellByPosition(0, i).String)
                except Exception:  # a lone surrogate: not a text that UNO can give
                    after.append(None)
                    print("not a text:", ext, repr(search), repr(with_), repr(TEXTS[i]))
            oracle["cases"].append(
                {"format": ext, "find": search, "with": with_, "match_case": match_case,
                 "entire_cells": entire, "result": after}
            )
    doc.close(True)
with open(os.path.join(HERE, "oracle.json"), "w") as f:
    json.dump(oracle, f, indent=1, ensure_ascii=False)
print(len(oracle["cases"]), "replacements")
