"""Makes the fixtures and `oracle.json` of `tests/writer_results.rs`: what the writers of `sheetrs-io` change of a result when they write it, against
`LibreOffice` 26.8.

THE RULE (`CONTRIBUTING.md`, *Parity and fidelity*): the reference is the file LibreOffice saves and opens again, not its memory. For each format (XLSX and ODS):

  1. LibreOffice makes `kinds.<ext>`: one formula (or one constant, in the sheet `C`) per case of `FORMULAS` and `CONSTANTS`, covering the kinds of result a writer
     can change: numbers (integers, fractions, very large and very small), text (empty, spaces, leading `=`, newlines, unicode), booleans, errors (every kind, and
     the `Err:5xx` of an ODS), dates, times, percentages, currencies and other formats, and an array formula;
  2. LibreOffice converts the formulas of `K` to values (`.uno:ConvertFormulaToValue`, *Data > Calculate > Formula to Value*), saves (`lo_converted.<ext>`) and opens
     the file again: `lo` is what it shows. Constants of the sheet `C` are what LibreOffice saves for the same constants (`lo_constants`);
  3. `sheetrs edit --formula-to-value K` (default mode and `--calculate-all`) writes `sheetrs_default.<ext>` / `sheetrs_all.<ext>`, and LibreOffice opens each of them:
     `default` and `all` are what it shows. `differences_*` lists the cells where that is not `lo`;
  4. `raw` holds the XML of each cell in the file of LibreOffice and in the files of `sheetrs edit` (what is written), for the document of the writers.

    scripts/libreoffice-server.sh
    cargo build -p sheetrs-cli
    /opt/libreoffice26.8/program/python make_oracle.py
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import zipfile

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale
from com.sun.star.util import Date

HERE = os.path.dirname(os.path.abspath(__file__))
SHEETRS = os.environ.get("SHEETRS", os.path.join(HERE, "..", "..", "..", "..", "target", "debug", "sheetrs"))
sys.path.insert(0, os.path.join(HERE, "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402

FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8"}

# (label, formula in the notation of the UNO API, number format code or None)
FORMULAS = [
    ("integer", "=2+3", None),
    ("fraction", "=1/3", None),
    ("float sum 0.1+0.2", "=0.1+0.2", None),
    ("negative", "=-5.5", None),
    ("zero", "=0", None),
    ("negative zero", "=-0", None),
    ("1E+300", "=1E+300", None),
    ("1E+15", "=1E+15", None),
    ("2^53+1", "=2^53+1", None),
    ("123456789012345678", "=123456789012345678", None),
    ("1E-300", "=1E-300", None),
    ("1E-15", "=1E-15", None),
    ("1.7976931348623157E+308", "=1.7976931348623157E+308", None),
    ("5E-324", "=4.94065645841247E-324", None),
    ("text", '="abc"', None),
    ("empty text", '=""', None),
    ("leading space", '=" lead"', None),
    ("trailing space", '="trail "', None),
    ("only spaces", '="   "', None),
    ("two inner spaces", '="a  b"', None),
    ("leading =", '="=1+1"', None),
    ("leading apostrophe", "=\"'q\"", None),
    ("newline", '="a"&CHAR(10)&"b"', None),
    ("tab", '="a"&CHAR(9)&"b"', None),
    ("unicode", '="Ünï😀"', None),
    ("numeric text", '="1"', None),
    ("text TRUE", '="TRUE"', None),
    ("text that looks like an error", '="#DIV/0!"', None),
    ("xml specials", '="<a>&amp;\'q\'"', None),
    ("300 characters", '=REPT("x";300)', None),
    ("true", "=TRUE()", None),
    ("false", "=FALSE()", None),
    ("comparison", "=1=1", None),
    ("#DIV/0!", "=1/0", None),
    ("#N/A", "=NA()", None),
    ("#NAME?", "=FOO(1)", None),
    ("#REF!", "=#REF!+1", None),
    ("#VALUE!", '=1+"x"', None),
    ("#NUM!", "=10^1000", None),
    ("#NULL!", "=#NULL!+1", None),
    ("Err:502", "=SQRT(-1)", None),
    ("Err:522 (circular)", "=A{n}", None),
    ("error through a reference", "=A{r1}+1", None),
    ("error in a cell with a format", "=1/0", "0.00%"),
    ("IFERROR gives a number", "=IFERROR(1/0;5)", None),
    ("reference to an empty cell", "=H1", None),
    ("date", "=DATE(2024;3;15)", "YYYY-MM-DD"),
    ("time", "=TIME(13;30;0)", "HH:MM:SS"),
    ("date and time", "=DATE(2024;3;15)+TIME(13;30;45)", "YYYY-MM-DD HH:MM:SS"),
    ("date zero", "=0", "YYYY-MM-DD"),
    ("date before the epoch", "=-1", "YYYY-MM-DD"),
    ("date 9999-12-31", "=DATE(9999;12;31)", "YYYY-MM-DD"),
    ("duration over 24 hours", "=1.5", "[HH]:MM:SS"),
    ("percentage", "=0.25", "0.00%"),
    ("currency", "=1234.5", "[$$-409]#,##0.00"),
    ("scientific format", "=12345.678", "0.00E+00"),
    ("fraction format", "=0.75", "# ?/?"),
    ("thousands separator", "=1234567.891", "#,##0.00"),
    ("text format", '="abc"', "@"),
    ("number in a text format", "=12", "@"),
    ("array formula", None, None),
]
# constants of the sheet `C`: (label, kind, value). The kinds are what `Constant` can be: a number, a text, a boolean (a number with the format TRUE/FALSE in
# LibreOffice), or an empty cell.
CONSTANTS = [
    ("integer", "number", 7),
    ("fraction", "number", 0.1),
    ("1E+300", "number", 1e300),
    ("1E-300", "number", 1e-300),
    ("5E-324", "number", 5e-324),
    ("1.7976931348623157E+308", "number", 1.7976931348623157e308),
    ("123456789012345678", "number", 123456789012345678.0),
    ("negative zero", "number", -0.0),
    ("2^53+1 as a double", "number", 9007199254740993.0),
    ("text", "text", "abc"),
    ("empty text", "text", ""),
    ("leading space", "text", " lead"),
    ("trailing space", "text", "trail "),
    ("only spaces", "text", "   "),
    ("leading =", "text", "=1+1"),
    ("leading apostrophe", "text", "'q"),
    ("newline", "text", "a\nb"),
    ("tab", "text", "a\tb"),
    ("unicode", "text", "Ünï😀"),
    ("numeric text", "text", "1"),
    ("text TRUE", "text", "TRUE"),
    ("text that looks like an error", "text", "#DIV/0!"),
    ("xml specials", "text", "<a>&amp;'q'"),
    ("300 characters", "text", "x" * 300),
    ("true", "boolean", True),
    ("false", "boolean", False),
    ("empty", "empty", None),
    ("NaN", "number", "NaN"),
    ("+inf", "number", "inf"),
    ("-inf", "number", "-inf"),
]


def pv(n, v):
    p = PropertyValue()
    p.Name, p.Value = n, v
    return p


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    smgr = ctx.ServiceManager
    return smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx), smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)


def format_key(doc, code):
    formats = doc.NumberFormats
    locale = Locale("en", "US", "")
    key = formats.queryKey(code, locale, False)
    if key == -1:
        key = formats.addNew(code, locale)
    return key


def fmt_info(doc, cell):
    f = doc.NumberFormats.getByKey(cell.NumberFormat)
    return {"code": f.FormatString, "type": f.Type}


def describe(doc, cell):
    kind = cell.getType().value
    value = cell.getValue()
    if value != value or value in (float("inf"), float("-inf")):
        value = str(value)  # JSON cannot hold it
    out = {"kind": kind, "string": cell.getString(), "value": value, "error": cell.getError(), "formula": cell.getFormula()}
    out["format"] = fmt_info(doc, cell)
    # what a formula shows: 1 number, 2 text, 4 error (`FormulaResultType2`)
    out["result"] = cell.FormulaResultType2 if kind == "FORMULA" else 0
    return out


def make_source(desktop, ext, path, null_date=None):
    doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
    if null_date:
        doc.NullDate = null_date
    doc.Sheets.getByIndex(0).Name = "K"
    doc.Sheets.insertNewByName("C", 1)
    doc.Sheets.insertNewByName("W", 2)
    k, c, w = doc.Sheets.getByName("K"), doc.Sheets.getByName("C"), doc.Sheets.getByName("W")
    for i, (label, formula, code) in enumerate(FORMULAS):
        cell = k.getCellByPosition(0, i)
        if formula is None:
            k.getCellRangeByPosition(0, i, 0, i + 1).setArrayFormula("={1|2}")
            continue
        cell.setFormula(formula.replace("{n}", str(i + 1)).replace("{r1}", str(next(j for j, x in enumerate(FORMULAS) if x[0] == "#DIV/0!") + 1)))
        if code:
            cell.NumberFormat = format_key(doc, code)
    for i, (label, kind, value) in enumerate(CONSTANTS):
        cell = c.getCellByPosition(0, i)
        w.getCellByPosition(0, i).setValue(0.0)  # the cell `sheetrs` writes over (it exists in the file)
        if kind == "number":
            try:
                cell.setValue(float(value))
            except Exception as error:  # LibreOffice cannot hold it
                print("LibreOffice cannot set", label, error)
        elif kind == "text":
            cell.setString(value)
        elif kind == "boolean":
            cell.setValue(1.0 if value else 0.0)
            cell.NumberFormat = format_key(doc, "BOOLEAN")
    doc.calculateAll()
    doc.storeToURL("file://" + path, (pv("FilterName", FILTERS[ext]),))
    doc.close(True)


def read_sheet(doc, name, rows):
    sheet = doc.Sheets.getByName(name)
    return [describe(doc, sheet.getCellByPosition(0, r)) for r in range(rows)]


def reread(desktop, path, name, rows):
    doc = desktop.loadComponentFromURL("file://" + path, "_blank", 0, ())
    out = read_sheet(doc, name, rows)
    doc.close(True)
    return out


def sheetrs_constants(src, out, ext):
    spec = []
    for i, (label, kind, value) in enumerate(CONSTANTS):
        spec.append({"row": i, "col": 0, "kind": kind, "value": value})
    run = subprocess.run(["cargo", "run", "-q", "-p", "sheetrs-edit", "--example", "write_constants", "--", src, out, "W"],
                         input=json.dumps(spec), capture_output=True, text=True, cwd=os.path.join(HERE, "..", "..", "..", ".."))
    return run


def lo_convert(desktop, dispatcher, src, dst, ext):
    doc = desktop.loadComponentFromURL("file://" + src, "_blank", 0, ())
    sheet = doc.Sheets.getByName("K")
    controller = doc.getCurrentController()
    controller.setActiveSheet(sheet)
    controller.select(sheet.getCellRangeByName(f"A1:A{len(FORMULAS) + 1}"))
    dispatcher.executeDispatch(controller.getFrame(), ".uno:ConvertFormulaToValue", "", 0, ())
    time.sleep(0.5)
    doc.storeToURL("file://" + dst, (pv("FilterName", FILTERS[ext]),))
    doc.close(True)


def raw_cells(path, sheet_name, rows):
    """The XML of the cells A1..A<rows> of a sheet, as the file writes it."""
    z = zipfile.ZipFile(path)
    out = []
    if path.endswith(".xlsx"):
        wb = z.read("xl/workbook.xml").decode()
        names = re.findall(r'<sheet name="([^"]+)"', wb)
        xml = z.read(f"xl/worksheets/sheet{names.index(sheet_name) + 1}.xml").decode()
        for r in range(1, rows + 1):
            m = re.search(rf'<c r="A{r}"[^>]*?(?:/>|>.*?</c>)', xml, re.S)
            out.append(m.group(0) if m else None)
        # the shared strings a cell points to
        if "xl/sharedStrings.xml" in z.namelist():
            shared = re.findall(r"<si>(.*?)</si>", z.read("xl/sharedStrings.xml").decode(), re.S)
            out = [re.sub(r'<v>(\d+)</v>', lambda mm: f"<v>{mm.group(1)}</v><!--shared: {shared[int(mm.group(1))][:60]}-->", x) if x and 't="s"' in x else x for x in out]
    else:
        xml = z.read("content.xml").decode()
        m = re.search(rf'<table:table table:name="{re.escape(sheet_name)}".*?</table:table>', xml, re.S)
        body = m.group(0)
        cells = re.findall(r"<table:table-row[^>]*?(?:/>|>.*?</table:table-row>)", body, re.S)
        # expand repeated rows
        for row in cells:
            n = int(re.search(r'table:number-rows-repeated="(\d+)"', row).group(1)) if 'table:number-rows-repeated' in row else 1
            first = re.search(r"<table:(?:covered-)?table-cell[^>]*?(?:/>|>.*?</table:(?:covered-)?table-cell>)", row, re.S)
            out.extend([first.group(0) if first else None] * n)
    return (out + [None] * rows)[:rows]


def compare(a, b):
    """The cells where two re-reads differ (index, LibreOffice's, the other)."""
    diffs = []
    for i, (x, y) in enumerate(zip(a, b)):
        keys = ("kind", "string", "value", "error", "formula", "result")
        if any(x[k] != y[k] for k in keys) or x["format"]["code"] != y["format"]["code"]:
            diffs.append(i)
    return diffs


def main():
    desktop, dispatcher = connect()
    n, nc = len(FORMULAS) + 1, len(CONSTANTS)
    oracle = {
        "formulas": [f[0] for f in FORMULAS],
        "constants": [c[0] for c in CONSTANTS],
        # the constants as `Constant` takes them (a number given as a text is one JSON cannot hold)
        "constants_spec": [{"kind": c[1], "value": c[2]} for c in CONSTANTS],
    }
    with tempfile.TemporaryDirectory() as tmp:
        for ext, variant in [(e, v) for v in ("", "_1904") for e in FILTERS]:
            src = os.path.join(HERE, f"kinds{variant}.{ext}")
            make_source(desktop, ext, src, Date(1, 1, 1904) if variant else None)
            fixture_determinism.normalize(src)
            lo = os.path.join(tmp, f"lo{variant}.{ext}")
            lo_convert(desktop, dispatcher, src, lo, ext)
            entry = {
                "source": reread(desktop, src, "K", n),
                "lo": reread(desktop, lo, "K", n),
                "lo_constants": reread(desktop, lo, "C", nc),
                "source_constants": reread(desktop, src, "C", nc),
                "raw_source": raw_cells(src, "K", n),
                "raw_lo": raw_cells(lo, "K", n),
                "raw_lo_constants": raw_cells(lo, "C", nc),
            }
            wrote = os.path.join(tmp, f"constants{variant}.{ext}")
            run = sheetrs_constants(src, wrote, ext)
            if run.returncode == 0:
                entry["sheetrs_constants"] = reread(desktop, wrote, "W", nc)
                entry["raw_sheetrs_constants"] = raw_cells(wrote, "W", nc)
                entry["differences_constants"] = compare(entry["source_constants"], entry["sheetrs_constants"])
            else:
                entry["sheetrs_constants"] = {"error": run.stderr[-500:]}
            for mode, flags in (("default", []), ("all", ["--calculate-all"])):
                out = os.path.join(tmp, f"sr_{mode}{variant}.{ext}")
                run = subprocess.run([SHEETRS, "edit", "-i", src, "--formula-to-value", "K", *flags, "-o", out], capture_output=True, text=True)
                if run.returncode != 0:
                    entry[mode] = {"error": run.stderr[-500:]}
                    continue
                entry[mode] = reread(desktop, out, "K", n)
                entry["raw_" + mode] = raw_cells(out, "K", n)
                entry["differences_" + mode] = compare(entry["lo"], entry[mode])
            # the results a calculation writes into cells that keep their formulas (`--calculate-all`), against `calculateAll()` and save
            recalc = os.path.join(tmp, f"lo_recalc{variant}.{ext}")
            doc = desktop.loadComponentFromURL("file://" + src, "_blank", 0, ())
            doc.calculateAll()
            doc.storeToURL("file://" + recalc, (pv("FilterName", FILTERS[ext]),))
            doc.close(True)
            entry["recalc_lo"] = reread(desktop, recalc, "K", n)
            entry["raw_recalc_lo"] = raw_cells(recalc, "K", n)
            out = os.path.join(tmp, f"sr_recalc{variant}.{ext}")
            run = subprocess.run([SHEETRS, "edit", "-i", src, "--hide-rows", "K!500", "--calculate-all", "-o", out], capture_output=True, text=True)
            if run.returncode == 0:
                entry["recalc"] = reread(desktop, out, "K", n)
                entry["raw_recalc"] = raw_cells(out, "K", n)
                entry["differences_recalc"] = compare(entry["recalc_lo"], entry["recalc"])
            else:
                entry["recalc"] = {"error": run.stderr[-500:]}
            oracle[ext + variant] = entry
    with open(os.path.join(HERE, "oracle.json"), "w") as f:
        json.dump(oracle, f, indent=1, sort_keys=True)


if __name__ == "__main__":
    main()
