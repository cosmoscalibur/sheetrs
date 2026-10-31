#!/usr/bin/env python3
"""Does LibreOffice compute the same formulas differently depending on the file format?

Loads tests/assets/semantics.xlsx in LibreOffice and recalculates it in full (X: how
LibreOffice executes an XLSX), builds the same formulas in a document created from
scratch, as a native ODS is (O: ODF defaults), copies X's document settings onto that
document (O2), and saves X as ODS and reloads it (R: what an XLSX exported to ODS does).
Writes one JSON line per formula with the four results, and prints the differences.

Run with LibreOffice's own Python (it has `uno`):
  SOFFICE=soffice26 PYLO=/opt/.../program/python scripts/compare_formats.py OUT.json [XLSX]
"""
import json
import os
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

SOFFICE = os.environ.get("SOFFICE", "soffice")
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "..", "tests", "assets", "semantics.xlsx")
OUT = sys.argv[1] if len(sys.argv) > 1 else "compare.json"
PORT = 2011


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def connect(profile):
    proc = subprocess.Popen([SOFFICE, "--headless", "--invisible", "--norestore", f"-env:UserInstallation=file://{profile}",
                             f"--accept=socket,host=localhost,port={PORT};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    for _ in range(90):
        try:
            r = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            return proc, r.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
        except Exception:
            time.sleep(1)
    raise SystemExit("cannot connect to soffice")


def load(desktop, path):
    return desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0, (pv("Hidden", True),))


def settings(doc):
    out = {}
    for p in doc.getPropertySetInfo().getProperties():
        try:
            v = doc.getPropertyValue(p.Name)
        except Exception:
            continue
        if isinstance(v, (bool, int, float, str)):
            out[p.Name] = v
        elif hasattr(v, "Year"):
            out[p.Name] = f"{v.Year}-{v.Month}-{v.Day}"
    return out


def result(cell):
    t = cell.getType().value  # EMPTY VALUE TEXT FORMULA
    if t != "FORMULA":
        return {"kind": t, "v": cell.getValue(), "s": cell.getString()}
    err = cell.getError()
    if err:
        return {"kind": "ERR", "err": err, "s": cell.getString()}
    rt = cell.FormulaResultType2  # 1 value, 2 string
    return {"kind": "NUM" if rt == 1 else "STR", "v": cell.getValue(), "s": cell.getString()}


def read_calc(doc, n):
    sh = doc.Sheets.getByName("Calc")
    return [result(sh.getCellByPosition(0, i)) for i in range(n)]


def main():
    profile = tempfile.mkdtemp(prefix="lo_cmp_")
    proc, ctx = connect(profile)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    try:
        x = load(desktop, SRC)
        x.calculateAll()
        calc = x.Sheets.getByName("Calc")
        cur = calc.createCursor(); cur.gotoEndOfUsedArea(False)
        n = cur.RangeAddress.EndRow + 1
        formulas = [calc.getCellByPosition(0, i).getFormula() for i in range(n)]
        labels = [calc.getCellByPosition(1, i).getString() for i in range(n)]
        data = x.Sheets.getByName("Data")
        dcur = data.createCursor(); dcur.gotoEndOfUsedArea(False)
        drows, dcols = dcur.RangeAddress.EndRow + 1, dcur.RangeAddress.EndColumn + 1
        drange = data.getCellRangeByPosition(0, 0, dcols - 1, drows - 1)
        ddata = drange.getFormulaArray()
        rx = read_calc(x, n)
        sx = settings(x)

        def native():
            d = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
            d.Sheets.insertNewByName("Data", 0)
            d.Sheets.insertNewByName("Calc", 1)
            for name in list(d.Sheets.ElementNames):
                if name not in ("Data", "Calc"):
                    d.Sheets.removeByName(name)
            d.Sheets.getByName("Data").getCellRangeByPosition(0, 0, dcols - 1, drows - 1).setFormulaArray(ddata)
            c = d.Sheets.getByName("Calc")
            for i, f in enumerate(formulas):
                c.getCellByPosition(0, i).setFormula(f)
            return d

        o = native(); o.calculateAll()
        ro = read_calc(o, n); so = settings(o)
        # O2: the settings of the imported document applied to the native one.
        for k, v in sx.items():
            if so.get(k) != v:
                try:
                    o.setPropertyValue(k, v)
                except Exception:
                    pass
        o.calculateAll()
        ro2 = read_calc(o, n)
        # R: save the imported document as ODS, reload, recalculate in full.
        tmp = os.path.join(profile, "x.ods")
        x.storeToURL(uno.systemPathToFileUrl(tmp), (pv("FilterName", "calc8"),))
        r = load(desktop, tmp); r.calculateAll()
        rr = read_calc(r, n); sr = settings(r)
        rows = [{"i": i + 1, "formula": labels[i], "api": formulas[i], "x": rx[i], "o": ro[i], "o2": ro2[i], "r": rr[i]} for i in range(n)]
        json.dump({"rows": rows, "settings": {"x": sx, "o": so, "r": sr}}, open(OUT, "w"), ensure_ascii=False, indent=1)
        for d in (x, o, r):
            d.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        proc.wait(timeout=60)


main()
