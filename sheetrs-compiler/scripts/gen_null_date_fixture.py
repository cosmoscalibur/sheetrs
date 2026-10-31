#!/usr/bin/env python3
"""Generate tests/assets/null_date_{1900,1904,1950}.ods: documents created in LibreOffice with
another null date (the date of serial 0; LibreOffice stores any date per document).

Each has a `Calc` sheet with a formula in column A and its text in column B, recalculated and
saved by LibreOffice, which tests/semantics/oracle.rs requires the engine to reproduce.

Run with LibreOffice's Python:  SOFFICE=soffice26 .../program/python scripts/gen_null_date_fixture.py
"""
import os
import subprocess
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.util import Date

SOFFICE = os.environ.get("SOFFICE", "soffice")
ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tests", "assets")
PORT = 2015
FORMULAS = [
    "DATE(2024;1;1)", "DATE(1900;1;1)", "DATE(1899;12;31)", "DATE(1900;3;1)", "DATE(1850;1;1)", "DATE(2024;2;29)+1",
    "YEAR(0)", "MONTH(0)", "DAY(0)", "YEAR(-1)", "YEAR(45000)", "WEEKDAY(0)", "WEEKDAY(DATE(2024;1;1))",
    "TEXT(0;\"yyyy-mm-dd\")", "TEXT(45000;\"yyyy-mm-dd\")", "TEXT(-1;\"yyyy-mm-dd\")",
    "DATEVALUE(\"2024-01-01\")", "DATEVALUE(\"1/1/1900\")", "VALUE(\"2024-01-01\")", "EDATE(DATE(2024;1;31);1)",
    "EOMONTH(DATE(2024;1;31);0)", "DAYS(DATE(2024;3;1);DATE(2024;2;1))", "DATEDIF(DATE(2020;1;1);DATE(2024;1;1);\"Y\")",
    "WEEKNUM(DATE(2024;1;1))", "NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;31))",
]
SYSTEMS = {"1900": (1900, 1, 1), "1904": (1904, 1, 1), "1950": (1950, 6, 15)}


def pv(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v; return p


def main():
    profile = tempfile.mkdtemp(prefix="lo_nd_")
    proc = subprocess.Popen([SOFFICE, "--headless", "--invisible", "--norestore", f"-env:UserInstallation=file://{profile}",
                             f"--accept=socket,host=localhost,port={PORT};urp;"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    ctx = None
    for _ in range(90):
        try:
            r = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            ctx = r.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext"); break
        except Exception:
            time.sleep(1)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    for name, (y, m, d) in SYSTEMS.items():
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
        nd = Date(); nd.Day, nd.Month, nd.Year = d, m, y
        doc.NullDate = nd
        doc.Sheets.insertNewByName("Calc", 1)
        doc.Sheets.removeByName(doc.Sheets.getByIndex(0).Name)
        sh = doc.Sheets.getByName("Calc")
        for i, f in enumerate(FORMULAS):
            sh.getCellByPosition(0, i).setFormula("=" + f)
            sh.getCellByPosition(1, i).setString(f.replace(";", ","))
        doc.calculateAll()
        out = os.path.join(ASSETS, f"null_date_{name}.ods")
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(out)), (pv("FilterName", "calc8"),))
        doc.close(True)
    try:
        desktop.terminate()
    except Exception:
        pass
    proc.wait(timeout=60)


main()
