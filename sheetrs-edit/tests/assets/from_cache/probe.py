"""What LibreOffice 26.8 gives for the cells that depend on what the engine cannot calculate: a macro of the document, an
add-in that is not installed, another workbook, `WEBSERVICE`, `FILTERXML` and `DDE`. The evidence of the table in
`sheetrs-compiler/docs/FIDELITY.md` (*What depends on what the engine cannot calculate*).

    /opt/libreoffice26.8/program/python probe.py build     # LibreOffice listening on port 2002 (`scripts/libreoffice-server.sh`)

`build` writes, in a temporary folder (`PROBE_OUT` names another), one workbook per format with each kind in `Calc!A:A` (`B` reads `A`; `Data!D1` is
read by the macro and the add-ins), every cached result a stale number (`raw.*`) or text (`rawS.ods`, which an ODS keeps where it calculates its numbers on
load); `rawN*` has no `Basic/` library. Then each file is loaded (macro security as installed, and `ALWAYS_EXECUTE_NO_WARN`) and the cells are read after
the default load, `calculate()`, `calculateAll()`, an edit of `D1`, the deletion of an unrelated row, of the row of `D1`, and each saved and opened again.
The table is printed and written to `result.json`."""
import os, re, sys, tempfile, time, zipfile, json
import uno
from com.sun.star.beans import PropertyValue

OUT = os.environ.get("PROBE_OUT") or tempfile.mkdtemp(prefix="sheetrs-probe-")
os.makedirs(OUT, exist_ok=True)
EXT = OUT + "/other.ods"


def pv(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v; return p


local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(30):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
FILT = {"ods": "calc8", "xlsx": "Calc MS Excel 2007 XML"}


def load(path, macro=None):
    props = [pv("Hidden", True)]
    if macro is not None:
        props.append(pv("MacroExecutionMode", macro))
    return desktop.loadComponentFromURL("file://" + path, "_blank", 0, tuple(props))


CASES = [
    ("macro", "=MYFUNC($Data.D1)"),
    ("addin_xll", "=_xll.MYADDIN($Data.D1)"),
    ("addin_sun", "=COM.SUN.STAR.SHEET.ADDIN.MYADDIN.GETX($Data.D1)"),
    ("external", f"='file://{EXT}'#$Sheet1.A1"),
    ("webservice", '=WEBSERVICE("http://127.0.0.1:9/x")'),
    ("filterxml", '=FILTERXML("<a><b>7</b></a>";"//b")'),
    ("dde", '=DDE("soffice";"x";"y")'),
    ("external_mix", f"='file://{EXT}'#$Sheet1.A1+$Data.D1"),
    ("control", "=$Data.D1*2"),
]


def build():
    d = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
    d.Sheets.getByIndex(0).Name = "Sheet1"; d.Sheets.getByIndex(0).getCellByPosition(0, 0).setValue(42)
    d.storeToURL("file://" + EXT, (pv("FilterName", "calc8"),)); d.close(True)
    for fmt, basic in (("ods", True), ("odsN", False), ("xlsx", False)):
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
        sh = doc.Sheets.getByIndex(0); sh.Name = "Data"
        sh.getCellByPosition(3, 0).setValue(5)
        for r in range(4):
            sh.getCellByPosition(0, r + 1).setValue(r + 1)
        doc.Sheets.insertNewByName("Calc", 1)
        c = doc.Sheets.getByName("Calc")
        if basic:
            lib = doc.BasicLibraries
            if not lib.hasByName("Standard"):
                lib.createLibrary("Standard")
            lib.getByName("Standard").insertByName("Module1", "Function MYFUNC(x)\n  MYFUNC = x * 3\nEnd Function\n")
        for i, (label, f) in enumerate(CASES):
            c.getCellByPosition(0, i).setFormula(f)
            c.getCellByPosition(1, i).setFormula(f"=A{i+1}+1")
        doc.calculateAll()
        base = {}
        for i, (label, f) in enumerate(CASES):
            base[label] = (c.getCellByPosition(0, i).getString(), c.getCellByPosition(0, i).getError())
        print(fmt, "computed at build:", base)
        real = "ods" if fmt == "odsN" else fmt
        name = {"ods": "raw", "odsN": "rawN", "xlsx": "raw"}[fmt]
        p = f"{OUT}/{name}.{real}"
        doc.storeToURL("file://" + p, (pv("FilterName", FILT[real]),)); doc.close(True)
        if real == "ods":
            import shutil
            shutil.copy(p, f"{OUT}/{name}S.ods")
            patch(f"{OUT}/{name}S.ods", "ods", True)
        patch(p, real)


def patch(p, fmt, text=False):
    tmp = p + ".tmp"
    with zipfile.ZipFile(p) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if fmt == "ods" and it.filename == "content.xml":
                s = data.decode()
                i = s.index('table:name="Calc"')
                head, tail = s[:i], s[i:]
                n = [0]

                def rep(m):
                    k = n[0]; n[0] += 1
                    row, col = divmod(k, 2)
                    stale = 777 + row if col == 0 else 1000 + row
                    if text:
                        return (f'<table:table-cell table:formula="{m.group(2)}" office:value-type="string" office:string-value="S{stale}" '
                                f'calcext:value-type="string"><text:p>S{stale}</text:p></table:table-cell>')
                    return (f'<table:table-cell table:formula="{m.group(2)}" office:value-type="float" office:value="{stale}" '
                            f'calcext:value-type="float"><text:p>{stale}</text:p></table:table-cell>')
                tail = re.sub(r'<table:table-cell\b([^>]*?table:formula="([^"]*)"[^>]*?)(/>|>.*?</table:table-cell>)', rep, tail, flags=re.S)
                print("ods patched cells", n[0])
                data = (head + tail).encode()
            if fmt == "xlsx" and it.filename == "xl/worksheets/sheet2.xml":
                s = data.decode()

                def rep(m):
                    ref = m.group(1); col = ref[0]; row = int(ref[1:]) - 1
                    stale = 777 + row if col == "A" else 1000 + row
                    return f'<c r="{ref}"><f>{m.group(3)}</f><v>{stale}</v></c>'
                s, n = re.subn(r'<c r="([A-Z]+[0-9]+)"([^>]*)><f[^>]*>(.*?)</f><v>.*?</v></c>', rep, s, flags=re.S)
                print("xlsx patched cells", n)
                data = s.encode()
            if fmt == "xlsx" and it.filename == "xl/workbook.xml":
                data = data.decode().replace(' fullCalcOnLoad="1"', "").encode()
            zout.writestr(it, data)
    os.replace(tmp, p)


def read(doc):
    c = doc.Sheets.getByName("Calc")
    out = {}
    for i, (label, f) in enumerate(CASES):
        a, b = c.getCellByPosition(0, i), c.getCellByPosition(1, i)

        def show(x):
            return f"E{x.getError()}:{x.getString()}" if x.getError() else x.getString()
        out[label] = f"{show(a)} | {show(b)}"
    return out


def store(d, q, fmt):
    d.storeToURL("file://" + q, (pv("FilterName", FILT[fmt]),))


def scenario(fmt, macro, raw=None):
    raw = raw or "raw"
    p = f"{OUT}/{raw}.{fmt}"
    res = {}
    if True:
        d = load(p, macro); c = d.Sheets.getByName("Calc")
        res["formulas as LO reads them"] = {CASES[i][0]: c.getCellByPosition(0, i).getFormula() for i in range(len(CASES))}; d.close(True)
    d = load(p, macro); res["L0 default load"] = read(d)
    d.calculate(); res["L1 calculate() soft"] = read(d)
    d.calculateAll(); res["L2 calculateAll()"] = read(d); d.close(True)
    d = load(p, macro); d.Sheets.getByName("Data").getCellByPosition(3, 0).setValue(6); res["L3 D1 edited (precedent)"] = read(d)
    q = f"{OUT}/saved3.{fmt}"; store(d, q, fmt); d.close(True)
    d = load(q, macro); res["L3b ... saved and reopened"] = read(d); d.close(True)
    d = load(p, macro); d.Sheets.getByName("Data").Rows.removeByIndex(2, 1); res["L4 row 3 of Data deleted (unrelated)"] = read(d)
    d.close(True)
    d = load(p, macro); q = f"{OUT}/saved.{fmt}"; store(d, q, fmt); d.close(True)
    d = load(q, macro); res["L5 untouched, saved and reopened"] = read(d); d.close(True)
    d = load(p, macro); d.calculateAll(); q = f"{OUT}/saved2.{fmt}"; store(d, q, fmt); d.close(True)
    d = load(q, macro); res["L6 calculateAll, saved and reopened"] = read(d); d.close(True)
    d = load(p, macro); d.Sheets.getByName("Data").Rows.removeByIndex(0, 1); res["L7 row 1 of Data deleted (D1 goes)"] = read(d); d.close(True)
    return res


if __name__ == "__main__":
    if "build" in sys.argv:
        build()
    out = {}
    for fmt, raw in (("ods", "rawS"), ("ods", "rawNS"), ("xlsx", "raw")):
        for macro, name in ((None, "macro security default"), (4, "ALWAYS_EXECUTE_NO_WARN")):
            out[f"{raw}.{fmt} / {name}"] = scenario(fmt, macro, raw)
    json.dump(out, open(OUT + "/result.json", "w"), indent=1)
    for k, v in out.items():
        print("=====", k)
        for step, r in v.items():
            print(" --", step)
            for label, val in r.items():
                print(f"     {label:11} {val}")
