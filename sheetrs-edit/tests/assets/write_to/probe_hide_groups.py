"""Which formulas LibreOffice 26.8 calculates when a row of `Data` is hidden, and why (a probe, not a test): the evidence of
`sheetrs_compiler::Cache::rows_hidden` and of the groups of formulas (`sheetrs-compiler/docs/FIDELITY.md`, *What hiding rows reaches*).

Every formula of `Calc` caches the text `S`; after hiding a row of `Data`, saving and opening the file again, a formula that shows
anything else was calculated. Blocks (one per column of `Calc`, each a list of formulas written in the cells of the rows it names):

    python3 probe_hide_groups.py gen DIR                                   # the XLSX and the ODS (needs openpyxl and soffice)
    /opt/libreoffice26.8/program/python probe_hide_groups.py run DIR [N]   # LibreOffice listening on port 2002
"""
import os
import re
import subprocess
import sys
import tempfile
import zipfile

# column of Calc: [(row, formula)]
BLOCKS = {
    "B": [(1, "=SUM(Data!A1:A3)"), (2, "=SUM(Data!A2:A4)")],                                      # a group of 2, ranges
    "C": [(1, "=SUM(Data!A1:A3)"), (2, "=SUM(Data!A2:A4)"), (3, "=SUM(Data!A3:A5)")],            # a group of 3
    "D": [(1, "=SUM(Data!A1:A3)"), (3, "=SUM(Data!A2:A4)")],                                      # one blank row between
    "E": [(1, "=Data!A2"), (2, "=Data!A3")],                                                      # a group of 2, single cells
    "F": [(1, "=SUM(Data!A1:A3)"), (2, "=SUM(Data!A1:A3)")],                                      # the same text twice
    "G": [(1, "=SUM(Data!A1:A3)"), (2, "=SUM(Data!A2:A4)"), (3, "=SUM(Data!B3:B5)")],            # two of a shape, then another
    "H": [(1, "=SUM(Data!A1:A3)")],                                                               # alone
    "I": [(1, "=Data!A1+Data!A2"), (2, "=Data!A2+Data!A3")],                                      # operators, group of 2
    "J": [(1, "=COUNT(Data!A1:A3)"), (2, "=COUNT(Data!A2:A4)")],                                  # another function
    "K": [(1, "=SUM(Data!A1:B3)"), (2, "=SUM(Data!A2:B4)")],                                      # a block of two columns
    "L": [(1, "=SUM(Data!A1:A3)"), (2, "=SUM(Data!A2:A4)"), (4, "=SUM(Data!A4:A6)"), (5, "=SUM(Data!A5:A7)")],  # two groups
    "M": [(1, "=SUM(Data!A1:A3)")],                                                               # a row of three: M1, N1, O1
    "N": [(1, "=SUM(Data!B1:B3)")],
    "O": [(1, "=SUM(Data!C1:C3)")],
    "P": [(1, "=SUM(Data!A1:A3)"), (2, "=SUM(Data!A2:A4)")],                                      # a group of 2 next to a group of 2 (Q)
    "Q": [(1, "=SUM(Data!B1:B3)"), (2, "=SUM(Data!B2:B4)")],
    "R": [(1, "=SUBTOTAL(109,Data!A1:A3)"), (2, "=SUBTOTAL(109,Data!A2:A4)")],                    # a group of SUBTOTAL
    "S": [(1, "=SUBTOTAL(109,Data!A2:A3)")],                                                      # alone
    "T": [(1, "=SUM(Data!A1:A3)"), (2, "=SUM(Data!A2:A4)"), (3, "=SUM(Data!A2:A4)")],            # the third repeats the second
    "U": [(1, "=SUM(Data!A:A)"), (2, "=SUM(Data!A:A)")],                                          # whole columns (same shape in both)
    "V": [(1, "=SUM(Data!A1:A3)+Data!B1"), (2, "=SUM(Data!A2:A4)+Data!B2")],                      # a range and a cell
    "W": [(1, "=SUM(Data!$A$1:$A$3)"), (2, "=SUM(Data!$A$1:$A$3)")],                              # absolute: the same shape
    "X": [(1, "=SUM(Data!1:3)"), (2, "=SUM(Data!2:4)")],                                          # whole rows
    "Y": [(1, "=SUM(Data!A2:C2)"), (2, "=SUM(Data!A3:C3)")],                                      # a row of cells
    "Z": [(1, "=SUM(Data!A1:A3)"), (2, "=SUM(Data!A2:A4)"), (3, "=Data!A1+1")],                   # a group of 2, then a single cell
    "AA": [(1, "=SUBTOTAL(109,Data!A1:A3)"), (2, "=SUM(Data!A2:A4)")],                            # two shapes
    "AB": [(1, "=SUM(Data!A1:A3)"), (2, "=SUM(Data!A2:A4)"), (3, "=SUM(Data!A3:A5)"), (4, "=SUM(Data!A4:A6)"),
           (5, "=SUM(Data!A5:A7)")],                                                              # a group of 5
    "AC": [(1, "=SUM(Data!A2:A2)"), (2, "=SUM(Data!A3:A3)")],                                      # ranges of one cell
}
SHOWS = [([2, 3], [3]), ([3, 4], [3]), ([], [3]), ([2], [2])]  # (hidden first, shown after)
HIDE_COLUMNS = [["A"], ["B"], ["C"]]
HIDES = [[1], [2], [3], [4], [5], [3, 4]]


def column_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def gen(folder):
    import openpyxl

    os.makedirs(folder, exist_ok=True)
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Data"
    for r in range(1, 9):
        for c in range(1, 4):
            data.cell(row=r, column=c, value=r * 10 + c)
    calc = wb.create_sheet("Calc")
    for column, cells in BLOCKS.items():
        for row, formula in cells:
            calc[f"{column}{row}"] = formula
    path = os.path.join(folder, "groups.xlsx")
    wb.save(path)
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            xml = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet2.xml":
                text, _ = re.subn(r'<c r="([A-Z]+\d+)"([^>]*)><f>(.*?)</f>(?:<v\s*/>|<v></v>)</c>',
                                  lambda m: '<c r="%s"%s t="str"><f>%s</f><v>S</v></c>' % (m.group(1), m.group(2), m.group(3)), xml.decode())
                xml = text.encode()
            if item.filename == "xl/workbook.xml":
                xml = xml.decode().replace(' fullCalcOnLoad="1"', "").encode()
            zout.writestr(item, xml)
    os.replace(tmp, path)
    with tempfile.TemporaryDirectory() as profile:
        subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "ods", "--outdir", folder, path],
                       check=True, capture_output=True)


def run(folder):
    import uno
    from com.sun.star.beans import PropertyValue

    def pv(name, value):
        p = PropertyValue()
        p.Name, p.Value = name, value
        return p

    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    filters = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8"}

    def load(p):
        return desktop.loadComponentFromURL("file://" + os.path.abspath(p), "_blank", 0, (pv("Hidden", True),))

    for extension in ("xlsx", "ods"):
        print(extension)
        for kind, hide in [("rows", h) for h in HIDES] + [("columns", h) for h in HIDE_COLUMNS] + [("hide then show", h) for h in SHOWS]:
            doc = load(os.path.join(folder, f"groups.{extension}"))
            if kind == "hide then show":
                for h in hide[0]:
                    doc.Sheets.getByName("Data").Rows.getByIndex(h - 1).IsVisible = False
                for h in hide[1]:
                    doc.Sheets.getByName("Data").Rows.getByIndex(h - 1).IsVisible = True
                hide = []
            for h in hide:
                if kind == "rows":
                    doc.Sheets.getByName("Data").Rows.getByIndex(h - 1).IsVisible = False
                else:
                    doc.Sheets.getByName("Data").Columns.getByIndex(column_index(h) - 1).IsVisible = False
            with tempfile.TemporaryDirectory() as t:
                out = os.path.join(t, f"x.{extension}")
                doc.storeToURL("file://" + out, (pv("FilterName", filters[extension]),))
                doc.close(True)
                again = load(out)
                sheet = again.Sheets.getByName("Calc")
                shown = []
                for column, cells in BLOCKS.items():
                    c = column_index(column) - 1
                    for row, formula in cells:
                        if sheet.getCellByPosition(c, row - 1).getString() != "S":
                            shown.append(f"{column}{row}")
                again.close(True)
            print(f"  hide Data {kind} {hide}: calculated {' '.join(shown) or '-'}")


if __name__ == "__main__":
    {"gen": gen, "run": run}[sys.argv[1]](sys.argv[2])
