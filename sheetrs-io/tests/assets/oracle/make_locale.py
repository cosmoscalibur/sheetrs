"""Makes `sheetrs-io/tests/assets/locale_es-CO.{xlsx,ods}`: numbers, dates, times, booleans and text with
formats whose shown text depends on the locale, saved by a LibreOffice running with the `es-CO` regional
setting (the ODS says it in its default cell style, the XLSX does not). `make_oracle.py` then records what
that LibreOffice shows for each cell (see the end of its docstring), and the test of `sheetrs-query`
(`the_display_of_a_book_follows_its_locale`) compares `display` with it: the XLSX with the locale given to
its `Book`, the ODS with the one the file says.

The built-in format ids 14 to 22 (the short date of the locale, issue #84) are not used: the codes
here are all written out.

Run it (needs `openpyxl` and `soffice`, the reference build LibreOffice 26.8):

    python3 make_locale.py
"""
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(HERE, "..", "..", "..", "..", "scripts"))
import fixture_determinism  # noqa: E402  (the same bytes on every run, see its docstring)

SOFFICE = os.environ.get("SOFFICE", "soffice")
LOCALE = "es-CO"

VALUES = [1234.5, -1234.5, 0.5, 1234567.891, 0.000123, 45367.75, 45658.0]
FORMATS = [
    "General", "0.00", "#,##0.00", "#,##0", "0.0%", "0.00E+00", '"$"#,##0.00', "#,##0;(#,##0)", "0.0",
    "dd/mm/yyyy", "dddd d mmmm yyyy", "d-mmm-yyyy", "mmmm", "hh:mm:ss", "yyyy-mm-dd hh:mm", "[hh]:mm",
]


def build(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Numbers"
    for row, value in enumerate(VALUES, 1):
        for col, code in enumerate(FORMATS, 1):
            cell = ws.cell(row, col, value)
            cell.number_format = code
    # Booleans, text and the text a formula makes of a number.
    ws["A10"], ws["B10"], ws["C10"] = True, False, "text"
    ws["A11"], ws["B11"] = "=1234.5&\"\"", "=0.5543*100&\"%\""
    ws["C11"], ws["D11"] = "=TEXT(1234.5,\"#,##0.00\")", "=TEXT(0.5,\"0.0%\")"
    ws["E11"], ws["F11"] = "=1/4", "=-1234.5"
    wb.save(path)


def profile(tmp):
    prof = os.path.join(tmp, "profile")
    os.makedirs(os.path.join(prof, "user"), exist_ok=True)
    item = (
        '<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" '
        f'oor:op="fuse"><value>{LOCALE}</value></prop></item>'
    )
    with open(os.path.join(prof, "user", "registrymodifications.xcu"), "w") as f:
        f.write(
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
            'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f"{item}</oor:items>"
        )
    return f"-env:UserInstallation=file://{prof}"


def main():
    name = f"locale_{LOCALE}"
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, f"{name}.xlsx")
        build(src)
        outdir = os.path.join(tmp, "out")
        for fmt in ("xlsx", "ods"):
            subprocess.run(
                [SOFFICE, profile(tmp), "--headless", "--convert-to", fmt, "--outdir", outdir, src],
                check=True, capture_output=True,
            )
            dest = os.path.join(OUT, f"{name}.{fmt}")
            shutil.move(os.path.join(outdir, f"{name}.{fmt}"), dest)
            fixture_determinism.normalize(dest)
            print("wrote", dest)


if __name__ == "__main__":
    main()
