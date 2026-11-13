#!/usr/bin/env python3
"""Ask LibreOffice what a list of formulas gives, over the same `Data` sheet as the oracle.

    scripts/lo_probe.py FORMULAS.txt [OUT.json]      # one formula per line, XLSX syntax, no `=`

Each formula goes to a row of the `Calc` sheet of a workbook built by `gen_semantics_fixture.build`
(so `Data!A1:A8`, the database block and the rest are the ones the oracle uses); LibreOffice 26.8
recalculates it and saves an ODS; the result of each row is read from the ODS: a number, a text, a
boolean or an error, as the text LibreOffice shows (`Err:502`, `#VALUE!`, `#N/A`…). With `OUT.json`
the answers are written as `[[formula, kind, result], …]`, otherwise printed one per line. With
`KEEP=probe.xlsx` the workbook is kept, to run the engine over the very same rows
(`sheetc calculate -i probe.xlsx`).

Needs `openpyxl` and `soffice` (`SOFFICE`, the reference build: LibreOffice 26.8). It is the tool the
per-function tables of `sheetrs-engine` (`arity`, see `probe_arity.py`) and the error codes of
FIDELITY.md are probed with.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gen_semantics_fixture as gen  # noqa: E402

NS = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
}
SOFFICE = os.environ.get("SOFFICE", "soffice")


def _q(prefix, name):
    return "{" + NS[prefix] + "}" + name


def probe(formulas, keep=None, cells=None):
    """`[(kind, result)]` for each formula: kind is `float`, `string`, `boolean` or `error`. With
    `keep`, the workbook that was built is also saved there (to run the engine over the same rows).
    `cells` adds values to the sheets (`{"Data!BN1": 5, …}`, text and numbers) before LibreOffice runs."""
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "probe.xlsx")
        gen.build(src, formulas)
        if cells:
            wb = openpyxl.load_workbook(src)
            for ref, value in cells.items():
                sheet, address = ref.split("!")
                wb[sheet][address] = value
            wb.save(src)
        if keep:
            shutil.copy(src, keep)
        out = os.path.join(tmp, "out")
        subprocess.run(
            [SOFFICE, f"-env:UserInstallation=file://{tmp}/profile", "--headless", "--convert-to", "ods", "--outdir", out, src],
            check=True, capture_output=True,
        )
        with zipfile.ZipFile(os.path.join(out, "probe.ods")) as z:
            root = ET.fromstring(z.read("content.xml"))
    for sheet in root.iter(_q("table", "table")):
        if sheet.get(_q("table", "name")) != "Calc":
            continue
        results = []
        for row in sheet.iter(_q("table", "table-row")):
            repeat = int(row.get(_q("table", "number-rows-repeated"), "1"))
            cell = next((c for c in row if c.tag in (_q("table", "table-cell"), _q("table", "covered-table-cell"))), None)
            if cell is None:
                results += [("empty", "")] * repeat
                continue
            kind = cell.get(_q("calcext", "value-type")) or cell.get(_q("office", "value-type")) or "empty"
            text = "\n".join("".join(p.itertext()) for p in cell.findall(_q("text", "p")))
            if kind == "float" or kind == "percentage" or kind == "currency":
                text = cell.get(_q("office", "value")) or text
            elif kind == "boolean":
                text = cell.get(_q("office", "boolean-value")) or text
            results += [(kind, text)] * repeat
        return results[: len(formulas)]
    raise SystemExit("the probe workbook has no Calc sheet")


def probe_all(formulas, batch=1000):
    """`probe` in batches; a batch that makes LibreOffice fail is split until the formula that does
    it is alone, and that one is answered `("crash", "")`."""
    out = []
    for i in range(0, len(formulas), batch):
        chunk = formulas[i : i + batch]
        try:
            out += probe(chunk)
        except subprocess.CalledProcessError:
            if len(chunk) == 1:
                out += [("crash", "")]
            else:
                half = len(chunk) // 2
                out += probe_all(chunk[:half], half) + probe_all(chunk[half:], len(chunk) - half)
    return out


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    formulas = [l.rstrip("\n") for l in open(sys.argv[1], encoding="utf-8") if l.strip()]
    answers = probe(formulas, os.environ.get("KEEP")) if os.environ.get("KEEP") else probe_all(formulas)
    if len(sys.argv) > 2:
        json.dump([[f, k, r] for f, (k, r) in zip(formulas, answers)], open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    else:
        for f, (k, r) in zip(formulas, answers):
            print(f"{f}\t{k}\t{r}")


if __name__ == "__main__":
    main()
