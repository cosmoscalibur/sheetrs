#!/usr/bin/env python3
"""The workbooks of `tests/from_cache.rs`: a workbook that depends on what the engine cannot calculate, in both formats.

    python3 make_fixtures.py     # needs openpyxl and `soffice` (LibreOffice) on the PATH
    /opt/libreoffice26.8/program/python make_external.py   # the external workbook, which needs a LibreOffice server

Each workbook has, in `Calc`, the cells

    A1  the formula that depends on what the engine cannot calculate, with a cached result
    B1  =A1+1 (cached 1000: it reads A1)      C1  =D1*2 (cached 999), with D1 = 5 (in `Data` for these two)
    E1  =C1+1 (cached 1000)

so that the result of a cell that was calculated (C1 is 10, E1 is 11, B1 reads A1) is told from the cached one
(999, 1000, and A1 itself, which the engine cannot calculate).

  macro      `A1 =MYFUNC(D1)` (cached 77) in a workbook that carries macros: `xl/vbaProject.bin` of an XLSX, a
             `Basic/` library of an ODS. No spreadsheet program has a function of that name: it may be the
             one a macro defines
  extension  `A1 =_xll.MYADDIN(D1)` (an add-in of an XLSX, which LibreOffice
             writes as `_xll.myaddin(...)` in an ODS),
             cached 77, and `F1 =WEBSERVICE("x")` (a function of LibreOffice that the engine does not
             implement), cached `abc`
  external   see `make_external.py`: `A1 =[1]Sheet1!A1` of an XLSX (an external workbook), cached 42

The cached results are not what a program would give, and every workbook is made byte-for-byte reproducible by
`fixture_determinism.normalize`.
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
USE_CACHE = os.path.join(HERE, "..", "cached")
sys.path.insert(0, USE_CACHE)
sys.path.insert(0, os.path.join(HERE, "..", "..", "..", "..", "scripts"))
import make_fixtures as base  # noqa: E402  (the helpers that build a workbook with the cached results)
import fixture_determinism  # noqa: E402

base.HERE = HERE  # the workbooks are written here, not next to the helpers

MACRO = [("A1", "=MYFUNC(Data!D1)", 77), ("B1", "=A1+1", 1000), ("C1", "=Data!D1*2", 999), ("E1", "=C1+1", 1000)]
EXTENSION = [("A1", "=_xll.MYADDIN(Data!D1)", 77), ("B1", "=A1+1", 1000), ("C1", "=Data!D1*2", 999), ("E1", "=C1+1", 1000),
             ("F1", '=_xlfn.WEBSERVICE("x")', "abc")]


def populate(data):
    data["D1"] = 5


def add_to_zip(path, entries, edit=None):
    """`path` with the entries added, and `edit(name, bytes)` applied to the others."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if edit:
                data = edit(item.filename, data)
            zout.writestr(item, data)
        for name, data in entries:
            zout.writestr(name, data)
    os.replace(tmp, path)


def vba(name, data):
    if name == "[Content_Types].xml":
        text = data.decode()
        text = text.replace("<Default ", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)
        return text.encode()
    return data


def main():
    for name, formulas in (("macro", MACRO), ("extension", EXTENSION)):
        path = os.path.join(HERE, f"{name}.xlsx")
        base.build(path, formulas, populate, ())
        base.to_ods(path)
        if name == "macro":
            add_to_zip(path, [("xl/vbaProject.bin", b"\0\0vba project")], vba)
            add_to_zip(os.path.join(HERE, "macro.ods"), [
                ("Basic/script-lc.xml", b'<?xml version="1.0" encoding="UTF-8"?><library:libraries xmlns:library="http://openoffice.org/2000/library"/>'),
                ("Basic/Standard/Module1.xml", b'<?xml version="1.0" encoding="UTF-8"?><script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1"/>'),
            ])
        for extension in ("xlsx", "ods"):
            fixture_determinism.normalize(os.path.join(HERE, f"{name}.{extension}"))


if __name__ == "__main__":
    main()
