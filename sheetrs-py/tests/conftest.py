"""Fixtures of the tests: the workspace's workbooks (XLSX and ODS, each judged on its own) and a
small writer of XLSX workbooks for the cases that need a cell no fixture has."""

import shutil
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import pytest

ROOT = Path(__file__).resolve().parents[2]
LINT_BOOKS = [ROOT / "tests/lint_book.xlsx", ROOT / "tests/lint_book.ods"]
# Sheets T, U, V with numbers and formulas, F with formulas over T; the names Gone, NameOne, NameTwo
# and NameThree.
STRUCTURE = ROOT / "sheetrs-edit/tests/assets/structure"
IO_ASSETS = ROOT / "sheetrs-io/tests/assets"


@pytest.fixture(params=["xlsx", "ods"])
def structure(request, tmp_path):
    """A copy of the structure workbook (so that a test can save over it), in each format."""
    source = STRUCTURE / f"book.{request.param}"
    copy = tmp_path / source.name
    shutil.copy(source, copy)
    return copy


@pytest.fixture(params=LINT_BOOKS, ids=lambda p: p.suffix[1:])
def lint_book(request):
    return request.param


def io_asset(name):
    return IO_ASSETS / name


def make_xlsx(path, sheets, names=None):
    """Writes a workbook. `sheets` is `{name: {"A1": value}}`, where a value is a number, a `str`
    (text), a `bool`, a `("=formula", cached)` pair or `("#N/A",)` for an error."""
    path = Path(path)
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, len(sheets) + 1)
        )
        + "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    defined = "".join(
        f'<definedName name="{escape(n)}">{escape(ref)}</definedName>'
        for n, ref in (names or {}).items()
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
        + "".join(
            f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
            for i, name in enumerate(sheets, 1)
        )
        + "</sheets>"
        + (f"<definedNames>{defined}</definedNames>" if defined else "")
        + "</workbook>"
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, len(sheets) + 1)
        )
        + "</Relationships>"
    )

    def cell_xml(address, value):
        if isinstance(value, tuple) and len(value) == 1:
            return f'<c r="{address}" t="e"><v>{escape(value[0])}</v></c>'
        if isinstance(value, tuple):
            formula, cached = value
            body = f"<f>{escape(formula.lstrip('='))}</f>"
            if isinstance(cached, str):
                return f'<c r="{address}" t="str">{body}<v>{escape(cached)}</v></c>'
            return f'<c r="{address}">{body}<v>{cached}</v></c>'
        if isinstance(value, bool):
            return f'<c r="{address}" t="b"><v>{int(value)}</v></c>'
        if isinstance(value, str):
            return f'<c r="{address}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
        return f'<c r="{address}"><v>{value}</v></c>'

    def column_of(address):
        letters = "".join(ch for ch in address if ch.isalpha())
        n = 0
        for ch in letters:
            n = n * 26 + ord(ch) - 64
        return n

    def sheet_xml(cells):
        rows = {}
        for address, value in cells.items():
            row = int("".join(ch for ch in address if ch.isdigit()))
            rows.setdefault(row, []).append((column_of(address), address, value))
        body = "".join(
            f'<row r="{row}">'
            + "".join(cell_xml(a, v) for _, a, v in sorted(entries))
            + "</row>"
            for row, entries in sorted(rows.items())
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"<sheetData>{body}</sheetData></worksheet>"
        )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        for i, cells in enumerate(sheets.values(), 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(cells))
    return path
