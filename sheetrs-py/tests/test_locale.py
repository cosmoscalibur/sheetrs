"""The locale of a workbook (`Book(..., locale=...)`): the regional settings that give a text to a number
(`TEXT`, the conversion of a number in a `&`) and read a number from a text, in `calculate`, in the
recalculation that follows an edit and in the queries. The cases are those of issues #84 and #106; the
oracle is `semantics_es-CO.{xlsx,ods}`, made by a LibreOffice running with the `es-CO` regional setting."""

import pytest

import sheetrs
from conftest import ROOT, _make_xlsx

# (formula, as LibreOffice shows it with es-CO, with en-US)
CASES = [
    ('TEXT(350*49799,"[$$-en-US]#.###.##0")', "$17.429.650", "$17429650.000.000"),
    ('"Componente inflacionario "&(0.5543*100)&"%"', "Componente inflacionario 55,43%", "Componente inflacionario 55.43%"),
    ('TEXT(0,"[$$-en-US]#,##0;([$$-en-US]#,##0)")', "$,000", "$0"),
]
ORACLES = ROOT / "sheetrs-compiler/tests/assets"


def _workbook(tmp_path):
    cells = {f"A{i}": ("=" + formula, "?") for i, (formula, _, _) in enumerate(CASES, 1)}
    return _make_xlsx(tmp_path / "cases.xlsx", {"S": cells})


def _values(book):
    block = book.calculate(ranges="S!A1:A3").value_ranges(calculate_all=True)["S!A1:A3"]
    return [row[0] for row in block.values]


def an_xlsx_takes_the_locale_of_the_book(tmp_path):
    path = _workbook(tmp_path)
    assert _values(sheetrs.Book(path, locale="es-CO")) == [es for _, es, _ in CASES]
    assert _values(sheetrs.Book(path, locale="en-US")) == [en for _, _, en in CASES]


def the_locale_is_accepted_with_bytes_and_streams(tmp_path):
    import io

    data = _workbook(tmp_path).read_bytes()
    for source in (data, bytearray(data), io.BytesIO(data)):
        assert _values(sheetrs.Book(source, locale="es_CO")) == [es for _, es, _ in CASES]


def an_ods_takes_the_locale_of_the_book():
    labels = [row[0] for row in sheetrs.Book(ORACLES / "semantics_es-CO.ods").range("Calc!B1:B1500")["Calc!B1:B1500"].values]
    for formula, es, en in CASES:
        asked = f"Calc!A{labels.index(formula) + 1}"
        for locale, want in (("es-CO", es), ("en-US", en)):
            book = sheetrs.Book(ORACLES / "semantics_es-CO.ods", locale=locale)
            assert book.calculate(ranges=asked).value_ranges(calculate_all=True).to_dict() == {asked: [[want]]}


def _run_python(code, env):
    """A child process whose environment has the locale variables of `env` and no others."""
    import os
    import subprocess
    import sys

    clean = {k: v for k, v in os.environ.items() if k not in ("LC_ALL", "LC_CTYPE", "LANG")}
    # Python itself sets `LC_CTYPE=C.UTF-8` in its environment when the locale it is given is not
    # installed on the machine (PEP 538), which would hide `LANG`: the tests ask it not to.
    clean["PYTHONCOERCECLOCALE"] = "0"
    out = subprocess.run([sys.executable, "-c", code], env={**clean, **env}, capture_output=True, text=True, check=True)
    return out.stdout.strip()


SYSTEM = """
import sys, sheetrs
path, locale = sys.argv[1], sys.argv[2] or None
book = sheetrs.Book(path, locale=locale).calculate(ranges="S!A1")
print(book.value_ranges(calculate_all=True).to_dict()["S!A1"][0][0])
"""


def the_given_locale_wins_over_the_system_one(tmp_path):
    path = _workbook(tmp_path)
    code = SYSTEM.replace("sys.argv[1], sys.argv[2]", repr(str(path)) + ", ''")
    given = SYSTEM.replace("sys.argv[1], sys.argv[2]", repr(str(path)) + ", 'en-US'")
    es = CASES[0][1]
    en = CASES[0][2]
    # Nothing given: the variables of the system, as LibreOffice reads them, for an XLSX...
    assert _run_python(code, {"LANG": "es_CO.UTF-8"}) == es
    assert _run_python(code, {}) == en
    # ...and what is given is not changed by them.
    assert _run_python(given, {"LANG": "es_CO.UTF-8"}) == en
    assert _run_python(given, {"LC_ALL": "es_CO.UTF-8"}) == en


def the_locale_is_normalized_and_none_by_default(tmp_path):
    path = _workbook(tmp_path)
    assert sheetrs.Book(path).locale is None
    assert sheetrs.Book(path, locale="es_co").locale == "es-CO"
    assert sheetrs.Book(path, locale="de").locale == "de"


def a_bad_locale_tag_is_a_value_error(tmp_path):
    with pytest.raises(ValueError, match="not a locale tag"):
        sheetrs.Book(_workbook(tmp_path), locale="not a locale")


def an_unknown_locale_is_an_error_that_names_it(tmp_path):
    book = sheetrs.Book(_workbook(tmp_path), locale="fr-FR").calculate(ranges="S!A1")
    with pytest.raises(sheetrs.Error, match="fr-FR"):
        book.value_ranges(calculate_all=True)


def save_keeps_the_locale(tmp_path):
    path = _make_xlsx(tmp_path / "in.xlsx", {"S": {"A1": 1, "B1": ("=TEXT(1234.5,\"#,##0.00\")", "?"), "C1": 2}})
    saved = sheetrs.Book(path, locale="es-CO").delete_columns("S!C").save(tmp_path / "out.xlsx", calculate_all=True)
    assert saved.locale == "es-CO"


def evaluate_takes_a_locale():
    assert sheetrs.evaluate('=TEXT(1234.5,"#,##0.00")', locale="es-CO") == "1.234,50"
    assert sheetrs.evaluate('=TEXT(1234.5,"#,##0.00")') == "1,234.50"
