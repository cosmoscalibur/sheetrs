"""The locale of a workbook (`Book(..., locale=...)`): the regional settings that give a text to a number
(`TEXT`, the conversion of a number in a `&`) and read a number from a text, in `calculate`, in the
recalculation that follows an edit and in the queries. The cases are those of issues #84 and #106; the
oracle is `semantics_es-CO.{xlsx,ods}`, made by a LibreOffice running with the `es-CO` regional setting."""

import pytest

import sheetrs
from conftest import ROOT, make_xlsx

# (formula, as LibreOffice shows it with es-CO, with en-US)
CASES = [
    ('TEXT(350*49799,"[$$-en-US]#.###.##0")', "$17.429.650", "$17429650.000.000"),
    ('"Componente inflacionario "&(0.5543*100)&"%"', "Componente inflacionario 55,43%", "Componente inflacionario 55.43%"),
    ('TEXT(0,"[$$-en-US]#,##0;([$$-en-US]#,##0)")', "$,000", "$0"),
]
ORACLES = ROOT / "sheetrs-compiler/tests/assets"


def workbook(tmp_path):
    cells = {f"A{i}": ("=" + formula, "?") for i, (formula, _, _) in enumerate(CASES, 1)}
    return make_xlsx(tmp_path / "cases.xlsx", {"S": cells})


def values(book):
    block = book.calculate(ranges="S!A1:A3").value_ranges()["S!A1:A3"]
    return [row[0] for row in block.values]


def test_an_xlsx_calculates_with_the_locale_given_to_the_book(tmp_path):
    path = workbook(tmp_path)
    assert values(sheetrs.Book(path, locale="es-CO")) == [es for _, es, _ in CASES]
    assert values(sheetrs.Book(path, locale="en-US")) == [en for _, _, en in CASES]


def test_the_locale_is_given_with_the_bytes_and_streams_too(tmp_path):
    import io

    data = workbook(tmp_path).read_bytes()
    for source in (data, bytearray(data), io.BytesIO(data)):
        assert values(sheetrs.Book(source, locale="es_CO")) == [es for _, es, _ in CASES]


def test_an_ods_calculates_with_the_locale_of_the_file():
    # Saved by a LibreOffice with the es-CO regional setting, which writes it in the default cell style.
    book = sheetrs.Book(ORACLES / "semantics_es-CO.ods")
    assert book.locale is None
    labels = [row[0] for row in book.range("Calc!B1:B1500")["Calc!B1:B1500"].values]
    for formula, es, en in CASES:
        row = labels.index(formula) + 1
        asked = f"Calc!A{row}"
        assert sheetrs.Book(ORACLES / "semantics_es-CO.ods").calculate(ranges=asked).value_ranges().to_dict() == {asked: [[es]]}
        # What is given to the book wins over what the file says.
        other = sheetrs.Book(ORACLES / "semantics_es-CO.ods", locale="en-US")
        assert other.calculate(ranges=asked).value_ranges().to_dict() == {asked: [[en]]}


def test_the_locale_is_normalized_and_none_when_not_given(tmp_path):
    path = workbook(tmp_path)
    assert sheetrs.Book(path).locale is None
    assert sheetrs.Book(path, locale="es_co").locale == "es-CO"
    assert sheetrs.Book(path, locale="de").locale == "de"


def test_a_locale_that_is_not_a_tag_is_refused_when_the_book_is_made(tmp_path):
    with pytest.raises(ValueError, match="not a locale tag"):
        sheetrs.Book(workbook(tmp_path), locale="not a locale")


def test_a_locale_the_engine_does_not_have_is_an_error_that_names_it(tmp_path):
    book = sheetrs.Book(workbook(tmp_path), locale="fr-FR").calculate(ranges="S!A1")
    with pytest.raises(sheetrs.Error, match="fr-FR"):
        book.value_ranges()


def test_the_book_that_save_returns_keeps_the_locale(tmp_path):
    path = make_xlsx(tmp_path / "in.xlsx", {"S": {"A1": 1, "B1": ("=TEXT(1234.5,\"#,##0.00\")", "?"), "C1": 2}})
    saved = sheetrs.Book(path, locale="es-CO").delete_columns("S!C").save(tmp_path / "out.xlsx")
    assert saved.locale == "es-CO"
    # The recalculation that follows the edit used it.
    assert sheetrs.Book(tmp_path / "out.xlsx").range("S!B1")["S!B1"].values == [["1.234,50"]]


def test_evaluate_with_a_locale_still_works_and_does_not_need_a_book():
    assert sheetrs.evaluate('=TEXT(1234.5,"#,##0.00")', locale="es-CO") == "1.234,50"
    assert sheetrs.evaluate('=TEXT(1234.5,"#,##0.00")') == "1,234.50"
