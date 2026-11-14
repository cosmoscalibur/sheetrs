"""`Book.find`: the search of LibreOffice's Find & Replace, over the cells."""

import pytest

import sheetrs
from conftest import make_xlsx

TEXTS = {
    "S": {
        "A1": "Apple",
        "A2": "apple pie",
        "A3": "Banana",
        "B1": 1.5,
        "B2": ('=A1&"!"', "Apple!"),
        "B3": "x1y",
        "C1": "APPLE",
        "C3": "bandana",
    },
    "R": {"A1": "apple", "B2": "pear"},
}


@pytest.fixture
def texts(tmp_path):
    return sheetrs.Book(make_xlsx(tmp_path / "texts.xlsx", TEXTS))


def cells(matches):
    return [f"{m.sheet}!{m.cell}" for m in matches]


def test_a_search_finds_the_cells_and_says_where(structure):
    found = sheetrs.Book(structure).find("101")
    assert cells(found) == ["U!A1", "V!B1", "V!A10"]
    match = found[0]
    assert isinstance(match, sheetrs.Match)
    assert (match.sheet, match.cell, match.row, match.col) == ("U", "A1", 1, 1)
    assert match.value == 101
    assert match.formula is None and match.display is None
    assert "U!A1" in repr(match)


def test_a_search_in_numbers_looks_at_the_text_of_the_input_line(structure):
    # 10 is in the cells that hold 10 and 100; the formula cells are searched only with include_formulas.
    book = sheetrs.Book(structure)
    assert cells(book.find("10", scope="T")) == ["T!B1", "T!A10", "T!B10"]
    assert "T!C10" in cells(book.find("10", scope="T", include_formulas=True))


def test_the_default_is_a_literal_text_without_case(texts):
    # The matches come sheet by sheet, each by rows.
    assert cells(texts.find("apple")) == ["S!A1", "S!C1", "S!A2", "R!A1"]


def test_match_case(texts):
    assert cells(texts.find("apple", match_case=True)) == ["S!A2", "R!A1"]


def test_entire_cells(texts):
    assert cells(texts.find("apple", entire_cells=True)) == ["S!A1", "S!C1", "R!A1"]
    assert cells(texts.find("apple", entire_cells=True, match_case=True)) == ["R!A1"]


def test_a_formula_cell_is_searched_by_its_result_only_when_asked(texts):
    assert "S!B2" not in cells(texts.find("Apple!"))
    assert cells(texts.find("Apple!", include_formulas=True)) == ["S!B2"]


def test_search_in_formulas(texts):
    found = texts.find("A1", search_in="formulas")
    assert cells(found) == ["S!B2"]
    found = texts.find("A1", search_in="formulas", properties=("value", "formula"))
    assert found[0].formula == 'A1&"!"'
    assert found[0].value == "Apple!"


def test_search_in_comments(texts):
    assert texts.find("anything", search_in="comments") == []


def test_regular_expressions(texts):
    assert cells(texts.find("ba.*na", regex=True)) == ["S!A3", "S!C3"]
    assert cells(texts.find("b(an)+a", regex=True)) == ["S!A3"]


def test_the_anchors_of_a_regular_expression_are_those_of_the_text(texts):
    assert cells(texts.find("^ban", regex=True)) == ["S!A3", "S!C3"]
    assert cells(texts.find("^b(an)+a$", regex=True, entire_cells=True)) == ["S!A3"]
    assert cells(texts.find("^(Banana|Apple pie)$", regex=True, match_case=True)) == ["S!A3"]


def test_a_regular_expression_outside_the_subset_is_a_value_error(texts):
    # What is not reproduced is refused, not answered differently.
    with pytest.raises(ValueError, match="unsupported search"):
        texts.find("\\bban", regex=True)


def test_wildcards(texts):
    assert cells(texts.find("ba?an?", wildcards=True)) == ["S!A3"]
    assert cells(texts.find("*an*", wildcards=True)) == ["S!A3", "S!C3"]
    assert cells(texts.find("ban*", wildcards=True, entire_cells=True)) == ["S!A3", "S!C3"]


def test_regex_and_wildcards_exclude_each_other(texts):
    with pytest.raises(ValueError, match="regex"):
        texts.find("a", regex=True, wildcards=True)


def test_the_options_are_checked(texts):
    with pytest.raises(ValueError, match="search_in"):
        texts.find("a", search_in="everything")
    with pytest.raises(ValueError, match="by"):
        texts.find("a", by="diagonal")
    with pytest.raises(ValueError, match="properties"):
        texts.find("a", properties=["colour"])


def test_a_bad_regular_expression_is_a_value_error(texts):
    with pytest.raises(ValueError):
        texts.find("(", regex=True)


def test_the_scope_is_a_sheet_a_range_or_a_list_of_them(texts):
    assert cells(texts.find("apple", scope="R")) == ["R!A1"]
    assert cells(texts.find("a", scope="S!A1:A2")) == ["S!A1", "S!A2"]
    assert cells(texts.find("apple", scope=["R", "S!A2:A2"])) == ["S!A2", "R!A1"]
    assert cells(texts.find("apple", scope=("S!C1",))) == ["S!C1"]


def test_the_matches_come_by_rows_or_by_columns(texts):
    assert cells(texts.find("e", scope="S")) == ["S!A1", "S!C1", "S!A2"]
    assert cells(texts.find("e", scope="S", by="columns")) == ["S!A1", "S!A2", "S!C1"]
    # By columns the sheets keep the order of the workbook.
    assert cells(texts.find("apple", by="columns")) == ["S!A1", "S!A2", "S!C1", "R!A1"]


def test_the_properties_asked_for_come_with_each_match(texts):
    match = texts.find("apple", properties=["value", "display", "number_format"])[0]
    assert match.value == "Apple"
    assert match.display == "Apple"
    assert match.number_format == "General"
    assert match.formula is None and match.comment is None


def test_the_text_a_cell_shows_can_be_searched(structure):
    found = sheetrs.Book(structure).find("1", scope="T!A1:A1", formatted=True)
    assert cells(found) == ["T!A1"]


def test_nothing_found_is_an_empty_list(texts):
    assert texts.find("zzz") == []


def test_a_search_with_operations_pending_raises(texts):
    texts.delete_sheets("R")
    with pytest.raises(sheetrs.PendingOperationsError):
        texts.find("apple")


def test_the_empty_search_text_finds_nothing_in_every_mode(texts):
    # LibreOffice: findAll with an empty search string reports no cell, in every mode.
    for options in [{}, {"regex": True}, {"wildcards": True}, {"entire_cells": True}, {"regex": True, "entire_cells": True, "match_case": True}]:
        assert texts.find("", **options) == []
    # Nothing is read, so not even a sheet that is not there is an error.
    assert texts.find("", scope="Nowhere") == []
