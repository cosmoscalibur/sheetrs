"""`save(recalculate=False)` and `to_bytes(recalculate=False)`: the formulas keep the results the file
caches (issue #126), and the workbook is not marked to be recalculated when it is opened."""

import zipfile

import pytest

import sheetrs
from conftest import make_xlsx

CACHED = {
    # Excel caches `x` for B2 (0<>"" is TRUE there); the engine, like LibreOffice, gives 0.
    "Sheet1": {"B1": ("=C1", 0), "B2": ('=IF(B1<>"","x",0)', "x")},
    "Data": {"A1": 1, "A2": 2, "A3": 3},
}


@pytest.fixture
def book(tmp_path):
    return make_xlsx(tmp_path / "book.xlsx", CACHED)


def results(path):
    return sheetrs.Book(path).range("Sheet1!B1:B2")["Sheet1!B1:B2"].values


def marked(path):
    with zipfile.ZipFile(path) as z:
        return b"fullCalcOnLoad" in z.read("xl/workbook.xml")


def test_by_default_an_edit_calculates_every_formula(book, tmp_path):
    sheetrs.Book(book).delete_columns("Data!D").save(tmp_path / "out.xlsx")
    assert results(tmp_path / "out.xlsx") == [[0], [0]]


@pytest.mark.parametrize(
    "edit",
    [
        lambda b: b.delete_columns("Data!D"),
        lambda b: b.delete_rows("Data!3"),
        lambda b: b.formula_to_value("Data"),
    ],
    ids=["delete_columns", "delete_rows", "formula_to_value"],
)
def test_without_recalculation_the_cached_results_stay(book, tmp_path, edit):
    out = tmp_path / "out.xlsx"
    edit(sheetrs.Book(book)).save(out, recalculate=False)
    assert results(out) == [[0], ["x"]]
    assert not marked(out)


def test_to_bytes_takes_the_option_too(book, tmp_path):
    out = tmp_path / "out.xlsx"
    out.write_bytes(sheetrs.Book(book).delete_columns("Data!D").to_bytes(recalculate=False))
    assert results(out) == [[0], ["x"]]
    assert not marked(out)


def test_explain_shows_the_plan_without_the_recalculation(book):
    plan = sheetrs.Book(book).delete_columns("Data!D")
    assert "recalculate" in plan.explain()
    assert "recalculate" not in plan.explain(recalculate=False)


def test_a_calculation_runs_whatever_the_option(book, tmp_path):
    out = tmp_path / "out.xlsx"
    sheetrs.Book(book).calculate().save(out, recalculate=False)
    assert results(out) == [[0], [0]]
