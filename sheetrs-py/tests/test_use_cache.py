"""`use_cache=True`: the recalculation after an edit calculates what the operations reach and leaves every other
formula with the result the file caches. The workbooks cache results the engine does not give (`from-excel` where
the engine gives `t`), in each format, and the 40 formulas `C1:C40` that no edit here reaches keep the share of
stale formulas where the walk is used."""

import shutil
from pathlib import Path

import pytest

import sheetrs

ROOT = Path(__file__).resolve().parents[2]
ASSETS = ROOT / "sheetrs-edit/tests/assets/use_cache"
CALC = "Calc!A1:A9"


@pytest.fixture(params=["xlsx", "ods"])
def book(request, tmp_path):
    source = ASSETS / f"book.{request.param}"
    copy = tmp_path / source.name
    shutil.copy(source, copy)
    return copy


def column(path):
    """The values of `Calc!A1:A9`, one per row."""
    return [row[0] for row in sheetrs.Book(path).range(CALC)[CALC].values]


def test_an_edit_leaves_the_results_of_the_formulas_it_does_not_reach(book, tmp_path):
    kept, full = tmp_path / f"kept{book.suffix}", tmp_path / f"full{book.suffix}"
    sheetrs.Book(book).delete_rows("Data!3").save(kept, use_cache=True)
    sheetrs.Book(book).delete_rows("Data!3").save(full)
    cached, with_cache, without = column(book), column(kept), column(full)
    # The range that shrank, the whole column the row leaves, `ROW()`: calculated, as a full run does.
    for index in (0, 4, 7):
        assert with_cache[index] == without[index] != cached[index]
    assert with_cache[0] == 180
    # `A7` reads a cell that stays: the result the file caches stays, where the engine gives `t`.
    assert with_cache[6] == cached[6] == "from-excel"
    assert without[6] == "t"


def test_it_is_a_keyword_of_to_bytes_and_explain(book, tmp_path):
    editor = sheetrs.Book(book).delete_rows("Data!3")
    assert "use the cache" in editor.explain(use_cache=True)
    assert "use the cache" not in editor.explain()
    out = tmp_path / f"out{book.suffix}"
    editor.save(out, use_cache=True)
    assert editor.to_bytes(use_cache=True) == out.read_bytes()
    assert editor.to_bytes() != out.read_bytes()


def test_the_report_of_the_issue(book, tmp_path):
    """An edit on another sheet does not replace what the application cached."""
    repro = ASSETS / f"repro.{book.suffix[1:]}"
    kept, full = tmp_path / f"kept{book.suffix}", tmp_path / f"full{book.suffix}"
    sheetrs.Book(repro).delete_columns("Data!D").save(kept, use_cache=True)
    sheetrs.Book(repro).delete_columns("Data!D").save(full)
    values = lambda path: [row[0] for row in sheetrs.Book(path).range("Calc!B1:B2")["Calc!B1:B2"].values]
    assert values(kept) == [0, "x"]
    assert values(full) == [0, 0]
