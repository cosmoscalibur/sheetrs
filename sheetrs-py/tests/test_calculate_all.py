"""The interface contract of the calculation mode. The logic is in Rust (`sheetrs-edit/tests/calculation_mode.rs`,
`write_to_libreoffice.rs` against `LibreOffice`, `sheetrs-cli/tests/edit_calculate_all.rs`): which formulas the plan reaches, what an ODS calculates
when it starts, what a deletion reaches. What the Python interface has to get right is the wiring, and that is all this file proves: `calculate_all` is
accepted by each method that exposes it, it is `False` by default, and it reaches the engine in both formats (a result that only the engine calculating
every formula gives changes with it); the errors of Rust are `sheetrs.Error`.

The workbooks cache results the engine does not give (`from-excel` where the engine gives `t`): a formula that kept its cache is told from one that
was calculated."""

import shutil
from pathlib import Path

import pytest

import sheetrs

ROOT = Path(__file__).resolve().parents[2]
ASSETS = ROOT / "sheetrs-edit/tests/assets/cached"
CALC = "Calc!A1:A9"


@pytest.fixture(params=["xlsx", "ods"])
def book(request, tmp_path):
    source = ASSETS / f"book.{request.param}"
    copy = tmp_path / source.name
    shutil.copy(source, copy)
    return copy


def _column(path):
    """The values of `Calc!A1:A9`, one per row."""
    return [row[0] for row in sheetrs.Book(path).range(CALC)[CALC].values]


def calculate_all_reaches_save(book, tmp_path):
    """The default keeps what the file caches for a formula the edit does not reach (`A7` reads a cell that stays); `calculate_all=True` calculates it.
    A pass-through that drops the keyword gives the default result for both and fails here."""
    kept, full = tmp_path / f"kept{book.suffix}", tmp_path / f"full{book.suffix}"
    sheetrs.Book(book).delete_rows("Data!3").save(kept)
    sheetrs.Book(book).delete_rows("Data!3").save(full, calculate_all=True)
    assert _column(kept)[6] == "from-excel"
    assert _column(full)[6] == "t"


def calculate_all_is_a_keyword_of_every_method_that_has_it(book, tmp_path):
    editor = sheetrs.Book(book).delete_rows("Data!3")
    # The default is the cache: the keyword is off unless given (what each mode does is tested in Rust).
    editor.explain(calculate_all=True)
    out = tmp_path / f"out{book.suffix}"
    editor.save(out, calculate_all=True)
    assert editor.to_bytes(calculate_all=True) == out.read_bytes()
    assert editor.to_bytes() != out.read_bytes()
    # `value_ranges` has no calculation to answer here: Rust's error is `sheetrs.Error`, with the keyword accepted.
    with pytest.raises(sheetrs.Error):
        editor.value_ranges(calculate_all=True)


def calculate_all_reaches_value_ranges(book):
    """A calculation by default changes only what its inputs reach; with `calculate_all` every formula is calculated."""
    ranges = ["Calc!A7"]
    kept = sheetrs.Book(book).calculate({"Data!A2": 25}, ranges).value_ranges().to_dict()
    full = sheetrs.Book(book).calculate({"Data!A2": 25}, ranges).value_ranges(calculate_all=True).to_dict()
    assert kept["Calc!A7"] == [["from-excel"]]
    assert full["Calc!A7"] == [["t"]]
