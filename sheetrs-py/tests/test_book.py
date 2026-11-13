"""The `Book`: how it is made, what it says about itself, and the exceptions."""

import os
from pathlib import Path

import pytest

import sheetrs
from conftest import LINT_BOOKS, STRUCTURE, make_xlsx


def test_version_is_the_workspace_version():
    major, minor, patch = sheetrs.__version__.split(".")
    assert (major + minor + patch).isdigit()


def test_a_book_is_made_from_a_path_a_pathlike_or_bytes(lint_book):
    by_str = sheetrs.Book(str(lint_book))
    by_path = sheetrs.Book(lint_book)
    by_bytes = sheetrs.Book(lint_book.read_bytes())
    assert by_str.stats().sheets.names() == by_path.stats().sheets.names() == by_bytes.stats().sheets.names()
    assert by_str.format == by_path.format == by_bytes.format == lint_book.suffix[1:]


def test_a_path_like_object_is_a_path(lint_book):
    class Wrapper(os.PathLike):
        def __fspath__(self):
            return str(lint_book)

    assert sheetrs.Book(Wrapper()).stats().sheets.names() == sheetrs.Book(lint_book).stats().sheets.names()


def test_the_format_is_the_one_of_the_content():
    for path in LINT_BOOKS:
        assert sheetrs.Book(path).format == path.suffix[1:]
        assert sheetrs.Book(path.read_bytes()).format == path.suffix[1:]


def test_what_is_not_a_workbook_is_a_format_error():
    with pytest.raises(sheetrs.FormatError):
        sheetrs.Book(b"not a workbook")


def test_the_extension_is_not_consulted(tmp_path):
    renamed = tmp_path / "book.ods"
    renamed.write_bytes(LINT_BOOKS[0].read_bytes())
    assert sheetrs.Book(renamed).format == LINT_BOOKS[0].suffix[1:]


def test_a_path_that_is_not_a_file_is_file_not_found(tmp_path):
    with pytest.raises(FileNotFoundError):
        sheetrs.Book(tmp_path / "absent.xlsx")
    (tmp_path / "folder.xlsx").mkdir()
    with pytest.raises(FileNotFoundError):
        sheetrs.Book(tmp_path / "folder.xlsx")


def test_a_file_that_is_not_a_workbook_is_a_format_error(tmp_path):
    other = tmp_path / "notes.txt"
    other.write_text("not a workbook")
    with pytest.raises(sheetrs.FormatError):
        sheetrs.Book(other)
    # A FormatError is an Error and a ValueError: an argument that is not allowed.
    with pytest.raises(ValueError):
        sheetrs.Book(other)


def test_anything_else_is_a_type_error():
    with pytest.raises(TypeError, match="path"):
        sheetrs.Book(3.5)


def test_a_workbook_that_cannot_be_parsed_is_an_error(tmp_path):
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"PK this is not a zip")
    with pytest.raises(sheetrs.Error):
        sheetrs.Book(broken).stats().sheets.names()


def test_the_hierarchy_of_the_exceptions():
    assert issubclass(sheetrs.Error, Exception)
    assert issubclass(sheetrs.FormatError, (sheetrs.Error, ValueError))
    assert issubclass(sheetrs.InvalidRangeError, (sheetrs.Error, ValueError))
    assert issubclass(sheetrs.UnavailableFunctionError, sheetrs.Error)
    assert issubclass(sheetrs.PendingOperationsError, sheetrs.Error)
    assert not issubclass(sheetrs.PendingOperationsError, ValueError)


def test_the_book_says_what_it_is_without_reading_it(structure):
    book = sheetrs.Book(structure)
    assert book.format == structure.suffix[1:]
    assert str(structure) in repr(book)
    assert "pending" not in repr(book)
    book.delete_rows("T!2")
    assert "1 pending operation" in repr(book)
    book.delete_columns("T!B")
    assert "2 pending operations" in repr(book)


def test_the_sheets_and_the_names(structure):
    book = sheetrs.Book(structure)
    assert book.stats().sheets.names() == ["T", "U", "V", "F"]
    names = book.stats().defined_names
    assert set(names) == {"Gone", "NameOne", "NameTwo", "NameThree"}
    assert names["NameTwo"].replace("$", "") == "T!A1:A10"


def test_a_book_from_bytes_reads_what_a_file_does(structure):
    from_bytes = sheetrs.Book(structure.read_bytes())
    assert from_bytes.range("T!A1:B3")["T!A1:B3"].values == (
        sheetrs.Book(structure).range("T!A1:B3")["T!A1:B3"].values
    )
    assert "bytes" in repr(from_bytes)


def test_a_workbook_written_by_hand_is_read(tmp_path):
    path = make_xlsx(tmp_path / "mini.xlsx", {"S": {"A1": 1}, "R": {"A1": 2}})
    assert sheetrs.Book(path).stats().sheets.names() == ["S", "R"]
    assert isinstance(Path(path).read_bytes(), bytes)
