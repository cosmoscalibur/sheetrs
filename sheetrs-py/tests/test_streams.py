"""A workbook in memory: a `Book` is made from a binary stream and `save` writes to one, and `save()`
with no destination replaces the stream the `Book` came from, as it replaces a file."""

import io

import pytest

import sheetrs

from conftest import STRUCTURE


def _bytes(extension):
    return (STRUCTURE / f"book.{extension}").read_bytes()


@pytest.fixture(params=["xlsx", "ods"])
def extension(request):
    return request.param


def test_a_book_is_made_from_bytes_a_bytearray_or_a_stream(extension):
    data = _bytes(extension)
    for source in (data, bytearray(data), io.BytesIO(data)):
        book = sheetrs.Book(source)
        assert book.format == extension and book.sheet_names == ["T", "U", "V", "F"]


def test_a_stream_is_read_from_its_start_even_when_it_is_at_the_end(extension):
    stream = io.BytesIO()
    stream.write(_bytes(extension))
    assert sheetrs.Book(stream).sheet_names == ["T", "U", "V", "F"]


def test_a_stream_that_gives_text_is_refused_with_the_way_to_open_it(extension, tmp_path):
    path = tmp_path / f"w.{extension}"
    path.write_bytes(_bytes(extension))
    with open(path, encoding="latin-1") as text:
        with pytest.raises(TypeError, match="binary"):
            sheetrs.Book(text)


def test_save_to_a_stream_writes_the_workbook_the_path_would_get(extension, tmp_path):
    book = sheetrs.Book(_bytes(extension)).delete_rows("T!2")
    stream = io.BytesIO()
    saved = book.save(stream)
    out = tmp_path / f"out.{extension}"
    book.save(out)
    assert stream.getvalue() == out.read_bytes()
    assert saved.format == extension and saved.sheets["T"].used_cells == sheetrs.Book(out).sheets["T"].used_cells
    # The Book that save returns is made on the stream: its own save() rewrites it.
    saved.delete_rows("T!2").save()
    assert sheetrs.Book(stream.getvalue()).sheets["T"].used_cells < sheetrs.Book(out).sheets["T"].used_cells


def test_save_with_no_destination_rewrites_the_stream_the_book_came_from(extension):
    stream = io.BytesIO(_bytes(extension))
    book = sheetrs.Book(stream)
    before = book.sheets["T"].used_cells
    returned = book.delete_rows("T!2").save()
    assert returned is book
    assert sheetrs.Book(stream.getvalue()).sheets["T"].used_cells < before
    assert book.sheets["T"].used_cells < before
    assert "pending" not in repr(book)


def test_a_stream_that_gets_shorter_is_cut_where_the_workbook_ends(extension):
    stream = io.BytesIO(_bytes(extension))
    sheetrs.Book(stream).delete_sheets(["U", "V"]).save()
    assert sheetrs.Book(stream.getvalue()).sheet_names == ["T", "F"]
    assert len(stream.getvalue()) < len(_bytes(extension))


def test_bytes_have_no_original_to_replace(extension):
    with pytest.raises(sheetrs.Error):
        sheetrs.Book(_bytes(extension)).delete_rows("T!2").save()


def test_the_value_ranges_are_written_as_json_to_a_stream_too(extension):
    stream = io.BytesIO()
    answer = io.BytesIO()
    sheetrs.Book(_bytes(extension)).calculate(ranges="F!A5:A6").save(stream, json=answer)
    assert answer.getvalue().endswith(b"\n") and b"valueRanges" in answer.getvalue()
    assert sheetrs.Book(stream.getvalue()).format == extension


def test_a_destination_that_is_neither_a_path_nor_a_stream_is_a_type_error(extension):
    with pytest.raises(TypeError, match="destination"):
        sheetrs.Book(_bytes(extension)).delete_rows("T!2").save(42)
