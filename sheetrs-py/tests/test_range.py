"""`Book.range`: one or several ranges read in one pass, as blocks of the properties asked for."""

import datetime

import pytest

import sheetrs
from conftest import io_asset, make_xlsx

T_A1_C4 = [[1, 10, 1], [2, 20, 40], [3, 30, 6], [4, 40, 160]]


def test_a_range_gives_a_block_of_values(structure):
    book = sheetrs.Book(structure)
    blocks = book.range("T!A1:C4")
    assert list(blocks) == ["T!A1:C4"]
    block = blocks["T!A1:C4"]
    assert isinstance(block, sheetrs.Block)
    assert block.values == T_A1_C4
    assert block.range == "T!A1:C4"
    assert block.major == "rows"


def test_several_ranges_of_several_sheets_come_in_the_order_asked(structure):
    book = sheetrs.Book(structure)
    blocks = book.range(["U!A1:B2", "T!A1:A2", "F!A1:A3", "V!B1"])
    assert list(blocks) == ["U!A1:B2", "T!A1:A2", "F!A1:A3", "V!B1"]
    assert blocks["U!A1:B2"].values == [[101, 110], [102, 120]]
    assert blocks["T!A1:A2"].values == [[1], [2]]
    assert blocks["F!A1:A3"].values == [[2], [4], [5]]
    assert blocks["V!B1"].values == [[1010]]


def test_overlapping_ranges_each_get_their_cells(structure):
    blocks = sheetrs.Book(structure).range(["T!A1:B2", "T!B2:C3"])
    assert blocks["T!A1:B2"].values == [[1, 10], [2, 20]]
    assert blocks["T!B2:C3"].values == [[20, 40], [30, 6]]


def test_a_range_without_a_sheet_is_on_the_first_one(structure):
    assert sheetrs.Book(structure).range("A1:B2")["A1:B2"].values == [[1, 10], [2, 20]]


def test_the_properties_asked_for_are_the_ones_given(structure):
    book = sheetrs.Book(structure)
    block = book.range(
        "T!A1:C2", properties=("value", "formula", "number_format", "display")
    )["T!A1:C2"]
    assert block.values == [[1, 10, 1], [2, 20, 40]]
    assert [row[2] is not None for row in block.formulas] == [True, True]
    assert block.formulas[0][0] is None
    assert block.number_formats[0][0] == "General"
    assert block.displays == [["1", "10", "1"], ["2", "20", "40"]]


def test_a_property_that_was_not_asked_for_raises(structure):
    block = sheetrs.Book(structure).range("T!A1:B2")["T!A1:B2"]
    for name, asked in (
        ("formulas", "formula"),
        ("number_formats", "number_format"),
        ("displays", "display"),
        ("comments", "comment"),
    ):
        with pytest.raises(AttributeError, match=asked):
            getattr(block, name)
    only_formulas = sheetrs.Book(structure).range("T!C1:C2", "formula")["T!C1:C2"]
    with pytest.raises(AttributeError, match="value"):
        only_formulas.values
    # What is iterated is the first property asked for when `value` was not.
    assert [row[0] is not None for row in only_formulas] == [True, True]


def test_the_properties_are_checked():
    with pytest.raises(ValueError, match="properties"):
        sheetrs.Book(io_asset("cells.xlsx")).range("T!A1", properties=("value", "colour"))
    with pytest.raises(ValueError, match="empty"):
        sheetrs.Book(io_asset("cells.xlsx")).range("T!A1", properties=())


def test_a_block_is_a_sequence_of_rows(structure):
    block = sheetrs.Book(structure).range("T!A1:C4")["T!A1:C4"]
    assert len(block) == 4
    assert list(block) == T_A1_C4
    assert block[0] == [1, 10, 1]
    assert block[-1] == [4, 40, 160]
    assert block[1:3] == [[2, 20, 40], [3, 30, 6]]
    assert block[2][1] == 30
    with pytest.raises(IndexError):
        block[4]
    assert "T!A1:C4" in repr(block)


def test_by_columns_the_lists_are_lists_of_columns(structure):
    block = sheetrs.Book(structure).range("T!A1:C2", major="columns")["T!A1:C2"]
    assert block.major == "columns"
    assert block.values == [[1, 2], [10, 20], [1, 40]]
    assert len(block) == 3
    with pytest.raises(ValueError, match="major"):
        sheetrs.Book(structure).range("T!A1", major="diagonal")


def test_an_empty_cell_is_none_inside_a_bounded_range(structure):
    block = sheetrs.Book(structure).range("T!C5:E6")["T!C5:E6"]
    assert block.values == [[12, 250, None], [360, None, None]]
    assert sheetrs.Book(structure).range("T!G9:H10")["T!G9:H10"].values == [
        [None, None],
        [None, None],
    ]


def test_a_range_open_at_one_end_is_cut_at_its_last_cell(structure):
    book = sheetrs.Book(structure)
    columns = book.range("T!A:C")["T!A:C"]
    assert len(columns) == 10
    assert columns[-1] == [10, 100, 1000]
    assert book.range("T!1:3")["T!1:3"].values == T_A1_C4[:3]
    assert book.range("T!Z:AA")["T!Z:AA"].values == []


def test_a_defined_name_that_is_a_range_is_accepted(structure):
    blocks = sheetrs.Book(structure).range(["NameTwo", "T!A1"])
    assert blocks["NameTwo"].values == [[n] for n in range(1, 11)]
    assert blocks["T!A1"].values == [[1]]
    with pytest.raises(sheetrs.InvalidRangeError, match="NoSuchName"):
        sheetrs.Book(structure).range("NoSuchName")


def test_a_range_or_a_sheet_that_does_not_exist_is_an_invalid_range(structure):
    book = sheetrs.Book(structure)
    with pytest.raises(sheetrs.InvalidRangeError, match="Nowhere"):
        book.range("Nowhere!A1")
    with pytest.raises(ValueError):  # it is also a ValueError
        book.range("T!A1:::B2")


def test_ranges_are_one_text_or_a_list_of_texts(structure):
    book = sheetrs.Book(structure)
    assert list(book.range(("T!A1", "T!A2"))) == ["T!A1", "T!A2"]
    assert list(book.range(r for r in ["T!A1"])) == ["T!A1"]
    with pytest.raises(TypeError, match="ranges"):
        book.range(3)


def test_numbers_are_floats_or_ints_when_exact(tmp_path):
    path = make_xlsx(
        tmp_path / "n.xlsx", {"S": {"A1": 3, "A2": 2.5, "A3": 1e20, "A4": -7}}
    )
    values = sheetrs.Book(path).range("S!A1:A4")["S!A1:A4"].values
    assert values == [[3], [2.5], [1e20], [-7]]
    assert [type(row[0]) for row in values] == [int, float, float, int]


def test_text_booleans_and_empty_cells_have_their_python_types(tmp_path):
    path = make_xlsx(tmp_path / "t.xlsx", {"S": {"A1": "hello", "A2": True, "A3": False}})
    values = sheetrs.Book(path).range("S!A1:A4")["S!A1:A4"].values
    assert values == [["hello"], [True], [False], [None]]
    assert [type(row[0]) for row in values] == [str, bool, bool, type(None)]


def test_an_error_is_a_cell_error_and_not_a_text(tmp_path):
    path = make_xlsx(
        tmp_path / "e.xlsx",
        {"S": {"A1": ("#N/A",), "A2": "#N/A", "A3": ("#DIV/0!",)}},
    )
    values = sheetrs.Book(path).range("S!A1:A3")["S!A1:A3"].values
    error, text, division = (row[0] for row in values)
    assert isinstance(error, sheetrs.CellError) and error.code == "#N/A"
    assert isinstance(text, str) and text == "#N/A"
    assert division.code == "#DIV/0!"
    assert error != text and text != error
    assert error == sheetrs.CellError("#N/A") and error != division


def test_errors_of_both_formats(request):
    for name in ("errors.xlsx", "errors.ods"):
        block = sheetrs.Book(io_asset(name)).range("Errors!A5:A6")["Errors!A5:A6"]
        assert block.values == [[sheetrs.CellError("#N/A")], [sheetrs.CellError("#DIV/0!")]]


def test_a_cell_error_is_a_value_with_its_code():
    error = sheetrs.CellError("Err:502")
    assert error.code == "Err:502"
    assert str(error) == "Err:502"
    assert repr(error) == "CellError('Err:502')"
    assert error == sheetrs.CellError("Err:502")
    assert hash(error) == hash(sheetrs.CellError("Err:502"))
    assert {error, sheetrs.CellError("Err:502"), sheetrs.CellError("#N/A")} == {
        error,
        sheetrs.CellError("#N/A"),
    }
    assert error != "Err:502"
    assert error != 502
    with pytest.raises(AttributeError):
        error.code = "x"


@pytest.mark.parametrize("name", ["cells.xlsx", "cells.ods"])
def test_dates_are_serial_numbers_unless_asked(name):
    book = sheetrs.Book(io_asset(name))
    serial = book.range("T!A1:A5")["T!A1:A5"].values
    assert serial[0] == [45000] and serial[3] == [1.0625]
    assert all(isinstance(row[0], (int, float)) for row in serial)
    assert book.range("T!A1:A5", dates=False)["T!A1:A5"].values == serial


@pytest.mark.parametrize("name", ["cells.xlsx", "cells.ods"])
def test_dates_as_dates_when_asked(name):
    book = sheetrs.Book(io_asset(name))
    values = book.range("T!A1:A5", dates=True)["T!A1:A5"].values
    assert values == [
        [datetime.date(2023, 3, 15)],
        [datetime.datetime(2023, 3, 15, 13, 45, 30)],
        [datetime.time(12, 30)],
        # An elapsed time ([hh]:mm:ss) is not a time of day: it stays a number.
        [1.0625],
        [datetime.date(1900, 1, 15)],
    ]
    assert type(values[0][0]) is datetime.date
    assert type(values[1][0]) is datetime.datetime
    assert type(values[2][0]) is datetime.time


@pytest.mark.parametrize(
    ("name", "origin"),
    [
        ("dates.xlsx", datetime.date(1899, 12, 30)),
        ("dates.ods", datetime.date(1899, 12, 30)),
        ("dates_1904.xlsx", datetime.date(1904, 1, 1)),
        ("dates_1904.ods", datetime.date(1904, 1, 1)),
        ("dates_1900.ods", datetime.date(1900, 1, 1)),
    ],
)
def test_the_date_system_is_the_one_of_the_file(name, origin):
    values = sheetrs.Book(io_asset(name)).range("Dates!A1:B3", dates=True)[
        "Dates!A1:B3"
    ].values
    assert values[0][0] == origin
    assert values[1][0] == origin + datetime.timedelta(days=1)
    assert values[2][1] == datetime.datetime.combine(origin, datetime.time()) + datetime.timedelta(days=2)


def test_a_date_system_that_is_none_of_the_three_is_an_error_that_names_the_date():
    """LibreOffice stores another null date (its API); SheetRS refuses it when the workbook is read,
    and a `Book` reads the workbook when it is opened."""
    with pytest.raises(sheetrs.UnsupportedDateSystemError) as raised:
        sheetrs.Book(io_asset("other_null_date.ods"))
    message = str(raised.value)
    assert "1950-06-15" in message
    for supported in ("1899-12-30", "1900-01-01", "1904-01-01"):
        assert supported in message
    assert isinstance(raised.value, sheetrs.FormatError)
    assert isinstance(raised.value, ValueError)


def test_fractions_of_a_second_are_kept_when_the_format_writes_them():
    row = sheetrs.Book(io_asset("dates.xlsx")).range("Dates!E1:E1", dates=True)[
        "Dates!E1:E1"
    ].values[0]
    assert row == [datetime.datetime(1899, 12, 30, 0, 0)]


def test_dates_is_a_flag_and_a_name_of_a_kind_of_date_is_refused():
    with pytest.raises(TypeError, match="dates|bool"):
        sheetrs.Book(io_asset("cells.xlsx")).range("T!A1", dates="datetime")


def test_comments_and_formats_are_read_when_asked():
    book = sheetrs.Book(io_asset("cells.xlsx"))
    block = book.range("T!A1:A2", properties=["number_format", "comment"])["T!A1:A2"]
    assert block.number_formats[0][0].endswith("yyyy-mm-dd")
    assert block.comments == [[None], [None]]
    assert len(block) == 2


def test_to_dicts_takes_the_first_row_as_the_header(tmp_path):
    path = make_xlsx(
        tmp_path / "d.xlsx",
        {"S": {"A1": "Name", "B1": "Qty", "A2": "pen", "B2": 3, "A3": "ink", "B3": 5}},
    )
    block = sheetrs.Book(path).range("S!A1:B3")["S!A1:B3"]
    assert block.to_dicts() == [{"Name": "pen", "Qty": 3}, {"Name": "ink", "Qty": 5}]
    by_columns = sheetrs.Book(path).range("S!A1:B3", major="columns")["S!A1:B3"]
    with pytest.raises(AttributeError, match="rows"):
        by_columns.to_dicts()
    formulas_only = sheetrs.Book(path).range("S!A1:B3", "formula")["S!A1:B3"]
    with pytest.raises(AttributeError, match="value"):
        formulas_only.to_dicts()


def test_an_empty_text_is_read_as_libreoffice_reads_each_file():
    # An XLSX empty text (a shared string, an inline string, a `t="str"` cell) is a blank cell for
    # LibreOffice: `None`. An ODS text cell with a paragraph, even an empty one, is an empty text, and
    # one with none (`office:string-value=""` alone) is blank. Rows 1 to 7 of `T` in each file
    # (`sheetrs-io/tests/assets/oracle/make_empty_text.py`).
    xlsx = sheetrs.Book(io_asset("empty_text.xlsx")).range("T!A1:A7")["T!A1:A7"].values
    assert xlsx == [[None]] * 7
    ods = sheetrs.Book(io_asset("empty_text.ods")).range("T!A1:A7")["T!A1:A7"].values
    assert ods == [[""], [""], [None], [None], [""], [""], [None]]
