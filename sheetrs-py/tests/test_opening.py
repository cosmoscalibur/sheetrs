"""What a `Book` says about the workbook it reads the first time it is described: the formats, the sheets, the
names, the date system, the macros and the totals; and the `Stats` and `Sheet` values that hold them."""

import datetime
import json
import zipfile

import pytest

import sheetrs

from conftest import STRUCTURE, io_asset, make_xlsx


def test_the_sheets_are_looked_up_by_name_or_by_position_and_iterate_in_order(structure):
    book = sheetrs.Book(structure)
    assert book.sheet_names == ["T", "U", "V", "F"]
    assert [s.name for s in book.sheets] == book.sheet_names
    assert len(book.sheets) == 4 and "T" in book.sheets and "Z" not in book.sheets
    assert book.sheets["F"].name == "F" == book.sheets[3].name == book.sheets[-1].name
    with pytest.raises(KeyError):
        book.sheets["Z"]
    with pytest.raises(IndexError):
        book.sheets[4]


def test_a_sheet_says_its_used_range_and_counts_what_the_stats_total(structure):
    book = sheetrs.Book(structure)
    sheet = book.sheets["F"]
    assert sheet.used_range.startswith("A1:")
    assert sheet.state == "visible"
    assert sheet.formulas > 0 and sheet.used_cells >= sheet.formulas
    assert sheet.empty_cells >= 0 and sheet.xml_size > 0
    sheets = list(book.sheets)
    assert book.total_sheets == len(sheets)
    assert book.total_used_cells == sum(s.used_cells for s in sheets)
    assert book.total_formulas == sum(s.formulas for s in sheets)
    assert book.total_errors == sum(s.errors for s in sheets)
    assert book.total_empty_cells == sum(s.empty_cells for s in sheets)
    assert book.total_comments == sum(s.comments for s in sheets)
    assert book.total_charts == sum(s.charts for s in sheets)
    assert book.total_tables == sum(s.tables for s in sheets)
    assert book.total_xml_size == sum(s.xml_size for s in sheets)
    assert book.total_file_size == structure.stat().st_size
    assert book.total_named_ranges == len(book.defined_names)


def test_a_sheet_without_cells_has_no_used_range(tmp_path):
    path = make_xlsx(tmp_path / "w.xlsx", {"Empty": {}, "Full": {"B3": 1}})
    book = sheetrs.Book(path)
    assert book.sheets["Empty"].used_range is None
    assert book.sheets["Full"].used_range == "A1:B3"


def test_a_hidden_sheet_says_its_state(tmp_path):
    path = make_xlsx(tmp_path / "w.xlsx", {"A": {"A1": 1}, "B": {"A1": 2}, "C": {"A1": 3}})
    patched = tmp_path / "hidden.xlsx"
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(patched, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = data.replace(b'name="B"', b'name="B" state="hidden"')
                data = data.replace(b'name="C"', b'name="C" state="veryHidden"')
            target.writestr(item, data)
    book = sheetrs.Book(patched)
    assert [s.state for s in book.sheets] == ["visible", "hidden", "very_hidden"]


def test_the_defined_names_are_a_dict_ordered_by_name(tmp_path):
    path = make_xlsx(tmp_path / "w.xlsx", {"Data": {"A1": 1}}, names={"Zeta": "Data!$A$1", "Alpha": "Data!$B$2"})
    book = sheetrs.Book(path)
    assert list(book.defined_names.items()) == [("Alpha", "Data!$B$2"), ("Zeta", "Data!$A$1")]
    assert book.total_named_ranges == 2


@pytest.mark.parametrize(
    ("name", "origin"),
    [
        ("dates.xlsx", datetime.date(1899, 12, 30)),
        ("dates_1904.xlsx", datetime.date(1904, 1, 1)),
        ("dates_1900.ods", datetime.date(1900, 1, 1)),
    ],
)
def test_the_date_system_is_the_date_of_serial_zero(name, origin):
    assert sheetrs.Book(io_asset(name)).date_system == origin


def test_the_format_and_the_macros_are_found_in_the_content(tmp_path):
    plain = make_xlsx(tmp_path / "plain.xlsx", {"A": {"A1": 1}})
    book = sheetrs.Book(plain)
    assert book.format == "xlsx" and book.has_macros is False
    with zipfile.ZipFile(plain) as source, zipfile.ZipFile(tmp_path / "macros.xlsx", "w") as target:
        for item in source.infolist():
            target.writestr(item, source.read(item.filename))
        target.writestr("xl/vbaProject.bin", b"\0")
    assert sheetrs.Book(tmp_path / "macros.xlsx").has_macros is True
    assert sheetrs.Book(io_asset("cells.ods")).format == "ods"


def test_the_properties_describe_the_file_so_they_wait_for_the_operations(structure):
    book = sheetrs.Book(structure)
    book.delete_rows("T!2")
    for read in (
        lambda: book.sheet_names,
        lambda: book.sheets,
        lambda: book.defined_names,
        lambda: book.total_formulas,
        lambda: book.stats,
        lambda: book.has_macros,
        lambda: book.date_system,
    ):
        with pytest.raises(sheetrs.PendingOperationsError):
            read()
    assert book.format in {"xlsx", "ods"}


def test_a_saved_book_describes_the_workbook_it_wrote(structure, tmp_path):
    book = sheetrs.Book(structure)
    before = book.sheets["T"].used_cells
    saved = book.delete_rows("T!2").save(tmp_path / f"out{structure.suffix}")
    assert saved.sheets["T"].used_cells < before


def test_the_book_that_saved_in_place_describes_the_file_as_it_is_now(structure):
    book = sheetrs.Book(structure)
    before = book.sheets["T"].used_cells
    book.delete_rows("T!2").save()
    assert book.sheets["T"].used_cells < before


def test_the_stats_are_typed_and_write_themselves_as_the_command_line_does(structure):
    stats = sheetrs.Book(structure).stats
    assert isinstance(stats, sheetrs.Stats)
    document = stats.to_dict()
    assert document["total_sheets"] == stats.total_sheets
    assert [s["name"] for s in document["sheets"]] == [s.name for s in stats.sheets]
    assert json.loads(stats.to_json()) == document
    sheet = stats.sheets["F"]
    assert sheet.to_dict() == next(s for s in document["sheets"] if s["name"] == "F")
    assert json.loads(sheet.to_json()) == sheet.to_dict()
