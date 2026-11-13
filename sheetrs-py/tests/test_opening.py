"""What `Book.stats()` says about the workbook: the sheets, the names, the macros and the totals; the
format of a `Book`; and the `Stats` and `Sheet` values that hold them."""

import datetime
import json
import zipfile

import pytest

import sheetrs

from conftest import STRUCTURE, io_asset, make_xlsx


def test_the_sheets_are_looked_up_by_name_or_by_position_and_iterate_in_order(structure):
    book = sheetrs.Book(structure)
    assert book.stats().sheets.names() == ["T", "U", "V", "F"]
    assert [s.name for s in book.stats().sheets] == book.stats().sheets.names()
    assert len(book.stats().sheets) == 4 and "T" in book.stats().sheets and "Z" not in book.stats().sheets
    assert book.stats().sheets["F"].name == "F" == book.stats().sheets[3].name == book.stats().sheets[-1].name
    with pytest.raises(KeyError):
        book.stats().sheets["Z"]
    with pytest.raises(IndexError):
        book.stats().sheets[4]


def test_a_sheet_says_its_used_range_and_counts_what_the_stats_total(structure):
    book = sheetrs.Book(structure)
    sheet = book.stats().sheets["F"]
    assert sheet.used_range is not None and ":" in sheet.used_range
    assert sheet.state == "visible"
    assert sheet.formulas > 0 and sheet.used_cells >= sheet.formulas
    assert sheet.empty_cells >= 0 and sheet.xml_size > 0
    sheets = list(book.stats().sheets)
    assert book.stats().total_sheets == len(sheets)
    assert book.stats().total_used_cells == sum(s.used_cells for s in sheets)
    assert book.stats().total_formulas == sum(s.formulas for s in sheets)
    assert book.stats().total_errors == sum(s.errors for s in sheets)
    assert book.stats().total_empty_cells == sum(s.empty_cells for s in sheets)
    assert book.stats().total_comments == sum(s.comments for s in sheets)
    assert book.stats().total_charts == sum(s.charts for s in sheets)
    assert book.stats().total_tables == sum(s.tables for s in sheets)
    assert book.stats().total_xml_size == sum(s.xml_size for s in sheets)
    assert book.stats().total_file_size == structure.stat().st_size
    assert book.stats().total_named_ranges == len(book.stats().defined_names)


def test_a_sheet_without_cells_has_no_used_range(tmp_path):
    path = make_xlsx(tmp_path / "w.xlsx", {"Empty": {}, "Full": {"B3": 1}})
    book = sheetrs.Book(path)
    assert book.stats().sheets["Empty"].used_range is None
    assert book.stats().sheets["Full"].used_range == "B3:B3"


def test_the_used_range_starts_at_the_first_used_cell(tmp_path):
    path = make_xlsx(tmp_path / "d.xlsx", {"Sheet1": {"D10": "x", "E11": "y"}})
    sheet = sheetrs.Book(path).stats().sheets["Sheet1"]
    assert sheet.used_range == "D10:E11"
    # The empty cells are counted from A1: 11 rows by 5 columns, 2 of them used.
    assert sheet.used_cells == 2
    assert sheet.empty_cells == 53


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
    assert [s.state for s in book.stats().sheets] == ["visible", "hidden", "very_hidden"]


def test_the_defined_names_are_a_dict_ordered_by_name(tmp_path):
    path = make_xlsx(tmp_path / "w.xlsx", {"Data": {"A1": 1}}, names={"Zeta": "Data!$A$1", "Alpha": "Data!$B$2"})
    book = sheetrs.Book(path)
    assert list(book.stats().defined_names.items()) == [("Alpha", "Data!$B$2"), ("Zeta", "Data!$A$1")]
    assert book.stats().total_named_ranges == 2


def test_the_format_and_the_macros_are_found_in_the_content(tmp_path):
    plain = make_xlsx(tmp_path / "plain.xlsx", {"A": {"A1": 1}})
    book = sheetrs.Book(plain)
    assert book.format == "xlsx" and book.stats().has_macros is False
    with zipfile.ZipFile(plain) as source, zipfile.ZipFile(tmp_path / "macros.xlsx", "w") as target:
        for item in source.infolist():
            target.writestr(item, source.read(item.filename))
        target.writestr("xl/vbaProject.bin", b"\0")
    assert sheetrs.Book(tmp_path / "macros.xlsx").stats().has_macros is True
    assert sheetrs.Book(io_asset("cells.ods")).format == "ods"


def test_stats_is_a_query_so_it_waits_for_the_operations(structure):
    book = sheetrs.Book(structure)
    book.delete_rows("T!2")
    with pytest.raises(sheetrs.PendingOperationsError):
        book.stats()
    assert book.format in {"xlsx", "ods"}


def test_a_saved_book_describes_the_workbook_it_wrote(structure, tmp_path):
    book = sheetrs.Book(structure)
    before = book.stats().sheets["T"].used_cells
    saved = book.delete_rows("T!2").save(tmp_path / f"out{structure.suffix}")
    assert saved.stats().sheets["T"].used_cells < before


def test_the_book_that_saved_in_place_describes_the_file_as_it_is_now(structure):
    book = sheetrs.Book(structure)
    before = book.stats().sheets["T"].used_cells
    book.delete_rows("T!2").save()
    assert book.stats().sheets["T"].used_cells < before


def test_the_stats_are_typed_and_write_themselves_as_the_command_line_does(structure):
    stats = sheetrs.Book(structure).stats()
    assert isinstance(stats, sheetrs.Stats)
    document = stats.to_dict()
    assert document["total_sheets"] == stats.total_sheets
    assert [s["name"] for s in document["sheets"]] == [s.name for s in stats.sheets]
    assert json.loads(stats.to_json()) == document
    sheet = stats.sheets["F"]
    assert sheet.to_dict() == next(s for s in document["sheets"] if s["name"] == "F")
    assert json.loads(sheet.to_json()) == sheet.to_dict()
