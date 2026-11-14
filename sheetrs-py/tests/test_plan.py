"""The operations of a `Book`: they accumulate, and `save` is the cut that plans and runs them."""

import io
import shutil
import zipfile

import pytest

import sheetrs
from conftest import ROOT, make_xlsx
from test_find import TEXTS


def parts(path):
    """Every part of a workbook, by name: the bytes the zip holds."""
    with zipfile.ZipFile(path) as z:
        return {name: z.read(name) for name in z.namelist()}


def whole(book):
    """What a workbook says: the cells of every sheet with value and formula, and the names."""
    out = {}
    for name in book.stats().sheets.names():
        key = f"{name}!A1:H20"
        block = book.range(key, properties=("value", "formula"))[key]
        out[name] = (block.values, block.formulas)
    return out, book.stats().defined_names


# ---- they accumulate; nothing runs until save --------------------------------------------------


def test_an_operation_returns_the_same_book_so_the_calls_chain(structure):
    book = sheetrs.Book(structure)
    assert book.delete_rows("T!2") is book
    assert (
        book.delete_columns("U!B")
        .delete_sheets("V")
        .delete_names("Gone")
        .formula_to_value("F")
        .replace("a", "b")
        is book
    )
    assert "6 pending operations" in repr(book)


def test_nothing_is_read_or_written_until_save(structure):
    before = structure.read_bytes()
    mtime = structure.stat().st_mtime_ns
    book = sheetrs.Book(structure).delete_rows("T!2").delete_sheets("V")
    book.explain()
    repr(book)
    assert structure.read_bytes() == before
    assert structure.stat().st_mtime_ns == mtime


def test_save_with_a_path_writes_it_and_gives_a_book_on_it(structure, tmp_path):
    out_path = tmp_path / f"out{structure.suffix}"
    book = sheetrs.Book(structure).delete_rows("T!2")
    saved = book.save(out_path)
    assert isinstance(saved, sheetrs.Book) and saved is not book
    assert out_path.is_file()
    assert "pending" not in repr(saved)
    # The rows moved up: row 2 of the result is what was row 3.
    assert saved.range("T!A1:A3")["T!A1:A3"].values == [[1], [3], [4]]
    # The original file is as it was, and the book keeps its operations: it can be saved elsewhere.
    assert sheetrs.Book(structure).range("T!A2")["T!A2"].values == [[2]]
    assert "1 pending operation" in repr(book)
    other = book.save(tmp_path / f"other{structure.suffix}")
    assert parts(other.path if hasattr(other, "path") else tmp_path / f"other{structure.suffix}")


def test_save_without_a_path_replaces_the_original_and_the_book_is_the_file_now(structure):
    original = structure.read_bytes()
    book = sheetrs.Book(structure).delete_rows("T!2")
    result = book.save()
    assert result is book
    assert structure.read_bytes() != original
    assert "pending" not in repr(book)
    assert book.range("T!A1:A3")["T!A1:A3"].values == [[1], [3], [4]]
    # No temporary file is left next to it.
    assert [p.name for p in structure.parent.iterdir()] == [structure.name]


def test_a_failed_save_leaves_nothing_behind(structure, tmp_path):
    out_path = tmp_path / f"never{structure.suffix}"
    with pytest.raises(sheetrs.Error):
        sheetrs.Book(structure).delete_rows("NoSuchSheet!2").save(out_path)
    assert not out_path.exists()
    before = structure.read_bytes()
    with pytest.raises(sheetrs.Error):
        sheetrs.Book(structure).delete_rows("NoSuchSheet!2").save()
    assert structure.read_bytes() == before
    assert [p.name for p in tmp_path.iterdir()] == [structure.name]


def test_a_book_from_bytes_has_no_file_to_replace(structure, tmp_path):
    book = sheetrs.Book(structure.read_bytes())
    book.delete_rows("T!2")
    with pytest.raises(sheetrs.Error, match="path"):
        book.save()
    saved = book.save(tmp_path / f"fromBytes{structure.suffix}")
    assert saved.range("T!A2")["T!A2"].values == [[3]]


def test_there_must_be_something_to_save(structure, tmp_path):
    with pytest.raises(sheetrs.Error):
        sheetrs.Book(structure).save(tmp_path / "x.xlsx")


def test_the_destination_cannot_be_the_file_being_edited(structure):
    with pytest.raises(sheetrs.Error):
        sheetrs.Book(structure).delete_rows("T!2").save(structure)


def test_a_path_like_destination_is_accepted(structure, tmp_path):
    saved = sheetrs.Book(structure).delete_rows("T!2").save(tmp_path / f"p{structure.suffix}")
    assert saved.stats().sheets.names() == ["T", "U", "V", "F"]


# ---- what each operation does ----------------------------------------------------------------


def test_delete_sheets(structure, tmp_path):
    saved = sheetrs.Book(structure).delete_sheets(["U", "V"]).save(tmp_path / f"s{structure.suffix}")
    assert saved.stats().sheets.names() == ["T", "F"]


def test_delete_names(structure, tmp_path):
    saved = (
        sheetrs.Book(structure)
        .delete_names(["Gone", "NameOne"])
        .save(tmp_path / f"n{structure.suffix}")
    )
    assert set(saved.stats().defined_names) == {"NameTwo", "NameThree"}


def test_delete_rows_moves_the_rest_up_and_the_references_to_them_are_an_error(structure, tmp_path):
    saved = sheetrs.Book(structure).delete_rows("T!2").save(tmp_path / f"r{structure.suffix}")
    values = saved.range("T!A1:A4")["T!A1:A4"].values
    assert values == [[1], [3], [4], [5]]
    # F!A1 referred to T!A2: it is #REF! now.
    first = saved.range("F!A1", properties=("value", "formula"))["F!A1"]
    assert first.formulas[0][0] is not None
    assert "REF" in str(first.values[0][0]) or "REF" in first.formulas[0][0]


def test_delete_columns(structure, tmp_path):
    saved = sheetrs.Book(structure).delete_columns("U!A").save(tmp_path / f"c{structure.suffix}")
    assert saved.range("U!A1:A2")["U!A1:A2"].values == [[110], [120]]


def test_a_deletion_that_would_break_a_merged_area_is_refused(structure, tmp_path):
    out_path = tmp_path / f"m{structure.suffix}"
    with pytest.raises(sheetrs.Error, match="merged"):
        sheetrs.Book(structure).delete_columns("T!A").save(out_path)
    assert not out_path.exists()


def test_the_scopes_of_one_call_are_simultaneous(structure, tmp_path):
    saved = (
        sheetrs.Book(structure).delete_rows(["T!2", "T!4"]).save(tmp_path / f"r2{structure.suffix}")
    )
    assert saved.range("T!A1:A3")["T!A1:A3"].values == [[1], [3], [5]]


def test_each_operation_addresses_the_workbook_as_it_is_and_steps_are_saves(structure, tmp_path):
    # In one plan the rows are those of the workbook: row 2 twice is row 2.
    once = tmp_path / f"once{structure.suffix}"
    saved = sheetrs.Book(structure).delete_rows("T!2").delete_rows("T!2").save(once)
    assert saved.range("T!A1:A3")["T!A1:A3"].values == [[1], [3], [4]]
    # Step by step is a save between the steps: the second addresses what the first left.
    again = sheetrs.Book(once).delete_rows("T!2").save(tmp_path / f"seq{structure.suffix}")
    assert again.range("T!A1:A3")["T!A1:A3"].values == [[1], [4], [5]]


def test_formula_to_value(structure, tmp_path):
    book = sheetrs.Book(structure)
    assert book.range("F!A1:A2", "formula")["F!A1:A2"].formulas != [[None], [None]]
    saved = book.formula_to_value("F").save(tmp_path / f"v{structure.suffix}")
    block = saved.range("F!A1:A3", properties=("value", "formula"))["F!A1:A3"]
    assert block.values == [[2], [4], [5]]
    assert block.formulas == [[None], [None], [None]]
    # A scope can be a range of columns or of rows.
    columns = sheetrs.Book(structure).formula_to_value("T!C:C").save(
        tmp_path / f"v2{structure.suffix}"
    )
    assert columns.range("T!C1:C2", "formula")["T!C1:C2"].formulas == [[None], [None]]


def test_formula_to_value_leaves_an_empty_text_where_the_formula_gave_an_error(structure, tmp_path):
    """LibreOffice's Formula to Value does (`F!A18` is `=NameTwo`, an error the file caches): the cell
    is an empty text in both formats, not a `CellError`."""
    book = sheetrs.Book(structure)
    before = book.range("F!A18", properties=("value", "formula"))["F!A18"]
    assert isinstance(before.values[0][0], sheetrs.CellError)
    saved = book.formula_to_value("F").save(tmp_path / f"e{structure.suffix}")
    after = saved.range("F!A17:A18", properties=("value", "formula"))["F!A17:A18"]
    # The readers give each file what LibreOffice has: an ODS cell of an empty text
    # (`office:value-type="string"` with an empty paragraph) is an empty text, and an XLSX one (an empty
    # shared or inline string) is a blank cell.
    assert after.values == [[7], [None if structure.suffix == ".xlsx" else ""]]
    assert after.formulas == [[None], [None]]


def test_replace_reports_what_it_did(tmp_path):
    book = sheetrs.Book(make_xlsx(tmp_path / "t.xlsx", TEXTS))
    saved = book.replace("apple", "cherry").save(tmp_path / "t2.xlsx")
    assert saved.report.replaced == 5 and saved.report.skipped == []
    assert saved.range("S!A1:C1")["S!A1:C1"].values == [["cherry", 1.5, "cherry"]]
    assert saved.range("R!A1")["R!A1"].values == [["cherry"]]
    assert book.report is None


def test_replace_with_regular_expression_groups_and_a_scope(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    saved = (
        sheetrs.Book(source)
        .replace("(an)(a)", "$2$1", regex=True, match_case=True, scope="S!A3:C3")
        .save(tmp_path / "g.xlsx")
    )
    assert saved.range("S!A3:C3")["S!A3:C3"].values == [["Baanna", "x1y", "bandaan"]]
    assert saved.report.replaced == 2


def test_replace_options(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    saved = (
        sheetrs.Book(source)
        .replace("apple", "X", entire_cells=True, match_case=True, scope=["R"])
        .replace("ban*", "Y", wildcards=True, entire_cells=True, scope="S")
        .save(tmp_path / "o.xlsx")
    )
    assert saved.range("R!A1")["R!A1"].values == [["X"]]
    assert saved.range("S!A3:C3")["S!A3:C3"].values == [["Y", "x1y", "Y"]]
    with pytest.raises(ValueError, match="regex"):
        sheetrs.Book(source).replace("a", "b", regex=True, wildcards=True)
    with pytest.raises(ValueError, match="search_in"):
        sheetrs.Book(source).replace("a", "b", search_in="nope")


def test_replace_in_formulas(structure, tmp_path):
    saved = (
        sheetrs.Book(structure)
        .replace("A2", "A3", search_in="formulas", scope="F!A1:A1", entire_cells=False)
        .save(tmp_path / f"f{structure.suffix}")
    )
    assert saved.report.replaced == 1
    assert "A3" in saved.range("F!A1", "formula")["F!A1"].formulas[0][0]


# ---- the plan: the fewest passes, the same result --------------------------------------------


def test_the_operations_that_can_share_a_pass_do(structure):
    book = (
        sheetrs.Book(structure)
        .delete_rows("T!2")
        .delete_columns("U!B")
        .delete_sheets("V")
        .delete_names("Gone")
    )
    text = book.explain()
    # One pass of the file for all of them, then the engine calculating what they left stale.
    assert "pass 1 of 2" in text
    for line in ("delete rows", "delete columns", "delete sheets", "delete names"):
        assert line in text


def test_the_operations_that_cannot_share_a_pass_say_why_by_their_passes(structure):
    book = (
        sheetrs.Book(structure)
        .delete_rows("T!2")
        .formula_to_value("F")
        .replace("9", "8", scope="U")
    )
    text = book.explain()
    # Whatever order they were called in: the replacement, then the conversion and the deletion together,
    # written against the workbook as it is (the engine calculates the cells to convert while the pass writes
    # them as values: no calculation of its own comes first).
    assert "pass 1 of 3" in text and "pass 3 of 3" in text
    assert (
        text.index("replace")
        < text.index("formula to value")
        < text.index("delete rows")
        < text.index("recalculate every formula")
    )
    # And what the replacement and the deletion left stale is calculated again at the end.
    assert text.count("recalculate every formula") == 1


def test_a_conversion_that_uses_the_cache_has_the_results_calculated_first(structure):
    book = (
        sheetrs.Book(structure)
        .delete_rows("T!2")
        .formula_to_value("F")
        .replace("9", "8", scope="U")
    )
    text = book.explain(use_cache=True)
    assert "pass 1 of 4" in text and "pass 4 of 4" in text
    assert (
        text.index("replace")
        < text.index("recalculate every formula")
        < text.index("formula to value")
        < text.index("delete rows")
    )
    assert text.rindex("recalculate every formula") > text.index("delete rows")


SEQUENCE = [
    lambda b: b.delete_rows("T!2"),
    lambda b: b.delete_columns("U!B"),
    lambda b: b.delete_sheets("V"),
    lambda b: b.delete_names("Gone"),
    lambda b: b.replace("9", "8", scope="U"),
]
# The order the planner runs them in: the replacement, then the deletions.
PLANNERS_ORDER = [SEQUENCE[4], *SEQUENCE[:4]]


def one_plan_and_one_at_a_time(structure, tmp_path, operations, planners_order):
    """The operations pushed in `operations` order in one plan, and saved one at a time, each step
    reading the file the one before wrote, in `planners_order`."""
    suffix = structure.suffix
    plan = sheetrs.Book(structure)
    for operation in operations:
        operation(plan)
    one = plan.save(tmp_path / f"one{suffix}")
    current = structure
    for i, operation in enumerate(planners_order):
        step = tmp_path / f"step{i}{suffix}"
        operation(sheetrs.Book(current)).save(step)
        current = step
    return one, tmp_path / f"one{suffix}", current


def test_one_plan_gives_the_bytes_of_the_operations_saved_one_at_a_time_in_the_planners_order(
    structure, tmp_path
):
    one, one_path, last = one_plan_and_one_at_a_time(
        structure, tmp_path, SEQUENCE, PLANNERS_ORDER
    )
    assert parts(one_path) == parts(last)
    assert whole(one) == whole(sheetrs.Book(last))


def test_the_same_with_formulas_made_values(structure, tmp_path):
    operations = [*SEQUENCE[:4], lambda b: b.formula_to_value("F"), SEQUENCE[4]]
    planners_order = [SEQUENCE[4], lambda b: b.formula_to_value("F"), *SEQUENCE[:4]]
    _, one_path, last = one_plan_and_one_at_a_time(
        structure, tmp_path, operations, planners_order
    )
    assert parts(one_path) == parts(last)


def test_the_order_the_operations_were_called_in_changes_nothing(structure, tmp_path):
    operations = [*SEQUENCE, lambda b: b.formula_to_value("F")]
    outputs = []
    for i, order in enumerate([operations, operations[::-1], operations[2:] + operations[:2]]):
        book = sheetrs.Book(structure)
        for operation in order:
            operation(book)
        outputs.append(book.to_bytes())
    assert outputs[0] == outputs[1] == outputs[2]


def test_two_deletions_of_the_same_row_are_that_row_once_and_adjacent_rows_are_one_band(
    structure,
):
    twice = sheetrs.Book(structure).delete_rows("T!3").delete_rows("T!3").to_bytes()
    once = sheetrs.Book(structure).delete_rows("T!3").to_bytes()
    assert twice == once
    apart = sheetrs.Book(structure).delete_rows("T!2").delete_rows("T!3").to_bytes()
    together = sheetrs.Book(structure).delete_rows("T!2:3").to_bytes()
    assert apart == together


def test_a_conversion_is_of_the_workbook_as_it_is_and_after_a_deletion_it_is_another_step(
    structure, tmp_path
):
    """The values LibreOffice gives after a deletion: the sum over a column reads one row less once a
    row went, so `SUM(T!A2:A3)` is 3 and not the 5 the file cached. Pushed with the deletion the
    conversion reads the file as it is (5); to convert what the deletion leaves, save it and convert
    the next step: its results are stale and are calculated again."""
    same = tmp_path / f"same{structure.suffix}"
    sheetrs.Book(structure).delete_rows("T!2").formula_to_value("F").save(same)
    block = sheetrs.Book(same).range("F!A2:A12", properties=("value",))["F!A2:A12"]
    assert [row[0] for row in block.values][3] == 5
    deleted = tmp_path / f"deleted{structure.suffix}"
    path = tmp_path / f"out{structure.suffix}"
    sheetrs.Book(structure).delete_rows("T!2").save(deleted)
    sheetrs.Book(deleted).formula_to_value("F").save(path)
    block = sheetrs.Book(path).range("F!A2:A12", properties=("value", "formula"))["F!A2:A12"]
    assert [row[0] for row in block.values] == [4, 5, 6, 3, 7, 9, 30, 33, 11, 53, 521]
    # No formula is left in what was converted.
    assert all(row[0] is None for row in block.formulas)


def test_a_conversion_the_engine_cannot_calculate_says_why_and_writes_nothing(tmp_path):
    """The `book` workbook has an external reference, which the engine does not calculate: the stale
    results of a saved deletion are not converted in the next step."""
    book = ROOT / "sheetrs-edit/tests/assets/book/book.xlsx"
    deleted = tmp_path / "deleted.xlsx"
    # The deletion is saved marked, with the warning: the engine cannot calculate what it left stale.
    with pytest.warns(sheetrs.RecalculationWarning, match="not calculated"):
        sheetrs.Book(book).delete_rows("Data!10:11").save(deleted)
    out = tmp_path / "out.xlsx"
    with pytest.raises(sheetrs.Error, match="could not be calculated again"):
        sheetrs.Book(deleted).formula_to_value("Data!D:E").save(out)
    assert not out.exists()


def test_an_operation_on_a_sheet_that_is_also_deleted_is_not_an_error(structure):
    book = sheetrs.Book(structure).delete_sheets("U").delete_rows("U!1")
    assert "delete sheets U" in book.explain()


def test_explain_says_when_there_is_nothing_to_do(structure):
    assert "nothing to do" in sheetrs.Book(structure).explain()


# ---- what is wrong is said where it is written ----------------------------------------------


def test_a_malformed_scope_is_an_invalid_range_where_it_is_written(structure):
    book = sheetrs.Book(structure)
    with pytest.raises(sheetrs.InvalidRangeError, match="T!B"):
        book.delete_rows("T!B")  # a column range where rows are deleted
    with pytest.raises(ValueError):
        book.delete_columns("T!3")
    # Nothing was added by the calls that failed.
    assert "pending" not in repr(book)


def test_an_operation_needs_its_targets(structure):
    """The array of targets is the argument: a call without it is a `TypeError`, an empty array a
    call that does nothing."""
    book = sheetrs.Book(structure)
    for name in ("delete_sheets", "delete_names", "delete_rows", "delete_columns", "formula_to_value"):
        with pytest.raises(TypeError):
            getattr(book, name)()
    assert "pending" not in repr(book)


# ---- a query waits for the operations to be saved ---------------------------------------------


QUERIES = {
    "range": lambda b: b.range("T!A1"),
    "find": lambda b: b.find("1"),
    "lint": lambda b: b.lint(),
    "stats": lambda b: b.stats(),
    "sheet_names": lambda b: b.stats().sheets.names(),
    "defined_names": lambda b: b.stats().defined_names,
}


@pytest.mark.parametrize("name", sorted(QUERIES))
def test_a_query_with_operations_pending_raises(structure, name):
    book = sheetrs.Book(structure).delete_rows("T!2").delete_sheets("V")
    with pytest.raises(sheetrs.PendingOperationsError) as error:
        QUERIES[name](book)
    message = str(error.value)
    assert "2 pending operations" in message
    # It says both ways out: save first, or ask before adding the operations.
    assert "save()" in message and "before adding" in message
    assert isinstance(error.value, sheetrs.Error)


def test_a_book_with_operations_still_says_what_it_is_and_what_it_will_do(structure):
    book = sheetrs.Book(structure).delete_rows("T!2")
    assert book.format == structure.suffix[1:]
    assert "pending" in repr(book)
    assert "delete rows" in book.explain()


def test_a_book_is_queried_before_its_operations_and_the_saved_book_after_them(structure, tmp_path):
    book = sheetrs.Book(structure)
    before = book.range("T!A2")["T!A2"].values
    book.delete_rows("T!2")
    saved = book.save(tmp_path / f"q{structure.suffix}")
    assert before == [[2]]
    assert saved.range("T!A2")["T!A2"].values == [[3]]
    assert book.report is None and saved.report.to_dict() == {"replaced": 0, "written": 0, "skipped": [], "diagnostics": []}


# ---- every operation takes the array of its targets ----------------------------------------------


def test_an_array_of_targets_is_the_sweep_and_one_str_is_one_target(structure, tmp_path):
    """One call with the array is the same plan as one call per target (the ranges of a call are
    simultaneous, so the rows that go together are given in one call)."""
    suffix = structure.suffix
    swept = (
        sheetrs.Book(structure)
        .delete_sheets(["U", "V"])
        .delete_names(("Gone", "NameOne"))
        .formula_to_value(["F", "T!C:C"])
        .save(tmp_path / f"swept{suffix}")
    )
    chained = (
        sheetrs.Book(structure)
        .delete_sheets("U")
        .delete_sheets("V")
        .delete_names("Gone")
        .delete_names("NameOne")
        .formula_to_value("F")
        .formula_to_value("T!C:C")
        .save(tmp_path / f"chained{suffix}")
    )
    assert whole(swept) == whole(chained)
    assert swept.stats().sheets.names() == ["T", "F"]


def test_any_iterable_of_str_is_an_array(structure, tmp_path):
    rows = sheetrs.Book(structure).delete_rows(r for r in ("T!2", "T!4"))
    saved = rows.save(tmp_path / f"g{structure.suffix}")
    assert saved.range("T!A1:A3")["T!A1:A3"].values == [[1], [3], [5]]
    as_tuple = sheetrs.Book(structure).delete_rows(("T!2", "T!4"))
    assert as_tuple.explain() == rows.explain()


def test_an_item_that_is_not_a_str_is_a_type_error_and_adds_nothing(structure):
    book = sheetrs.Book(structure)
    for call in (
        lambda: book.delete_rows(["T!2", 3]),
        lambda: book.delete_rows(3),
        lambda: book.delete_columns([b"T!B"]),
        lambda: book.delete_sheets(None),
        lambda: book.delete_names(["Gone", None]),
        lambda: book.formula_to_value([1]),
    ):
        with pytest.raises(TypeError, match="iterable of str"):
            call()
    assert "nothing to do" in book.explain()


def test_a_malformed_target_in_an_array_names_it_and_adds_nothing(structure):
    book = sheetrs.Book(structure)
    with pytest.raises(sheetrs.InvalidRangeError, match="T!B"):
        book.delete_rows(["T!2", "T!B"])
    assert "nothing to do" in book.explain()


def test_an_empty_array_is_a_call_that_does_nothing(structure):
    book = sheetrs.Book(structure)
    assert (
        book.delete_sheets([])
        .delete_names(())
        .formula_to_value([])
        .delete_rows(iter(()))
        .delete_columns([])
        is book
    )
    assert "nothing to do" in book.explain()
    book.delete_rows("T!2")
    book.delete_rows([])
    assert book.explain().count("delete rows") == 1


def test_the_scope_and_the_outputs_are_arrays_too(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    saved = (
        sheetrs.Book(source)
        .replace("apple", "X", entire_cells=True, match_case=True, scope=(s for s in ["R", "S!A2"]))
        .save(tmp_path / "s.xlsx")
    )
    assert saved.range("R!A1")["R!A1"].values == [["X"]]


# ---- replace: pairs for a sweep ---------------------------------------------------------------------


def test_a_mapping_of_pairs_is_one_replace_per_pair_in_order(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    sweep = sheetrs.Book(source).replace({"apple": "cherry", "cherry": "plum"})
    chained = sheetrs.Book(source).replace("apple", "cherry").replace("cherry", "plum")
    assert sweep.explain() == chained.explain()
    assert sweep.explain().count("replace") == 2
    one = sweep.save(tmp_path / "sweep.xlsx")
    two = chained.save(tmp_path / "chained.xlsx")
    assert parts(tmp_path / "sweep.xlsx") == parts(tmp_path / "chained.xlsx")
    assert one.range("R!A1")["R!A1"].values == [["plum"]]
    assert two.report.replaced == one.report.replaced


def test_the_order_of_the_pairs_matters(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    forward = sheetrs.Book(source).replace([("apple", "cherry"), ("cherry", "plum")])
    backward = sheetrs.Book(source).replace([("cherry", "plum"), ("apple", "cherry")])
    forward.save(tmp_path / "f.xlsx")
    backward.save(tmp_path / "b.xlsx")
    assert parts(tmp_path / "f.xlsx") != parts(tmp_path / "b.xlsx")
    assert sheetrs.Book(tmp_path / "b.xlsx").range("R!A1")["R!A1"].values == [["cherry"]]


def test_the_options_and_the_scope_are_shared_by_the_pairs(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    saved = (
        sheetrs.Book(source)
        .replace([("apple", "X"), ("pear", "Y")], entire_cells=True, match_case=True, scope=["R"])
        .save(tmp_path / "o.xlsx")
    )
    assert saved.range("R!A1:B2")["R!A1:B2"].values == [["X", None], [None, "Y"]]
    assert saved.range("S!A1")["S!A1"].values == [["Apple"]]  # outside the scope


def test_the_pairs_are_pairs_of_str(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    book = sheetrs.Book(source)
    for text, replacement in (
        ("a", None),  # a str needs its replacement
        ({"a": "b"}, "c"),  # a mapping carries its own
        ([("a", "b", "c")], None),
        ([("a", 1)], None),
        ([("a",)], None),
        (["ab"], None),  # a str is not a pair
        (3, None),
    ):
        with pytest.raises(TypeError, match="replace"):
            book.replace(text, replacement)
    assert "nothing to do" in book.explain()
    assert book.replace({}) is book and book.replace([]) is book
    assert "nothing to do" in book.explain()


def test_a_mistake_in_a_later_pair_adds_no_pair(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    book = sheetrs.Book(source)
    with pytest.raises(TypeError):
        book.replace([("apple", "cherry"), ("pear", 2)])
    assert "nothing to do" in book.explain()


# ---- to_bytes: the cut to memory ------------------------------------------------------------


def a_plan(book):
    return book.delete_rows("T!2").formula_to_value("F").replace("9", "8", scope="U")


def test_to_bytes_gives_the_bytes_save_writes_for_the_same_plan(structure, tmp_path):
    book = a_plan(sheetrs.Book(structure))
    data = book.to_bytes()
    assert isinstance(data, bytes)
    saved = tmp_path / f"saved{structure.suffix}"
    a_plan(sheetrs.Book(structure)).save(saved)
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        in_memory = {name: z.read(name) for name in z.namelist()}
    assert in_memory == parts(saved)


def test_a_book_made_from_the_bytes_reads_what_the_saved_one_reads(structure, tmp_path):
    book = a_plan(sheetrs.Book(structure))
    again = sheetrs.Book(book.to_bytes())
    saved = a_plan(sheetrs.Book(structure)).save(tmp_path / f"saved{structure.suffix}")
    assert whole(again) == whole(saved)


def test_to_bytes_keeps_the_operations_and_can_be_repeated(structure):
    book = sheetrs.Book(structure).delete_rows("T!2")
    first = book.to_bytes()
    assert "delete rows" in book.explain()
    assert book.to_bytes() == first
    with pytest.raises(sheetrs.PendingOperationsError):
        book.range("T!A1:B2")


def test_a_book_made_from_bytes_gives_the_edited_workbook_as_bytes_and_leaves_the_original(structure):
    data = structure.read_bytes()
    book = sheetrs.Book(data).delete_sheets("V")
    out = sheetrs.Book(book.to_bytes())
    assert "V" not in out.stats().sheets.names()
    assert "V" in sheetrs.Book(data).stats().sheets.names()


def test_to_bytes_with_no_operations_is_the_workbook_as_it_is(structure):
    data = structure.read_bytes()
    assert sheetrs.Book(structure).to_bytes() == data
    assert sheetrs.Book(data).to_bytes() == data


def test_to_bytes_sets_the_report(structure):
    book = sheetrs.Book(structure).replace("9", "8", scope="U")
    assert book.report is None
    book.to_bytes()
    assert book.report is not None and book.report.replaced >= 0


def test_to_bytes_leaves_no_file_behind(structure, tmp_path):
    before = sorted(p.name for p in structure.parent.iterdir())
    sheetrs.Book(structure).delete_rows("T!2").to_bytes()
    assert sorted(p.name for p in structure.parent.iterdir()) == before


def test_to_bytes_of_a_file_that_went(structure):
    book = sheetrs.Book(structure)
    structure.unlink()
    with pytest.raises(OSError):  # no operations: the file is read as it is
        book.to_bytes()
    with pytest.raises(sheetrs.Error):  # operations: the editor cannot read it
        book.delete_rows("T!2").to_bytes()


def test_hide_and_show_rows_columns_and_sheets(structure, tmp_path):
    hidden = tmp_path / f"h{structure.suffix}"
    sheetrs.Book(structure).hide_rows("T!3").hide_columns("T!B").hide_sheets("V").save(hidden)
    # What was hidden is shown again by the opposite operations.
    shown = tmp_path / f"s{structure.suffix}"
    sheetrs.Book(hidden).show_rows("T!3").show_columns("T!B").show_sheets("V").save(shown)
    assert sheetrs.Book(shown).range("T!A1:A4")["T!A1:A4"].values == sheetrs.Book(structure).range("T!A1:A4")["T!A1:A4"].values


def test_replace_with_the_empty_text_replaces_nothing_in_every_mode(tmp_path):
    # LibreOffice: an empty search string matches nothing (replaceAll changes no cell), in every mode.
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    for options in [{}, {"regex": True}, {"wildcards": True}, {"entire_cells": True}, {"regex": True, "entire_cells": True, "match_case": True}]:
        saved = sheetrs.Book(source).replace("", "X", **options).save(tmp_path / "e.xlsx")
        assert saved.report.replaced == 0 and saved.report.skipped == []
        assert saved.range("S!A1:C1")["S!A1:C1"].values == [["Apple", 1.5, "APPLE"]]
    # A sheet that is not there is no error either: nothing is read.
    saved = sheetrs.Book(source).replace("", "X", scope="Nowhere").save(tmp_path / "n.xlsx")
    assert saved.report.replaced == 0


def test_replace_with_an_empty_mapping_key_replaces_nothing(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    saved = sheetrs.Book(source).replace({"": "X", "apple": "cherry"}).save(tmp_path / "m.xlsx")
    assert saved.report.replaced == 5
    assert saved.range("S!A1:C1")["S!A1:C1"].values == [["cherry", 1.5, "cherry"]]


def test_replace_reads_the_patterns_that_match_the_empty_text(tmp_path):
    source = make_xlsx(tmp_path / "t.xlsx", TEXTS)
    # `$` adds at the end of the text; `^` and `x*` change nothing where they match no characters.
    saved = sheetrs.Book(source).replace("$", "!", regex=True, scope="S!A1:A3").save(tmp_path / "d.xlsx")
    assert saved.range("S!A1:A3")["S!A1:A3"].values == [["Apple!"], ["apple pie!"], ["Banana!"]]
    assert saved.report.replaced == 3
    saved = sheetrs.Book(source).replace("^", "!", regex=True).save(tmp_path / "c.xlsx")
    assert saved.report.replaced == 0
