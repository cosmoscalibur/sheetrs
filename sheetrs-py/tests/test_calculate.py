"""`Book.calculate`, a step of the plan, with `value_ranges()` and `save(json=...)` for its value ranges, and
`sheetrs.evaluate`."""

import json
import shutil
import warnings
import zipfile

import pytest

import sheetrs
from conftest import ROOT, STRUCTURE, make_xlsx


def is_ref_error(value):
    return isinstance(value, sheetrs.CellError) and value.code in {"#REF!", "Err:524"}


# ---- the value ranges in memory: calculate(ranges=...).to_dict() ---------------------------------


def value_ranges(book, **kwargs):
    """The value ranges of a calculation, in memory."""
    return book.calculate(**kwargs).value_ranges().to_dict()


def test_major_columns_lists_each_range_as_columns_and_rows_is_the_default(structure):
    rows = value_ranges(sheetrs.Book(structure), ranges="F!A1:A3")
    columns = value_ranges(sheetrs.Book(structure), ranges="F!A1:A3", major="columns")
    assert rows == {"F!A1:A3": [[2], [4], [5]]}
    assert columns == {"F!A1:A3": [[2, 4, 5]]}
    block = sheetrs.Book(structure).calculate(ranges="F!A1:A3", major="columns").value_ranges()["F!A1:A3"]
    assert block.major == "columns"
    with pytest.raises(ValueError, match="major"):
        sheetrs.Book(structure).calculate(ranges="F!A1", major="diagonal")


def test_the_values_of_the_value_ranges(structure):
    result = value_ranges(sheetrs.Book(structure), ranges=["F!A1:A3", "T!C1"])
    assert result == {"F!A1:A3": [[2], [4], [5]], "T!C1": [[1]]}
    assert list(result) == ["F!A1:A3", "T!C1"]


def test_one_output_can_be_a_text(structure):
    assert value_ranges(sheetrs.Book(structure), ranges="F!A1:A2") == {"F!A1:A2": [[2], [4]]}


def test_without_ranges_they_are_the_formula_cells_of_the_workbook(structure):
    result = value_ranges(sheetrs.Book(structure))
    assert result["T!C1:C4"][:2] == [[1], [40]]
    assert "F!A1:A21" in result


def test_the_inputs_are_written_before_calculating(structure):
    book = sheetrs.Book(structure)
    result = value_ranges(book, inputs={"T!A1": 100}, ranges=["T!C1", "F!A5"])
    assert result == {"T!C1": [[100]], "F!A5": [[5]]}
    # Nothing is written: value_ranges() runs the operations and writes the workbook nowhere.
    assert sheetrs.Book(structure).range("T!A1")["T!A1"].values == [[1]]


def test_an_input_is_a_number_a_text_a_boolean_or_a_list_of_rows(structure):
    # The operations accumulate on a Book: each calculation is asked of a Book of its own.
    assert value_ranges(sheetrs.Book(structure), inputs={"T!A1": 2.5}, ranges="T!C1") == {"T!C1": [[2.5]]}
    rows = {"T!A1:B2": [[1, 10], [2, 20]]}
    assert value_ranges(sheetrs.Book(structure), inputs=rows, ranges="T!C2") == {"T!C2": [[40]]}
    with pytest.raises(TypeError, match="an input"):
        sheetrs.Book(structure).calculate(inputs={"T!A1": object()}, ranges="T!C1")


def test_an_error_in_the_values_is_a_cell_error(tmp_path):
    path = make_xlsx(
        tmp_path / "e.xlsx",
        {"S": {"A1": 1, "B1": ("=FOO(A1)", "#NAME?"), "C1": ("=A1/0", ("#DIV/0!")), "D1": ("=A1+1", 2)}},
    )
    values = value_ranges(sheetrs.Book(path), ranges="S!B1:D1")["S!B1:D1"][0]
    assert values == [sheetrs.CellError("#NAME?"), sheetrs.CellError("#DIV/0!"), 2]
    assert isinstance(values[0], sheetrs.CellError)


def test_what_is_not_an_output_or_an_input_is_an_error(structure):
    # Each case on a Book of its own: a calculation that was asked stays in the plan.
    # An empty cell outside the outputs is read as the workbook has it: empty.
    assert value_ranges(sheetrs.Book(structure), ranges="F!B1") == {"F!B1": [[None]]}
    with pytest.raises(sheetrs.Error, match="not an input"):
        value_ranges(sheetrs.Book(structure), inputs={"T!C1": 1}, ranges="F!A1")
    with pytest.raises(sheetrs.Error, match="Nowhere"):
        value_ranges(sheetrs.Book(structure), inputs={"Nowhere!A1": 1}, ranges="F!A1")
    with pytest.raises(sheetrs.Error):
        value_ranges(sheetrs.Book(structure), ranges=["NoSuchName"])
    with pytest.raises(TypeError, match="ranges"):
        sheetrs.Book(structure).calculate(ranges=3)


def test_to_dict_is_a_terminal_that_runs_the_pending_operations(structure):
    """`value_ranges()` is allowed with operations pending, as `save()` is: it runs them in memory."""
    book = sheetrs.Book(structure).delete_rows("T!2").calculate(ranges="F!A5:A6")
    # The calculation is of the workbook as it is: the deletion is written against it too and runs
    # after (the values are in the cells the workbook has, F!A5 is SUM(T!A2:A3) = 5).
    assert book.value_ranges().to_dict() == {"F!A5:A6": [[5], [9]]}
    # The Book still has its operations: nothing was written and nothing cleared.
    assert "2 pending operations" in repr(book)


def test_to_dict_without_a_calculate_step_is_an_error(structure):
    with pytest.raises(sheetrs.Error, match="no calculate step"):
        sheetrs.Book(structure).value_ranges()
    with pytest.raises(sheetrs.Error, match="no calculate step"):
        sheetrs.Book(structure).delete_rows("T!2").value_ranges()


def test_a_function_the_engine_does_not_have_cannot_give_its_value_ranges(tmp_path):
    path = make_xlsx(
        tmp_path / "u.xlsx", {"S": {"A1": ('=WEBSERVICE("x")', "x"), "B1": 1}}
    )
    with pytest.raises(sheetrs.UnavailableFunctionError, match="WEBSERVICE"):
        value_ranges(sheetrs.Book(path), ranges="S!A1")
    with pytest.raises(sheetrs.Error):  # it is an Error too
        value_ranges(sheetrs.Book(path), ranges="S!A1")


# ---- the value ranges as JSON: save(json=...) ------------------------------------------------------


def test_save_writes_the_value_ranges_as_json_besides_the_workbook(structure, tmp_path):
    out = tmp_path / f"o{structure.suffix}"
    answer = tmp_path / "answer.json"
    book = sheetrs.Book(structure).delete_rows("T!2").calculate(ranges="F!A5:A6")
    book.save(out, json=answer)
    assert json.loads(answer.read_text())["valueRanges"][0]["values"] == [[5], [9]]
    assert out.exists()
    # The deletion was made too: the row 2 of T is gone.
    assert sheetrs.Book(out).range("T!A1:A3")["T!A1:A3"].values == [[1], [3], [4]]


def test_save_json_without_a_calculation_is_an_error_and_writes_nothing(structure, tmp_path):
    out = tmp_path / f"o{structure.suffix}"
    answer = tmp_path / "answer.json"
    with pytest.raises(sheetrs.Error, match="no calculate step"):
        sheetrs.Book(structure).delete_rows("T!2").save(out, json=answer)
    assert not out.exists() and not answer.exists()


def test_save_json_in_place(structure, tmp_path):
    answer = tmp_path / "answer.json"
    book = sheetrs.Book(structure).calculate({"T!A1": 7}, ranges="T!C1")
    assert book.save(json=answer) is book
    assert json.loads(answer.read_text())["valueRanges"][0]["values"] == [[7]]
    assert sheetrs.Book(structure).range("T!A1")["T!A1"].values == [[7]]


# ---- as an operation: calculate() -----------------------------------------------------------


def test_calculate_is_an_operation_that_returns_the_book(structure):
    book = sheetrs.Book(structure)
    assert book.calculate() is book
    assert book.calculate({"T!A1": 3}) is book
    assert book.calculate(ranges="F!A1") is book
    assert "3 pending operations" in repr(book)
    assert book.explain().count("calculate every formula") == 3


def test_the_engine_recalculates_what_a_saved_step_left(structure, tmp_path):
    # Two steps: the deletion is saved, and the calculation is of the workbook it left.
    deleted = sheetrs.Book(structure).delete_rows("T!2").save(tmp_path / f"d{structure.suffix}")
    saved = deleted.calculate().save(tmp_path / f"c{structure.suffix}")
    values = saved.range("F!A1:A3")["F!A1:A3"].values
    # F!A1 referred to the row that is gone: an error. The cells that referred to others moved with
    # them and have the value of what they refer to now.
    assert is_ref_error(values[0][0])
    assert values[1:] == [[4], [5]]
    assert saved.range("T!A1:C3")["T!A1:C3"].values[1][0] == 3


def test_the_inputs_are_written_and_what_depends_on_them_is_calculated(structure, tmp_path):
    saved = (
        sheetrs.Book(structure)
        .calculate(inputs={"T!A1": 100})
        .save(tmp_path / f"i{structure.suffix}")
    )
    block = saved.range("T!A1:C1", properties=("value", "formula"))["T!A1:C1"]
    assert block.values == [[100, 10, 100]]
    assert block.formulas[0][2] is not None  # the formula is still there


def test_a_calculation_with_named_ranges_saves_the_workbook_with_the_inputs_written(
    structure, tmp_path
):
    book = sheetrs.Book(structure).calculate(inputs={"T!A1": 100}, ranges="T!C1")
    saved = book.save(tmp_path / f"n{structure.suffix}")
    assert saved.range("T!A1:C1")["T!A1:C1"].values == [[100, 10, 100]]
    assert book.value_ranges().to_dict() == {"T!C1": [[100]]}


def test_the_operations_after_a_calculation_address_what_it_leaves(structure, tmp_path):
    saved = (
        sheetrs.Book(structure)
        .calculate(inputs={"T!A1": 100})
        .delete_rows("T!1")
        .save(tmp_path / f"o{structure.suffix}")
    )
    assert saved.range("T!A1:A2")["T!A1:A2"].values == [[2], [3]]


def test_the_calculation_runs_before_the_deletions_whatever_the_order_they_were_called_in(
    structure, tmp_path
):
    outputs = []
    for i, order in enumerate(("after", "before")):
        book = sheetrs.Book(structure)
        if order == "after":
            book.delete_rows("T!1").calculate().delete_columns("U!B")
        else:
            book.calculate().delete_columns("U!B").delete_rows("T!1")
        text = book.explain()
        assert text.index("calculate") < text.index("delete rows") < text.index("delete columns")
        outputs.append(book.save(tmp_path / f"b{i}{structure.suffix}").to_bytes())
    assert outputs[0] == outputs[1]


def test_calculating_in_place(structure):
    book = sheetrs.Book(structure).calculate({"T!A1": 7})
    assert book.save() is book
    assert sheetrs.Book(structure).range("T!A1")["T!A1"].values == [[7]]
    assert "pending" not in repr(book)


def test_a_calculation_with_nothing_else_is_a_plan_of_one_step(structure, tmp_path):
    saved = sheetrs.Book(structure).calculate().save(tmp_path / f"n{structure.suffix}")
    assert saved.range("F!A1:A3")["F!A1:A3"].values == [[2], [4], [5]]


# ---- the mark to recalculate when the workbook is opened ---------------------------------------


EXTERNAL = ROOT / "sheetrs-edit/tests/assets/book/book.xlsx"


def marked(path):
    """Whether an XLSX asks the application that opens it to recalculate every formula."""
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/workbook.xml").decode()
    return 'fullCalcOnLoad="1"' in xml or 'fullCalcOnLoad="true"' in xml


def test_a_workbook_is_never_marked_unless_the_engine_cannot_calculate_it(tmp_path):
    source = tmp_path / "s.xlsx"
    shutil.copy(STRUCTURE / "book.xlsx", source)
    # What a deletion leaves stale is calculated again by the engine: no mark, whatever the order of
    # the calls, with or without a calculation.
    for name, build in {
        "calculated": lambda b: b.calculate(),
        "deleted_then_calculated": lambda b: b.delete_rows("T!2").calculate(),
        "calculated_then_deleted": lambda b: b.calculate().delete_rows("T!2"),
        "deleted": lambda b: b.delete_rows("T!2"),
    }.items():
        out = tmp_path / f"{name}.xlsx"
        build(sheetrs.Book(source)).save(out)
        assert not marked(out), name
    # A workbook that comes marked is not once the engine calculated it.
    marked_source = tmp_path / "marked.xlsx"
    with zipfile.ZipFile(source) as archive, zipfile.ZipFile(marked_source, "w") as target:
        for item in archive.infolist():
            data = archive.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = data.replace(b"<calcPr ", b'<calcPr fullCalcOnLoad="1" ', 1)
            target.writestr(item, data)
    assert marked(marked_source)
    again = tmp_path / "again.xlsx"
    sheetrs.Book(marked_source).calculate().save(again)
    assert not marked(again)


def test_a_workbook_the_engine_cannot_calculate_is_saved_marked_with_a_warning(tmp_path):
    out = tmp_path / "out.xlsx"
    with pytest.warns(sheetrs.RecalculationWarning, match="spreadsheet program that opens") as record:
        saved = sheetrs.Book(EXTERNAL).delete_rows("Data!10:11").calculate().save(out)
    assert len(record) == 1
    assert marked(out)
    assert [d.message for d in saved.report.diagnostics] == [str(record[0].message)]
    assert isinstance(sheetrs.RecalculationWarning("x"), UserWarning)


def test_a_calculation_the_engine_can_run_gives_no_warning(structure, tmp_path):
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        sheetrs.Book(structure).delete_rows("T!2").calculate().save(tmp_path / f"w{structure.suffix}")


def test_the_warning_can_be_an_error_and_the_workbook_is_written_all_the_same(tmp_path):
    out = tmp_path / "out.xlsx"
    with warnings.catch_warnings():
        warnings.simplefilter("error", sheetrs.RecalculationWarning)
        with pytest.raises(sheetrs.RecalculationWarning):
            sheetrs.Book(EXTERNAL).calculate().save(out)
    assert out.exists() and marked(out)


def test_the_value_ranges_of_a_workbook_the_engine_cannot_calculate_are_an_error_not_a_warning(tmp_path):
    out = tmp_path / "out.xlsx"
    answer = tmp_path / "answer.json"
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        with pytest.raises(sheetrs.Error, match="could not be calculated"):
            sheetrs.Book(EXTERNAL).calculate().save(out, json=answer)
        with pytest.raises(sheetrs.Error, match="could not be calculated"):
            sheetrs.Book(EXTERNAL).calculate().value_ranges()
    assert not out.exists() and not answer.exists()


# ---- evaluate -------------------------------------------------------------------------------


def test_evaluate_a_formula():
    assert sheetrs.evaluate("=SUM(1,2,3)") == 6
    assert sheetrs.evaluate("=1.5*2") == 3
    assert isinstance(sheetrs.evaluate("=1.5*2"), int)
    assert sheetrs.evaluate('="a"&"b"') == "ab"
    assert sheetrs.evaluate("=TRUE()") is True
    assert sheetrs.evaluate("=1/3") == pytest.approx(1 / 3)


def test_evaluate_an_error_is_a_cell_error():
    assert sheetrs.evaluate("=1/0") == sheetrs.CellError("#DIV/0!")
    assert sheetrs.evaluate("=FOO(1)") == sheetrs.CellError("#NAME?")
    assert sheetrs.evaluate("=SQRT(-1)").code == "Err:502"


def test_evaluate_with_the_settings_of_a_locale():
    assert sheetrs.evaluate('=TEXT(1234.5,"#,##0.00")') == "1,234.50"
    assert sheetrs.evaluate('=TEXT(1234.5,"#,##0.00")', locale="es-ES") == "1.234,50"
    with pytest.raises(sheetrs.Error, match="locale"):
        sheetrs.evaluate("=1", locale="zz-notalocale")


def test_evaluate_a_formula_the_engine_cannot_run():
    with pytest.raises(sheetrs.UnavailableFunctionError):
        sheetrs.evaluate('=WEBSERVICE("x")')
    # A formula that is not valid is the error LibreOffice shows for it, not an exception.
    assert isinstance(sheetrs.evaluate("=1+"), sheetrs.CellError)


# ---- the typed results ---------------------------------------------------------------------


def test_the_value_ranges_are_a_mapping_of_blocks_that_writes_itself_as_dict_and_json(structure):
    answer = sheetrs.Book(structure).calculate(ranges=["F!A5:A6", "F!A1"]).value_ranges()
    assert isinstance(answer, sheetrs.ValueRanges)
    assert len(answer) == 2 and list(answer) == answer.keys() == ["F!A5:A6", "F!A1"]
    assert "F!A5:A6" in answer and "F!A9" not in answer
    block = answer["F!A5:A6"]
    assert isinstance(block, sheetrs.Block) and block.values == [[5], [9]]
    assert [name for name, _ in answer.items()] == answer.keys()
    assert [b.range for b in answer.values()] == ["F!A5:A6", "F!A1"]
    assert answer.to_dict() == {"F!A5:A6": [[5], [9]], "F!A1": [[2]]}
    document = json.loads(answer.to_json())
    assert [r["range"] for r in document["valueRanges"]] == ["F!A5:A6", "F!A1"]
    with pytest.raises(KeyError):
        answer["F!A9"]


def test_the_json_of_save_is_the_json_of_the_value_ranges(structure, tmp_path):
    book = sheetrs.Book(structure).calculate(ranges="F!A5:A6")
    answer = tmp_path / "answer.json"
    book.save(tmp_path / f"o{structure.suffix}", json=answer)
    assert answer.read_text() == book.value_ranges().to_json() + "\n"


def test_the_report_of_a_save_is_typed_and_writes_itself(structure, tmp_path):
    saved = sheetrs.Book(structure).replace("1", "2").save(tmp_path / f"o{structure.suffix}")
    report = saved.report
    assert isinstance(report, sheetrs.Report)
    assert report.replaced == len(report.to_dict()["skipped"]) + report.replaced >= 1
    assert all(isinstance(s, sheetrs.Skipped) for s in report.skipped)
    assert json.loads(report.to_json()) == report.to_dict()
    assert set(report.to_dict()) == {"replaced", "written", "skipped", "diagnostics"}


# ---- now and seed: the volatile functions fixed by a calculation ---------------------------------

NOW = 45366.75  # 2024-03-15 18:00


@pytest.fixture(params=["xlsx", "ods"])
def volatile(request, tmp_path):
    """A copy of the workbook with `V!A1 =NOW()`, `V!B1 =TODAY()`, `V!C1 =RAND()` and rows to edit."""
    source = ROOT / "sheetrs-edit/tests/assets/volatile" / f"volatile.{request.param}"
    copy = tmp_path / source.name
    shutil.copy(source, copy)
    return copy


def volatile_row(book):
    return book.value_ranges().to_dict()["V!A1:C1"][0]


def test_now_fixes_now_and_today(volatile):
    now, today, _ = volatile_row(sheetrs.Book(volatile).calculate(ranges="V!A1:C1", now=NOW))
    assert now == pytest.approx(NOW)
    assert today == 45366


def test_seed_fixes_rand(volatile):
    def rand(seed):
        return volatile_row(sheetrs.Book(volatile).calculate(ranges="V!A1:C1", seed=seed))[2]

    assert 0 <= rand(42) < 1
    assert rand(42) == rand(42)
    assert rand(42) != rand(44)  # an odd seed and the even one before it are the same stream


def test_without_now_the_clock_is_the_systems(volatile):
    now = volatile_row(sheetrs.Book(volatile).calculate(ranges="V!A1:C1"))[0]
    assert now > 45292 and now != pytest.approx(NOW)


def test_the_recalculation_after_a_deletion_runs_with_the_now_and_the_seed_of_the_step(
    volatile, tmp_path
):
    saved = (
        sheetrs.Book(volatile)
        .calculate(now=NOW, seed=7)
        .delete_rows("V!3")
        .save(tmp_path / f"out{volatile.suffix}")
    )
    now, today, rand = saved.range("V!A1:C1")["V!A1:C1"].values[0]
    assert now == pytest.approx(NOW)
    assert today == 45366
    alone = volatile_row(sheetrs.Book(volatile).calculate(ranges="V!A1:C1", now=NOW, seed=7))
    assert rand == alone[2]


def test_now_must_be_a_number_and_seed_a_non_negative_integer(volatile):
    with pytest.raises(TypeError):
        sheetrs.Book(volatile).calculate(now="today")
    with pytest.raises((TypeError, OverflowError)):
        sheetrs.Book(volatile).calculate(seed=-1)


def test_now_and_seed_are_properties_of_the_book(volatile):
    book = sheetrs.Book(volatile, now=NOW, seed=42)
    assert (book.now, book.seed) == (NOW, 42)
    assert (sheetrs.Book(volatile).now, sheetrs.Book(volatile).seed) == (None, None)
    now, today, rand = volatile_row(book.calculate(ranges="V!A1:C1"))
    assert now == pytest.approx(NOW)
    assert today == 45366
    same = volatile_row(sheetrs.Book(volatile).calculate(ranges="V!A1:C1", seed=42))
    assert rand == same[2]


def test_now_and_seed_of_the_book_fix_the_recalculation_an_edit_leaves(volatile, tmp_path):
    saved = (
        sheetrs.Book(volatile, now=NOW, seed=7)
        .delete_rows("V!3")
        .save(tmp_path / f"out{volatile.suffix}")
    )
    now, today, rand = saved.range("V!A1:C1")["V!A1:C1"].values[0]
    assert now == pytest.approx(NOW)
    assert today == 45366
    alone = volatile_row(sheetrs.Book(volatile).calculate(ranges="V!A1:C1", now=NOW, seed=7))
    assert rand == alone[2]
    # The book that save returns keeps them.
    assert (saved.now, saved.seed) == (NOW, 7)


def test_now_and_seed_of_a_call_win_over_the_ones_of_the_book_field_by_field(volatile):
    book = sheetrs.Book(volatile, now=40000, seed=42)
    now, _, rand = volatile_row(book.calculate(ranges="V!A1:C1", now=NOW))
    assert now == pytest.approx(NOW)
    assert rand == volatile_row(sheetrs.Book(volatile).calculate(ranges="V!A1:C1", seed=42))[2]
    other = volatile_row(sheetrs.Book(volatile, now=40000, seed=42).calculate(ranges="V!A1:C1", seed=44))
    assert other[0] == pytest.approx(40000)
    assert other[2] == volatile_row(sheetrs.Book(volatile).calculate(ranges="V!A1:C1", seed=44))[2]


def test_the_book_takes_now_and_seed_from_bytes_and_streams_too(volatile):
    import io

    data = volatile.read_bytes()
    for source in (data, bytearray(data), io.BytesIO(data)):
        book = sheetrs.Book(source, now=NOW, seed=42)
        assert volatile_row(book.calculate(ranges="V!A1:C1"))[0] == pytest.approx(NOW)


def test_now_and_seed_of_the_book_must_be_a_number_and_a_non_negative_integer(volatile):
    with pytest.raises(TypeError):
        sheetrs.Book(volatile, now="today")
    with pytest.raises((TypeError, OverflowError)):
        sheetrs.Book(volatile, seed=-1)


def test_evaluate_takes_now_and_seed():
    assert sheetrs.evaluate("=NOW()", now=NOW) == pytest.approx(NOW)
    assert sheetrs.evaluate("=RAND()", seed=42) == sheetrs.evaluate("=RAND()", seed=42)
    assert sheetrs.evaluate("=RAND()", seed=42) != sheetrs.evaluate("=RAND()", seed=44)


# ---- the ranges of a calculation: the declared outputs, and what lies outside them ---------------

LOANS = [ROOT / "sheetrs-compiler/tests/assets/loan.xlsx", ROOT / "sheetrs-compiler/tests/assets/loan.ods"]


@pytest.fixture(params=LOANS, ids=lambda p: p.suffix[1:])
def loan(request):
    """The loan declares the outputs `payment` (`Loan!B4`), `interest` (`Loan!B6`) and `balance` (`Loan!B8`);
    `Loan!B5` is a formula nothing declares and `Loan!B2` a constant."""
    return request.param


def test_declared_and_inside_the_outputs_the_ranges_are_a_selection(loan):
    answer = sheetrs.Book(loan).calculate(ranges=["Loan!B6", "payment"]).value_ranges()
    assert [block.range for block in answer.values()] == ["Loan!B6", "Loan!B4"]
    assert answer["payment"].values[0][0] == pytest.approx(1073.6432460242795)


def test_declared_and_outside_the_outputs_the_range_is_read_from_the_written_workbook(loan):
    # Embedded: the workbook is its own template, so a range that is not an output is read, in memory and
    # with no recalculation, from the workbook just written (the input is in it), in the order asked.
    answer = (
        sheetrs.Book(loan)
        .calculate({"principal": 1000}, ranges=["Loan!B5", "payment", "Loan!B2", "Loan!B1"])
        .value_ranges()
    )
    assert [block.range for block in answer.values()] == ["Loan!B5", "Loan!B4", "Loan!B2", "Loan!B1"]
    assert answer["payment"].values[0][0] == pytest.approx(5.368216230121398)
    assert answer["Loan!B2"].values[0][0] == pytest.approx(0.05)
    assert answer["Loan!B1"].values[0][0] == 1000
    # A formula the outputs left out keeps the result the template stored: it is read, not calculated.
    assert answer["Loan!B5"].values[0][0] == pytest.approx(386511.5, abs=1)


def test_a_range_outside_the_outputs_is_read_from_the_written_workbook_in_memory(loan):
    # The calculation writes the workbook (the input is in it, the outputs are calculated) and `range`
    # reads it from those bytes, with no recalculation and no file.
    data = sheetrs.Book(loan).calculate({"principal": 1000}).to_bytes()
    blocks = sheetrs.Book(data).range(["Loan!B1", "Loan!B2", "Loan!B4"])
    assert blocks["Loan!B1"].values[0][0] == 1000
    assert blocks["Loan!B2"].values[0][0] == pytest.approx(0.05)
    assert blocks["Loan!B4"].values[0][0] == pytest.approx(5.368216230121398)
    assert "Loan!A1" in [cell.sheet + "!" + cell.cell for cell in sheetrs.Book(data).find("Principal")]


def test_a_range_that_is_not_one_of_the_workbook_is_an_error(loan):
    with pytest.raises(sheetrs.Error, match="Nowhere!A1"):
        sheetrs.Book(loan).calculate(ranges="Nowhere!A1").value_ranges()


def test_with_none_declared_the_formula_cells_and_the_constants_are_answered(structure):
    answer = sheetrs.Book(structure).calculate(ranges=["F!A5:A6", "T!A1:A3"]).value_ranges()
    assert answer["F!A5:A6"].values == [[5], [9]]
    assert answer["T!A1:A3"].values[0][0] == 1
