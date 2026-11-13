"""`Book.inspect()`: what `sheetc inspect` says about a workbook, as a typed value."""

import json

import sheetrs


def test_inspect_says_what_the_compiler_sees_without_calculating(structure):
    inspection = sheetrs.Book(structure).inspect()
    assert isinstance(inspection, sheetrs.Inspection)
    assert inspection.sheets == ["T", "U", "V", "F"]
    assert inspection.date_system == "1899-12-30"
    assert inspection.formulas > 0 and inspection.cells >= inspection.formulas
    assert inspection.circular_groups == 0 and inspection.run_time_references is False
    assert inspection.unsupported_functions == []
    assert inspection.diagnostics == [] or all(isinstance(d, str) for d in inspection.diagnostics)


def test_inspect_lists_the_functions_used_most_used_first(structure):
    functions = sheetrs.Book(structure).inspect().functions
    assert functions and all(isinstance(f, sheetrs.FunctionUse) for f in functions)
    assert [f.calls for f in functions] == sorted((f.calls for f in functions), reverse=True)
    assert any(f.name == "SUM" and f.implemented for f in functions)


def test_inspect_writes_itself_as_a_dict_and_as_json(structure):
    inspection = sheetrs.Book(structure).inspect()
    document = inspection.to_dict()
    assert document["sheets"] == inspection.sheets
    assert [f["name"] for f in document["functions"]] == [f.name for f in inspection.functions]
    assert json.loads(inspection.to_json()) == document


def test_inspect_waits_for_the_pending_operations(structure):
    book = sheetrs.Book(structure).delete_rows("T!2")
    try:
        book.inspect()
    except sheetrs.PendingOperationsError:
        return
    raise AssertionError("inspect() must not answer about a workbook the operations have not made")
