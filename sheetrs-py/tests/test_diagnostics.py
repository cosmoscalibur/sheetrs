"""What a run reports besides its result: `Diagnostic` records with a stable code and context fields,
in the report of the `Book`, in the `sheetrs` logger of `logging` and, the warnings, in `warnings`."""

import json
import logging
import warnings

import pytest

import sheetrs

from conftest import ROOT

EXTERNAL = ROOT / "sheetrs-edit/tests/assets/book/book.xlsx"


def _run_that_cannot_be_calculated(tmp_path):
    with pytest.warns(sheetrs.RecalculationWarning):
        return sheetrs.Book(EXTERNAL).calculate().save(tmp_path / "out.xlsx")


def test_a_calculation_the_engine_cannot_do_is_a_diagnostic_with_a_stable_code(tmp_path):
    saved = _run_that_cannot_be_calculated(tmp_path)
    (diagnostic,) = saved.report.diagnostics
    assert isinstance(diagnostic, sheetrs.Diagnostic)
    assert diagnostic.level == "warning"
    assert diagnostic.code == "RECALC_NOT_SUPPORTED"
    assert diagnostic.context["operation"] == "calculate"
    assert diagnostic.context["reason"]
    assert "will be done by the spreadsheet program" in diagnostic.message
    assert str(diagnostic) == f"warning: {diagnostic.message}"


def test_a_diagnostic_writes_itself_as_a_dict_and_as_json(tmp_path):
    (diagnostic,) = _run_that_cannot_be_calculated(tmp_path).report.diagnostics
    record = diagnostic.to_dict()
    assert set(record) == {"level", "code", "message", "context"}
    assert record["context"] == diagnostic.context
    assert json.loads(diagnostic.to_json()) == record


def test_the_diagnostic_goes_to_the_logger_with_its_code_and_fields_in_extra(tmp_path, caplog):
    with caplog.at_level(logging.DEBUG, logger="sheetrs"):
        saved = _run_that_cannot_be_calculated(tmp_path)
    (record,) = [r for r in caplog.records if r.name == "sheetrs"]
    (diagnostic,) = saved.report.diagnostics
    assert record.levelno == logging.WARNING
    assert record.getMessage() == diagnostic.message
    assert record.code == "RECALC_NOT_SUPPORTED"
    assert record.operation == "calculate" and record.reason == diagnostic.context["reason"]


def test_the_warning_is_a_recalculation_warning_with_the_message_of_the_diagnostic(tmp_path):
    with pytest.warns(sheetrs.RecalculationWarning) as caught:
        saved = sheetrs.Book(EXTERNAL).calculate().save(tmp_path / "out.xlsx")
    assert [str(w.message) for w in caught] == [d.message for d in saved.report.diagnostics]
    assert caught[0].filename == __file__  # the warning points at the caller of save()


def test_a_run_with_nothing_to_report_logs_and_warns_nothing(structure, tmp_path, caplog):
    with caplog.at_level(logging.DEBUG, logger="sheetrs"), warnings.catch_warnings():
        warnings.simplefilter("error")
        saved = sheetrs.Book(structure).delete_rows("T!2").save(tmp_path / f"o{structure.suffix}")
    assert saved.report.diagnostics == []
    assert [r for r in caplog.records if r.name == "sheetrs"] == []


def test_to_bytes_reports_the_same_way_as_save(tmp_path, caplog):
    book = sheetrs.Book(EXTERNAL).calculate()
    with caplog.at_level(logging.WARNING, logger="sheetrs"), pytest.warns(sheetrs.RecalculationWarning):
        book.to_bytes()
    assert [d.code for d in book.report.diagnostics] == ["RECALC_NOT_SUPPORTED"]
    assert [r.code for r in caplog.records if r.name == "sheetrs"] == ["RECALC_NOT_SUPPORTED"]


def test_a_logging_handler_that_raises_does_not_hide_the_workbook_that_was_written(tmp_path):
    out = tmp_path / "out.xlsx"
    with warnings.catch_warnings():
        warnings.simplefilter("error", sheetrs.RecalculationWarning)
        with pytest.raises(sheetrs.RecalculationWarning):
            sheetrs.Book(EXTERNAL).calculate().save(out)
    assert out.exists()
