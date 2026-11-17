"""What the engine cannot calculate (another workbook, a macro, an extension) has no option: the engine does what `LibreOffice` does. The logic is in Rust
(`sheetrs-edit/tests/from_cache.rs`, `sheetrs-compiler/tests/macros.rs`: which cell keeps the cache, which is `#NAME?`, which is an error, and the refusal
of a workbook with macros by the build). What the Python interface has to get right is the mapping, and that is all this file proves: the warning of the
report reaches Python as a `UserWarning` and a log record with its code, the refusal of the build is a `sheetrs.Error` that names secure mode, the
keyword `secure_mode` reaches it, and the three options of an earlier version are gone from `Book` and `calculate`.
The workbooks are in `sheetrs-edit/tests/assets/from_cache` (`A1` depends on it and caches a result)."""

import logging

import pytest

import sheetrs
from conftest import ROOT

ASSETS = ROOT / "sheetrs-edit/tests/assets/from_cache"


def the_report_warning_is_a_python_warning_and_a_log_record(tmp_path, caplog):
    out = tmp_path / "out.xlsx"
    with caplog.at_level(logging.WARNING, logger="sheetrs"), pytest.warns(UserWarning, match="keep the result the file caches"):
        saved = sheetrs.Book(ASSETS / "extension.xlsx").calculate().save(out)
    assert [d.code for d in saved.report.diagnostics] == ["EXTENSION_FROM_CACHE"]
    assert any("EXTENSION_FROM_CACHE" in record.getMessage() or "keep the result" in record.getMessage() for record in caplog.records)


def a_workbook_with_macros_is_a_sheetrs_error_that_names_secure_mode(tmp_path):
    # The build refuses it, with or without `calculate_all`, and a save leaves no file.
    for calculate_all in (False, True):
        with pytest.raises(sheetrs.Error, match="secure mode"):
            sheetrs.Book(ASSETS / "macro.xlsx").calculate().save(tmp_path / "out.xlsx", calculate_all=calculate_all)
        assert not (tmp_path / "out.xlsx").exists()


def inspect_reads_a_workbook_with_macros():
    assert sheetrs.Book(ASSETS / "macro.xlsx").inspect()


def secure_mode_turns_the_macro_into_an_error_cell_and_a_python_warning(tmp_path, caplog):
    # The keyword reaches the build: the same run that raised above goes on, the warning of the report is a `UserWarning` and a log record.
    out = tmp_path / "out.xlsx"
    with caplog.at_level(logging.WARNING, logger="sheetrs"), pytest.warns(UserWarning, match="secure mode"):
        saved = sheetrs.Book(ASSETS / "macro.xlsx").calculate().save(out, calculate_all=True, secure_mode=True)
    assert [d.code for d in saved.report.diagnostics] == ["UNSUPPORTED_AS_ERROR"]
    assert any("UNSUPPORTED_AS_ERROR" in record.getMessage() or "secure mode" in record.getMessage() for record in caplog.records)
    assert out.exists()


@pytest.mark.filterwarnings("ignore::UserWarning")
def secure_mode_is_a_keyword_of_every_terminal_that_calculates(tmp_path):
    book = sheetrs.Book(ASSETS / "macro.ods").calculate()
    assert book.to_bytes(calculate_all=True, secure_mode=True)
    book.value_ranges(calculate_all=True, secure_mode=True)
    # Off by default: the same calls raise.
    with pytest.raises(sheetrs.Error, match="secure mode"):
        book.to_bytes(calculate_all=True)
    with pytest.raises(sheetrs.Error, match="secure mode"):
        book.value_ranges(calculate_all=True)
    # And a save raises, and leaves no file.
    with pytest.raises(sheetrs.Error, match="secure mode"):
        book.save(tmp_path / "out.ods", calculate_all=True)
    assert not (tmp_path / "out.ods").exists()


def removed_options_are_refused(tmp_path):
    for option in ("external_from_cache", "macro_from_cache", "extension_from_cache"):
        with pytest.raises(TypeError):
            sheetrs.Book(ASSETS / "external.xlsx", **{option: True})
        with pytest.raises(TypeError):
            sheetrs.Book(ASSETS / "external.xlsx").calculate(**{option: True})
        assert not hasattr(sheetrs.Book(ASSETS / "external.xlsx"), option)
