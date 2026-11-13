"""`Book.lint`, `Book.stats` and `sheetrs.lint_rules`."""

import json

import pytest

import sheetrs


def test_lint_gives_the_violations(lint_book):
    violations = sheetrs.Book(lint_book).lint()
    assert violations
    first = violations[0]
    assert isinstance(first, sheetrs.Violation)
    assert first.severity in {"error", "warning", "info"}
    assert first.scope in {"book", "sheet", "cell"}
    assert first.rule_id and first.rule_name and first.message
    assert first.to_dict()["rule_id"] == first.rule_id
    assert first.rule_id in repr(first) and str(first) == repr(first)
    scopes = [v.scope for v in violations]
    order = {"book": 0, "sheet": 1, "cell": 2}
    assert scopes == sorted(scopes, key=order.get)


def test_a_cell_violation_says_where(lint_book):
    cell = next(v for v in sheetrs.Book(lint_book).lint() if v.scope == "cell")
    assert cell.sheet and cell.cell and cell.row is not None and cell.col is not None
    book_level = [v for v in sheetrs.Book(lint_book).lint() if v.scope == "book"]
    assert all(v.sheet is None and v.cell is None for v in book_level)


def test_lint_from_bytes_gives_what_a_file_does(lint_book):
    from_file = [v.rule_id for v in sheetrs.Book(lint_book).lint()]
    from_bytes = [
        v.rule_id
        for v in sheetrs.Book(lint_book.read_bytes()).lint()
    ]
    assert from_bytes == from_file


def test_the_text_of_a_run_in_each_format(lint_book):
    book = sheetrs.Book(lint_book)
    human = book.lint().to_text()
    assert isinstance(human, str) and human.strip()
    assert "\x1b[" not in human  # no terminal styles

    document = json.loads(book.lint().to_json())
    assert {"file", "summary", "violations"} <= set(document)
    assert document["file"] == str(lint_book)
    assert len(document["violations"]) == len(book.lint())

    sarif = json.loads(book.lint().to_sarif())
    assert sarif["version"] == "2.1.0"
    assert len(sarif["runs"][0]["results"]) == len(book.lint())
    assert sarif["runs"][0]["tool"]["driver"]["name"] == "sheetrs lint"


def test_the_name_a_document_gives_the_workbook(lint_book):
    book = sheetrs.Book(lint_book)
    assert json.loads(book.lint().to_json(file="mine.xlsx"))["file"] == "mine.xlsx"
    sarif = json.loads(book.lint().to_sarif(file="mine.xlsx"))
    uri = sarif["runs"][0]["results"][0]["locations"][0]["physicalLocation"]["artifactLocation"]["uri"]
    assert uri == "mine.xlsx"
    bytes_book = sheetrs.Book(lint_book.read_bytes())
    assert json.loads(bytes_book.lint().to_json())["file"] == f"workbook.{lint_book.suffix[1:]}"


def test_the_report_is_a_sequence_and_its_dict_is_the_json_document(lint_book):
    report = sheetrs.Book(lint_book).lint()
    assert len(report) == len(report.violations) == len(list(report))
    assert report[0].rule_id == report.violations[0].rule_id
    assert [v.rule_id for v in report[:2]] == [v.rule_id for v in report.violations[:2]]
    assert report.to_dict() == json.loads(report.to_json())


def test_the_configuration_is_a_file(lint_book, tmp_path):
    book = sheetrs.Book(lint_book)
    rule = sorted({v.rule_id for v in book.lint()})[0]
    config = tmp_path / "sheetrs.toml"
    config.write_text(f'[global]\ndisabled_rules = ["{rule}"]\n')
    assert rule not in {v.rule_id for v in book.lint(config)}
    assert rule not in {v.rule_id for v in book.lint(str(config))}
    assert json.loads(book.lint(config).to_json())["violations"]


def test_an_unknown_rule_or_a_file_that_is_not_there_is_a_value_error(lint_book, tmp_path):
    bad = tmp_path / "bad.toml"
    bad.write_text('[global]\ndisabled_rules = ["ERR1O2"]\n')
    with pytest.raises(ValueError, match="configuration"):
        sheetrs.Book(lint_book).lint(bad)
    with pytest.raises(ValueError, match="configuration"):
        sheetrs.Book(lint_book).lint(tmp_path / "absent.toml")


def test_the_rules_and_what_a_configuration_enables(tmp_path):
    rules = sheetrs.lint_rules()
    ids = [r["id"] for r in rules]
    assert ids == sorted(ids) and len(ids) > 20
    assert set(rules[0]) == {"id", "name", "category", "default", "enabled"}
    assert all(r["enabled"] == r["default"] for r in rules)
    config = tmp_path / "c.toml"
    config.write_text('[global]\ndisabled_rules = ["REF301", "CPX"]\n')
    configured = {r["id"]: r for r in sheetrs.lint_rules(config)}
    assert not configured["REF301"]["enabled"]
    assert not any(r["enabled"] for r in configured.values() if r["category"] == "CPX")
    bad = tmp_path / "bad.toml"
    bad.write_text('[global]\ndisabled_rules = ["ERR1O2"]\n')
    with pytest.raises(ValueError):
        sheetrs.lint_rules(bad)


def test_stats_total_the_sheets_and_cells_and_the_size_of_a_file_but_not_of_bytes(lint_book):
    stats = sheetrs.Book(lint_book).stats
    sheets = list(stats.sheets)
    assert stats.total_sheets == len(sheets) >= 1
    assert stats.total_used_cells == sum(s.used_cells for s in sheets)
    assert stats.total_file_size == lint_book.stat().st_size
    from_bytes = sheetrs.Book(lint_book.read_bytes()).stats
    assert from_bytes.total_file_size == 0
    assert from_bytes.total_sheets == stats.total_sheets


def test_a_workbook_that_cannot_be_read_is_an_error(tmp_path):
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"PK this is not a zip")
    with pytest.raises(sheetrs.Error):
        sheetrs.Book(broken)
