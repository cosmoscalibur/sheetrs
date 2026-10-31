"""Exercises the built extension on the workspace's parity fixtures (XLSX and ODS)."""

from pathlib import Path

import pytest

import sheetrs

TESTS = Path(__file__).resolve().parents[2] / "tests"
FIXTURES = [TESTS / "minimal_test.xlsx", TESTS / "minimal_test.ods"]


@pytest.fixture(params=FIXTURES, ids=lambda p: p.suffix[1:])
def workbook(request):
    return request.param


def test_version_is_the_workspace_version():
    major, minor, patch = sheetrs.__version__.split(".")
    assert (major + minor + patch).isdigit()


def test_rules_catalog():
    rules = sheetrs.rules()
    ids = [rule["id"] for rule in rules]
    assert ids == sorted(ids)
    assert {"ERR102", "CALC202", "REF301"} <= set(ids)
    assert {"id", "name", "category", "default"} == set(rules[0])


def test_lint_returns_violations(workbook):
    violations = sheetrs.lint(workbook)
    assert violations
    first = violations[0]
    assert first.rule_id and first.message
    assert first.severity in {"error", "warning", "info"}
    assert first.scope in {"book", "sheet", "cell"}
    assert first.to_dict()["rule_id"] == first.rule_id


def test_lint_parity_between_formats():
    # The ODS fixture, exported from LibreOffice, carries a macro library the XLSX twin lacks
    # (VBA1101): a property of the fixtures, not of the bindings.
    xlsx, ods = (
        {v.rule_id for v in sheetrs.lint(p)} - {"VBA1101"} for p in FIXTURES
    )
    assert xlsx == ods


def test_lint_accepts_str_pathlike_and_bytes(workbook):
    from_path = [v.to_dict() for v in sheetrs.lint(workbook)]
    assert [v.to_dict() for v in sheetrs.lint(str(workbook))] == from_path
    data = workbook.read_bytes()
    assert [
        v.to_dict() for v in sheetrs.lint(data, extension=workbook.suffix[1:])
    ] == from_path


def test_bytes_need_an_extension(workbook):
    with pytest.raises(ValueError, match="extension"):
        sheetrs.lint(workbook.read_bytes())


def test_wrong_source_type():
    with pytest.raises(TypeError):
        sheetrs.lint(42)


def test_cell_violations_carry_a_position(workbook):
    cells = [v for v in sheetrs.lint(workbook) if v.scope == "cell"]
    assert cells
    for v in cells:
        assert v.sheet and v.cell and v.row is not None and v.col is not None
        assert v.cell.endswith(str(v.row + 1))


def test_config_disables_rules(workbook):
    everything = {v.rule_id for v in sheetrs.lint(workbook)}
    rule = sorted(everything)[0]
    toml = f'[global]\ndisabled_rules = ["{rule}"]\n'
    assert rule not in {v.rule_id for v in sheetrs.lint(workbook, config=toml)}


def test_config_path_and_exclusive_config(workbook, tmp_path):
    path = tmp_path / "sheetlint.toml"
    path.write_text("[global]\n")
    assert sheetrs.lint(workbook, config_path=path)
    with pytest.raises(ValueError, match="not both"):
        sheetrs.lint(workbook, config="", config_path=path)


def test_invalid_config_is_a_value_error(workbook):
    with pytest.raises(ValueError, match="configuration"):
        sheetrs.lint(workbook, config="global = 3")


def test_missing_file_is_a_file_not_found_error(tmp_path):
    with pytest.raises(FileNotFoundError):
        sheetrs.lint(tmp_path / "missing.xlsx")


def test_corrupt_workbook_is_a_sheetrs_error():
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.lint(b"not a zip", extension="xlsx")


def test_stats(workbook):
    stats = sheetrs.stats(workbook)
    assert stats["total_sheets"] == len(stats["sheets"]) >= 1
    assert stats["total_used_cells"] == sum(s["used_cells"] for s in stats["sheets"])
    assert stats["total_file_size"] == workbook.stat().st_size


def test_stats_from_bytes(workbook):
    stats = sheetrs.stats(workbook.read_bytes(), extension=workbook.suffix[1:])
    assert stats["total_file_size"] == 0
    assert stats["sheets"] == sheetrs.stats(workbook)["sheets"]


def test_edit_removes_a_sheet(workbook, tmp_path):
    names = [s["name"] for s in sheetrs.stats(workbook)["sheets"]]
    assert len(names) > 1
    output = tmp_path / f"out{workbook.suffix}"
    sheetrs.edit(workbook, output, remove_sheets=[names[-1]])
    left = [s["name"] for s in sheetrs.stats(output)["sheets"]]
    assert left == names[:-1]


def test_edit_never_overwrites_its_input(workbook):
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.edit(workbook, workbook, recalc_on_open=True)


def test_edit_rejects_a_malformed_scope(workbook, tmp_path):
    with pytest.raises(ValueError, match="scope"):
        sheetrs.edit(
            workbook, tmp_path / f"o{workbook.suffix}", convert_to_values=["!A"]
        )
