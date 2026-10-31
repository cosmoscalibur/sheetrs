"""Exercises the built extension on the workspace's parity fixtures (XLSX and ODS)."""

from pathlib import Path

import pytest

import sheetrs

TESTS = Path(__file__).resolve().parents[2] / "tests"
FIXTURES = [TESTS / "minimal_test.xlsx", TESTS / "minimal_test.ods"]
# Sheets T, U, V with numbers and formulas, and F with formulas over them (also across sheets).
STRUCTURE = Path(__file__).resolve().parents[2] / "sheetrs-edit/tests/assets/structure"


@pytest.fixture(params=FIXTURES, ids=lambda p: p.suffix[1:])
def workbook(request):
    return request.param


@pytest.fixture(params=["xlsx", "ods"])
def structure(request):
    return STRUCTURE / f"book.{request.param}"


def test_version_is_the_workspace_version():
    major, minor, patch = sheetrs.__version__.split(".")
    assert (major + minor + patch).isdigit()


def test_rules_catalog_reflects_the_configuration():
    default = sheetrs.linter.Linter().rules()
    ids = [rule["id"] for rule in default]
    assert ids == sorted(ids)
    assert {"ERR102", "CALC202", "REF301"} <= set(ids)
    assert {"id", "name", "category", "default", "enabled"} == set(default[0])
    assert all(rule["enabled"] == rule["default"] for rule in default)

    config = '[global]\ndisabled_rules = ["REF301", "CPX"]\n'
    configured = {r["id"]: r for r in sheetrs.linter.Linter(config=config).rules()}
    assert not configured["REF301"]["enabled"]
    assert not any(r["enabled"] for r in configured.values() if r["category"] == "CPX")
    assert configured["ERR102"]["enabled"] == configured["ERR102"]["default"]


def test_unknown_rules_in_the_configuration_are_rejected():
    with pytest.raises(ValueError, match="ERR1O2"):
        sheetrs.linter.Linter(config='[global]\ndisabled_rules = ["ERR1O2"]\n')
    with pytest.raises(ValueError, match="ERR1O2"):
        sheetrs.linter.lint(
            FIXTURES[0], config='[global]\ndisabled_rules = ["ERR1O2"]\n'
        )


def test_a_linter_is_reusable(workbook):
    linter = sheetrs.linter.Linter()
    first = [v.to_dict() for v in linter.lint(workbook)]
    assert first == [v.to_dict() for v in linter.lint(workbook)]
    assert first == [v.to_dict() for v in sheetrs.linter.lint(workbook)]


def test_lint_returns_violations(workbook):
    violations = sheetrs.linter.lint(workbook)
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
        {v.rule_id for v in sheetrs.linter.lint(p)} - {"VBA1101"} for p in FIXTURES
    )
    assert xlsx == ods


def test_lint_accepts_str_pathlike_and_bytes(workbook):
    from_path = [v.to_dict() for v in sheetrs.linter.lint(workbook)]
    assert [v.to_dict() for v in sheetrs.linter.lint(str(workbook))] == from_path
    data = workbook.read_bytes()
    assert [
        v.to_dict() for v in sheetrs.linter.lint(data, extension=workbook.suffix[1:])
    ] == from_path


def test_bytes_need_an_extension(workbook):
    with pytest.raises(ValueError, match="extension"):
        sheetrs.linter.lint(workbook.read_bytes())


def test_wrong_source_type():
    with pytest.raises(TypeError):
        sheetrs.linter.lint(42)


def test_cell_violations_carry_a_position(workbook):
    cells = [v for v in sheetrs.linter.lint(workbook) if v.scope == "cell"]
    assert cells
    for v in cells:
        assert v.sheet and v.cell and v.row is not None and v.col is not None
        assert v.cell.endswith(str(v.row + 1))


def test_config_disables_rules(workbook):
    everything = {v.rule_id for v in sheetrs.linter.lint(workbook)}
    rule = sorted(everything)[0]
    toml = f'[global]\ndisabled_rules = ["{rule}"]\n'
    assert rule not in {v.rule_id for v in sheetrs.linter.lint(workbook, config=toml)}


def test_config_path_and_exclusive_config(workbook, tmp_path):
    path = tmp_path / "sheetlint.toml"
    path.write_text("[global]\n")
    assert sheetrs.linter.lint(workbook, config_path=path)
    with pytest.raises(ValueError, match="not both"):
        sheetrs.linter.lint(workbook, config="", config_path=path)


def test_invalid_config_is_a_value_error(workbook):
    with pytest.raises(ValueError, match="configuration"):
        sheetrs.linter.lint(workbook, config="global = 3")


def test_missing_file_is_a_file_not_found_error(tmp_path):
    with pytest.raises(FileNotFoundError):
        sheetrs.linter.lint(tmp_path / "missing.xlsx")


def test_corrupt_workbook_is_a_sheetrs_error():
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.linter.lint(b"not a zip", extension="xlsx")


def test_stats(workbook):
    stats = sheetrs.stats.compute(workbook)
    assert stats["total_sheets"] == len(stats["sheets"]) >= 1
    assert stats["total_used_cells"] == sum(s["used_cells"] for s in stats["sheets"])
    assert stats["total_file_size"] == workbook.stat().st_size


def test_stats_from_bytes(workbook):
    stats = sheetrs.stats.compute(workbook.read_bytes(), extension=workbook.suffix[1:])
    assert stats["total_file_size"] == 0
    assert stats["sheets"] == sheetrs.stats.compute(workbook)["sheets"]


def sheet_names(path):
    return [s["name"] for s in sheetrs.stats.compute(path)["sheets"]]


def test_lazy_plan_applies_everything_in_one_pass(workbook, tmp_path):
    names = sheet_names(workbook)
    assert len(names) > 2
    output = tmp_path / f"out{workbook.suffix}"
    plan = (
        sheetrs.edit.Plan(workbook)
        .remove_sheets(names[-1])
        .remove_sheets(names[-2], names[-1])  # a repeat counts once
        .recalc_on_open()
    )
    assert plan.operations == [
        f"remove sheet {names[-1]}",
        f"remove sheet {names[-2]}",
        "recalculate when opened",
    ]
    assert len(plan) == 3
    assert not output.exists()  # lazy: nothing ran yet
    plan.save(output)
    assert sheet_names(output) == names[:-2]


def test_plan_can_be_saved_twice(workbook, tmp_path):
    plan = sheetrs.edit.Plan(workbook).recalc_on_open()
    for n in (1, 2):
        plan.save(tmp_path / f"o{n}{workbook.suffix}")
        assert (tmp_path / f"o{n}{workbook.suffix}").exists()


def test_empty_plan_does_not_save(workbook, tmp_path):
    with pytest.raises(sheetrs.SheetrsError, match="No operations"):
        sheetrs.edit.Plan(workbook).save(tmp_path / f"o{workbook.suffix}")


def test_plan_validates_scopes_when_they_are_added(workbook):
    with pytest.raises(ValueError, match="scope"):
        sheetrs.edit.Plan(workbook).delete_rows("!A")


def test_plan_never_overwrites_its_input(workbook):
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.edit.Plan(workbook).recalc_on_open().save(workbook)


def test_plan_with_a_missing_input_fails_on_save(tmp_path):
    plan = sheetrs.edit.Plan(tmp_path / "missing.xlsx").recalc_on_open()
    with pytest.raises(FileNotFoundError):
        plan.save(tmp_path / "out.xlsx")


def test_eager_operations_run_at_once(workbook, tmp_path):
    names = sheet_names(workbook)
    output = tmp_path / f"out{workbook.suffix}"
    sheetrs.edit.remove_sheets(workbook, output, names[-1])
    assert sheet_names(output) == names[:-1]


def test_eager_operations_chain_through_their_outputs(workbook, tmp_path):
    names = sheet_names(workbook)
    first = tmp_path / f"one{workbook.suffix}"
    second = tmp_path / f"two{workbook.suffix}"
    sheetrs.edit.remove_sheets(workbook, first, names[-1])
    sheetrs.edit.remove_sheets(first, second, names[-2])
    assert sheet_names(second) == names[:-2]


def test_eager_scopes_are_validated(workbook, tmp_path):
    with pytest.raises(ValueError, match="scope"):
        sheetrs.edit.delete_columns(workbook, tmp_path / f"o{workbook.suffix}", "!A")


def test_submodules_import_as_modules():
    from sheetrs.edit import Plan, remove_sheets
    from sheetrs.linter import Linter, Violation, lint
    from sheetrs.stats import compute

    assert callable(lint) and callable(compute) and callable(remove_sheets)
    assert (Plan.__name__, Linter.__name__, Violation.__name__) == (
        "Plan",
        "Linter",
        "Violation",
    )


def test_plan_schedule_is_canonical_and_independent_of_the_order_added(structure):
    first = (
        sheetrs.edit.Plan(structure)
        .delete_rows("T!5", "T!3:6", "U!2")
        .delete_columns("T!D")
        .convert_to_values("F")
        .remove_ranges("NameOne")
        .recalc_on_open()
    )
    second = (
        sheetrs.edit.Plan(structure)
        .recalc_on_open()
        .remove_ranges("NameOne")
        .delete_columns("T!D")
        .convert_to_values("F")
        .delete_rows("U!2", "T!3:6")
    )
    assert first.schedule == second.schedule
    assert first.schedule == [
        "convert formulas to values in F",
        "delete rows T!3:6",  # merged from T!5 and T!3:6, the later rows first
        "delete rows U!2:2",
        "delete columns T!D:D",
        "remove named range NameOne",
        "recalculate when opened",
    ]
    assert len(first.operations) == 7  # as added


def test_plan_gives_the_same_workbook_in_any_order(structure, tmp_path):
    ops = [
        ("delete_rows", ("T!2",)),
        ("delete_rows", ("U!4:5",)),
        ("convert_to_values", ("F",)),
        ("remove_sheets", ("V",)),
    ]
    outputs = []
    for i, order in enumerate([ops, ops[::-1], ops[2:] + ops[:2]]):
        plan = sheetrs.edit.Plan(structure)
        for name, args in order:
            getattr(plan, name)(*args)
        output = tmp_path / f"o{i}{structure.suffix}"
        plan.save(output)
        outputs.append(sheetrs.stats.compute(output)["sheets"])
    assert outputs[0] == outputs[1] == outputs[2]


def test_rows_need_a_row_range_and_columns_a_column_range(workbook, tmp_path):
    with pytest.raises(ValueError, match="row range"):
        sheetrs.edit.Plan(workbook).delete_rows("T!B:D")
    with pytest.raises(ValueError, match="column range"):
        sheetrs.edit.Plan(workbook).delete_columns("T!3:10")
    with pytest.raises(ValueError, match="column range"):
        sheetrs.edit.delete_columns(workbook, tmp_path / f"o{workbook.suffix}", "T")
