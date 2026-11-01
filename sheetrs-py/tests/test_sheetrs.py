"""Exercises the built extension on the workspace's parity fixtures (XLSX and ODS)."""

from pathlib import Path
from zipfile import ZipFile

import pytest

import sheetrs

TESTS = Path(__file__).resolve().parents[2] / "tests"
FIXTURES = [TESTS / "minimal_test.xlsx", TESTS / "minimal_test.ods"]
# Sheets T, U, V with numbers and formulas, and F with formulas over T.
STRUCTURE = Path(__file__).resolve().parents[2] / "sheetrs-edit/tests/assets/structure"


@pytest.fixture(params=FIXTURES, ids=lambda p: p.suffix[1:])
def workbook(request):
    return request.param


@pytest.fixture(params=["xlsx", "ods"])
def structure(request):
    return STRUCTURE / f"book.{request.param}"


def sheet_names(path):
    return [s["name"] for s in sheetrs.open(path).stats()["sheets"]]


# ---- the API ------------------------------------------------------------------------------------


def test_version_is_the_workspace_version():
    major, minor, patch = sheetrs.__version__.split(".")
    assert (major + minor + patch).isdigit()


def test_the_api_is_flat_with_one_way_to_do_each_thing():
    assert set(sheetrs.__all__) >= {"open", "Source", "EditPlan", "Linter", "Violation"}
    for gone in ("lint", "stats", "edit", "read", "scan", "Workbook", "LazyWorkbook"):
        assert not hasattr(sheetrs, gone)


# ---- open gives the input of every tool; it is not an editable workbook ----------------------------


def test_open_gives_a_source_that_has_no_save(workbook):
    source = sheetrs.open(workbook)
    assert isinstance(source, sheetrs.Source)
    # Opening is not editing: a source cannot be saved, only given to a tool or edited.
    assert not hasattr(source, "save")
    with pytest.raises(AttributeError):
        source.save("out.xlsx")


def test_the_first_operation_makes_the_edit_plan_that_has_save(workbook, tmp_path):
    plan = sheetrs.open(workbook).remove_sheets(sheet_names(workbook)[-1])
    assert isinstance(plan, sheetrs.EditPlan)
    assert hasattr(plan, "save")
    out = tmp_path / f"out{workbook.suffix}"
    plan.save(out)
    assert sheet_names(out) == sheet_names(workbook)[:-1]


def test_open_reads_nothing_and_checks_the_path_and_the_format(workbook, tmp_path):
    with pytest.raises(FileNotFoundError):
        sheetrs.open(tmp_path / "missing.xlsx")
    with pytest.raises(ValueError, match="XLSX, XLSM or ODS"):
        sheetrs.open(tmp_path)
    with pytest.raises(TypeError):
        sheetrs.open(42)
    with pytest.raises(ValueError, match="extension"):
        sheetrs.open(workbook.read_bytes())
    with pytest.raises(ValueError, match="XLSX, XLSM or ODS"):
        sheetrs.open(workbook.read_bytes(), extension="csv")
    # A path, a str and bytes give the same source of data.
    by_path = sheetrs.open(workbook).stats()
    assert sheetrs.open(str(workbook)).stats() == by_path
    from_bytes = sheetrs.open(workbook.read_bytes(), extension=workbook.suffix[1:]).stats()
    assert from_bytes["sheets"] == by_path["sheets"]
    assert from_bytes["total_file_size"] == 0


def test_a_source_is_given_to_every_tool(workbook):
    source = sheetrs.open(workbook)
    assert sheetrs.Linter().lint(source)
    assert source.stats()["total_sheets"] >= 1
    assert source.delete_rows("NoSuchSheet!1").operations == ["delete rows NoSuchSheet!1:1"]


# ---- Linter ---------------------------------------------------------------------------------------


def test_rules_catalog_reflects_the_configuration():
    default = sheetrs.Linter().rules()
    ids = [rule["id"] for rule in default]
    assert ids == sorted(ids)
    assert {"ERR102", "CALC202", "REF301"} <= set(ids)
    assert {"id", "name", "category", "default", "enabled"} == set(default[0])
    assert all(rule["enabled"] == rule["default"] for rule in default)

    config = '[global]\ndisabled_rules = ["REF301", "CPX"]\n'
    configured = {r["id"]: r for r in sheetrs.Linter(config=config).rules()}
    assert not configured["REF301"]["enabled"]
    assert not any(r["enabled"] for r in configured.values() if r["category"] == "CPX")
    assert configured["ERR102"]["enabled"] == configured["ERR102"]["default"]


def test_unknown_rules_in_the_configuration_are_rejected():
    with pytest.raises(ValueError, match="ERR1O2"):
        sheetrs.Linter(config='[global]\ndisabled_rules = ["ERR1O2"]\n')


def test_lint_returns_violations_and_a_linter_is_reusable(workbook):
    linter = sheetrs.Linter()
    violations = linter.lint(sheetrs.open(workbook))
    assert violations
    first = violations[0]
    assert first.rule_id and first.message
    assert first.severity in {"error", "warning", "info"}
    assert first.scope in {"book", "sheet", "cell"}
    assert first.to_dict()["rule_id"] == first.rule_id
    assert [v.to_dict() for v in violations] == [
        v.to_dict() for v in linter.lint(sheetrs.open(workbook))
    ]


def test_lint_parity_between_formats():
    # The ODS fixture, exported from LibreOffice, carries a macro library the XLSX twin lacks
    # (VBA1101): a property of the fixtures, not of the bindings.
    xlsx, ods = (
        {v.rule_id for v in sheetrs.Linter().lint(sheetrs.open(p))} - {"VBA1101"}
        for p in FIXTURES
    )
    assert xlsx == ods


def test_lint_a_path_or_bytes_gives_the_same_violations(workbook):
    linter = sheetrs.Linter()
    from_path = [v.to_dict() for v in linter.lint(sheetrs.open(workbook))]
    data = workbook.read_bytes()
    opened = sheetrs.open(data, extension=workbook.suffix[1:])
    assert [v.to_dict() for v in linter.lint(opened)] == from_path


def test_lint_takes_a_source_and_nothing_else(workbook):
    with pytest.raises(TypeError):
        sheetrs.Linter().lint(workbook)  # a path is opened first
    with pytest.raises(TypeError):
        sheetrs.Linter().lint(42)


def test_cell_violations_carry_a_position(workbook):
    cells = [v for v in sheetrs.Linter().lint(sheetrs.open(workbook)) if v.scope == "cell"]
    assert cells
    for v in cells:
        assert v.sheet and v.cell and v.row is not None and v.col is not None
        assert v.cell.endswith(str(v.row + 1))


def test_config_disables_rules(workbook):
    source = sheetrs.open(workbook)
    rule = sorted({v.rule_id for v in sheetrs.Linter().lint(source)})[0]
    toml = f'[global]\ndisabled_rules = ["{rule}"]\n'
    assert rule not in {v.rule_id for v in sheetrs.Linter(config=toml).lint(source)}


def test_config_path_and_exclusive_config(workbook, tmp_path):
    path = tmp_path / "sheetlint.toml"
    path.write_text("[global]\n")
    assert sheetrs.Linter(config_path=path).lint(sheetrs.open(workbook))
    with pytest.raises(ValueError, match="not both"):
        sheetrs.Linter(config="", config_path=path)
    with pytest.raises(ValueError, match="configuration"):
        sheetrs.Linter(config="global = 3")


def test_a_corrupt_workbook_is_a_sheetrs_error():
    source = sheetrs.open(b"not a zip", extension="xlsx")
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.Linter().lint(source)
    with pytest.raises(sheetrs.SheetrsError):
        source.stats()


def test_stats(workbook):
    stats = sheetrs.open(workbook).stats()
    assert stats["total_sheets"] == len(stats["sheets"]) >= 1
    assert stats["total_used_cells"] == sum(s["used_cells"] for s in stats["sheets"])
    assert stats["total_file_size"] == workbook.stat().st_size


# ---- editing --------------------------------------------------------------------------------------


def test_nothing_is_read_or_written_until_save(workbook, tmp_path):
    before = workbook.read_bytes()
    out = tmp_path / f"out{workbook.suffix}"
    plan = sheetrs.open(workbook).remove_sheets(sheet_names(workbook)[-1])
    assert not out.exists()  # recorded, not run
    plan.save(out)
    plan.save(tmp_path / f"again{workbook.suffix}")  # a plan can be saved more than once
    assert workbook.read_bytes() == before  # the original is never modified


def test_methods_return_a_new_plan_and_leave_the_old_one(workbook):
    base = sheetrs.open(workbook).remove_sheets("X")
    two = base.remove_ranges("N")
    assert base.operations == ["remove sheets X"]
    assert two.operations == ["remove sheets X", "remove named ranges N"]


def test_errors_of_a_scope_show_where_it_is_written_and_those_of_the_file_at_save(
    workbook, tmp_path
):
    source = sheetrs.open(workbook)
    with pytest.raises(ValueError, match="row range"):
        source.delete_rows("T!B:D")
    with pytest.raises(ValueError, match="column range"):
        source.delete_columns("T!3:10")
    with pytest.raises(ValueError, match="invalid scope"):
        source.convert_to_values("!A")
    missing = source.delete_rows("NoSuchSheet!1")  # recorded
    with pytest.raises(sheetrs.SheetrsError):
        missing.save(tmp_path / f"out{workbook.suffix}")


def test_a_plan_never_overwrites_the_file_it_was_opened_from(workbook):
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.open(workbook).remove_sheets("X").save(workbook)


def test_an_edit_plan_runs_on_bytes_as_on_the_file(structure, tmp_path):
    ops = lambda s: s.delete_rows("T!4:5").remove_sheets("V")
    from_file, from_bytes = tmp_path / f"a{structure.suffix}", tmp_path / f"b{structure.suffix}"
    ops(sheetrs.open(structure)).save(from_file)
    ops(sheetrs.open(structure.read_bytes(), extension=structure.suffix[1:])).save(from_bytes)
    assert sheetrs.open(from_file).stats()["sheets"] == sheetrs.open(from_bytes).stats()["sheets"]


def test_operations_are_sequential_and_the_scopes_of_a_call_are_simultaneous(
    structure, tmp_path
):
    def stats_of(plan, name):
        out = tmp_path / f"{name}{structure.suffix}"
        plan.save(out)
        return sheetrs.open(out).stats()["sheets"]

    source = sheetrs.open(structure)
    # Row 3 deleted twice is the rows 3 and 4; "T!3, T!4" at once is the same two rows.
    twice = source.delete_rows("T!3").delete_rows("T!3")
    at_once = source.delete_rows("T!3", "T!4")
    assert stats_of(twice, "a") == stats_of(at_once, "b")
    # Not the same as two deletions of the input's row 3.
    other = source.delete_rows("T!3", "T!3")
    assert stats_of(other, "c") != stats_of(at_once, "d")


def test_the_plan_marks_the_saved_file_for_recalculation_when_it_leaves_stale_results(
    structure, tmp_path
):
    def saved(plan, name, **options):
        out = tmp_path / f"{name}{structure.suffix}"
        plan.save(out, **options)
        return out.read_bytes()

    source = sheetrs.open(structure)
    deleting = source.delete_rows("T!4:5")
    # Deleting leaves stale results, so the plan marks the file without being asked; the user may
    # avoid the mark, and may force it where the plan would not mark.
    assert "recalculate when opened" in deleting.explain()
    assert saved(deleting, "auto") != saved(deleting, "never", recalc_on_open=False)
    assert saved(deleting, "auto") == saved(deleting, "always", recalc_on_open=True)
    assert "recalculate" not in deleting.explain(recalc_on_open=False)
    converting = source.convert_to_values("F")
    assert "recalculate" not in converting.explain()  # converting changes no other result
    assert saved(converting, "c-auto") == saved(converting, "c-never", recalc_on_open=False)
    assert saved(converting, "c-auto") != saved(converting, "c-always", recalc_on_open=True)
    # The mark is never one of the operations the user wrote.
    assert "recalculate" not in "".join(deleting.operations)


def test_explain_shows_the_passes_and_a_conversion_after_a_change_needs_another(structure):
    source = sheetrs.open(structure)
    one = source.delete_rows("T!2").delete_columns("U!B").remove_sheets("V")
    assert one.explain().startswith("pass 1 of 1:")
    two = source.delete_rows("T!2").convert_to_values("F")
    assert two.explain().startswith("pass 1 of 2:")
    assert "pass 2 of 2:" in two.explain()
    gone = source.remove_sheets("U").delete_rows("U!2")
    with pytest.raises(ValueError, match="removed"):
        gone.explain()


def test_a_3d_reference_stops_a_structural_edit_and_names_the_issue(tmp_path):
    broken = tmp_path / "span.xlsx"
    with ZipFile(STRUCTURE / "book.xlsx") as zin, ZipFile(broken, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet4.xml":
                data = data.replace(
                    b'<f aca="false">T!A2</f>', b'<f aca="false">SUM(T:V!A2:A4)</f>', 1
                )
                assert b"T:V!" in data
            zout.writestr(item, data)
    source = sheetrs.open(broken)
    with pytest.raises(sheetrs.SheetrsError, match="issues/75"):
        source.delete_rows("T!4:5").save(tmp_path / "out.xlsx")
    # What changes no structure is not stopped.
    source.convert_to_values("F").save(tmp_path / "ok.xlsx")
