"""Exercises the built extension on the workspace's parity fixtures (XLSX and ODS)."""

from pathlib import Path
from zipfile import ZipFile

import pytest

import sheetrs
from sheetrs.edit import Editor
from sheetrs.linter import Linter
from sheetrs.stats import collect

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
    return [s["name"] for s in collect(sheetrs.open(path))["sheets"]]


# ---- the API ------------------------------------------------------------------------------------


def test_version_is_the_workspace_version():
    major, minor, patch = sheetrs.__version__.split(".")
    assert (major + minor + patch).isdigit()


def test_the_package_is_laid_out_as_the_crates_are():
    assert set(sheetrs.__all__) >= {"open", "Source", "SheetrsError", "linter", "stats", "edit"}
    assert sheetrs.linter.Linter is Linter
    assert sheetrs.stats.collect is collect
    assert sheetrs.edit.Editor is Editor
    # One home for each thing: no tool is also at the top.
    for elsewhere in ("Linter", "Violation", "Editor", "EditPlan", "collect", "lint", "edit_plan"):
        assert not hasattr(sheetrs, elsewhere)


# ---- open gives the input of every tool; it is not an editable workbook ----------------------------


def test_open_gives_a_source_with_no_behavior_of_its_own(workbook):
    source = sheetrs.open(workbook)
    assert isinstance(source, sheetrs.Source)
    # A source is data for the tools: it cannot be saved, measured, linted or edited by itself.
    for method in ("save", "stats", "lint", "delete_rows", "remove_sheets"):
        assert not hasattr(source, method), method


def test_the_editor_has_save_and_a_source_is_given_to_it(workbook, tmp_path):
    editor = Editor(sheetrs.open(workbook)).remove_sheets(sheet_names(workbook)[-1])
    assert isinstance(editor, Editor)
    out = tmp_path / f"out{workbook.suffix}"
    editor.save(out)
    assert sheet_names(out) == sheet_names(workbook)[:-1]
    with pytest.raises(TypeError):
        Editor(workbook)  # a path is not a source: open it first


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
    by_path = collect(sheetrs.open(workbook))
    assert collect(sheetrs.open(str(workbook))) == by_path
    from_bytes = collect(sheetrs.open(workbook.read_bytes(), extension=workbook.suffix[1:]))
    assert from_bytes["sheets"] == by_path["sheets"]
    assert from_bytes["total_file_size"] == 0


def test_a_source_is_given_to_every_tool(workbook):
    source = sheetrs.open(workbook)
    assert Linter().lint(source)
    assert collect(source)["total_sheets"] >= 1
    assert "delete rows NoSuchSheet!1:1" in Editor(source).delete_rows("NoSuchSheet!1").explain()


# ---- Linter ---------------------------------------------------------------------------------------


def test_rules_catalog_reflects_the_configuration():
    default = Linter().rules()
    ids = [rule["id"] for rule in default]
    assert ids == sorted(ids)
    assert {"ERR102", "CALC202", "REF301"} <= set(ids)
    assert {"id", "name", "category", "default", "enabled"} == set(default[0])
    assert all(rule["enabled"] == rule["default"] for rule in default)

    config = '[global]\ndisabled_rules = ["REF301", "CPX"]\n'
    configured = {r["id"]: r for r in Linter(config=config).rules()}
    assert not configured["REF301"]["enabled"]
    assert not any(r["enabled"] for r in configured.values() if r["category"] == "CPX")
    assert configured["ERR102"]["enabled"] == configured["ERR102"]["default"]


def test_unknown_rules_in_the_configuration_are_rejected():
    with pytest.raises(ValueError, match="ERR1O2"):
        Linter(config='[global]\ndisabled_rules = ["ERR1O2"]\n')


def test_lint_returns_violations_and_a_linter_is_reusable(workbook):
    linter = Linter()
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
        {v.rule_id for v in Linter().lint(sheetrs.open(p))} - {"VBA1101"}
        for p in FIXTURES
    )
    assert xlsx == ods


def test_lint_a_path_or_bytes_gives_the_same_violations(workbook):
    linter = Linter()
    from_path = [v.to_dict() for v in linter.lint(sheetrs.open(workbook))]
    data = workbook.read_bytes()
    opened = sheetrs.open(data, extension=workbook.suffix[1:])
    assert [v.to_dict() for v in linter.lint(opened)] == from_path


def test_lint_takes_a_source_and_nothing_else(workbook):
    with pytest.raises(TypeError):
        Linter().lint(workbook)  # a path is opened first
    with pytest.raises(TypeError):
        Linter().lint(42)


def test_cell_violations_carry_a_position(workbook):
    cells = [v for v in Linter().lint(sheetrs.open(workbook)) if v.scope == "cell"]
    assert cells
    for v in cells:
        assert v.sheet and v.cell and v.row is not None and v.col is not None
        assert v.cell.endswith(str(v.row + 1))


def test_config_disables_rules(workbook):
    source = sheetrs.open(workbook)
    rule = sorted({v.rule_id for v in Linter().lint(source)})[0]
    toml = f'[global]\ndisabled_rules = ["{rule}"]\n'
    assert rule not in {v.rule_id for v in Linter(config=toml).lint(source)}


def test_config_path_and_exclusive_config(workbook, tmp_path):
    path = tmp_path / "sheetlint.toml"
    path.write_text("[global]\n")
    assert Linter(config_path=path).lint(sheetrs.open(workbook))
    with pytest.raises(ValueError, match="not both"):
        Linter(config="", config_path=path)
    with pytest.raises(ValueError, match="configuration"):
        Linter(config="global = 3")


def test_a_corrupt_workbook_is_a_sheetrs_error():
    source = sheetrs.open(b"not a zip", extension="xlsx")
    with pytest.raises(sheetrs.SheetrsError):
        Linter().lint(source)
    with pytest.raises(sheetrs.SheetrsError):
        collect(source)


def test_stats(workbook):
    stats = collect(sheetrs.open(workbook))
    assert stats["total_sheets"] == len(stats["sheets"]) >= 1
    assert stats["total_used_cells"] == sum(s["used_cells"] for s in stats["sheets"])
    assert stats["total_file_size"] == workbook.stat().st_size


# ---- editing --------------------------------------------------------------------------------------


def test_nothing_is_read_or_written_until_save(workbook, tmp_path):
    before = workbook.read_bytes()
    out = tmp_path / f"out{workbook.suffix}"
    plan = Editor(sheetrs.open(workbook)).remove_sheets(sheet_names(workbook)[-1])
    assert not out.exists()  # recorded, not run
    plan.save(out)
    plan.save(tmp_path / f"again{workbook.suffix}")  # an editor can be saved more than once
    assert workbook.read_bytes() == before  # the original is never modified


def test_the_editor_accumulates_changes_and_returns_itself(workbook):
    editor = Editor(sheetrs.open(workbook))
    assert editor.remove_sheets("X") is editor
    assert editor.remove_ranges("N") is editor
    text = editor.explain()
    assert "remove sheets X" in text and "remove named ranges N" in text


def test_an_editor_with_no_changes_does_not_save(workbook, tmp_path):
    out = tmp_path / f"out{workbook.suffix}"
    with pytest.raises(sheetrs.SheetrsError, match="No operations"):
        Editor(sheetrs.open(workbook)).save(out)
    assert not out.exists()


def test_errors_of_a_scope_show_where_it_is_written_and_those_of_the_file_at_save(
    workbook, tmp_path
):
    source = sheetrs.open(workbook)
    with pytest.raises(ValueError, match="row range"):
        Editor(source).delete_rows("T!B:D")
    with pytest.raises(ValueError, match="column range"):
        Editor(source).delete_columns("T!3:10")
    with pytest.raises(ValueError, match="invalid scope"):
        Editor(source).convert_to_values("!A")
    missing = Editor(source).delete_rows("NoSuchSheet!1")  # recorded
    with pytest.raises(sheetrs.SheetrsError):
        missing.save(tmp_path / f"out{workbook.suffix}")


def test_an_editor_never_overwrites_the_file_it_was_opened_from(workbook):
    with pytest.raises(sheetrs.SheetrsError):
        Editor(sheetrs.open(workbook)).remove_sheets("X").save(workbook)


def test_an_editor_runs_on_bytes_as_on_the_file(structure, tmp_path):
    def edited(source, out):
        Editor(source).delete_rows("T!4:5").remove_sheets("V").save(out)
        return collect(sheetrs.open(out))["sheets"]

    from_file = edited(sheetrs.open(structure), tmp_path / f"a{structure.suffix}")
    data = structure.read_bytes()
    from_bytes = edited(
        sheetrs.open(data, extension=structure.suffix[1:]), tmp_path / f"b{structure.suffix}"
    )
    assert from_file == from_bytes


def test_operations_are_sequential_and_the_scopes_of_a_call_are_simultaneous(
    structure, tmp_path
):
    def stats_of(plan, name):
        out = tmp_path / f"{name}{structure.suffix}"
        plan.save(out)
        return collect(sheetrs.open(out))["sheets"]

    source = sheetrs.open(structure)
    # Row 3 deleted twice is the rows 3 and 4; "T!3, T!4" at once is the same two rows.
    twice = Editor(source).delete_rows("T!3").delete_rows("T!3")
    at_once = Editor(source).delete_rows("T!3", "T!4")
    assert stats_of(twice, "a") == stats_of(at_once, "b")
    # Not the same as two deletions of the input's row 3.
    other = Editor(source).delete_rows("T!3", "T!3")
    assert stats_of(other, "c") != stats_of(at_once, "d")


def is_marked_for_recalculation(path):
    """The mark: a flag in an XLSX; in an ODS, the formulas have no cached result left."""
    with ZipFile(path) as z:
        if path.suffix == ".xlsx":
            return b'fullCalcOnLoad="1"' in z.read("xl/workbook.xml")
        content = z.read("content.xml").decode()
    formulas = [c for c in content.split("<table:table-cell")[1:] if "table:formula=" in c]
    return bool(formulas) and not any("office:value-type" in c.split(">")[0] for c in formulas)


def test_the_editor_always_decides_the_mark_for_recalculation_and_there_is_no_option(
    structure, tmp_path
):
    def saved(plan, name):
        out = tmp_path / f"{name}{structure.suffix}"
        plan.save(out)
        return out

    source = sheetrs.open(structure)
    # Deleting leaves stale results, so the editor marks the file; nobody has to ask for it.
    deleting = Editor(source).delete_rows("T!4:5")
    assert "recalculate when opened" in deleting.explain()
    assert is_marked_for_recalculation(saved(deleting, "deleting"))
    # Converting changes no other result: no mark.
    converting = Editor(source).convert_to_values("F")
    assert "recalculate" not in converting.explain()
    assert not is_marked_for_recalculation(saved(converting, "converting"))
    # The mark goes last, after everything, so a conversion that comes before is not affected.
    both = Editor(source).convert_to_values("F").delete_rows("T!4:5")
    assert "recalculate" in both.explain()
    assert both.explain().count("recalculate") == 1  # said once, at the end, not as an operation
    # And it cannot be asked for or avoided: a user who does not know what it is cannot choose.
    with pytest.raises(TypeError):
        deleting.save(tmp_path / f"x{structure.suffix}", recalc_on_open=False)
    with pytest.raises(TypeError):
        deleting.explain(recalc_on_open=True)


def test_explain_shows_the_passes_and_a_conversion_after_a_change_needs_another(structure):
    source = sheetrs.open(structure)
    one = Editor(source).delete_rows("T!2").delete_columns("U!B").remove_sheets("V")
    assert one.explain().startswith("pass 1 of 1:")
    two = Editor(source).delete_rows("T!2").convert_to_values("F")
    assert two.explain().startswith("pass 1 of 2:")
    assert "pass 2 of 2:" in two.explain()
    gone = Editor(source).remove_sheets("U").delete_rows("U!2")
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
        Editor(source).delete_rows("T!4:5").save(tmp_path / "out.xlsx")
    # What changes no structure is not stopped.
    Editor(source).convert_to_values("F").save(tmp_path / "ok.xlsx")
