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


def test_a_linter_is_reusable(workbook):
    linter = sheetrs.Linter()
    first = [v.to_dict() for v in linter.lint(workbook)]
    assert first == [v.to_dict() for v in linter.lint(workbook)]


def test_lint_returns_violations(workbook):
    violations = sheetrs.Linter().lint(workbook)
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
        {v.rule_id for v in sheetrs.Linter().lint(p)} - {"VBA1101"} for p in FIXTURES
    )
    assert xlsx == ods


def test_lint_accepts_str_pathlike_and_bytes(workbook):
    from_path = [v.to_dict() for v in sheetrs.Linter().lint(workbook)]
    assert [v.to_dict() for v in sheetrs.Linter().lint(str(workbook))] == from_path
    data = workbook.read_bytes()
    assert [
        v.to_dict() for v in sheetrs.Linter().lint(data, extension=workbook.suffix[1:])
    ] == from_path


def test_bytes_need_an_extension(workbook):
    with pytest.raises(ValueError, match="extension"):
        sheetrs.Linter().lint(workbook.read_bytes())


def test_wrong_source_type():
    with pytest.raises(TypeError):
        sheetrs.Linter().lint(42)


def test_cell_violations_carry_a_position(workbook):
    cells = [v for v in sheetrs.Linter().lint(workbook) if v.scope == "cell"]
    assert cells
    for v in cells:
        assert v.sheet and v.cell and v.row is not None and v.col is not None
        assert v.cell.endswith(str(v.row + 1))


def test_config_disables_rules(workbook):
    everything = {v.rule_id for v in sheetrs.Linter().lint(workbook)}
    rule = sorted(everything)[0]
    toml = f'[global]\ndisabled_rules = ["{rule}"]\n'
    assert rule not in {v.rule_id for v in sheetrs.Linter(config=toml).lint(workbook)}


def test_config_path_and_exclusive_config(workbook, tmp_path):
    path = tmp_path / "sheetlint.toml"
    path.write_text("[global]\n")
    assert sheetrs.Linter(config_path=path).lint(workbook)
    with pytest.raises(ValueError, match="not both"):
        sheetrs.Linter(config="", config_path=path)


def test_invalid_config_is_a_value_error(workbook):
    with pytest.raises(ValueError, match="configuration"):
        sheetrs.Linter(config="global = 3")


def test_missing_file_is_a_file_not_found_error(tmp_path):
    with pytest.raises(FileNotFoundError):
        sheetrs.Linter().lint(tmp_path / "missing.xlsx")


def test_corrupt_workbook_is_a_sheetrs_error():
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.Linter().lint(b"not a zip", extension="xlsx")


def test_stats(workbook):
    stats = sheetrs.stats(workbook)
    assert stats["total_sheets"] == len(stats["sheets"]) >= 1
    assert stats["total_used_cells"] == sum(s["used_cells"] for s in stats["sheets"])
    assert stats["total_file_size"] == workbook.stat().st_size


def test_stats_from_bytes(workbook):
    stats = sheetrs.stats(workbook.read_bytes(), extension=workbook.suffix[1:])
    assert stats["total_file_size"] == 0
    assert stats["sheets"] == sheetrs.stats(workbook)["sheets"]


def sheet_names(path):
    return [s["name"] for s in sheetrs.stats(path)["sheets"]]


def test_the_api_is_flat_and_has_one_way_to_do_each_thing():
    assert set(sheetrs.__all__) >= {
        "Linter", "Violation", "stats", "read", "scan", "Workbook", "LazyWorkbook",
    }
    for gone in ("lint", "edit", "EditPlan", "linter"):
        assert not hasattr(sheetrs, gone)


# ---- eager and lazy: the same methods, the same meaning --------------------------------------


def test_read_loads_and_never_modifies_the_file(workbook, tmp_path):
    before = workbook.read_bytes()
    wb = sheetrs.read(workbook).remove_sheets(sheet_names(workbook)[-1])
    wb.write(tmp_path / f"out{workbook.suffix}")
    assert workbook.read_bytes() == before


def test_eager_runs_when_called_and_the_result_can_be_inspected(workbook):
    names = sheet_names(workbook)
    original = sheetrs.read(workbook)
    smaller = original.remove_sheets(names[-1])
    # The call ran: the result is a workbook with one sheet less, and it can be measured as is.
    assert sheet_names(smaller) == names[:-1]
    assert sheetrs.Linter().lint(smaller) is not None
    # The workbook it was called on does not change.
    assert sheet_names(original) == names


def test_eager_errors_show_where_they_are_made(workbook):
    wb = sheetrs.read(workbook)
    with pytest.raises(sheetrs.SheetrsError):
        wb.delete_rows("NoSuchSheet!1")  # raised here, not at a later write
    with pytest.raises(ValueError, match="row range"):
        wb.delete_rows("T!B:D")


def test_lazy_reads_nothing_until_it_runs(workbook, tmp_path):
    names = sheet_names(workbook)
    lazy = sheetrs.scan(workbook).remove_sheets(names[-1])
    output = tmp_path / f"out{workbook.suffix}"
    assert not output.exists()
    lazy.sink(output)
    assert sheet_names(output) == names[:-1]
    # The plan is not consumed, and what it was called on is not changed.
    assert sheetrs.scan(workbook).operations == []
    lazy.sink(tmp_path / f"again{workbook.suffix}")


def test_lazy_errors_show_when_the_plan_runs(workbook, tmp_path):
    lazy = sheetrs.scan(workbook).delete_rows("NoSuchSheet!1")  # recorded, not run
    with pytest.raises(sheetrs.SheetrsError):
        lazy.sink(tmp_path / f"out{workbook.suffix}")
    with pytest.raises(sheetrs.SheetrsError, match="No operations"):
        sheetrs.scan(workbook).sink(tmp_path / f"o{workbook.suffix}")


def test_scan_checks_only_the_path(tmp_path):
    with pytest.raises(FileNotFoundError):
        sheetrs.scan(tmp_path / "missing.xlsx")
    with pytest.raises(ValueError, match="XLSX, XLSM or ODS"):
        sheetrs.scan(tmp_path)


def test_read_checks_the_file_and_takes_bytes(workbook, tmp_path):
    with pytest.raises(FileNotFoundError):
        sheetrs.read(tmp_path / "missing.xlsx")
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not a zip")
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.read(broken)
    wb = sheetrs.read(workbook.read_bytes(), extension=workbook.suffix[1:])
    assert sheet_names(wb) == sheet_names(workbook)
    with pytest.raises(ValueError, match="extension"):
        sheetrs.read(workbook.read_bytes())


def test_a_workbook_never_overwrites_the_file_it_reads(workbook):
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.read(workbook).write(workbook)
    with pytest.raises(sheetrs.SheetrsError):
        sheetrs.scan(workbook).recalc_on_open().sink(workbook)


def test_the_scopes_of_a_call_are_simultaneous_and_calls_are_sequential(structure, tmp_path):
    def rows(wb):
        return {k: v for k, v in sorted(cells(wb).items()) if k[0] == "F"}

    # Row 3 deleted twice is the rows 3 and 4; "T!3, T!4" at once is the same two rows.
    twice = sheetrs.read(structure).delete_rows("T!3").delete_rows("T!3")
    at_once = sheetrs.read(structure).delete_rows("T!3", "T!4")
    assert rows(twice) == rows(at_once)


def cells(wb):
    # The statistics are enough to tell two workbooks apart here.
    return {("F", i, s["name"]): (s["used_cells"], s["formulas"], s["errors"])
            for i, s in enumerate(sheetrs.stats(wb)["sheets"])}


def test_lazy_and_eager_give_the_same_workbook(structure, tmp_path):
    steps = [
        ("convert_to_values", ("F",)),
        ("delete_rows", ("T!2",)),
        ("delete_rows", ("T!3",)),   # the input's row 4
        ("delete_columns", ("U!B",)),
        ("remove_ranges", ("NameOne",)),
        ("remove_sheets", ("V",)),
        ("recalc_on_open", ()),
    ]
    eager = sheetrs.read(structure)
    lazy = sheetrs.scan(structure)
    for name, args in steps:
        eager = getattr(eager, name)(*args)
        lazy = getattr(lazy, name)(*args)
    assert lazy.operations[0].startswith("convert formulas to values")
    out = tmp_path / f"lazy{structure.suffix}"
    lazy.sink(out)
    assert sheetrs.stats(out)["sheets"] == sheetrs.stats(eager)["sheets"]
    assert sheetrs.stats(lazy.collect())["sheets"] == sheetrs.stats(eager)["sheets"]


def test_explain_shows_the_passes_and_a_conversion_after_a_change_needs_another(structure):
    one = sheetrs.scan(structure).delete_rows("T!2").delete_columns("U!B").remove_sheets("V")
    assert one.explain().startswith("pass 1 of 1:")
    two = sheetrs.scan(structure).delete_rows("T!2").convert_to_values("F")
    assert two.explain().startswith("pass 1 of 2:")
    assert "pass 2 of 2:" in two.explain()
    gone = sheetrs.scan(structure).remove_sheets("U").delete_rows("U!2")
    with pytest.raises(ValueError, match="removed"):
        gone.explain()


def test_a_3d_reference_stops_a_structural_edit_and_names_the_issue(tmp_path):
    from zipfile import ZipFile

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
    with pytest.raises(sheetrs.SheetrsError, match="issues/75"):
        sheetrs.read(broken).delete_rows("T!4:5")
    with pytest.raises(sheetrs.SheetrsError, match="issues/75"):
        sheetrs.scan(broken).delete_rows("T!4:5").sink(tmp_path / "out.xlsx")
    # What changes no structure is not stopped.
    sheetrs.read(broken).recalc_on_open().write(tmp_path / "ok.xlsx")
