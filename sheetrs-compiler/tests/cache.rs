//! `Cache` tells a recalculation what the plan before it changed: only the formulas those changes reach,
//! and the ones that depend on them, are calculated; every other formula keeps the result the workbook
//! caches. The formulas that remain in the program (`Program::formula_cells`) are the ones calculated,
//! and `Report::formulas_cached` counts the ones that keep their result.

use serde_json::json;
use sheetrs_compiler::memory::Options;
use sheetrs_compiler::model::ModelCell;
use sheetrs_compiler::{Cache, Model, Spec, compile};
use sheetrs_engine::addr::CellRef;
use sheetrs_engine::ir::Mode;
use sheetrs_engine::value::Value;
use std::collections::BTreeSet;

const DATA: u16 = 0;
const CALC: u16 = 1;

/// A formula cell: `(sheet, row, column)`, 0-based.
type At = (u16, u32, u32);

/// The result a formula caches in these workbooks: not what any of them gives.
fn cached() -> Value {
    Value::Number(999.0)
}

/// `Data` holds 1, 2, 3 in `A1:A3` and `Calc` has no formula yet.
fn model() -> Model {
    let mut model = Model::new();
    let data = model.add_sheet("Data");
    model.add_sheet("Calc");
    for (row, value) in ["A1", "A2", "A3"].into_iter().zip([1.0, 2.0, 3.0]) {
        model.set(data, row, Value::Number(value));
    }
    model
}

/// A formula in `Calc` (`B1`...) that caches `value`.
fn put(model: &mut Model, a1: &str, text: &str, value: Value) {
    let (row, col) = position(a1);
    model.cells.insert(
        CellRef::new(CALC, row, col),
        ModelCell {
            value,
            formula: Some(text.to_string()),
            array: false,
            recalculated_on_load: false,
        },
    );
}

fn position(a1: &str) -> (u32, u32) {
    let area = sheetrs_engine::addr::parse_area(a1).unwrap();
    (area.r0, area.c0)
}

fn calc(a1: &str) -> At {
    let (row, col) = position(a1);
    (CALC, row, col)
}

/// Forty formulas that nothing reaches, so that what a test adds is far from a quarter of them.
fn fillers(model: &mut Model) {
    for k in 1..=40 {
        put(model, &format!("F{k}"), "Data!A3*2", cached());
    }
}

/// The rows `first..=last` (0-based) of `sheet` were deleted.
fn rows_gone(sheet: &str, first: u32, last: u32) -> Cache {
    Cache {
        rows_deleted: vec![(sheet.to_string(), first, last)],
        ..Cache::default()
    }
}

/// The columns `first..=last` (0-based) of `sheet` were deleted.
fn columns_gone(sheet: &str, first: u32, last: u32) -> Cache {
    Cache {
        columns_deleted: vec![(sheet.to_string(), first, last)],
        ..Cache::default()
    }
}

fn change(sheet: &str, r0: u32, c0: u32, r1: u32, c1: u32) -> Cache {
    Cache {
        changed: vec![(sheet.to_string(), r0, c0, r1, c1)],
        ..Cache::default()
    }
}

/// The formulas the recalculation calculates, and how many keep their result.
fn run(model: &Model, cache: Cache) -> (BTreeSet<At>, usize) {
    let spec = Spec {
        mode: Mode::Recalculate,
        cache,
        ..Spec::default()
    };
    let (program, report) = compile(model, &spec).unwrap();
    let reached = program
        .formula_cells
        .iter()
        .flat_map(sheetrs_engine::ir::CellRun::cells)
        .map(|at| (at.sheet, at.row, at.col))
        .collect();
    (reached, report.formulas_cached)
}

fn reached(model: &Model, cache: Cache) -> BTreeSet<At> {
    run(model, cache).0
}

fn all(model: &Model) -> BTreeSet<At> {
    model
        .cells
        .iter()
        .filter(|(_, cell)| cell.formula.is_some())
        .map(|(at, _)| (at.sheet, at.row, at.col))
        .collect()
}

fn set(cells: &[&str]) -> BTreeSet<At> {
    cells.iter().map(|a1| calc(a1)).collect()
}

#[test]
fn a_cache_with_no_change_keeps_every_formula() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", cached());
    put(&mut model, "B2", "B1+1", cached());
    let (reached, kept) = run(&model, Cache::default());
    assert!(reached.is_empty(), "{reached:?}");
    assert_eq!(kept, 2);
}

#[test]
fn a_changed_cell_reaches_its_readers_and_their_dependents() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", cached());
    put(&mut model, "B2", "B1+1", cached());
    put(&mut model, "B3", "SUM(B1:B2)", cached());
    put(&mut model, "B4", "Data!A3*2", cached());
    fillers(&mut model);
    // Data!A1
    let (reached, kept) = run(&model, change("Data", 0, 0, 0, 0));
    assert_eq!(reached, set(&["B1", "B2", "B3"]));
    assert_eq!(kept, 41);
}

#[test]
fn a_changed_cell_reaches_the_formulas_that_read_a_range_that_holds_it() {
    let mut model = model();
    put(&mut model, "B1", "SUM(Data!A1:A3)", cached());
    put(&mut model, "B2", "SUM(Data!A1:A1)", cached());
    fillers(&mut model);
    // Data!A2
    assert_eq!(reached(&model, change("Data", 1, 0, 1, 0)), set(&["B1"]));
}

#[test]
fn a_changed_rectangle_of_another_sheet_reaches_nothing() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", cached());
    fillers(&mut model);
    // Calc!A1:A3, which no formula reads.
    assert!(reached(&model, change("Calc", 0, 0, 2, 0)).is_empty());
    // A sheet that does not exist: nothing to reach.
    assert!(reached(&model, change("Nowhere", 0, 0, 2, 0)).is_empty());
}

#[test]
fn a_formula_inside_a_changed_rectangle_is_calculated() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", cached());
    put(&mut model, "B2", "Data!A2*2", cached());
    fillers(&mut model);
    // Calc!B1 itself (a replaced formula, a result written into it).
    assert_eq!(reached(&model, change("Calc", 0, 1, 0, 1)), set(&["B1"]));
}

#[test]
fn a_formula_with_no_result_is_calculated_with_what_depends_on_it() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", Value::Empty);
    put(&mut model, "B2", "B1+1", cached());
    put(&mut model, "B3", "Data!A2*2", cached());
    fillers(&mut model);
    assert_eq!(reached(&model, Cache::default()), set(&["B1", "B2"]));
}

#[test]
fn volatile_formulas_are_calculated_with_what_depends_on_them() {
    for volatile in ["NOW()", "TODAY()", "RAND()", "RANDBETWEEN(1,6)"] {
        let mut model = model();
        put(&mut model, "B1", volatile, cached());
        put(&mut model, "B2", "B1+1", cached());
        put(&mut model, "B3", "Data!A2*2", cached());
        fillers(&mut model);
        assert_eq!(
            reached(&model, Cache::default()),
            set(&["B1", "B2"]),
            "{volatile}"
        );
    }
}

/// The rows `first..=last` (0-based) of `Data` changed their visibility.
fn hidden(first: u32, last: u32) -> Cache {
    Cache {
        rows_hidden: vec![("Data".to_string(), first, last)],
        ..Cache::default()
    }
}

#[test]
fn subtotal_is_calculated_when_a_row_it_reads_is_hidden_and_aggregate_is_not() {
    let mut model = model();
    put(&mut model, "B1", "SUBTOTAL(109,Data!A1:A3)", cached());
    put(&mut model, "B2", "AGGREGATE(9,5,Data!A1:A3)", cached());
    put(&mut model, "B3", "SUM(Data!A1:A3)", cached());
    put(&mut model, "B4", "B1+1", cached());
    put(&mut model, "B6", "SUBTOTAL(109,Data!A2:A3)", cached());
    fillers(&mut model);
    assert!(reached(&model, Cache::default()).is_empty());
    // Row 2 of `Data` (index 1): `LibreOffice` calculates both SUBTOTAL (B1 reads it and so does B6) and what reads B1.
    assert_eq!(reached(&model, hidden(1, 1)), set(&["B1", "B4", "B6"]));
    // Row 1: B6 does not read it.
    assert_eq!(reached(&model, hidden(0, 0)), set(&["B1", "B4"]));
    // A row of another place: nothing.
    assert!(reached(&model, hidden(10, 12)).is_empty());
}

/// OBSERVED with `LibreOffice` 26.8 in both formats (`write_to/probe_hide_groups.py`): of the formulas that are not SUBTOTAL, only
/// the ones in a group (adjacent formulas of a column that read the same relative cells) are calculated, and only those that read a range
/// with a row that changed; a formula alone, a group of single cells and cells that are not adjacent are not.
#[test]
fn hiding_a_row_calculates_the_formulas_of_a_group_that_read_it_and_no_other() {
    let mut model = model();
    // A group of two: `SUM(A1:A3)` and `SUM(A2:A4)` read the same cells relative to themselves.
    put(&mut model, "C1", "SUM(Data!A1:A3)", cached());
    put(&mut model, "C2", "SUM(Data!A2:A4)", cached());
    // The same two, with a row between: not adjacent.
    put(&mut model, "E1", "SUM(Data!A1:A3)", cached());
    put(&mut model, "E3", "SUM(Data!A2:A4)", cached());
    // Cells of a group of two that read single cells.
    put(&mut model, "G1", "Data!A2", cached());
    put(&mut model, "G2", "Data!A3", cached());
    // The same text twice is not the same shape (the references are relative).
    put(&mut model, "I1", "SUM(Data!A1:A3)", cached());
    put(&mut model, "I2", "SUM(Data!A1:A3)", cached());
    // The same text with absolute references is.
    put(&mut model, "K1", "SUM(Data!$A$1:$A$3)", cached());
    put(&mut model, "K2", "SUM(Data!$A$1:$A$3)", cached());
    fillers(&mut model);
    assert_eq!(
        reached(&model, hidden(1, 1)),
        set(&["C1", "C2", "K1", "K2"])
    );
    // Only the member that reads the row.
    assert_eq!(reached(&model, hidden(0, 0)), set(&["C1", "K1", "K2"]));
}

#[test]
fn whole_columns_of_a_sheet_that_lost_rows_are_calculated() {
    let mut model = model();
    put(&mut model, "B1", "COUNT(Data!A:A)", cached());
    put(&mut model, "B2", "SUM(Data!A1:A3)", cached());
    put(&mut model, "B3", "COUNT(Calc!H:H)", cached());
    fillers(&mut model);
    assert_eq!(reached(&model, rows_gone("Data", 1, 1)), set(&["B1"]));
    assert_eq!(reached(&model, rows_gone("Calc", 1, 1)), set(&["B3"]));
    assert!(reached(&model, rows_gone("Nowhere", 1, 1)).is_empty());
    // Columns that went do not shorten a column.
    assert!(reached(&model, columns_gone("Data", 1, 1)).is_empty());
}

#[test]
fn whole_rows_of_a_sheet_that_lost_columns_are_calculated() {
    let mut model = model();
    put(&mut model, "B1", "SUM(Data!1:1)", cached());
    put(&mut model, "B2", "SUM(Data!A1:C1)", cached());
    put(&mut model, "B3", "SUM(Calc!4:4)", cached());
    fillers(&mut model);
    assert_eq!(reached(&model, columns_gone("Data", 1, 1)), set(&["B1"]));
    assert_eq!(reached(&model, columns_gone("Calc", 1, 1)), set(&["B3"]));
    assert!(reached(&model, columns_gone("Nowhere", 1, 1)).is_empty());
    // Rows that went do not shorten a row.
    assert!(reached(&model, rows_gone("Data", 1, 1)).is_empty());
}

/// Observed (the file saved and opened again, both formats, `sheetrs-edit/tests/assets/write_to`, `dirty`): `ROW` and `COLUMN`
/// are calculated again when the cell of the formula or the cell or range they are given moved (or went), and
/// no other function of position is: `ROWS`, `COLUMNS`, `AREAS`, `SHEET` and `ADDRESS` keep the result when the cells only move.
#[test]
fn row_and_column_are_calculated_when_their_cell_or_what_they_read_moved() {
    // (formula, the deletion that moves what it reads or its own cell, a deletion that moves nothing of it)
    let cases = [
        ("ROW()", rows_gone("Calc", 2, 2), rows_gone("Calc", 9, 9)),
        (
            "COLUMN()",
            columns_gone("Calc", 0, 0),
            columns_gone("Calc", 3, 3),
        ),
        (
            "ROW(Data!A3)",
            rows_gone("Data", 1, 1),
            rows_gone("Data", 4, 4),
        ),
        (
            "COLUMN(Data!C1)",
            columns_gone("Data", 1, 1),
            columns_gone("Data", 3, 3),
        ),
        (
            "ROW(Data!A1:A3)",
            rows_gone("Data", 0, 0),
            rows_gone("Data", 4, 4),
        ),
        // The sheet of the cell is the one that matters.
        // And the cell of the formula moves, whatever it is given (`LibreOffice`: `COLUMN(Data!B1)` of a moved cell).
        (
            "ROW(Data!A3)",
            rows_gone("Calc", 1, 1),
            rows_gone("Calc", 9, 9),
        ),
    ];
    for (text, moves, elsewhere) in cases {
        let mut model = model();
        put(&mut model, "B5", text, cached());
        put(&mut model, "C5", "B5&\"\"", cached());
        put(&mut model, "B6", "Data!A2*2", cached());
        fillers(&mut model);
        assert!(reached(&model, Cache::default()).is_empty(), "{text}");
        assert_eq!(reached(&model, moves), set(&["B5", "C5"]), "{text}");
        assert!(reached(&model, elsewhere).is_empty(), "{text}");
    }
}

#[test]
fn the_other_functions_of_position_keep_their_result_when_cells_only_move() {
    let others = [
        "ROWS(Data!A3:A4)",
        "COLUMNS(Data!B1:C1)",
        "ADDRESS(1,1)",
        "SHEET()",
        "SHEETS()",
        "AREAS(Data!A3)",
        "Data!A3",
    ];
    for text in others {
        let mut model = model();
        put(&mut model, "B5", text, cached());
        fillers(&mut model);
        for gone in [
            rows_gone("Data", 1, 1),
            columns_gone("Data", 0, 0),
            rows_gone("Calc", 1, 1),
        ] {
            assert!(reached(&model, gone).is_empty(), "{text}");
        }
    }
}

/// `LibreOffice` calculates every formula of the workbook again when a sheet goes (the formula `ROW()` and the one that
/// reads no cell of the sheet included).
#[test]
fn a_deleted_sheet_reaches_every_formula() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", cached());
    put(&mut model, "B2", "ROW()", cached());
    put(&mut model, "B3", "5+5", cached());
    fillers(&mut model);
    let gone = Cache {
        sheets_deleted: true,
        ..Cache::default()
    };
    let (reached, kept) = run(&model, gone);
    assert_eq!(reached, all(&model));
    assert_eq!(kept, 0);
}

#[test]
fn a_formula_that_reads_through_a_name_is_calculated_when_the_name_is_stale() {
    let mut model = model();
    model.names.insert("rate".into(), "Data!$A$1".into());
    put(&mut model, "B1", "rate*2", cached());
    put(&mut model, "B2", "RATE+1", cached());
    put(&mut model, "B3", "Data!A2*2", cached());
    put(&mut model, "B4", "B1+1", cached());
    fillers(&mut model);
    assert!(reached(&model, Cache::default()).is_empty());
    let stale = |name: &str| Cache {
        stale_names: vec![name.to_string()],
        ..Cache::default()
    };
    // The name (written in any case), and not another one: a layout change alone leaves the formulas as they are.
    assert_eq!(reached(&model, stale("Rate")), set(&["B1", "B2", "B4"]));
    assert!(reached(&model, stale("other")).is_empty());
    assert!(reached(&model, rows_gone("Data", 1, 1)).is_empty());
}

/// A name that goes reaches every formula that reads through a name (any name of the workbook), and no other.
#[test]
fn a_deleted_name_reaches_every_formula_that_reads_through_a_name() {
    let mut model = model();
    model.names.insert("rate".into(), "Data!$A$1".into());
    model.names.insert("other".into(), "Data!$A$2".into());
    put(&mut model, "B1", "rate*2", cached());
    put(&mut model, "B2", "other+1", cached());
    put(&mut model, "B3", "Data!A2*2", cached());
    put(&mut model, "B4", "B2+1", cached());
    fillers(&mut model);
    let deleted = Cache {
        names_deleted: true,
        ..Cache::default()
    };
    assert_eq!(reached(&model, deleted), set(&["B1", "B2", "B4"]));
}

/// A deleted name: in an XLSX `LibreOffice` also calculates the formulas that hold a bare error `#REF!` (it reads it as a name), and
/// not the ones that hold a qualified one (`Data!#REF!`); an ODS has no such formula.
#[test]
fn a_deleted_name_reaches_the_formulas_with_a_bare_ref_error_of_an_xlsx() {
    let build = |xlsx: bool| {
        let mut model = model();
        model.from_xlsx = xlsx;
        put(&mut model, "B1", "COUNT(#REF!)", cached());
        put(&mut model, "B2", "COUNT(Data!#REF!)", cached());
        put(&mut model, "B3", "SUM(Data!A1:A3)", cached());
        fillers(&mut model);
        model
    };
    let deleted = Cache {
        names_deleted: true,
        ..Cache::default()
    };
    assert_eq!(reached(&build(true), deleted.clone()), set(&["B1"]));
    assert!(reached(&build(false), deleted).is_empty());
    assert!(reached(&build(true), Cache::default()).is_empty());
}

/// A formula that builds a reference at run time is calculated, as `LibreOffice` does when it opens the file (it is volatile for it),
/// and what depends on it; the formulas that read nothing it changes keep their result: what it may read holds what the file caches.
#[test]
fn a_reference_built_at_run_time_is_calculated_and_the_rest_is_kept() {
    for dynamic in ["INDIRECT(\"Data!A\"&Data!A1)", "OFFSET(Data!A1,Data!A2,0)"] {
        let mut model = model();
        put(&mut model, "B1", dynamic, cached());
        put(&mut model, "B2", "Data!A2*2", cached());
        fillers(&mut model);
        let (reached, kept) = run(&model, change("Data", 0, 0, 0, 0));
        assert_eq!(reached, set(&["B1"]), "{dynamic}");
        assert_eq!(kept, 41, "{dynamic}");
        let (reached, _) = run(&model, Cache::default());
        assert_eq!(reached, set(&["B1"]), "{dynamic}");
    }
}

/// With arguments that are known when the program is built the reference is a plain one, and the
/// formula is reached as any other that reads that cell.
#[test]
fn a_reference_known_at_build_time_is_an_ordinary_one() {
    let mut model = model();
    put(&mut model, "B1", "INDIRECT(\"Data!A1\")", cached());
    put(&mut model, "B2", "OFFSET(Data!A1,1,0)", cached());
    fillers(&mut model);
    let (first, kept) = run(&model, change("Data", 0, 0, 0, 0));
    assert_eq!(first, set(&["B1"]));
    assert_eq!(kept, 41);
    assert_eq!(reached(&model, change("Data", 1, 0, 1, 0)), set(&["B2"]));
}

/// Twenty formulas over `Data!B1:B20`: the quarter of them that a change may reach.
fn twenty() -> Model {
    let mut model = model();
    for row in 1..=20 {
        model.set(DATA, &format!("B{row}"), Value::Number(1.0));
        put(
            &mut model,
            &format!("D{row}"),
            &format!("Data!B{row}*2"),
            cached(),
        );
    }
    model
}

/// However many formulas a change reaches, the others keep their result: `LibreOffice` does not calculate the whole
/// workbook because many formulas are stale (an earlier version did where more than a quarter were).
#[test]
fn unreached_formulas_keep_their_result() {
    let model = twenty();
    for cells in [4, 5, 6, 7, 19] {
        let (reached, kept) = run(&model, change("Data", 0, 1, cells - 1, 1));
        assert_eq!(reached.len(), cells as usize, "{cells} cells");
        assert_eq!(kept, 20 - cells as usize, "{cells} cells");
    }
}

#[test]
fn the_formulas_with_a_result_keep_it_however_many_have_none() {
    let without = |count: usize| {
        let mut model = twenty();
        for row in 1..=count {
            put(
                &mut model,
                &format!("D{row}"),
                &format!("Data!B{row}*2"),
                Value::Empty,
            );
        }
        run(&model, Cache::default())
    };
    for count in [4, 5, 6, 7, 19] {
        let (reached, kept) = without(count);
        assert_eq!(reached.len(), count, "{count} formulas");
        assert_eq!(kept, 20 - count, "{count} formulas");
    }
}

/// The cells of an array formula are calculated together: the block is one formula whose cells the
/// engine fills at once, and what is stale for one cell is for all of them.
#[test]
fn the_cells_of_an_array_formula_are_reached_together() {
    let build = |master_result: Value, member_result: Value| {
        let mut model = model();
        let cell = |value: Value| ModelCell {
            value,
            formula: Some("Data!A1:A3*2".into()),
            array: true,
            recalculated_on_load: false,
        };
        // `Calc!B1:B3`: the first cell, and the two that show their own element of the result.
        model
            .cells
            .insert(CellRef::new(CALC, 0, 1), cell(master_result));
        for (row, offset) in [(1, 1), (2, 2)] {
            let at = CellRef::new(CALC, row, 1);
            model.cells.insert(at, cell(member_result.clone()));
            model.array_members.insert(at, (offset, 0));
        }
        put(&mut model, "D1", "B2+1", cached());
        put(&mut model, "D2", "Data!A2*2", cached());
        fillers(&mut model);
        model
    };
    // The writer drops the result of the first cell only: the members come along.
    let model = build(Value::Empty, cached());
    assert_eq!(
        reached(&model, Cache::default()),
        set(&["B1", "B2", "B3", "D1"])
    );
    // And the other way round: a member without a result reaches the first cell.
    let model = build(cached(), Value::Empty);
    assert_eq!(
        reached(&model, Cache::default()),
        set(&["B1", "B2", "B3", "D1"])
    );
    // With nothing stale the block keeps its results.
    let model = build(cached(), cached());
    assert!(reached(&model, Cache::default()).is_empty());
}

/// A name may hold a dot or an underscore, and a number in a formula that reads through none holds a dot.
#[test]
fn a_name_with_a_dot_is_found_in_the_formulas_that_read_through_it() {
    let mut model = model();
    model.names.insert("tax.rate".into(), "Data!$A$1".into());
    model.names.insert("tax_rate".into(), "Data!$A$2".into());
    put(&mut model, "B1", "tax.rate*2", cached());
    put(&mut model, "B4", "tax_rate*2", cached());
    put(&mut model, "B2", "Data!A2*0.5", cached());
    put(&mut model, "B3", "SUM(tax.rate,0.5)", cached());
    fillers(&mut model);
    let stale = Cache {
        stale_names: vec!["tax.rate".into(), "tax_rate".into()],
        ..Cache::default()
    };
    assert_eq!(reached(&model, stale), set(&["B1", "B3", "B4"]));
}

/// The cells of an array formula that cannot be placed in their block (the first cell is no formula, or
/// is outside the sheet) are calculated as any other formula.
#[test]
fn an_array_cell_with_no_block_to_join_is_left_to_its_own_result() {
    let mut model = model();
    let cell = |value: Value| ModelCell {
        value,
        formula: Some("Data!A1:A3*2".into()),
        array: true,
        recalculated_on_load: false,
    };
    // `Calc!B2` says its block starts one row above, where there is no formula; `Calc!B3`, five rows above.
    for (row, offset) in [(1, 1), (2, 5)] {
        let at = CellRef::new(CALC, row, 1);
        model.cells.insert(at, cell(Value::Empty));
        model.array_members.insert(at, (offset, 0));
    }
    put(&mut model, "D1", "Data!A2*2", cached());
    fillers(&mut model);
    assert_eq!(reached(&model, Cache::default()), set(&["B2", "B3"]));
}

#[test]
fn calculate_all_calculates_every_formula_of_a_recalculation() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", cached());
    fillers(&mut model);
    let spec = Spec {
        mode: Mode::Recalculate,
        calculate_all: true,
        ..Spec::default()
    };
    let (program, report) = compile(&model, &spec).unwrap();
    assert_eq!(report.formulas_cached, 0);
    assert_eq!(
        program
            .formula_cells
            .iter()
            .map(|run| run.cells().count())
            .sum::<usize>(),
        41
    );
}

#[test]
fn the_options_of_json_have_no_changes() {
    let options = Options::from_json(&serde_json::Value::Null).unwrap();
    assert_eq!(options.cache, Cache::default());
    let options = Options::from_json(&json!({"now": 45000.5})).unwrap();
    assert_eq!(options.cache, Cache::default());
    assert_eq!(options.config.now, Some(45000.5));
}

#[test]
fn formulas_that_ask_what_a_cell_holds_are_calculated_when_a_cell_changes() {
    let mut model = model();
    put(&mut model, "B1", "ISFORMULA(Data!A1)", cached());
    put(&mut model, "B2", "FORMULA(Data!A1)", cached());
    put(&mut model, "B3", "FORMULATEXT(Data!A1)", cached());
    put(&mut model, "B4", "_xlfn.ISFORMULA(Data!A1)&\"\"", cached());
    put(&mut model, "B5", "Data!A2*2", cached());
    put(&mut model, "B6", "B1&\"\"", cached());
    fillers(&mut model);
    let expected = set(&["B1", "B2", "B3", "B4", "B6"]);
    // `LibreOffice` calculates `FORMULATEXT` (`FORMULA` of an ODS) when it opens any file, and not `ISFORMULA` (OBSERVED, both formats).
    assert_eq!(reached(&model, Cache::default()), set(&["B2", "B3"]));
    // Observed (the file saved and opened again): cells that only move leave the result of `ISFORMULA` as it is.
    assert_eq!(reached(&model, rows_gone("Data", 1, 1)), set(&["B2", "B3"]));
    // The cell they ask about changed: the formulas that ask and what depends on them (B6 reads B1).
    assert_eq!(reached(&model, change("Data", 0, 0, 0, 0)), expected);
    // Another cell changed: only what reads it (`LibreOffice` keeps `ISFORMULA(Data!A1)` when `Data!A2` is replaced).
    assert_eq!(
        reached(&model, change("Data", 1, 0, 1, 0)),
        set(&["B2", "B3", "B5"])
    );
}

/// Both lists the analysis walks are in the order of the cells, and the formulas left as the file caches them (a function
/// of `LibreOffice` that the engine lacks, that nothing reaches) are not in one of them: each formula is judged by its own
/// cell, not by the one at its index.
#[test]
fn a_formula_left_as_cached_does_not_shift_the_judgment_of_the_others() {
    let mut model = model();
    put(&mut model, "B1", "WEBSERVICE(\"x\")", cached());
    put(&mut model, "B2", "Data!A1*2", cached());
    put(&mut model, "B3", "Data!A2*2", Value::Empty);
    fillers(&mut model);
    let spec = Spec {
        mode: Mode::Recalculate,
        ..Spec::default()
    };
    let (program, _) = compile(&model, &spec).unwrap();
    let reached: BTreeSet<At> = program
        .formula_cells
        .iter()
        .flat_map(sheetrs_engine::ir::CellRun::cells)
        .map(|at| (at.sheet, at.row, at.col))
        .collect();
    // Only the formula with no result is calculated.
    assert_eq!(reached, set(&["B3"]));
}

mod names {
    use sheetrs_compiler::compile::names_reached;
    use std::collections::HashMap;

    fn names(list: &[(&str, &str)]) -> HashMap<String, String> {
        list.iter()
            .map(|(name, text)| ((*name).to_string(), (*text).to_string()))
            .collect()
    }

    fn bands(sheet: &str, first: u32, last: u32) -> Vec<(String, u32, u32)> {
        vec![(sheet.to_string(), first, last)]
    }

    /// Observed (the file saved and opened again, both formats, `dirty`): a name is reached when a band begins at or before
    /// the last row (or column) of its target (it moved, shrank or went); not when the band is after it or on another sheet.
    #[test]
    fn a_name_is_reached_when_its_target_moved_shrank_or_went() {
        let defined = names(&[
            ("cell3", "Data!$A$3"),
            ("cell5", "Data!$A$5"),
            ("cell1", "Data!$A$1"),
            ("rng16", "Data!$A$1:$A$6"),
            ("rng12", "Data!$A$1:$A$2"),
            ("other", "Other!$A$1"),
            ("constant", "5"),
        ]);
        // Delete row 3 (0-based 2).
        assert_eq!(
            names_reached(&defined, &bands("Data", 2, 2), &[]),
            ["cell3", "cell5", "rng16"]
        );
        // Delete row 6 (0-based 5): only what reaches it.
        assert_eq!(
            names_reached(&defined, &bands("Data", 5, 5), &[]),
            ["rng16"]
        );
        // Delete row 7 and below: nothing of these.
        assert!(names_reached(&defined, &bands("Data", 6, 7), &[]).is_empty());
        // Delete row 1: every target of Data.
        assert_eq!(
            names_reached(&defined, &bands("Data", 0, 0), &[]),
            ["cell1", "cell3", "cell5", "rng12", "rng16"]
        );
    }

    #[test]
    fn a_name_is_reached_by_the_columns_that_go_before_or_inside_its_target() {
        let defined = names(&[
            ("b1", "Data!$B$1"),
            ("c1", "Data!$C$1"),
            ("d3", "Data!$D$3"),
            ("a_to_a", "Data!$A:$A"),
            ("row3", "Data!$3:$3"),
        ]);
        // Delete column C (2): C1 went, D3 moved, a whole row is shorter, B1 and column A are not touched.
        assert_eq!(
            names_reached(&defined, &[], &bands("Data", 2, 2)),
            ["c1", "d3", "row3"]
        );
        // Delete column A: every target is in it or after it.
        assert_eq!(
            names_reached(&defined, &[], &bands("Data", 0, 0)),
            ["a_to_a", "b1", "c1", "d3", "row3"]
        );
        // Rows that go make a whole column shorter (the last row of its target is the last of the sheet).
        assert_eq!(
            names_reached(&defined, &bands("Data", 9, 9), &[]),
            ["a_to_a"]
        );
    }

    /// A name made of other names is reached when one of them loses cells, not when it only moves.
    #[test]
    fn a_name_of_names_is_reached_by_lost_cells_not_by_moves() {
        let defined = names(&[
            ("cell_five", "Data!$A$5"),
            ("range_one_six", "Data!$A$1:$A$6"),
            ("by_cell", "cell_five+1"),
            ("by_range", "SUM(range_one_six)"),
            ("both", "cell_five+range_one_six"),
        ]);
        // Delete row 3: cell_five moves (reached, it is direct), the range shrinks (reached), the name that uses the
        // cell only sees it move.
        assert_eq!(
            names_reached(&defined, &bands("Data", 2, 2), &[]),
            ["both", "by_range", "cell_five", "range_one_six"]
        );
        // Delete row 5: the cell goes, so the names that use it are reached.
        assert_eq!(
            names_reached(&defined, &bands("Data", 4, 4), &[]),
            ["both", "by_cell", "by_range", "cell_five", "range_one_six"]
        );
    }
}

/// OBSERVED with `LibreOffice` 26.8, the default load of both formats (`write_to/probe_dirty.py classes`): it calculates `HYPERLINK`, every
/// type of `CELL` and `INFO` when it opens any file, and what reads them; `ROW`, `COLUMN`, a read of a whole column and `ISFORMULA` keep
/// the result the file caches.
#[test]
fn only_what_libreoffice_calculates_on_load_is_calculated() {
    let mut model = model();
    put(
        &mut model,
        "B1",
        "HYPERLINK(\"http://x.y\",\"t\")",
        cached(),
    );
    put(&mut model, "B2", "CELL(\"row\",Data!A1)", cached());
    put(&mut model, "B3", "INFO(\"system\")", cached());
    put(&mut model, "B4", "B1&\"!\"", cached());
    put(&mut model, "B5", "ROW()", cached());
    put(&mut model, "B6", "COUNT(Data!A:A)", cached());
    put(&mut model, "B7", "COLUMN(Data!B1)", cached());
    put(&mut model, "B8", "ISFORMULA(Data!A1)", cached());
    fillers(&mut model);
    assert_eq!(
        reached(&model, Cache::default()),
        set(&["B1", "B2", "B3", "B4"])
    );
}

#[test]
fn a_formula_calculated_on_load_does_not_reach_its_readers() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1+1", cached());
    put(&mut model, "B2", "B1+1", cached());
    put(&mut model, "B3", "Data!A2+1", cached());
    // `LibreOffice` calculates B1 when it opens the file; B2, which reads it and caches a result it keeps, shows what the file caches.
    model.cells.insert(
        CellRef::new(CALC, 0, 1),
        ModelCell {
            value: cached(),
            formula: Some("Data!A1+1".to_string()),
            array: false,
            recalculated_on_load: true,
        },
    );
    fillers(&mut model);
    assert_eq!(reached(&model, Cache::default()), set(&["B1"]));
}
