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
        cache: Some(cache),
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
fn a_changed_cell_reaches_the_formulas_that_read_it_and_those_that_depend_on_them() {
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
fn a_formula_inside_a_changed_rectangle_is_calculated_though_nothing_else_reads_it() {
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

#[test]
fn subtotal_and_aggregate_are_calculated_when_rows_are_hidden() {
    let mut model = model();
    put(&mut model, "B1", "SUBTOTAL(109,Data!A1:A3)", cached());
    put(&mut model, "B2", "AGGREGATE(9,5,Data!A1:A3)", cached());
    put(&mut model, "B3", "SUM(Data!A1:A3)", cached());
    put(&mut model, "B4", "B1+1", cached());
    fillers(&mut model);
    assert!(reached(&model, Cache::default()).is_empty());
    let hidden = Cache {
        rows_hidden: true,
        ..Cache::default()
    };
    assert_eq!(reached(&model, hidden), set(&["B1", "B2", "B4"]));
}

#[test]
fn whole_columns_of_a_sheet_that_lost_rows_are_calculated() {
    let mut model = model();
    put(&mut model, "B1", "COUNT(Data!A:A)", cached());
    put(&mut model, "B2", "SUM(Data!A1:A3)", cached());
    put(&mut model, "B3", "COUNT(Calc!H:H)", cached());
    fillers(&mut model);
    let lost = |sheet: &str| Cache {
        rows_deleted: vec![sheet.to_string()],
        ..Cache::default()
    };
    assert_eq!(reached(&model, lost("Data")), set(&["B1"]));
    assert_eq!(reached(&model, lost("Calc")), set(&["B3"]));
    assert!(reached(&model, lost("Nowhere")).is_empty());
    // Columns that went do not shorten a column.
    let columns = Cache {
        columns_deleted: vec!["Data".to_string()],
        ..Cache::default()
    };
    assert!(reached(&model, columns).is_empty());
}

#[test]
fn whole_rows_of_a_sheet_that_lost_columns_are_calculated() {
    let mut model = model();
    put(&mut model, "B1", "SUM(Data!1:1)", cached());
    put(&mut model, "B2", "SUM(Data!A1:C1)", cached());
    put(&mut model, "B3", "SUM(Calc!4:4)", cached());
    fillers(&mut model);
    let lost = |sheet: &str| Cache {
        columns_deleted: vec![sheet.to_string()],
        ..Cache::default()
    };
    assert_eq!(reached(&model, lost("Data")), set(&["B1"]));
    assert_eq!(reached(&model, lost("Calc")), set(&["B3"]));
    assert!(reached(&model, lost("Nowhere")).is_empty());
    // Rows that went do not shorten a row.
    let rows = Cache {
        rows_deleted: vec!["Data".to_string()],
        ..Cache::default()
    };
    assert!(reached(&model, rows).is_empty());
}

#[test]
fn the_functions_of_position_are_calculated_when_the_layout_changed() {
    let positional = [
        "ROW()",
        "ROWS(Data!A1:A3)",
        "COLUMN()",
        "COLUMNS(Data!A1:C1)",
        "ADDRESS(1,1)",
        "CELL(\"address\",Data!A1)",
        "SHEET()",
        "SHEETS()",
        "AREAS(Data!A1)",
    ];
    let layout = Cache {
        layout: true,
        ..Cache::default()
    };
    for text in positional {
        let mut model = model();
        put(&mut model, "B1", text, cached());
        put(&mut model, "B2", "B1&\"\"", cached());
        put(&mut model, "B3", "Data!A2*2", cached());
        fillers(&mut model);
        assert!(
            reached(&model, Cache::default()).is_empty(),
            "{text} without a change of layout"
        );
        assert_eq!(
            reached(&model, layout.clone()),
            set(&["B1", "B2"]),
            "{text}"
        );
    }
}

#[test]
fn a_formula_that_reads_through_a_name_is_calculated_when_the_layout_or_the_name_changed() {
    let mut model = model();
    model.names.insert("rate".into(), "Data!$A$1".into());
    put(&mut model, "B1", "rate*2", cached());
    put(&mut model, "B2", "RATE+1", cached());
    put(&mut model, "B3", "Data!A2*2", cached());
    put(&mut model, "B4", "B1+1", cached());
    fillers(&mut model);
    assert!(reached(&model, Cache::default()).is_empty());
    let layout = Cache {
        layout: true,
        ..Cache::default()
    };
    assert_eq!(reached(&model, layout), set(&["B1", "B2", "B4"]));
    let deleted = |name: &str| Cache {
        names_deleted: vec![name.to_string()],
        ..Cache::default()
    };
    // The name that went (written in any case), and not another one.
    assert_eq!(reached(&model, deleted("Rate")), set(&["B1", "B2", "B4"]));
    assert!(reached(&model, deleted("other")).is_empty());
}

/// A formula that builds a reference at run time can read any cell: nothing is out of reach, and every
/// formula is calculated, as without a cache.
#[test]
fn a_reference_built_at_run_time_calculates_the_whole_workbook() {
    for dynamic in ["INDIRECT(\"Data!A\"&Data!A1)", "OFFSET(Data!A1,Data!A2,0)"] {
        let mut model = model();
        put(&mut model, "B1", dynamic, cached());
        put(&mut model, "B2", "Data!A2*2", cached());
        fillers(&mut model);
        let (reached, kept) = run(&model, change("Data", 0, 0, 0, 0));
        assert_eq!(reached, all(&model), "{dynamic}");
        assert_eq!(kept, 0, "{dynamic}");
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

#[test]
fn more_than_a_quarter_of_the_formulas_reached_calculates_them_all() {
    let model = twenty();
    // 4 and 5 of 20: the cache is used, 5 is the quarter.
    for cells in [4, 5] {
        let (reached, kept) = run(&model, change("Data", 0, 1, cells - 1, 1));
        assert_eq!(reached.len(), cells as usize);
        assert_eq!(kept, 20 - cells as usize);
    }
    // 6 of 20: the walk would only add its cost.
    for cells in [6, 7] {
        let (reached, kept) = run(&model, change("Data", 0, 1, cells - 1, 1));
        assert_eq!(reached, all(&model), "{cells} cells");
        assert_eq!(kept, 0, "{cells} cells");
    }
}

#[test]
fn more_than_a_quarter_of_the_formulas_without_a_result_calculates_them_all() {
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
    for count in [4, 5] {
        let (reached, kept) = without(count);
        assert_eq!(reached.len(), count);
        assert_eq!(kept, 20 - count);
    }
    for count in [6, 7] {
        let (reached, kept) = without(count);
        assert_eq!(reached.len(), 20, "{count} formulas");
        assert_eq!(kept, 0, "{count} formulas");
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
    let layout = Cache {
        layout: true,
        ..Cache::default()
    };
    assert_eq!(reached(&model, layout), set(&["B1", "B3", "B4"]));
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

/// Only the recalculation of an edit reads the cache: any other program is built whole.
#[test]
fn a_program_that_is_not_a_recalculation_calculates_every_formula() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", cached());
    fillers(&mut model);
    for mode in [Mode::Values, Mode::Embedded] {
        let spec = Spec {
            mode,
            cache: Some(Cache::default()),
            ..Spec::default()
        };
        let (_, report) = compile(&model, &spec).unwrap();
        assert_eq!(report.formulas_cached, 0, "{mode:?}");
    }
}

#[test]
fn a_run_without_a_cache_calculates_every_formula() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", cached());
    fillers(&mut model);
    let spec = Spec {
        mode: Mode::Recalculate,
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
fn the_options_of_json_have_no_cache() {
    let options = Options::from_json(&serde_json::Value::Null).unwrap();
    assert!(options.cache.is_none());
    let options = Options::from_json(&json!({"now": 45000.5})).unwrap();
    assert!(options.cache.is_none());
    assert_eq!(options.config.now, Some(45000.5));
}

/// `FORMULA`, `ISFORMULA` and `FORMULATEXT` of a literal reference are answered while the formula is
/// lowered, so it keeps no reference to the cell: it is stale when any cell changes or moves.
#[test]
fn formulas_that_ask_what_a_cell_holds_are_calculated_when_a_cell_changes_or_moves() {
    let mut model = model();
    put(&mut model, "B1", "ISFORMULA(Data!A1)", cached());
    put(&mut model, "B2", "FORMULA(Data!A1)", cached());
    put(&mut model, "B3", "FORMULATEXT(Data!A1)", cached());
    put(&mut model, "B4", "_xlfn.ISFORMULA(Data!A1)&\"\"", cached());
    put(&mut model, "B5", "Data!A2*2", cached());
    put(&mut model, "B6", "B1&\"\"", cached());
    fillers(&mut model);
    let expected = set(&["B1", "B2", "B3", "B4", "B6"]);
    assert!(reached(&model, Cache::default()).is_empty());
    let layout = Cache {
        layout: true,
        ..Cache::default()
    };
    assert_eq!(reached(&model, layout), expected.clone());
    // A replaced cell, wherever it is: the formula does not say which cell it asks about.
    let mut with_reader = expected;
    with_reader.insert(calc("B5"));
    assert_eq!(reached(&model, change("Data", 1, 0, 1, 0)), with_reader);
}
