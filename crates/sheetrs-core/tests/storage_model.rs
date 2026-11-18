use std::collections::BTreeMap;

use proptest::prelude::*;
use sheetrs_core::{Cell, CellPos, ErrorKind, Formula, Value, Visibility, Workbook};

fn value() -> impl Strategy<Value = Value> {
    prop_oneof![
        Just(Value::Empty),
        any::<f64>()
            .prop_filter("not a number", |number| !number.is_nan())
            .prop_map(Value::Number),
        any::<bool>().prop_map(Value::Bool),
        prop::sample::select(vec!["", "a", "b", "a longer text", "ñ"])
            .prop_map(|text| Value::Text(text.into())),
        prop_oneof![
            Just(ErrorKind::Null),
            Just(ErrorKind::DivZero),
            Just(ErrorKind::Value),
            Just(ErrorKind::Ref),
            Just(ErrorKind::Name),
            Just(ErrorKind::Num),
            Just(ErrorKind::Na),
            prop::sample::select(vec!["#SPILL!", "Err:509"])
                .prop_map(|text| ErrorKind::Other(text.into())),
        ]
        .prop_map(Value::Error),
    ]
}

fn formula() -> impl Strategy<Value = Option<Formula>> {
    prop_oneof![
        Just(None),
        prop::sample::select(vec!["A1+1", "SUM(A1:A3)", ""])
            .prop_map(|text| Some(Formula::Text(text.to_owned()))),
        prop_oneof![0_u32..5, Just(1 << 31), Just(u32::MAX)]
            .prop_map(|index| Some(Formula::Shared { index })),
    ]
}

fn cell() -> impl Strategy<Value = Cell> {
    (value(), formula(), any::<bool>()).prop_map(|(value, formula, stale)| Cell {
        value,
        formula,
        stale,
    })
}

fn position() -> impl Strategy<Value = CellPos> {
    prop_oneof![(0_u32..40, 0_u32..12), (0_u32..1_048_576, 0_u32..16_384)]
        .prop_map(|(row, col)| CellPos { row, col })
}

/// Cells compare by their debug text, which tells `-0.0` from `0.0` and keeps NaN payloads.
fn shown(cell: &Cell) -> String {
    format!("{cell:?}")
}

proptest! {
    #[test]
    fn columns_behave_like_an_ordered_map(
        entries in prop::collection::vec((position(), cell()), 0..200),
        absent in prop::collection::vec(position(), 0..20),
    ) {
        let mut workbook = Workbook::new();
        let mut builder = workbook.sheet_builder("S1");
        let mut reference: BTreeMap<CellPos, Cell> = BTreeMap::new();
        for (pos, cell) in entries {
            builder.set(pos, &cell).expect("within limits");
            reference.insert(pos, cell);
        }
        builder.finish();
        let sheet = workbook.sheet("S1").expect("the sheet was added");

        prop_assert_eq!(sheet.len(), reference.len());
        prop_assert_eq!(sheet.is_empty(), reference.is_empty());
        for (pos, cell) in &reference {
            let view = sheet.get(*pos).expect("a stored cell is found");
            prop_assert_eq!(shown(&view.to_cell()), shown(cell));
        }
        for pos in absent {
            if !reference.contains_key(&pos) {
                prop_assert!(sheet.get(pos).is_none());
            }
        }

        let by_row: Vec<(CellPos, String)> = sheet
            .cells()
            .map(|(pos, view)| (pos, shown(&view.to_cell())))
            .collect();
        let expected_by_row: Vec<(CellPos, String)> = reference
            .iter()
            .map(|(pos, cell)| (*pos, shown(cell)))
            .collect();
        prop_assert_eq!(by_row, expected_by_row.clone());

        let mut expected_by_column = expected_by_row;
        expected_by_column.sort_by_key(|(pos, _)| (pos.col, pos.row));
        let by_column: Vec<(CellPos, String)> = sheet
            .cells_by_column()
            .map(|(pos, view)| (pos, shown(&view.to_cell())))
            .collect();
        prop_assert_eq!(by_column, expected_by_column);
    }
}

fn single_cell(content: &Cell) -> Workbook {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    builder
        .set(CellPos { row: 0, col: 0 }, content)
        .expect("within limits");
    builder.finish();
    workbook
}

fn number(value: f64) -> Cell {
    Cell {
        value: Value::Number(value),
        formula: None,
        stale: false,
    }
}

fn first_number(workbook: &Workbook) -> Option<f64> {
    let sheet = workbook.sheet("S1")?;
    match sheet.get(CellPos { row: 0, col: 0 })?.to_cell().value {
        Value::Number(value) => Some(value),
        _ => None,
    }
}

#[test]
fn numbers_keep_every_bit() {
    for bits in [
        (-0.0_f64).to_bits(),
        0.0_f64.to_bits(),
        f64::MIN_POSITIVE.to_bits() / 2,
        f64::INFINITY.to_bits(),
        f64::NEG_INFINITY.to_bits(),
        0x7ff8_0000_0000_0001,
        0.1_f64.to_bits(),
    ] {
        let workbook = single_cell(&number(f64::from_bits(bits)));
        assert_eq!(first_number(&workbook).map(f64::to_bits), Some(bits));
    }
}

#[test]
fn a_later_cell_replaces_the_earlier_one_at_the_same_address() {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    let pos = CellPos { row: 3, col: 2 };
    builder.set(pos, &number(1.0)).expect("within limits");
    builder.set(pos, &number(2.0)).expect("within limits");
    builder.finish();
    let sheet = workbook.sheet("S1").expect("the sheet was added");
    assert_eq!(sheet.len(), 1);
    assert_eq!(
        sheet.get(pos).map(|view| view.to_cell().value),
        Some(Value::Number(2.0))
    );
}

#[test]
fn sheets_keep_their_order_names_and_visibility() {
    let mut workbook = Workbook::new();
    for (name, visibility) in [
        ("S1", Visibility::Visible),
        ("S2", Visibility::Hidden),
        ("S3", Visibility::VeryHidden),
    ] {
        let mut builder = workbook.sheet_builder(name);
        builder.set_visibility(visibility);
        builder.finish();
    }
    let summary: Vec<(String, Visibility)> = workbook
        .sheets()
        .map(|sheet| (sheet.name().to_owned(), sheet.visibility()))
        .collect();
    assert_eq!(
        summary,
        vec![
            ("S1".to_owned(), Visibility::Visible),
            ("S2".to_owned(), Visibility::Hidden),
            ("S3".to_owned(), Visibility::VeryHidden),
        ]
    );
    assert!(workbook.sheet("S4").is_none());
}

#[test]
fn shared_formula_anchors_are_kept_per_sheet() {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    builder.set_shared_anchor(2, CellPos { row: 4, col: 1 });
    builder.finish();
    let sheet = workbook.sheet("S1").expect("the sheet was added");
    assert_eq!(sheet.shared_anchor(2), Some(CellPos { row: 4, col: 1 }));
    assert_eq!(sheet.shared_anchor(3), None);
}

#[test]
fn a_builder_reports_whether_a_cell_already_holds_a_formula() {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    let pos = CellPos { row: 0, col: 0 };
    assert!(!builder.has_formula_at(pos));
    builder
        .set(
            pos,
            &Cell {
                value: Value::Empty,
                formula: Some(Formula::Text("1+1".to_owned())),
                stale: true,
            },
        )
        .expect("within limits");
    assert!(builder.has_formula_at(pos));
    builder
        .set(CellPos { row: 1, col: 0 }, &number(1.0))
        .expect("within limits");
    assert!(!builder.has_formula_at(CellPos { row: 1, col: 0 }));
}

#[test]
fn absorbing_a_sheet_of_another_workbook_translates_its_texts_and_formulas() {
    let text = |value: &str, formula: Option<Formula>, stale: bool| Cell {
        value: Value::Text(value.into()),
        formula,
        stale,
    };
    let mut other = Workbook::new();
    let mut builder = other.sheet_builder("P");
    builder
        .set(
            CellPos { row: 0, col: 0 },
            &text("b", Some(Formula::Text("A1+1".to_owned())), true),
        )
        .expect("within limits");
    builder
        .set(
            CellPos { row: 1, col: 0 },
            &Cell {
                value: Value::Error(ErrorKind::Other("#SPILL!".into())),
                formula: Some(Formula::Shared { index: u32::MAX }),
                stale: false,
            },
        )
        .expect("within limits");
    builder
        .set(CellPos { row: 2, col: 0 }, &text("a", None, false))
        .expect("within limits");
    builder.set_shared_anchor(7, CellPos { row: 1, col: 0 });
    builder.finish();

    let mut main = Workbook::new();
    let mut builder = main.sheet_builder("S1");
    builder
        .set(
            CellPos { row: 0, col: 1 },
            &text("a", Some(Formula::Text("X".to_owned())), false),
        )
        .expect("within limits");
    builder
        .set(CellPos { row: 2, col: 0 }, &text("old", None, false))
        .expect("within limits");
    builder
        .absorb(&other.sheet("P").expect("the sheet was added"))
        .expect("within limits");
    builder.finish();

    let sheet = main.sheet("S1").expect("the sheet was added");
    let at = |row, col| {
        sheet
            .get(CellPos { row, col })
            .map(|view| shown(&view.to_cell()))
    };
    let expected = |sheet: &str, row, col| {
        other
            .sheet(sheet)
            .and_then(|view| view.get(CellPos { row, col }))
            .map(|view| shown(&view.to_cell()))
    };
    assert_eq!(at(0, 0), expected("P", 0, 0));
    assert_eq!(at(1, 0), expected("P", 1, 0));
    assert_eq!(at(2, 0), expected("P", 2, 0));
    assert_eq!(
        at(0, 1),
        Some(shown(&text(
            "a",
            Some(Formula::Text("X".to_owned())),
            false
        )))
    );
    assert_eq!(sheet.shared_anchor(7), Some(CellPos { row: 1, col: 0 }));
    assert_eq!(sheet.len(), 4);
}

#[test]
fn rows_written_in_descending_order_keep_the_cost_of_a_sorted_insert() {
    let rows = 300_000_u32;
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    for row in (0..rows).rev() {
        builder
            .set(CellPos { row, col: 0 }, &number(f64::from(row)))
            .expect("within limits");
    }
    builder
        .set(CellPos { row: 7, col: 0 }, &number(-1.0))
        .expect("within limits");
    builder.finish();
    let sheet = workbook.sheet("S1").expect("the sheet was added");
    assert_eq!(sheet.len(), usize::try_from(rows).expect("small"));
    let rows_read: Vec<u32> = sheet.cells().map(|(pos, _)| pos.row).collect();
    assert!(rows_read.iter().copied().eq(0..rows));
    assert_eq!(
        sheet
            .get(CellPos { row: 7, col: 0 })
            .map(|view| view.to_cell().value),
        Some(Value::Number(-1.0))
    );
}

#[test]
fn a_formula_in_a_cell_written_out_of_order_is_seen_by_the_builder() {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    let formula_cell = Cell {
        value: Value::Empty,
        formula: Some(Formula::Text("1+1".to_owned())),
        stale: true,
    };
    builder
        .set(CellPos { row: 5, col: 0 }, &number(1.0))
        .expect("within limits");
    builder
        .set(CellPos { row: 2, col: 0 }, &formula_cell)
        .expect("within limits");
    assert!(builder.has_formula_at(CellPos { row: 2, col: 0 }));
    assert!(!builder.has_formula_at(CellPos { row: 5, col: 0 }));
    builder
        .set(CellPos { row: 2, col: 0 }, &number(2.0))
        .expect("within limits");
    assert!(!builder.has_formula_at(CellPos { row: 2, col: 0 }));
}

proptest! {
    #[test]
    fn absorbing_equals_storing_cell_by_cell(
        base in prop::collection::vec((position(), cell()), 0..60),
        other in prop::collection::vec((position(), cell()), 0..60),
    ) {
        let mut source = Workbook::new();
        let mut builder = source.sheet_builder("P");
        for (pos, cell) in &other {
            builder.set(*pos, cell).expect("within limits");
        }
        builder.finish();
        let source_sheet = source.sheet("P").expect("the sheet was added");

        let mut absorbed = Workbook::new();
        let mut builder = absorbed.sheet_builder("S1");
        for (pos, cell) in &base {
            builder.set(*pos, cell).expect("within limits");
        }
        builder.absorb(&source_sheet).expect("within limits");
        builder.finish();

        let mut direct = Workbook::new();
        let mut builder = direct.sheet_builder("S1");
        for (pos, cell) in &base {
            builder.set(*pos, cell).expect("within limits");
        }
        for (pos, view) in source_sheet.cells() {
            builder.set(pos, &view.to_cell()).expect("within limits");
        }
        builder.finish();

        let left = absorbed.sheet("S1").expect("the sheet was added");
        let right = direct.sheet("S1").expect("the sheet was added");
        prop_assert_eq!(left.len(), right.len());
        let shown_cells = |sheet: &sheetrs_core::SheetView<'_>| -> Vec<(CellPos, String)> {
            sheet.cells().map(|(pos, view)| (pos, shown(&view.to_cell()))).collect()
        };
        prop_assert_eq!(shown_cells(&left), shown_cells(&right));
    }
}
