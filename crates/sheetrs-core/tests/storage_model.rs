use std::collections::BTreeMap;

use proptest::prelude::*;
use sheetrs_core::{
    AbsorbError, Cell, CellPos, ErrorKind, Formula, FormulaRef, SharedMember, SheetBuilder, Value,
    ValueRef, Visibility, Workbook,
};

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
        .set(CellPos { row: 3, col: 0 }, &text("old", None, false))
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
    assert_eq!(sheet.len(), 5);
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

proptest! {
    #[test]
    fn absorbing_equals_storing_cell_by_cell(
        base in prop::collection::vec((position(), cell()), 0..60),
        other in prop::collection::vec((position(), cell()), 0..60),
    ) {
        let taken: std::collections::BTreeSet<CellPos> = base.iter().map(|(pos, _)| *pos).collect();
        let other: Vec<(CellPos, Cell)> = other
            .into_iter()
            .filter(|(pos, _)| !taken.contains(pos))
            .collect();
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

fn shared(index: u32, text: Option<&str>) -> Cell {
    Cell {
        value: Value::Number(1.0),
        formula: Some(match text {
            Some(text) => Formula::Text(text.to_owned()),
            None => Formula::Shared { index },
        }),
        stale: true,
    }
}

fn member(group: u32, anchor_covers_itself: Option<bool>) -> SharedMember {
    SharedMember {
        group,
        anchor_covers_itself,
    }
}

/// Loads `cell` the way the reader does, with the shared group it names, if any.
fn load(
    builder: &mut SheetBuilder<'_>,
    pos: CellPos,
    cell: &Cell,
    shared: impl Into<Option<SharedMember>>,
) {
    let value = match &cell.value {
        Value::Empty => ValueRef::Empty,
        Value::Number(number) => ValueRef::Number(*number),
        Value::Text(text) => ValueRef::Text(text),
        other => panic!("not used here: {other:?}"),
    };
    let formula = cell.formula.as_ref().map(|formula| match formula {
        Formula::Text(text) => FormulaRef::Text(text),
        Formula::Shared { index } => FormulaRef::Shared { index: *index },
    });
    builder
        .load_view(pos, &value, formula, cell.stale, shared.into())
        .expect("within limits");
}

fn at(row: u32, col: u32) -> CellPos {
    CellPos { row, col }
}

/// A sheet with group 0 anchored validly, group 1 anchored without a valid `ref` and group 2
/// that has only dependents.
fn grouped_sheet() -> Workbook {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    load(
        &mut builder,
        at(0, 0),
        &shared(0, Some("1+1")),
        member(0, Some(true)),
    );
    load(
        &mut builder,
        at(1, 0),
        &shared(0, None),
        Some(member(0, None)),
    );
    load(
        &mut builder,
        at(0, 1),
        &shared(1, Some("2+2")),
        member(1, Some(false)),
    );
    load(
        &mut builder,
        at(1, 1),
        &shared(1, None),
        Some(member(1, None)),
    );
    load(
        &mut builder,
        at(0, 2),
        &shared(2, None),
        Some(member(2, None)),
    );
    builder.set_shared_anchor(0, at(0, 0));
    builder.set_shared_anchor(1, at(0, 1));
    builder.finish_loaded();
    workbook
}

#[test]
fn a_shared_group_without_a_valid_anchor_loses_all_its_cells() {
    let workbook = grouped_sheet();
    let sheet = workbook.sheet("S1").expect("the sheet was added");
    let kept: Vec<CellPos> = sheet.cells().map(|(pos, _)| pos).collect();
    assert_eq!(kept, vec![at(0, 0), at(1, 0)]);
    assert_eq!(sheet.len(), 2);
    assert_eq!(sheet.shared_anchor(0), Some(at(0, 0)));
    assert_eq!(sheet.shared_anchor(1), None);
}

#[test]
fn plain_finish_keeps_the_cells_of_every_shared_group() {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    builder
        .set(at(0, 0), &shared(7, None))
        .expect("within limits");
    builder.finish();
    let sheet = workbook.sheet("S1").expect("the sheet was added");
    assert_eq!(sheet.len(), 1);
}

#[test]
fn group_validity_is_decided_over_the_blocks_that_were_absorbed() {
    let mut first = Workbook::new();
    let mut builder = first.sheet_builder("P");
    load(
        &mut builder,
        at(0, 0),
        &shared(0, None),
        Some(member(0, None)),
    );
    builder.finish();
    let mut second = Workbook::new();
    let mut builder = second.sheet_builder("P");
    load(
        &mut builder,
        at(5, 0),
        &shared(0, Some("1+1")),
        member(0, Some(true)),
    );
    builder.finish();

    let mut whole = Workbook::new();
    let mut builder = whole.sheet_builder("S1");
    for block in [&first, &second] {
        builder
            .absorb(&block.sheet("P").expect("the sheet was added"))
            .expect("within limits");
    }
    builder.finish_loaded();
    let sheet = whole.sheet("S1").expect("the sheet was added");
    assert_eq!(sheet.len(), 2);
}

fn constant(value: f64) -> Cell {
    number(value)
}

fn plain_formula(text: &str) -> Cell {
    Cell {
        value: Value::Number(1.0),
        formula: Some(Formula::Text(text.to_owned())),
        stale: true,
    }
}

/// The cell at row 0 of a one-column sheet of a dropped group after `steps` were loaded there.
fn dropped_group_leaves(steps: &[(Cell, Option<SharedMember>)]) -> Option<Cell> {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    for (cell, shared) in steps {
        load(&mut builder, at(0, 0), cell, *shared);
    }
    builder.finish_loaded();
    workbook
        .sheet("S1")?
        .get(at(0, 0))
        .map(|view| view.to_cell())
}

#[test]
fn a_dropped_member_leaves_the_cell_it_shared_an_address_with() {
    let dropped_anchor = (shared(0, Some("1+1")), Some(member(0, Some(false))));
    let dropped_dependent = (shared(0, None), Some(member(0, None)));
    let constant_cell = (constant(4.0), None);
    let formula_cell = (plain_formula("4"), None);
    for (steps, expected) in [
        (
            vec![constant_cell.clone(), dropped_dependent.clone()],
            Some(constant(4.0)),
        ),
        (
            vec![dropped_dependent.clone(), constant_cell.clone()],
            Some(constant(4.0)),
        ),
        (
            vec![formula_cell.clone(), dropped_dependent.clone()],
            Some(plain_formula("4")),
        ),
        (
            vec![dropped_anchor.clone(), formula_cell.clone()],
            Some(plain_formula("4")),
        ),
        (
            vec![dropped_anchor.clone(), constant_cell.clone()],
            Some(constant(4.0)),
        ),
        (vec![dropped_anchor.clone()], None),
        (vec![dropped_dependent.clone()], None),
    ] {
        let left = dropped_group_leaves(&steps).map(|cell| shown(&cell));
        assert_eq!(left, expected.map(|cell| shown(&cell)), "{steps:?}");
    }
}

#[test]
fn a_cell_that_changes_group_at_the_same_address_follows_the_last_group() {
    let invalid = (shared(0, Some("1+1")), Some(member(0, Some(false))));
    let valid = (shared(1, Some("5+5")), Some(member(1, Some(true))));
    let left = dropped_group_leaves(&[invalid, valid.clone()]);
    assert_eq!(left.map(|cell| shown(&cell)), Some(shown(&valid.0)));
}

#[test]
fn absorbing_a_cell_at_an_address_that_is_taken_is_reported() {
    let mut other = Workbook::new();
    let mut builder = other.sheet_builder("P");
    builder.set(at(2, 0), &number(1.0)).expect("within limits");
    builder.finish();
    let mut main = Workbook::new();
    let mut builder = main.sheet_builder("S1");
    builder.set(at(2, 0), &number(2.0)).expect("within limits");
    builder.set(at(9, 0), &number(3.0)).expect("within limits");
    let outcome = builder.absorb(&other.sheet("P").expect("the sheet was added"));
    assert_eq!(outcome, Err(AbsorbError::Overlap));
}

#[test]
fn a_formula_is_not_replaced_by_a_cell_without_one_even_out_of_order() {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    load(&mut builder, at(5, 0), &constant(1.0), None);
    load(&mut builder, at(2, 0), &plain_formula("1+1"), None);
    load(&mut builder, at(2, 0), &constant(2.0), None);
    load(&mut builder, at(5, 0), &plain_formula("2+2"), None);
    builder.finish_loaded();
    let sheet = workbook.sheet("S1").expect("the sheet was added");
    let formula_at = |row| sheet.get(at(row, 0)).map(|view| view.to_cell().formula);
    assert_eq!(formula_at(2), Some(Some(Formula::Text("1+1".to_owned()))));
    assert_eq!(formula_at(5), Some(Some(Formula::Text("2+2".to_owned()))));
}

/// A sheet where group 0 is anchored validly at column 1, and `steps` are loaded at row 0 of
/// column 0.
fn leaves_with_a_valid_group(steps: &[(Cell, Option<SharedMember>)]) -> Option<Cell> {
    let mut workbook = Workbook::new();
    let mut builder = workbook.sheet_builder("S1");
    load(
        &mut builder,
        at(0, 1),
        &shared(0, Some("0+0")),
        Some(member(0, Some(true))),
    );
    for (cell, shared) in steps {
        load(&mut builder, at(0, 0), cell, *shared);
    }
    builder.finish_loaded();
    workbook
        .sheet("S1")?
        .get(at(0, 0))
        .map(|view| view.to_cell())
}

#[test]
fn a_dropped_member_over_a_kept_one_leaves_the_kept_one() {
    let kept_anchor = (shared(0, Some("1+1")), Some(member(0, Some(true))));
    let kept_dependent = (shared(0, None), Some(member(0, None)));
    let dropped_anchor = (shared(1, Some("5+5")), Some(member(1, Some(false))));
    let orphan_dependent = (shared(2, None), Some(member(2, None)));
    let constant_cell = (constant(3.0), None);
    for (steps, expected) in [
        (
            vec![kept_anchor.clone(), dropped_anchor.clone()],
            Some(kept_anchor.0.clone()),
        ),
        (
            vec![kept_dependent.clone(), orphan_dependent.clone()],
            Some(kept_dependent.0.clone()),
        ),
        (
            vec![orphan_dependent.clone(), kept_dependent.clone()],
            Some(kept_dependent.0.clone()),
        ),
        (
            vec![
                kept_anchor.clone(),
                dropped_anchor.clone(),
                constant_cell.clone(),
            ],
            Some(kept_anchor.0.clone()),
        ),
        (
            vec![dropped_anchor.clone(), kept_anchor.clone()],
            Some(kept_anchor.0.clone()),
        ),
    ] {
        let left = leaves_with_a_valid_group(&steps).map(|cell| shown(&cell));
        assert_eq!(left, expected.map(|cell| shown(&cell)), "{steps:?}");
    }
}

fn text_cell(text: &str) -> Cell {
    Cell {
        value: Value::Text(text.into()),
        formula: None,
        stale: false,
    }
}

/// The writes of a sheet where group 0 is valid, group 1 is dropped and an address is
/// written several times with texts and formula texts.
fn contested_writes() -> Vec<(CellPos, Cell, Option<SharedMember>)> {
    vec![
        (
            at(0, 0),
            shared(0, Some("0+0")),
            Some(member(0, Some(true))),
        ),
        (at(5, 0), text_cell("first"), None),
        (
            at(5, 0),
            shared(1, Some("5+5")),
            Some(member(1, Some(false))),
        ),
        (at(5, 0), text_cell("second"), None),
        (
            at(6, 0),
            shared(1, Some("6+6")),
            Some(member(1, Some(false))),
        ),
        (at(6, 0), shared(0, None), Some(member(0, None))),
        (at(6, 0), text_cell("third"), None),
    ]
}

fn finish_loaded_sheet(blocks: &[&[(CellPos, Cell, Option<SharedMember>)]]) -> Vec<String> {
    let mut parts = Vec::new();
    for block in blocks {
        let mut workbook = Workbook::new();
        let mut builder = workbook.sheet_builder("P");
        for (pos, cell, shared) in *block {
            load(&mut builder, *pos, cell, *shared);
        }
        builder.finish();
        parts.push(workbook);
    }
    let mut whole = Workbook::new();
    let mut builder = whole.sheet_builder("S1");
    for part in &parts {
        builder
            .absorb(&part.sheet("P").expect("the sheet was added"))
            .expect("the blocks do not overlap");
    }
    builder.finish_loaded();
    let sheet = whole.sheet("S1").expect("the sheet was added");
    sheet
        .cells()
        .map(|(pos, view)| format!("{pos:?} {:?}", view.to_cell()))
        .collect()
}

#[test]
fn absorbed_blocks_decide_contested_addresses_as_one_sheet_does() {
    let writes = contested_writes();
    let whole = finish_loaded_sheet(&[&writes]);
    let (head, tail) = writes.split_at(1);
    assert_eq!(finish_loaded_sheet(&[head, tail]), whole);
    assert!(whole.iter().any(|cell| cell.contains("second")));
    // The constant after a kept dependent is blocked by its formula.
    assert!(
        whole
            .iter()
            .any(|cell| cell.contains("row: 6") && cell.contains("Shared { index: 0 }"))
    );
}
