//! A calculation (`sheetc calculate`, `Book.calculate`) is built and run in one go with the inputs it is given, so what it
//! calculates is the cone of those inputs; a build keeps the cone of every declared input.

use sheetrs_compiler::model::ModelCell;
use sheetrs_compiler::{Model, Spec, compile};
use sheetrs_engine::addr::CellRef;
use sheetrs_engine::ir::Mode;
use sheetrs_engine::value::Value;

const CALC: u16 = 1;

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

/// A formula in `Calc` that caches `value`.
fn put(model: &mut Model, a1: &str, text: &str, value: Value) {
    let area = sheetrs_engine::addr::parse_area(a1).unwrap();
    model.cells.insert(
        CellRef::new(CALC, area.r0, area.c0),
        ModelCell {
            value,
            formula: Some(text.to_string()),
            array: false,
            recalculated_on_load: false,
        },
    );
}

/// Forty formulas that nothing reaches.
fn fillers(model: &mut Model) {
    for k in 1..=40 {
        put(model, &format!("F{k}"), "Data!A3*2", cached());
    }
}

/// A calculation is built and run in one go with the inputs it is given, so the only cells that can change are those: the formulas
/// that read another declared input keep the result the workbook caches. A build is run later with any declared input, and keeps
/// the cone of all of them.
#[test]
fn a_calculation_reaches_its_inputs_and_a_build_every_declared_one() {
    let mut model = model();
    put(&mut model, "B1", "Data!A1*2", cached());
    put(&mut model, "B2", "Data!A2*2", cached());
    put(&mut model, "B3", "B1+1", cached());
    fillers(&mut model);
    let kept = |given: Option<Vec<String>>| {
        let spec = Spec {
            mode: Mode::Embedded,
            inputs: vec![
                sheetrs_compiler::BindingSpec::parse("Data!A1"),
                sheetrs_compiler::BindingSpec::parse("Data!A2"),
            ],
            inputs_declared: true,
            given_inputs: given,
            ..Spec::default()
        };
        let (_, report) = compile(&model, &spec).unwrap();
        (report.formulas_kept, report.formulas_cached)
    };
    // A build: the three formulas that read an input are in the program, the forty fillers keep their result.
    assert_eq!(kept(None), (3, 40));
    // A calculation given `Data!A1`: B1 and what reads it. Given nothing: no formula.
    assert_eq!(kept(Some(vec!["Data!A1".into()])), (2, 41));
    assert_eq!(kept(Some(vec!["Data!A2".into()])), (1, 42));
    assert_eq!(kept(Some(Vec::new())), (0, 43));
    // A key that is no cell of the model: any declared input may be what it means.
    assert_eq!(kept(Some(vec!["nowhere".into()])), (3, 40));
}
