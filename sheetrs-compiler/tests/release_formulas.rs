//! A model passed to `compile` by value loses the text of its formulas once they are lowered; one
//! passed by reference is the caller's and is not touched. The program is the same either way.

use sheetrs_compiler::model::Model;
use sheetrs_compiler::{BindingSpec, Spec, compile};
use sheetrs_engine::value::Value;

fn model() -> Model {
    let mut m = Model::new();
    let s = m.add_sheet("Sheet1");
    m.set(s, "A1", Value::Number(2.0));
    m.set_formula(s, "B1", "A1*3");
    m.set_formula(s, "C1", "ISFORMULA(B1)");
    m.set_formula(s, "D1", "FORMULATEXT(B1)");
    m
}

fn spec() -> Spec {
    Spec {
        inputs: vec![BindingSpec::parse("Sheet1!A1")],
        outputs: ["B1", "C1", "D1"]
            .map(|c| BindingSpec::parse(&format!("Sheet1!{c}")))
            .into(),
        ..Spec::default()
    }
}

#[test]
fn an_owned_model_gives_the_program_of_a_borrowed_one() {
    let kept = model();
    let (borrowed, _) = compile(&kept, &spec()).unwrap();
    let (owned, _) = compile(model(), &spec()).unwrap();
    assert_eq!(
        serde_json::to_string(&borrowed).unwrap(),
        serde_json::to_string(&owned).unwrap()
    );
    let answer = owned
        .calculate(
            &serde_json::Value::Null,
            &sheetrs_engine::Overrides::default(),
        )
        .unwrap();
    assert_eq!(answer["Sheet1!B1"], 6.0);
    assert_eq!(answer["Sheet1!C1"], true);
    assert_eq!(answer["Sheet1!D1"], "=A1*3");
}

#[test]
fn a_borrowed_model_keeps_its_formulas() {
    let kept = model();
    compile(&kept, &spec()).unwrap();
    let texts: Vec<_> = kept
        .cells
        .iter()
        .filter_map(|(_, c)| c.formula.clone())
        .collect();
    assert_eq!(texts, ["A1*3", "ISFORMULA(B1)", "FORMULATEXT(B1)"]);
}

#[test]
fn releasing_keeps_which_cells_are_formulas() {
    let mut kept = model();
    kept.cells.release_formulas();
    let formulas = kept
        .cells
        .iter()
        .filter(|(_, c)| c.formula.is_some())
        .count();
    assert_eq!(formulas, 3);
    assert!(
        kept.cells
            .iter()
            .all(|(_, c)| c.formula.as_deref().is_none_or(str::is_empty))
    );
}
