//! Comparison of what the oracle observed in a cell with what a reader put in the model.

use sheetrs_core::{Cell, Value};

use crate::case::CellValue;

/// What a reader may do differently from the oracle in one cell, as the format decides.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Leniency {
    /// The cell has to match, formula presence included.
    Strict,
    /// The reader may omit the formula the oracle shows.
    FormulaOptional,
    /// The cell has to match, and its value is compared even when the reader flags it stale.
    Values,
    /// The cell does not match, whatever else is compared.
    Mismatch,
}

/// The formula the oracle shows for a cell, if any.
#[must_use]
pub fn formula_of(value: &CellValue) -> Option<&String> {
    match value {
        CellValue::Number { formula, .. }
        | CellValue::String { formula, .. }
        | CellValue::Error { formula, .. }
        | CellValue::NonFinite { formula, .. } => formula.as_ref(),
        CellValue::Empty => None,
    }
}

/// Whether the cell read matches the oracle's, formula presence included unless
/// `formula_optional`. With `pending_stale`, the stale value of a formula cell is not compared:
/// the oracle recalculates it, which needs the formula engine.
#[must_use]
pub fn matches(
    expected: &CellValue,
    actual: Option<&Cell>,
    pending_stale: bool,
    formula_optional: bool,
) -> bool {
    let actual_has_formula = actual.is_some_and(|cell| cell.formula.is_some());
    if (formula_of(expected).is_some() != actual_has_formula) && !formula_optional {
        return false;
    }
    if pending_stale && actual.is_some_and(|cell| cell.stale) {
        return true;
    }
    match (expected, actual.map(|cell| &cell.value)) {
        (CellValue::Empty, None | Some(Value::Empty)) => true,
        (CellValue::Number { v, .. }, Some(Value::Bool(flag))) => {
            v.to_bits() == f64::from(u8::from(*flag)).to_bits()
        }
        (CellValue::NonFinite { v, .. }, Some(Value::Number(number))) => match v.as_str() {
            "nan" => number.is_nan(),
            "inf" => number.is_infinite() && number.is_sign_positive(),
            "-inf" => number.is_infinite() && number.is_sign_negative(),
            _ => false,
        },
        (CellValue::Number { v, .. }, Some(Value::Number(number))) => {
            v.to_bits() == number.to_bits()
        }
        (CellValue::String { v, .. }, Some(Value::Text(text))) => v.as_str() == &**text,
        (CellValue::Error { text, .. }, Some(Value::Error(kind))) => kind.as_text() == text,
        _ => false,
    }
}
