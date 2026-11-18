//! Test cases and observed values, in the JSON shape exchanged with the oracle script.

use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

/// A synthetic workbook, the operations applied to it and the ranges observed afterwards.
#[derive(Debug, Serialize)]
pub struct Case {
    pub name: String,
    pub sheets: Vec<SheetSpec>,
    pub ops: Vec<Op>,
    pub reads: Vec<Read>,
}

/// A sheet of a case: its name and its cells by A1 address.
#[derive(Debug, Serialize)]
pub struct SheetSpec {
    pub name: String,
    pub cells: BTreeMap<String, CellInput>,
}

/// The content of an input cell.
#[derive(Debug, Clone, Serialize)]
#[serde(untagged)]
pub enum CellInput {
    Number(f64),
    /// Text. A string starting with `=` is a formula in the UNO API grammar
    /// (English function names, `;` as argument separator).
    Text(String),
}

/// An operation applied to the workbook, in the order listed by the case.
///
/// Rows (`"2"`, `"2:3"`) are one-based and columns (`"B"`, `"B:C"`) are letters.
#[derive(Debug, Serialize)]
#[serde(tag = "op", rename_all = "snake_case")]
pub enum Op {
    DeleteRows {
        sheet: String,
        rows: String,
    },
    DeleteColumns {
        sheet: String,
        columns: String,
    },
    DeleteSheet {
        sheet: String,
    },
    ToValues {
        sheet: String,
        range: String,
    },
    Recalculate,
    RecalculateHard,
    /// Replaces the workbook with the one loaded from `path`, like File > Open. A text file
    /// is loaded with the import `filter` the case names, never with the one `LibreOffice`
    /// would guess.
    Open {
        path: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        filter: Option<TextFilter>,
    },
    /// Saves the workbook as it is at this point, like File > Save As.
    SaveAs {
        format: Format,
        path: String,
    },
}

/// What the text import filter makes of a field that is quoted.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum QuotedFields {
    /// A quoted field that looks like a number is a number.
    Detected,
    /// A quoted field is text.
    Text,
}

/// What the text import filter makes of dates, times, percentages, currencies and the other
/// numbers that are not plain decimals.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum SpecialNumbers {
    /// They are text.
    Text,
    /// They are numbers.
    Detected,
}

/// What the text import filter makes of a field that starts with `=`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum FormulaFields {
    /// It is text.
    Text,
    /// It is a formula.
    Evaluated,
}

/// What the text import filter does with the spaces around a field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum Spaces {
    /// They stay.
    Kept,
    /// They are dropped.
    Trimmed,
}

/// The options of the text import filter, which the oracle sets explicitly.
#[derive(Debug, Clone, Serialize)]
pub struct TextFilter {
    /// The field separator, as a character code.
    pub separator: u32,
    /// The text delimiter, as a character code; 0 for none.
    pub delimiter: u32,
    /// The character set code: 76 is UTF-8.
    pub charset: u32,
    /// The language of numbers and dates: 1033 is English (United States).
    pub language: u32,
    pub quoted_fields: QuotedFields,
    pub special_numbers: SpecialNumbers,
    pub formula_fields: FormulaFields,
    pub spaces: Spaces,
}

impl TextFilter {
    /// The filter that reads a comma-separated file of UTF-8 text with `"` as delimiter, for the
    /// English (United States) language, with every option off.
    #[must_use]
    pub const fn comma_separated() -> Self {
        Self {
            separator: 44,
            delimiter: 34,
            charset: 76,
            language: 1033,
            quoted_fields: QuotedFields::Detected,
            special_numbers: SpecialNumbers::Text,
            formula_fields: FormulaFields::Text,
            spaces: Spaces::Kept,
        }
    }
}

/// A file format the oracle can save.
#[derive(Debug, Clone, Copy, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum Format {
    Xlsx,
}

/// A range whose cells are observed after the operations.
#[derive(Debug, Serialize)]
pub struct Read {
    pub sheet: String,
    pub range: String,
}

/// The values the oracle observed for one case.
#[derive(Debug, Deserialize)]
pub struct Observed {
    pub name: String,
    pub reads: Vec<ReadResult>,
    /// The sheets of the workbook after the operations, in order.
    #[serde(default)]
    pub sheets: Vec<SheetState>,
}

/// A sheet of the workbook the oracle observed.
#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct SheetState {
    pub name: String,
    /// Whether the sheet is shown.
    pub visible: bool,
}

/// The cells of one [`Read`], by row and then column.
#[derive(Debug, Deserialize)]
pub struct ReadResult {
    pub sheet: String,
    pub range: String,
    pub cells: Vec<Vec<CellValue>>,
}

/// An observed cell. `formula` is present when the cell holds a formula.
#[derive(Debug, Clone, PartialEq, Deserialize)]
#[serde(tag = "t", rename_all = "snake_case")]
pub enum CellValue {
    Empty,
    Number {
        v: f64,
        formula: Option<String>,
    },
    /// An infinity or NaN, which JSON cannot hold: `v` is `inf`, `-inf` or `nan`.
    NonFinite {
        v: String,
        formula: Option<String>,
    },
    String {
        v: String,
        formula: Option<String>,
    },
    /// A formula that evaluates to an error: the oracle's error `code` and the text it displays.
    Error {
        code: u32,
        text: String,
        formula: Option<String>,
    },
}

#[cfg(test)]
mod tests {

    use serde_json::json;

    use super::*;

    #[test]
    fn case_serializes_to_the_shape_the_oracle_script_reads() -> Result<(), serde_json::Error> {
        let case = Case {
            name: "delete_row".to_owned(),
            sheets: vec![SheetSpec {
                name: "S1".to_owned(),
                cells: BTreeMap::from([
                    ("A1".to_owned(), CellInput::Number(1.0)),
                    ("B1".to_owned(), CellInput::Text("=A1".to_owned())),
                ]),
            }],
            ops: vec![
                Op::DeleteRows {
                    sheet: "S1".to_owned(),
                    rows: "2:3".to_owned(),
                },
                Op::Recalculate,
            ],
            reads: vec![Read {
                sheet: "S1".to_owned(),
                range: "A1:B1".to_owned(),
            }],
        };
        assert_eq!(
            serde_json::to_value(&case)?,
            json!({
                "name": "delete_row",
                "sheets": [{"name": "S1", "cells": {"A1": 1.0, "B1": "=A1"}}],
                "ops": [
                    {"op": "delete_rows", "sheet": "S1", "rows": "2:3"},
                    {"op": "recalculate"}
                ],
                "reads": [{"sheet": "S1", "range": "A1:B1"}]
            })
        );
        Ok(())
    }

    #[test]
    fn open_serializes_with_the_path() -> Result<(), serde_json::Error> {
        let op = Op::Open {
            path: "/tmp/book.xlsx".to_owned(),
            filter: None,
        };
        assert_eq!(
            serde_json::to_value(&op)?,
            json!({"op": "open", "path": "/tmp/book.xlsx"})
        );
        Ok(())
    }

    #[test]
    fn open_with_a_text_filter_serializes_its_options() -> Result<(), serde_json::Error> {
        let op = Op::Open {
            path: "/tmp/book.csv".to_owned(),
            filter: Some(TextFilter::comma_separated()),
        };
        let value = serde_json::to_value(&op)?;
        assert_eq!(value["filter"]["separator"], 44);
        assert_eq!(value["filter"]["special_numbers"], "text");
        Ok(())
    }

    #[test]
    fn save_as_serializes_with_the_format_and_path() -> Result<(), serde_json::Error> {
        let op = Op::SaveAs {
            format: Format::Xlsx,
            path: "/tmp/book.xlsx".to_owned(),
        };
        assert_eq!(
            serde_json::to_value(&op)?,
            json!({"op": "save_as", "format": "xlsx", "path": "/tmp/book.xlsx"})
        );
        Ok(())
    }

    #[test]
    fn observed_numbers_keep_every_bit() -> Result<(), Box<dyn std::error::Error>> {
        for text in [
            "6.934698245112694e276",
            "0.30000000000000004",
            "1.0000000000000002",
            "5e-324",
            "1.7976931348623157e308",
        ] {
            let expected: f64 = text.parse()?;
            let cell: CellValue = serde_json::from_str(&format!(r#"{{"t":"number","v":{text}}}"#))?;
            let CellValue::Number { v, .. } = cell else {
                return Err(format!("not a number: {cell:?}").into());
            };
            assert_eq!(v.to_bits(), expected.to_bits(), "{text}");
        }
        Ok(())
    }

    #[test]
    fn observed_cells_deserialize_by_type() -> Result<(), serde_json::Error> {
        let cells: Vec<CellValue> = serde_json::from_value(json!([
            {"t": "empty"},
            {"t": "number", "v": 2.5},
            {"t": "number", "v": 4.0, "formula": "=SUM(A1:A2)"},
            {"t": "string", "v": "x"},
            {"t": "error", "code": 524, "text": "#REF!", "formula": "=$S2.A1"}
        ]))?;
        assert_eq!(
            cells,
            vec![
                CellValue::Empty,
                CellValue::Number {
                    v: 2.5,
                    formula: None
                },
                CellValue::Number {
                    v: 4.0,
                    formula: Some("=SUM(A1:A2)".to_owned())
                },
                CellValue::String {
                    v: "x".to_owned(),
                    formula: None
                },
                CellValue::Error {
                    code: 524,
                    text: "#REF!".to_owned(),
                    formula: Some("=$S2.A1".to_owned())
                },
            ]
        );
        Ok(())
    }
}
