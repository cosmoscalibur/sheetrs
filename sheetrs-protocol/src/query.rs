//! Reading the cells of a workbook.

use crate::find::Find;
use crate::reference::Reference;
use serde::{Deserialize, Deserializer, Serialize};
use serde_json::Value;
use std::sync::Arc;

/// What to say about each cell. Only what is asked is read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Property {
    /// What the cell holds: a number, text, a boolean, or an error, which is the object
    /// `{"error": "#DIV/0!"}` (as the answer of `sheetc calculate` gives it), so that it is not taken for a
    /// text that starts with `#`.
    Value,
    /// The formula, without the `=`, as the file writes it: Excel's syntax in an XLSX,
    /// `OpenFormula` (`[.A1]+[$T.B2]`, `;` between arguments) in an ODS.
    Formula,
    /// The format code of the cell (`0.00%`, `yyyy-mm-dd`).
    NumberFormat,
    /// What a person sees in the cell: the value written with its format.
    Display,
    /// The text of the note of the cell (a cell with a note is returned even if it has no value).
    Comment,
}

/// A date, a time of day or both, as ISO 8601 text with its type: what the `value` of a cell is
/// with `dates`. In JSON it is `{"type": "datetime", "value": "2024-03-15T10:30:00"}`, so
/// that a consumer that reads JSON tells it from a text cell.
///
/// The text is `YYYY-MM-DD` for a date, `hh:mm:ss` for a time (`hh:mm:ss.fff` when the format of the
/// cell writes fractions of a second) and `YYYY-MM-DDThh:mm:ss` for both. The year has four digits.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "lowercase")]
#[non_exhaustive]
pub enum DateValue {
    /// `2024-03-15`.
    Date(String),
    /// `10:30:00`: the time of day, without the day.
    Time(String),
    /// `2024-03-15T10:30:00`.
    Datetime(String),
}

impl DateValue {
    /// The ISO 8601 text, without the type.
    #[must_use]
    pub fn text(&self) -> &str {
        match self {
            Self::Date(t) | Self::Time(t) | Self::Datetime(t) => t,
        }
    }

    /// The JSON the `value` of a cell holds.
    #[must_use]
    pub fn to_json(&self) -> Value {
        serde_json::to_value(self).unwrap_or(Value::Null)
    }
}

/// Where to read.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Scope {
    /// Every sheet.
    Workbook,
    /// A sheet by name.
    Sheet(String),
    /// A range, `Sheet1!A1:C9` (without a sheet, the first one).
    Range(Reference),
}

impl From<&str> for Scope {
    /// A sheet (`Data`, `'My Sheet'`) or a range (`Data!A1:C9`, `A1:C9` on the first sheet), as the
    /// command line and the bindings write a scope.
    fn from(text: &str) -> Self {
        match text.parse::<Reference>() {
            Ok(reference) if text.contains('!') || reference.sheet.is_none() => {
                Self::Range(reference)
            }
            _ => Self::Sheet(text.trim_matches('\'').to_string()),
        }
    }
}

fn one_or_many<'de, D: Deserializer<'de>>(d: D) -> Result<Vec<Scope>, D::Error> {
    #[derive(Deserialize)]
    #[serde(untagged)]
    enum Shape {
        One(Scope),
        Many(Vec<Scope>),
    }
    Ok(match Shape::deserialize(d)? {
        Shape::One(s) => vec![s],
        Shape::Many(m) => m,
    })
}

fn the_workbook() -> Vec<Scope> {
    vec![Scope::Workbook]
}

fn the_value() -> Vec<Property> {
    vec![Property::Value]
}

/// A read of a workbook: where, which properties, and optionally which cells (the ones a [`Find`]
/// matches). The cells that are empty are not returned.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Query {
    /// One scope or a list of them; the whole workbook when omitted.
    #[serde(default = "the_workbook", deserialize_with = "one_or_many")]
    /// Where to read; the whole workbook when omitted.
    pub scope: Vec<Scope>,
    /// The value when omitted.
    #[serde(default = "the_value")]
    pub properties: Vec<Property>,
    /// Keep only the cells where this is found.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub find: Option<Find>,
    /// The locale that writes `display` and the dates that a search by value sees (`en-US` when
    /// omitted).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub locale: Option<String>,
    /// Give the numbers whose format is a date, a time or both as a [`DateValue`] (ISO 8601 text with
    /// a type tag, worked out from the serial with the date system of the file: the 1900, the 1904 and
    /// any `table:null-date` of an ODS) instead of the serial number the file stores (`45366.4375`),
    /// which is the default. Every other cell keeps its value.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub dates: bool,
}

impl Default for Query {
    fn default() -> Self {
        Self {
            scope: the_workbook(),
            properties: the_value(),
            find: None,
            locale: None,
            dates: false,
        }
    }
}

/// A cell of the answer, with the properties that were asked.
///
/// The sheet is shared by every cell of it (an `Arc<str>`, one copy per sheet and not one per cell) and
/// the cell in A1 notation is worked out from `row` and `col` when it is asked for
/// ([`cell`](Self::cell)). Written as JSON it has `sheet`, `cell`, `row`, `col` and the properties
/// that were asked.
#[derive(Debug, Clone, PartialEq)]
pub struct CellRecord {
    /// The sheet of the cell.
    pub sheet: Arc<str>,
    /// One-based, as the cell is written.
    pub row: u32,
    /// One-based, as the column is written (`B` is 2).
    pub col: u32,
    /// The value: a number (or a typed date with `dates`), text, boolean or `{"error": …}`.
    pub value: Option<Value>,
    /// The formula, without the `=`, in the syntax of the file.
    pub formula: Option<String>,
    /// The format code of the cell.
    pub number_format: Option<String>,
    /// The value written with its number format, in the locale of the query.
    pub display: Option<String>,
    /// The text of the note of the cell.
    pub comment: Option<String>,
}

/// The JSON of a [`CellRecord`].
#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct CellRecordJson {
    sheet: String,
    cell: String,
    row: u32,
    col: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    value: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    number_format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    display: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    comment: Option<String>,
}

impl Serialize for CellRecord {
    fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        CellRecordJson {
            sheet: self.sheet.to_string(),
            cell: self.cell(),
            row: self.row,
            col: self.col,
            value: self.value.clone(),
            formula: self.formula.clone(),
            number_format: self.number_format.clone(),
            display: self.display.clone(),
            comment: self.comment.clone(),
        }
        .serialize(serializer)
    }
}

impl<'de> Deserialize<'de> for CellRecord {
    fn deserialize<D: serde::Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        let json = CellRecordJson::deserialize(deserializer)?;
        Ok(Self {
            sheet: json.sheet.into(),
            row: json.row,
            col: json.col,
            value: json.value,
            formula: json.formula,
            number_format: json.number_format,
            display: json.display,
            comment: json.comment,
        })
    }
}

impl CellRecord {
    /// The cell in A1 notation (`B2`), worked out from `row` and `col`.
    #[must_use]
    pub fn cell(&self) -> String {
        sheetrs_formula::addr::a1(self.row.saturating_sub(1), self.col.saturating_sub(1))
    }

    /// The `value` of the cell as a [`DateValue`], when the query asked for `dates` and the
    /// cell is a date or a time.
    #[must_use]
    pub fn date_value(&self) -> Option<DateValue> {
        serde_json::from_value(self.value.clone()?).ok()
    }
}

/// The answer to a [`Query`]: the cells, sheet by sheet in the order of the workbook, each sheet by
/// rows.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct QueryResponse {
    /// The cells that were asked for and are not empty, in the order of the sheet.
    pub cells: Vec<CellRecord>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_query_without_fields_reads_the_values_of_the_workbook() {
        let q: Query = serde_json::from_str("{}").unwrap();
        assert_eq!(q, Query::default());
    }

    #[test]
    fn dates_are_serial_unless_asked_and_a_date_value_says_its_type() {
        let q: Query = serde_json::from_str("{}").unwrap();
        assert!(!q.dates);
        assert!(!serde_json::to_string(&q).unwrap().contains("dates"));
        let q: Query = serde_json::from_str(r#"{"dates": true}"#).unwrap();
        assert!(q.dates);
        assert!(serde_json::from_str::<Query>(r#"{"dates": "datetime"}"#).is_err());
        let v = DateValue::Datetime("2024-03-15T10:30:00".into());
        assert_eq!(
            v.to_json(),
            serde_json::json!({"type": "datetime", "value": "2024-03-15T10:30:00"})
        );
        assert_eq!(v.text(), "2024-03-15T10:30:00");
        let record = CellRecord {
            sheet: "T".into(),
            row: 1,
            col: 1,
            value: Some(v.to_json()),
            formula: None,
            number_format: None,
            display: None,
            comment: None,
        };
        assert_eq!(record.date_value(), Some(v));
        let number = CellRecord {
            value: Some(serde_json::json!(45366.4375)),
            ..record
        };
        assert_eq!(number.date_value(), None);
    }

    #[test]
    fn a_scope_is_one_or_several_and_is_written_as_what_it_is() {
        let q: Query = serde_json::from_str(r#"{"scope": {"sheet": "Data"}}"#).unwrap();
        assert_eq!(q.scope, vec![Scope::Sheet("Data".into())]);
        let q: Query = serde_json::from_str(
            r#"{"scope": ["workbook", {"range": "Data!A1:B2"}], "properties": ["formula", "number_format"]}"#,
        )
        .unwrap();
        assert_eq!(q.scope.len(), 2);
        assert_eq!(q.properties, [Property::Formula, Property::NumberFormat]);
        assert!(serde_json::from_str::<Query>(r#"{"scope": {"range": "nope"}}"#).is_err());
        assert!(serde_json::from_str::<Query>(r#"{"properties": ["array"]}"#).is_err());
    }
}
