//! Reading the cells of a workbook.

use crate::find::Find;
use crate::reference::Reference;
use serde::{Deserialize, Deserializer, Serialize};
use serde_json::Value;

/// What to say about each cell. Only what is asked is read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Property {
    /// What the cell holds: a number, text, a boolean, an error (`"#DIV/0!"`).
    Value,
    /// The formula, without the `=`, in Excel's syntax (also for an ODS).
    Formula,
    /// The format code of the cell (`0.00%`, `yyyy-mm-dd`).
    NumberFormat,
    /// What a person sees in the cell: the value written with its format.
    Display,
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
    pub scope: Vec<Scope>,
    /// The value when omitted.
    #[serde(default = "the_value")]
    pub properties: Vec<Property>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub find: Option<Find>,
    /// The locale that writes `display` and the dates that a search by value sees (`en-US` when
    /// omitted).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub locale: Option<String>,
}

impl Default for Query {
    fn default() -> Self {
        Self {
            scope: the_workbook(),
            properties: the_value(),
            find: None,
            locale: None,
        }
    }
}

/// A cell of the answer, with the properties that were asked.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct CellRecord {
    pub sheet: String,
    /// `B2`.
    pub cell: String,
    /// One-based, as the cell is written.
    pub row: u32,
    pub col: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display: Option<String>,
}

/// The answer to a [`Query`]: the cells, sheet by sheet in the order of the workbook, each sheet by
/// rows.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct QueryResponse {
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
