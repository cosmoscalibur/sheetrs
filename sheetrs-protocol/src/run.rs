//! What a compiled program is asked and what it answers.

use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};

/// Orientation of the arrays returned for ranges (Google Sheets' `majorDimension`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "UPPERCASE")]
pub enum MajorDimension {
    /// An array of rows.
    #[default]
    Rows,
    /// An array of columns.
    Columns,
}

impl MajorDimension {
    #[must_use]
    pub fn parse(s: &str) -> Option<Self> {
        match s.to_ascii_uppercase().as_str() {
            "ROWS" => Some(Self::Rows),
            "COLUMNS" => Some(Self::Columns),
            _ => None,
        }
    }
}

/// A request to a program: every field is optional.
///
/// Without `ranges` the answer is the outputs the program declared; with them, exactly those
/// ranges (a program built with `outputs = any` has no default answer and requires `ranges`).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields, rename_all = "camelCase")]
pub struct Request {
    /// The value of each declared input, keyed by its name.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inputs: Option<Map<String, Value>>,
    /// Ranges (or names) to read.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ranges: Option<Vec<String>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_dimension: Option<MajorDimension>,
}

/// The values of one requested range: always an array of rows (or of columns), whatever its size.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct ValueRange {
    /// The range, canonical (`Sheet1!B2:C9`).
    pub range: String,
    /// The defined name or output name it was asked by.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub values: Vec<Vec<Value>>,
}

/// The answer to a [`Request`].
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum Response {
    /// The ranges that were asked for.
    #[serde(rename_all = "camelCase")]
    ValueRanges { value_ranges: Vec<ValueRange> },
    /// The outputs the program declared, keyed by name.
    Outputs { outputs: Map<String, Value> },
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_request_is_what_the_cli_documents() {
        let r: Request = serde_json::from_str(
            r#"{"inputs": {"a": 1}, "ranges": ["S!A1"], "majorDimension": "COLUMNS"}"#,
        )
        .unwrap();
        assert_eq!(r.major_dimension, Some(MajorDimension::Columns));
        assert_eq!(r.ranges.as_deref(), Some(&["S!A1".to_string()][..]));
        assert!(serde_json::from_str::<Request>(r#"{"a": 1}"#).is_err());
    }

    #[test]
    fn a_response_is_one_of_the_two_shapes() {
        let r: Response = serde_json::from_str(
            r#"{"valueRanges": [{"range": "S!A1", "name": "x", "values": [[1]]}]}"#,
        )
        .unwrap();
        assert!(matches!(r, Response::ValueRanges { .. }));
        let o: Response = serde_json::from_str(r#"{"outputs": {"y": [[2]]}}"#).unwrap();
        assert!(matches!(o, Response::Outputs { .. }));
    }
}
