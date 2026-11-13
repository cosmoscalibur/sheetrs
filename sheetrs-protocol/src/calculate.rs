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
/// Without `ranges` the answer is the ranges of the outputs the program declared (each one named);
/// with them, exactly those ranges. A key of `inputs` is an input, an alias (a defined name inside the inputs) or a range
/// of constants of the program.
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
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct ValueRange {
    /// The defined name or output name it was asked by. It is written first, as the engine writes
    /// it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// The range, canonical (`Sheet1!B2:C9`).
    pub range: String,
    pub values: Vec<Vec<Value>>,
}

/// What a calculation answers: the values of the ranges that were asked for (or, when none was, of
/// the outputs the program declares), in the shape of the protocol: `{"valueRanges": [...]}`, as the
/// batch read of a spreadsheet API answers.
///
/// It is typed for the layers that hand it over as an object ([`ValueRanges::to_entries`]) and for the
/// ones that write it as JSON ([`ValueRanges::to_json`], [`ValueRanges::write_json`]): the same content
/// in two forms, as the `to_dict` and `to_json` of a table are.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields, rename_all = "camelCase")]
pub struct ValueRanges {
    /// One entry per range, in the order of the request.
    pub value_ranges: Vec<ValueRange>,
}

impl ValueRanges {
    /// The value ranges of the JSON answer of a calculation (`{"valueRanges": [...]}`).
    ///
    /// # Errors
    ///
    /// When the JSON is not that shape.
    pub fn from_json(answer: Value) -> Result<Self, serde_json::Error> {
        serde_json::from_value(answer)
    }

    /// The answer with its arrays laid out as `major` says: the arrays of rows the engine gives
    /// (`Rows`) or their transposition, arrays of columns (`Columns`).
    #[must_use]
    pub fn with_major(mut self, major: MajorDimension) -> Self {
        if major == MajorDimension::Rows {
            return self;
        }
        fn transpose(rows: &[Vec<Value>]) -> Vec<Vec<Value>> {
            let width = rows.iter().map(Vec::len).max().unwrap_or(0);
            (0..width)
                .map(|c| {
                    rows.iter()
                        .map(|r| r.get(c).cloned().unwrap_or(Value::Null))
                        .collect()
                })
                .collect()
        }
        for range in &mut self.value_ranges {
            range.values = transpose(&range.values);
        }
        self
    }

    /// The JSON answer of the protocol, as `sheetc calculate --json` writes it.
    #[must_use]
    pub fn to_json(&self) -> Value {
        serde_json::to_value(self).unwrap_or(Value::Null)
    }

    /// Writes the JSON answer followed by a newline.
    ///
    /// # Errors
    ///
    /// When `out` fails.
    pub fn write_json(&self, mut out: impl std::io::Write) -> std::io::Result<()> {
        serde_json::to_writer(&mut out, self).map_err(std::io::Error::from)?;
        out.write_all(b"\n")
    }

    /// The values of each range, by name, in order: `(name, rows)` with the rows as a JSON array of
    /// arrays. With `asked` (the ranges the calculation was asked for, as they were written), each
    /// range is named by the text it was asked with; otherwise by its name, or by its range when it
    /// has none.
    #[must_use]
    pub fn to_entries(&self, asked: Option<&[String]>) -> Vec<(String, Value)> {
        self.value_ranges
            .iter()
            .enumerate()
            .map(|(i, range)| {
                let name = asked
                    .and_then(|asked| asked.get(i).cloned())
                    .or_else(|| range.name.clone())
                    .unwrap_or_else(|| range.range.clone());
                (name, Value::from(range.values.clone()))
            })
            .collect()
    }
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
    fn a_response_is_the_value_ranges_of_the_protocol() {
        let r: ValueRanges = serde_json::from_str(
            r#"{"valueRanges": [{"range": "S!A1", "name": "x", "values": [[1]]}]}"#,
        )
        .unwrap();
        assert_eq!(r.value_ranges.len(), 1);
        assert!(serde_json::from_str::<ValueRanges>(r#"{"outputs": {"y": [[2]]}}"#).is_err());
    }

    #[test]
    fn value_ranges_are_named_by_what_was_asked_then_by_their_name_then_by_their_range() {
        let answer = serde_json::json!({"valueRanges": [
            {"range": "S!A1", "name": "x", "values": [[1]]},
            {"range": "S!B1:B2", "values": [[2], [3]]},
        ]});
        let value_ranges = ValueRanges::from_json(answer.clone()).unwrap();
        assert_eq!(value_ranges.to_json(), answer);
        let names = |asked: Option<&[String]>| {
            value_ranges
                .to_entries(asked)
                .into_iter()
                .map(|(name, _)| name)
                .collect::<Vec<_>>()
        };
        assert_eq!(names(None), ["x", "S!B1:B2"]);
        assert_eq!(names(Some(&["a".to_string(), "b".to_string()])), ["a", "b"]);
    }

    #[test]
    fn columns_are_the_transposition_of_the_rows() {
        let rows = ValueRanges::from_json(serde_json::json!({"valueRanges": [
            {"range": "S!A1:B2", "values": [[1, 2], [3, 4]]},
        ]}))
        .unwrap();
        assert_eq!(
            rows.with_major(MajorDimension::Columns).to_json(),
            serde_json::json!({"valueRanges": [
                {"range": "S!A1:B2", "values": [[1, 3], [2, 4]]},
            ]})
        );
    }

    #[test]
    fn the_json_of_the_value_ranges_ends_with_a_newline() {
        let value_ranges = ValueRanges::from_json(serde_json::json!({"valueRanges": [
            {"range": "S!A1", "values": [[2]]},
        ]}))
        .unwrap();
        let mut text = Vec::new();
        value_ranges.write_json(&mut text).unwrap();
        assert_eq!(
            text,
            b"{\"valueRanges\":[{\"range\":\"S!A1\",\"values\":[[2]]}]}\n"
        );
    }
}
