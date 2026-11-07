//! What is searched for and how, as `LibreOffice`'s Find & Replace describes it.
//!
//! The same [`Find`] drives a read-only search (`sheetrs-query`) and a replacement
//! (`sheetrs-edit`), so the two agree on which cells match.

use serde::{Deserialize, Serialize};

/// How the text is read.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Mode {
    /// The text as it is.
    #[default]
    Literal,
    /// A regular expression, as ICU reads it (`LibreOffice`'s engine).
    Regex,
    /// `*` any text, `?` one character, `\` takes the next character as it is (the escape of the
    /// Find & Replace of `LibreOffice`; `~` is the one of `COUNTIF`).
    Wildcards,
}

/// What of a cell is searched.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SearchIn {
    /// The value of the cell, as text.
    #[default]
    Value,
    /// The formula of a formula cell (and the text of any other cell).
    Formula,
}

/// A search.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Find {
    pub text: String,
    #[serde(default)]
    pub mode: Mode,
    /// Distinguish upper and lower case.
    #[serde(default)]
    pub match_case: bool,
    /// The text must be the whole content of the cell, not a part of it.
    #[serde(default)]
    pub entire_cells: bool,
    #[serde(default)]
    pub search_in: SearchIn,
    /// When searching values: also look at the result of the cells that hold a formula.
    /// Off by default in a query, so a search by value reaches what was typed, not what was
    /// computed. A replacement ignores it: `LibreOffice` always reaches formula cells there.
    #[serde(default)]
    pub include_formulas: bool,
}

impl Find {
    #[must_use]
    pub fn literal(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            mode: Mode::Literal,
            match_case: false,
            entire_cells: false,
            search_in: SearchIn::Value,
            include_formulas: false,
        }
    }
}

/// A search and what takes the place of what it finds. In a regular expression, `$1` is a group
/// and `&` the whole match, as in `LibreOffice`.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Replace {
    #[serde(flatten)]
    pub find: Find,
    pub with: String,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_search_needs_only_its_text() {
        let f: Find = serde_json::from_str(r#"{"text": "a"}"#).unwrap();
        assert_eq!(f, Find::literal("a"));
        assert!(serde_json::from_str::<Find>(r#"{"text": "a", "regex": true}"#).is_err());
    }

    #[test]
    fn a_replacement_is_a_search_and_a_text() {
        let r: Replace = serde_json::from_str(
            r#"{"text": "(a)", "mode": "regex", "match_case": true, "with": "$1$1"}"#,
        )
        .unwrap();
        assert_eq!(r.find.mode, Mode::Regex);
        assert_eq!(r.with, "$1$1");
    }
}
