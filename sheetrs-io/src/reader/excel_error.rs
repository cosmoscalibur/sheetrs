//! The error values a cell can hold (`#DIV/0!`, `#REF!`, …): part of the workbook model.

use serde::{Deserialize, Serialize};
use std::fmt;
use std::str::FromStr;

/// Excel calculation error types.
///
/// Represents the finite set of error values that can appear in Excel/ODS cells.
/// Using an enum instead of raw strings saves memory (1 byte vs 8+ bytes for `Arc<str>`)
/// and enables integer-based comparisons at rule time.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ExcelError {
    /// `#NULL!` — Intersection of two ranges that don't intersect.
    Null,
    /// `#DIV/0!` — Division by zero.
    DivZero,
    /// `#VALUE!` — Wrong type of argument or operand.
    Value,
    /// `#REF!` — Invalid cell reference.
    Ref,
    /// `#NAME?` — Unrecognized formula name or text.
    Name,
    /// `#NUM!` — Invalid numeric value.
    Num,
    /// `#N/A` — Value not available.
    NA,
    /// `#SPILL!` — Spill range is not blank (dynamic arrays).
    Spill,
    /// `#CALC!` — Calculation engine error (dynamic arrays).
    Calc,
    /// An error of `LibreOffice` that Excel does not have, as an ODS keeps it: `Err:502` (invalid
    /// argument), `Err:522` (circular reference), `Err:523` (the iteration does not converge), …
    /// `LibreOffice` writes all of them to an XLSX as `#VALUE!`; an ODS says which it is, and so
    /// does the cell here. The ones that have an Excel name (`Err:524` is `#REF!`) read as that.
    Libre(u16),
}

impl ExcelError {
    /// All known error literals for formula-string scanning.
    pub const ERROR_LITERALS: &[&str] = &[
        "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!",
    ];

    /// Parse a cell error string (e.g. from `CellValue::Error` text content).
    ///
    /// Returns `None` for unrecognized patterns. Handles both exact matches
    /// (e.g. `"#REF!"`) and the codes of `LibreOffice` (`"Err:522"`).
    #[must_use]
    pub fn from_cell_str(s: &str) -> Option<Self> {
        s.parse().ok()
    }
}

impl fmt::Display for ExcelError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let s = match self {
            Self::Null => "#NULL!",
            Self::DivZero => "#DIV/0!",
            Self::Value => "#VALUE!",
            Self::Ref => "#REF!",
            Self::Name => "#NAME?",
            Self::Num => "#NUM!",
            Self::NA => "#N/A",
            Self::Spill => "#SPILL!",
            Self::Calc => "#CALC!",
            Self::Libre(code) => return write!(f, "Err:{code}"),
        };
        write!(f, "{s}")
    }
}

/// The text is not the name of an error value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParseExcelErrorError(pub String);

impl fmt::Display for ParseExcelErrorError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "unknown error value: {}", self.0)
    }
}

impl std::error::Error for ParseExcelErrorError {}

impl FromStr for ExcelError {
    type Err = ParseExcelErrorError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "#NULL!" => Ok(Self::Null),
            "#DIV/0!" => Ok(Self::DivZero),
            "#VALUE!" => Ok(Self::Value),
            "#REF!" => Ok(Self::Ref),
            "#NAME?" => Ok(Self::Name),
            "#NUM!" => Ok(Self::Num),
            "#N/A" => Ok(Self::NA),
            "#SPILL!" => Ok(Self::Spill),
            "#CALC!" => Ok(Self::Calc),
            _ => match s.strip_prefix("Err:").and_then(|c| c.parse::<u16>().ok()) {
                // The codes `LibreOffice` shows under an Excel name.
                Some(503) => Ok(Self::Num),
                Some(519) => Ok(Self::Value),
                Some(524) => Ok(Self::Ref),
                Some(525) => Ok(Self::Name),
                Some(532) => Ok(Self::DivZero),
                Some(code) => Ok(Self::Libre(code)),
                None => Err(ParseExcelErrorError(s.to_string())),
            },
        }
    }
}

impl Serialize for ExcelError {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        serializer.serialize_str(&self.to_string())
    }
}

impl<'de> Deserialize<'de> for ExcelError {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        s.parse().map_err(serde::de::Error::custom)
    }
}
