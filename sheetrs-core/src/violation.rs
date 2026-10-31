//! Violation reporting system with hierarchical structure

use crate::reader::Workbook;
use crate::reader::parser_utils::column_letters;
use serde::ser::SerializeStruct;
use serde::{Deserialize, Serialize};
use std::cmp::Ordering;
use std::fmt;
use std::str::FromStr;
use std::sync::Arc;

/// Unique identifier for each linter rule.
///
/// Each variant corresponds to a specific rule implementation. The enum provides
/// type-safe rule identification with zero-cost string conversion.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum RuleId {
    // Excel Errors (1xx)
    /// ERR101: Broken Named Ranges
    Err101,
    /// ERR102: Error Cells
    Err102,
    /// ERR103: Reference to Error
    Err103,

    // Unreliable Calculations (2xx)
    /// CALC201: Hardcoded Number
    Calc201,
    /// CALC202: Circular Reference
    Calc202,
    /// CALC203: Double Operator
    Calc203,
    /// CALC204: Approximate Lookup
    Calc204,
    /// CALC205: Double Count
    Calc205,

    // Reference Issues (3xx)
    /// REF301: Unused Named Ranges
    Ref301,
    /// REF302: Duplicate Names
    Ref302,
    /// REF303: Empty Sheets
    Ref303,
    /// REF304: Large Used Range
    Ref304,
    /// REF305: Blank Rows/Columns
    Ref305,
    /// REF306: Unused Sheets
    Ref306,
    /// REF307: Whole Column/Row References
    Ref307,
    /// REF309: Reference to Empty Cell
    Ref309,
    /// REF310: Longer Reference Expected
    Ref310,

    // Formula Interruptions (4xx)
    /// INT401: Interrupted by Data
    Int401,
    /// INT402: Interrupted by Empty
    Int402,
    /// INT403: Interrupted by Other
    Int403,

    // Complexity (5xx)
    /// CPX501: Excessive Sheet Counts
    Cpx501,
    /// CPX502: Merged Cells
    Cpx502,
    /// CPX503: Excessive Conditional Formatting
    Cpx503,
    /// CPX504: Deep IF Nesting
    Cpx504,
    /// CPX505: Deep Formula Nesting
    Cpx505,
    /// CPX506: Many Operations
    Cpx506,
    /// CPX507: Multiple Sheet References
    Cpx507,
    /// CPX508: Many References
    Cpx508,
    /// CPX509: Long Formula
    Cpx509,

    // Vulnerable Formulas (6xx)
    /// VUL601: Duplicate Formulas
    Vul601,
    /// VUL602: Volatile Functions
    Vul602,
    /// VUL603: Empty String Test
    Vul603,
    /// VUL604: Error-Prone Functions
    Vul604,
    /// VUL605: Legacy Array
    Vul605,
    /// VUL606: Deprecated Function
    Vul606,

    // Data Issues (7xx)
    /// DATA701: Non-Descriptive Sheet Name
    Data701,
    /// DATA702: Inconsistent Number Format
    Data702,
    /// DATA703: Inconsistent Date Format
    Data703,
    /// DATA704: Long Text Cell
    Data704,
    /// DATA705: Unnecessary Space
    Data705,
    /// DATA706: Numeric Text in Calculation
    Data706,

    // External References (8xx)
    /// EXT801: External Workbook
    Ext801,
    /// EXT802: Web URLs
    Ext802,

    // Hidden Information (9xx)
    /// HID901: Hidden Worksheet
    Hid901,
    /// HID902: Hidden Columns/Rows
    Hid902,

    // Files & Settings (10xx)
    /// FILE1001: Large File Size
    File1001,
    /// FILE1002: Old Spreadsheet (time since the last modification)
    File1002,
    /// FILE1003: 1904 Date System
    File1003,

    // VBA Issues (11xx)
    /// VBA1101: Has Macros
    Vba1101,
}

impl RuleId {
    /// Returns the string representation of the rule ID (e.g., "ERR101").
    #[must_use]
    pub const fn as_str(&self) -> &'static str {
        match self {
            Self::Err101 => "ERR101",
            Self::Err102 => "ERR102",
            Self::Err103 => "ERR103",
            Self::Calc201 => "CALC201",
            Self::Calc202 => "CALC202",
            Self::Calc203 => "CALC203",
            Self::Calc204 => "CALC204",
            Self::Calc205 => "CALC205",
            Self::Ref301 => "REF301",
            Self::Ref302 => "REF302",
            Self::Ref303 => "REF303",
            Self::Ref304 => "REF304",
            Self::Ref305 => "REF305",
            Self::Ref306 => "REF306",
            Self::Ref307 => "REF307",
            Self::Ref309 => "REF309",
            Self::Ref310 => "REF310",
            Self::Int401 => "INT401",
            Self::Int402 => "INT402",
            Self::Int403 => "INT403",
            Self::Cpx501 => "CPX501",
            Self::Cpx502 => "CPX502",
            Self::Cpx503 => "CPX503",
            Self::Cpx504 => "CPX504",
            Self::Cpx505 => "CPX505",
            Self::Cpx506 => "CPX506",
            Self::Cpx507 => "CPX507",
            Self::Cpx508 => "CPX508",
            Self::Cpx509 => "CPX509",
            Self::Vul601 => "VUL601",
            Self::Vul602 => "VUL602",
            Self::Vul603 => "VUL603",
            Self::Vul604 => "VUL604",
            Self::Vul605 => "VUL605",
            Self::Vul606 => "VUL606",
            Self::Data701 => "DATA701",
            Self::Data702 => "DATA702",
            Self::Data703 => "DATA703",
            Self::Data704 => "DATA704",
            Self::Data705 => "DATA705",
            Self::Data706 => "DATA706",
            Self::Ext801 => "EXT801",
            Self::Ext802 => "EXT802",
            Self::Hid901 => "HID901",
            Self::Hid902 => "HID902",
            Self::File1001 => "FILE1001",
            Self::File1002 => "FILE1002",
            Self::File1003 => "FILE1003",
            Self::Vba1101 => "VBA1101",
        }
    }
}

impl fmt::Display for RuleId {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.as_str())
    }
}

/// Error type for parsing invalid rule ID strings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParseRuleIdError(pub String);

impl fmt::Display for ParseRuleIdError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "unknown rule ID: {}", self.0)
    }
}

impl std::error::Error for ParseRuleIdError {}

impl FromStr for RuleId {
    type Err = ParseRuleIdError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "ERR101" => Ok(Self::Err101),
            "ERR102" => Ok(Self::Err102),
            "ERR103" => Ok(Self::Err103),
            "CALC201" => Ok(Self::Calc201),
            "CALC202" => Ok(Self::Calc202),
            "CALC203" => Ok(Self::Calc203),
            "CALC204" => Ok(Self::Calc204),
            "CALC205" => Ok(Self::Calc205),
            "REF301" => Ok(Self::Ref301),
            "REF302" => Ok(Self::Ref302),
            "REF303" => Ok(Self::Ref303),
            "REF304" => Ok(Self::Ref304),
            "REF305" => Ok(Self::Ref305),
            "REF306" => Ok(Self::Ref306),
            "REF307" => Ok(Self::Ref307),
            "REF309" => Ok(Self::Ref309),
            "REF310" => Ok(Self::Ref310),
            "INT401" => Ok(Self::Int401),
            "INT402" => Ok(Self::Int402),
            "INT403" => Ok(Self::Int403),
            "CPX501" => Ok(Self::Cpx501),
            "CPX502" => Ok(Self::Cpx502),
            "CPX503" => Ok(Self::Cpx503),
            "CPX504" => Ok(Self::Cpx504),
            "CPX505" => Ok(Self::Cpx505),
            "CPX506" => Ok(Self::Cpx506),
            "CPX507" => Ok(Self::Cpx507),
            "CPX508" => Ok(Self::Cpx508),
            "CPX509" => Ok(Self::Cpx509),
            "VUL601" => Ok(Self::Vul601),
            "VUL602" => Ok(Self::Vul602),
            "VUL603" => Ok(Self::Vul603),
            "VUL604" => Ok(Self::Vul604),
            "VUL605" => Ok(Self::Vul605),
            "VUL606" => Ok(Self::Vul606),
            "DATA701" => Ok(Self::Data701),
            "DATA702" => Ok(Self::Data702),
            "DATA703" => Ok(Self::Data703),
            "DATA704" => Ok(Self::Data704),
            "DATA705" => Ok(Self::Data705),
            "DATA706" => Ok(Self::Data706),
            "EXT801" => Ok(Self::Ext801),
            "EXT802" => Ok(Self::Ext802),
            "HID901" => Ok(Self::Hid901),
            "HID902" => Ok(Self::Hid902),
            "FILE1001" => Ok(Self::File1001),
            "FILE1002" => Ok(Self::File1002),
            "FILE1003" => Ok(Self::File1003),
            "VBA1101" => Ok(Self::Vba1101),
            _ => Err(ParseRuleIdError(s.to_string())),
        }
    }
}

impl Serialize for RuleId {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        serializer.serialize_str(self.as_str())
    }
}

impl<'de> Deserialize<'de> for RuleId {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        s.parse().map_err(serde::de::Error::custom)
    }
}

/// Excel calculation error types.
///
/// Represents the finite set of error values that can appear in Excel/ODS cells.
/// Using an enum instead of raw strings saves memory (1 byte vs 8+ bytes for `Arc<str>`)
/// and enables integer-based comparisons at rule time.
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
}

impl ExcelError {
    /// All known error literals for formula-string scanning.
    pub const ERROR_LITERALS: &[&str] = &[
        "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!",
    ];

    /// Parse a cell error string (e.g. from `CellValue::Error` text content).
    ///
    /// Returns `None` for unrecognized patterns. Handles both exact matches
    /// (e.g. `"#REF!"`) and ODS error codes (e.g. `"Err:508"`).
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
        };
        write!(f, "{s}")
    }
}

impl FromStr for ExcelError {
    type Err = ParseRuleIdError;

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
            _ => Err(ParseRuleIdError(s.to_string())),
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

/// Severity level of a violation
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
pub enum Severity {
    Info,
    Warning,
    Error,
}

/// Scope of a violation (book, sheet, or cell level)
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum ViolationScope {
    /// Book-level violation
    Book,
    /// Sheet-level violation with 0-based sheet index
    Sheet(u16),
    /// Cell-level violation with 0-based sheet index
    Cell(u16, CellReference),
}

impl ViolationScope {
    /// Get the sheet index if this is a sheet or cell scope
    pub fn sheet_index(&self) -> Option<u16> {
        match self {
            ViolationScope::Book => None,
            ViolationScope::Sheet(idx) => Some(*idx),
            ViolationScope::Cell(idx, _) => Some(*idx),
        }
    }
}

impl PartialOrd for ViolationScope {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for ViolationScope {
    fn cmp(&self, other: &Self) -> Ordering {
        match (self, other) {
            (ViolationScope::Book, ViolationScope::Book) => Ordering::Equal,
            (ViolationScope::Book, _) => Ordering::Less,
            (_, ViolationScope::Book) => Ordering::Greater,
            (ViolationScope::Sheet(a), ViolationScope::Sheet(b)) => a.cmp(b),
            (ViolationScope::Sheet(_), ViolationScope::Cell(_, _)) => Ordering::Less,
            (ViolationScope::Cell(_, _), ViolationScope::Sheet(_)) => Ordering::Greater,
            (ViolationScope::Cell(sheet_a, cell_a), ViolationScope::Cell(sheet_b, cell_b)) => {
                sheet_a.cmp(sheet_b).then_with(|| cell_a.cmp(cell_b))
            }
        }
    }
}

/// Cell reference (e.g., A1, B2)
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CellReference {
    pub row: u32,
    pub col: u32,
}

impl CellReference {
    pub fn new(row: u32, col: u32) -> Self {
        Self { row, col }
    }
}

impl PartialOrd for CellReference {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for CellReference {
    fn cmp(&self, other: &Self) -> Ordering {
        self.row
            .cmp(&other.row)
            .then_with(|| self.col.cmp(&other.col))
    }
}

impl std::fmt::Display for CellReference {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}{}", column_letters(self.col), self.row + 1)
    }
}

/// Context provided to [`ViolationData`] for human-readable message formatting.
///
/// The writer/formatter constructs this context once and passes it when
/// rendering violation messages, enabling index-to-name resolution.
pub struct FormatContext<'a> {
    /// Reference to the workbook for resolving sheet indexes to names.
    pub workbook: &'a Workbook,
}

/// Trait for compact violation incident data.
///
/// Each walker rule defines its own struct implementing this trait. The data
/// struct stores only observed facts (integer indexes, tuples) — not
/// thresholds, configuration, or pre-formatted strings. The human-readable
/// message is produced at output time via [`format_message`](ViolationData::format_message).
pub trait ViolationData: fmt::Debug + Send + Sync {
    /// Produce the human-readable message, resolving indexes via context.
    fn format_message(&self, ctx: &FormatContext<'_>) -> String;

    /// Downcast support for testing and programmatic access.
    fn as_any(&self) -> &dyn std::any::Any;
}

/// Internal storage for violation messages.
///
/// Supports two paths: pre-formatted strings and deferred data
/// (used by rules implementing [`ViolationData`]).
#[derive(Debug, Clone)]
enum ViolationMessageInner {
    /// Pre-formatted message string.
    Legacy(String),
    /// Compact incident data with deferred formatting (walker path).
    Data(Arc<dyn ViolationData>),
}

/// A linter violation
#[derive(Debug, Clone)]
pub struct Violation {
    /// Unique rule identifier
    pub rule_id: RuleId,
    /// Scope of the violation
    pub scope: ViolationScope,
    /// Severity level
    pub severity: Severity,
    /// Violation content (legacy string or deferred data).
    message_inner: ViolationMessageInner,
}

impl Violation {
    /// Creates a new violation with a pre-formatted message string.
    ///
    /// Prefer [`with_data`](Violation::with_data) for compact, deferred
    /// formatting when a [`ViolationData`] struct is available.
    pub fn new(
        rule_id: RuleId,
        scope: ViolationScope,
        message: impl Into<String>,
        severity: Severity,
    ) -> Self {
        Self {
            rule_id,
            scope,
            severity,
            message_inner: ViolationMessageInner::Legacy(message.into()),
        }
    }

    /// Creates a new violation with compact incident data.
    ///
    /// The human-readable message is deferred to output time via
    /// [`format_message`](Violation::format_message).
    pub fn with_data(
        rule_id: RuleId,
        scope: ViolationScope,
        data: impl ViolationData + 'static,
        severity: Severity,
    ) -> Self {
        Self {
            rule_id,
            scope,
            severity,
            message_inner: ViolationMessageInner::Data(Arc::new(data)),
        }
    }

    /// Produce the human-readable message.
    ///
    /// For legacy violations, returns the stored string. For data-backed
    /// violations, calls [`ViolationData::format_message`] with the context.
    pub fn format_message(&self, ctx: &FormatContext<'_>) -> String {
        match &self.message_inner {
            ViolationMessageInner::Legacy(msg) => msg.clone(),
            ViolationMessageInner::Data(data) => data.format_message(ctx),
        }
    }

    /// Returns the legacy message string, if present.
    ///
    /// Returns `None` for data-backed violations (use
    /// [`format_message`](Violation::format_message) instead).
    pub fn legacy_message(&self) -> Option<&str> {
        match &self.message_inner {
            ViolationMessageInner::Legacy(msg) => Some(msg),
            ViolationMessageInner::Data(_) => None,
        }
    }

    /// Attempt to downcast the incident data to a concrete type.
    ///
    /// Returns `None` if the violation uses a legacy message or if the
    /// concrete type does not match.
    pub fn data<T: ViolationData + 'static>(&self) -> Option<&T> {
        match &self.message_inner {
            ViolationMessageInner::Data(data) => data.as_any().downcast_ref::<T>(),
            ViolationMessageInner::Legacy(_) => None,
        }
    }

    /// Convenience accessor: renders the human-readable message.
    ///
    /// For legacy violations, returns the stored string.
    /// For data-backed violations, renders using a default [`Workbook`] context.
    /// Prefer [`format_message`](Violation::format_message) when a real workbook
    /// is available (e.g., in the formatter) for accurate name resolution.
    pub fn message(&self) -> String {
        match &self.message_inner {
            ViolationMessageInner::Legacy(msg) => msg.clone(),
            ViolationMessageInner::Data(data) => {
                let wb = Workbook::default();
                let ctx = FormatContext { workbook: &wb };
                data.format_message(&ctx)
            }
        }
    }
}

impl PartialEq for Violation {
    fn eq(&self, other: &Self) -> bool {
        self.rule_id == other.rule_id
            && self.scope == other.scope
            && self.severity == other.severity
    }
}

impl Eq for Violation {}

impl PartialOrd for Violation {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for Violation {
    fn cmp(&self, other: &Self) -> Ordering {
        self.scope
            .cmp(&other.scope)
            .then_with(|| self.rule_id.cmp(&other.rule_id))
    }
}

impl Serialize for Violation {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        let mut state = serializer.serialize_struct("Violation", 4)?;
        state.serialize_field("rule_id", &self.rule_id)?;
        state.serialize_field("scope", &self.scope)?;
        // For legacy messages, serialize the stored string.
        // For data-backed violations, serialize a placeholder (the
        // formatter should pre-render via format_message before JSON output).
        let message = match &self.message_inner {
            ViolationMessageInner::Legacy(msg) => msg.clone(),
            ViolationMessageInner::Data(data) => format!("{data:?}"),
        };
        state.serialize_field("message", &message)?;
        state.serialize_field("severity", &self.severity)?;
        state.end()
    }
}
