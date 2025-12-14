//! Violation reporting system with hierarchical structure

use serde::{Deserialize, Serialize};
use std::cmp::Ordering;

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
    /// Sheet-level violation
    Sheet(String),
    /// Cell-level violation
    Cell(String, CellReference),
}

impl ViolationScope {
    /// Get the sheet name if this is a sheet or cell scope
    pub fn sheet_name(&self) -> Option<&str> {
        match self {
            ViolationScope::Book => None,
            ViolationScope::Sheet(name) => Some(name),
            ViolationScope::Cell(name, _) => Some(name),
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

    /// Convert to Excel-style reference (e.g., "A1")
    pub fn to_excel_ref(&self) -> String {
        format!("{}{}", Self::col_to_letter(self.col), self.row + 1)
    }

    /// Convert column number to letter (0 -> A, 1 -> B, etc.)
    fn col_to_letter(mut col: u32) -> String {
        let mut result = String::new();
        loop {
            result.insert(0, (b'A' + (col % 26) as u8) as char);
            if col < 26 {
                break;
            }
            col = col / 26 - 1;
        }
        result
    }
}

impl PartialOrd for CellReference {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for CellReference {
    fn cmp(&self, other: &Self) -> Ordering {
        self.row.cmp(&other.row).then_with(|| self.col.cmp(&other.col))
    }
}

impl std::fmt::Display for CellReference {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.to_excel_ref())
    }
}

/// A linter violation
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Violation {
    /// Rule ID (e.g., "ERR001")
    pub rule_id: String,
    /// Scope of the violation
    pub scope: ViolationScope,
    /// Human-readable message
    pub message: String,
    /// Severity level
    pub severity: Severity,
}

impl Violation {
    pub fn new(rule_id: impl Into<String>, scope: ViolationScope, message: impl Into<String>, severity: Severity) -> Self {
        Self {
            rule_id: rule_id.into(),
            scope,
            message: message.into(),
            severity,
        }
    }
}

impl PartialOrd for Violation {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for Violation {
    fn cmp(&self, other: &Self) -> Ordering {
        self.scope.cmp(&other.scope)
            .then_with(|| self.rule_id.cmp(&other.rule_id))
    }
}
