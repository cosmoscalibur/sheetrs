//! EXT802: Web URLs
//!
//! A web, FTP or file address stored in a cell is an outbound link: it can break, point at a
//! resource that changes, leak an internal location or be a phishing target. The rule makes such
//! links visible so they can be checked.
//!
//! # Reports
//!
//! One violation per contiguous block of cells containing the same URL (unit: a range, scoped to
//! the sheet; a single cell is reported as a cell). A URL is any text starting `http://`,
//! `https://`, `ftp://` or `file://` and running to the next whitespace, found in the text value
//! of a cell (including the saved result of a formula). Cells that touch by row or column
//! adjacency and hold the same URL form one block, and a cell with several URLs belongs to a block
//! of each. The message gives the URL and the bounding range.
//!
//! The URL is cut at whitespace only, so punctuation that follows it in a sentence (`.`, `)`, `"`)
//! is part of the URL as reported.
//!
//! # Does not report
//!
//! - Numbers, dates and other non-text values, and the text of a formula itself: only the stored
//!   text value is searched, so a `HYPERLINK("https://...")` formula is seen only if its saved
//!   result contains the address.
//! - Hyperlinks stored outside the cell value (hyperlink relationships, the display text of a
//!   link).
//! - Addresses without a scheme (`www.example.com`, `mailto:`).
//! - Links to other workbooks, which are EXT801.
//!
//! # Parameters
//!
//! - `url_links_status` (string): `ALL` reports every URL; `INVALID` checks each distinct URL with
//!   an HTTP `HEAD` request and reports only those that do not answer with a success status (the
//!   message then says "Invalid external URL"). Default `ALL`; any other value is treated as `ALL`.
//! - `url_timeout_seconds` (integer): timeout of each check, in seconds. Default `5`.
//!
//! The check needs the `link-validation` feature (on by default) and is not available for WASM; if
//! the feature is off the check always succeeds, so with `INVALID` nothing is reported.
//!
//! # Severity
//!
//! Always `Warning`.

use super::helpers::{bounding_box, find_contiguous_ranges};
use super::{LinterContext, RuleCategory, WalkerRule};
use crate::config::LinterConfig;
use crate::reader::{Cell, Sheet};
use crate::violation::{
    CellReference, FormatContext, RuleId, Severity, Violation, ViolationData, ViolationScope,
};
use regex::Regex;
use std::collections::{BTreeMap, HashMap};
use std::sync::LazyLock;

/// Per-sheet URL cell collection: sheet_index → Vec<(row, col, url)>.
type SheetUrlCellMap = HashMap<u16, Vec<(u32, u32, String)>>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkStatus {
    All,     // Report all URLs
    Invalid, // Only report invalid/inaccessible URLs
}

impl LinkStatus {
    fn from_str(s: &str) -> Self {
        match s.to_uppercase().as_str() {
            "INVALID" => LinkStatus::Invalid,
            "ALL" => LinkStatus::All,
            _ => LinkStatus::All, // Default: ALL
        }
    }
}

/// Rule that checks for external web URLs in cell values.
///
/// External links can be a security risk or simply undesirable in certain contexts.
/// This rule can be configured to report all URLs or only broken ones (if link validation is enabled).
///
/// # Configuration
///
/// * `url_links_status` - Status to check: "ALL" (default, all URLs) or "INVALID" (only broken URLs).
/// * `url_timeout_seconds` - Timeout for link validation in seconds (default: 5).
pub struct WebUrlsRule {
    status: LinkStatus,
    timeout_secs: u64,
    /// Cells with URLs per sheet.
    sheet_cells: SheetUrlCellMap,
}

impl WebUrlsRule {
    /// Create a new instance configured from the linter config.
    pub fn new(config: &LinterConfig) -> Self {
        let status = config
            .get_param_str("url_links_status", None)
            .map(LinkStatus::from_str)
            .unwrap_or(LinkStatus::All);

        let timeout_secs = config
            .get_param_int("url_timeout_seconds", None)
            .unwrap_or(5) as u64;

        Self {
            status,
            timeout_secs,
            sheet_cells: HashMap::new(),
        }
    }
}

impl Default for WebUrlsRule {
    fn default() -> Self {
        Self {
            status: LinkStatus::All,
            timeout_secs: 5,
            sheet_cells: HashMap::new(),
        }
    }
}

/// Incident data for EXT802.
#[derive(Debug)]
pub struct WebUrlData {
    /// The detected URL.
    pub url: String,
    /// Range as (start_row, start_col, end_row, end_col).
    pub range: (u32, u32, u32, u32),
    /// Whether the URL was determined to be invalid/inaccessible.
    pub is_invalid: bool,
}

impl ViolationData for WebUrlData {
    fn format_message(&self, _ctx: &FormatContext<'_>) -> String {
        let (sr, sc, er, ec) = self.range;
        let range_str = if sr == er && sc == ec {
            CellReference::new(sr, sc).to_string()
        } else {
            format!(
                "{}:{}",
                CellReference::new(sr, sc),
                CellReference::new(er, ec)
            )
        };

        if self.is_invalid {
            format!(
                "Invalid external URL '{}' (not accessible) in range: {}",
                self.url, range_str
            )
        } else {
            format!("External URL '{}' found in range: {}", self.url, range_str)
        }
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

static URL_PATTERN: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(https?://|ftp://|file://)[^\s]+").unwrap());

/// Extract URLs from text using regex
fn extract_urls(text: &str) -> Vec<String> {
    URL_PATTERN
        .find_iter(text)
        .map(|m| m.as_str().to_string())
        .collect()
}

/// Check if a URL is accessible (returns true if accessible, false otherwise)
#[cfg(all(feature = "link-validation", not(target_arch = "wasm32")))]
fn check_url_status(url: &str, timeout_secs: u64) -> bool {
    use reqwest::blocking::Client;
    use std::time::Duration;

    let Ok(client) = Client::builder()
        .timeout(Duration::from_secs(timeout_secs))
        .build()
    else {
        return false;
    };

    // Use HEAD request to avoid downloading content
    match client.head(url).send() {
        Ok(response) => response.status().is_success(),
        Err(_) => false,
    }
}

#[cfg(any(not(feature = "link-validation"), target_arch = "wasm32"))]
fn check_url_status(_url: &str, _timeout_secs: u64) -> bool {
    // Fallback: assume valid if feature not enabled
    true
}

impl WalkerRule for WebUrlsRule {
    fn id(&self) -> RuleId {
        RuleId::Ext802
    }

    fn name(&self) -> &'static str {
        "Web URLs"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::External
    }

    fn on_cell(
        &mut self,
        sheet: &Sheet,
        at: (u32, u32),
        cell: &Cell,
        _ctx: &mut LinterContext,
        _out: &mut Vec<Violation>,
    ) {
        if let crate::reader::workbook::CellValue::Text(text) = &cell.value {
            let urls = extract_urls(text);
            if !urls.is_empty() {
                self.sheet_cells
                    .entry(sheet.sheet_index)
                    .or_default()
                    .extend(urls.into_iter().map(|url| (at.0, at.1, url)));
            }
        }
    }

    fn on_sheet_end(&mut self, sheet: &Sheet, _ctx: &mut LinterContext, out: &mut Vec<Violation>) {
        let Some(mut cells) = self.sheet_cells.remove(&sheet.sheet_index) else {
            return;
        };
        cells.sort_unstable();
        // Group cells by URL
        let mut grouped: BTreeMap<String, Vec<(u32, u32)>> = BTreeMap::new();
        for (row, col, url) in cells {
            grouped.entry(url).or_default().push((row, col));
        }

        for (url, group_cells) in grouped {
            // Validate URL status if configured for INVALID-only mode
            if matches!(self.status, LinkStatus::Invalid)
                && check_url_status(&url, self.timeout_secs)
            {
                continue; // Skip valid URLs
            }

            let is_invalid = matches!(self.status, LinkStatus::Invalid);

            let ranges = find_contiguous_ranges(&group_cells);
            for range in ranges {
                out.push(Violation::with_data(
                    RuleId::Ext802,
                    ViolationScope::Sheet(sheet.sheet_index),
                    WebUrlData {
                        url: url.clone(),
                        range: bounding_box(&range),
                        is_invalid,
                    },
                    Severity::Warning,
                ));
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use crate::rules::LinterContext;
    use crate::rules::collect_violations;
    use std::collections::HashMap;
    use std::sync::Arc;

    #[test]
    fn test_url_in_text_cell() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                value: CellValue::Text(Arc::from("https://example.com")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let mut rule = WebUrlsRule::default();
        let mut ctx = LinterContext::default();

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }

        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, RuleId::Ext802);
    }

    #[test]
    fn test_multiple_urls_in_cell() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                formula: None,
                is_array: false,
                num_fmt: None,
                value: CellValue::Text(Arc::from("Check https://example.com and https://test.org")),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            sheet_index: 0,
            cells: cells.into(),
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let mut rule = WebUrlsRule::default();
        let mut ctx = LinterContext::default();

        for (at, cell) in sheet.cells.iter() {
            rule.on_cell(&sheet, at, cell, &mut ctx, &mut Vec::new());
        }

        let violations = collect_violations(|out| rule.on_sheet_end(&sheet, &mut ctx, out));

        // Should detect both URLs
        assert_eq!(violations.len(), 2);
    }
}
