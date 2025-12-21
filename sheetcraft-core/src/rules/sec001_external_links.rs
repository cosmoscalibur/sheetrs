//! SEC001: External links detection

use super::{LinterRule, RuleCategory};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkDetectionMode {
    Url,
    Workbook,
    All,
}

impl LinkDetectionMode {
    fn from_str(s: &str) -> Self {
        match s.to_uppercase().as_str() {
            "URL" => LinkDetectionMode::Url,
            "WORKBOOK" => LinkDetectionMode::Workbook,
            "ALL" => LinkDetectionMode::All,
            _ => LinkDetectionMode::Workbook, // Default
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkStatus {
    All,     // Report all links
    Invalid, // Only report invalid links
}

impl LinkStatus {
    fn from_str(s: &str) -> Self {
        match s.to_uppercase().as_str() {
            "INVALID" => LinkStatus::Invalid,
            "ALL" => LinkStatus::All,
            _ => LinkStatus::All, // Default
        }
    }
}

pub struct ExternalLinksRule {
    mode: LinkDetectionMode,
    status: LinkStatus,
    timeout_secs: u64,
}

impl ExternalLinksRule {
    pub fn new(config: &LinterConfig) -> Self {
        let mode = config
            .get_param_str("external_links_type", None)
            .map(LinkDetectionMode::from_str)
            .unwrap_or(LinkDetectionMode::Workbook);

        let status = config
            .get_param_str("external_links_status", None)
            .map(LinkStatus::from_str)
            .unwrap_or(LinkStatus::All);

        let timeout_secs = config
            .get_param_int("url_timeout_seconds", None)
            .unwrap_or(5) as u64;

        Self {
            mode,
            status,
            timeout_secs,
        }
    }
}

impl Default for ExternalLinksRule {
    fn default() -> Self {
        Self {
            mode: LinkDetectionMode::Workbook,
            status: LinkStatus::All,
            timeout_secs: 5,
        }
    }
}

impl LinterRule for ExternalLinksRule {
    fn id(&self) -> &str {
        "SEC001"
    }

    fn name(&self) -> &str {
        "External links"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::SecurityAndPrivacy
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();
        let mut seen_workbooks = std::collections::HashSet::new();
        let mut seen_urls = std::collections::HashSet::new();

        for sheet in &workbook.sheets {
            // Collect all cells with external workbook references
            let mut workbook_cells: Vec<(u32, u32, String)> = Vec::new();
            let mut url_cells: Vec<(u32, u32, String)> = Vec::new();

            for cell in sheet.all_cells() {
                // Check for external links in formulas (workbook references)
                if matches!(
                    self.mode,
                    LinkDetectionMode::Workbook | LinkDetectionMode::All
                ) {
                    if let Some(formula) = cell.value.as_formula() {
                        let workbook_names = extract_external_workbook_names(formula);
                        for workbook_name in workbook_names {
                            workbook_cells.push((cell.row, cell.col, workbook_name));
                        }
                    }
                }

                // Check for external links in text values (URLs)
                if matches!(self.mode, LinkDetectionMode::Url | LinkDetectionMode::All) {
                    if let crate::reader::workbook::CellValue::Text(text) = &cell.value {
                        if is_external_url(text) {
                            url_cells.push((cell.row, cell.col, text.clone()));
                        }
                    }
                }
            }

            // Group workbook references by workbook name and create range-based violations
            if !workbook_cells.is_empty() {
                let grouped = group_cells_by_value(workbook_cells);
                for (workbook_name, cells) in grouped {
                    // Skip if already processed (deduplication)
                    if seen_workbooks.contains(&workbook_name) {
                        continue;
                    }
                    seen_workbooks.insert(workbook_name.clone());

                    // Validate workbook existence if status is INVALID
                    if matches!(self.status, LinkStatus::Invalid) {
                        if check_workbook_exists(&workbook.path, &workbook_name) {
                            continue; // Skip valid workbooks
                        }
                    }

                    let ranges = find_contiguous_ranges(&cells);

                    // Create a separate violation for each contiguous range
                    for range in ranges {
                        let range_str = format_single_range(&range);
                        let message = if matches!(self.status, LinkStatus::Invalid) {
                            format!(
                                "Invalid external workbook reference {} (file not found) in range: {}",
                                workbook_name, range_str
                            )
                        } else {
                            format!(
                                "External workbook reference {} found in range: {}",
                                workbook_name, range_str
                            )
                        };
                        violations.push(Violation::new(
                            self.id(),
                            ViolationScope::Sheet(sheet.name.clone()),
                            message,
                            Severity::Warning,
                        ));
                    }
                }
            }

            // Group URLs and create range-based violations
            if !url_cells.is_empty() {
                let grouped = group_cells_by_value(url_cells);
                for (url, cells) in grouped {
                    // Skip if already processed (deduplication)
                    if seen_urls.contains(&url) {
                        continue;
                    }
                    seen_urls.insert(url.clone());

                    // Validate URL status if status is INVALID
                    if matches!(self.status, LinkStatus::Invalid) {
                        if check_url_status(&url, self.timeout_secs) {
                            continue; // Skip valid URLs
                        }
                    }

                    let ranges = find_contiguous_ranges(&cells);

                    // Create a separate violation for each contiguous range
                    for range in ranges {
                        let range_str = format_single_range(&range);
                        let message = if matches!(self.status, LinkStatus::Invalid) {
                            format!(
                                "Invalid external URL '{}' (not accessible) in range: {}",
                                url, range_str
                            )
                        } else {
                            format!("External URL '{}' found in range: {}", url, range_str)
                        };
                        violations.push(Violation::new(
                            self.id(),
                            ViolationScope::Sheet(sheet.name.clone()),
                            message,
                            Severity::Warning,
                        ));
                    }
                }
            }
        }

        // Check for external links found in metadata (book-level)
        for link in &workbook.external_links {
            if seen_workbooks.contains(link) || seen_urls.contains(link) {
                continue;
            }

            let message = format!("External link '{}' found in workbook metadata.", link);
            violations.push(Violation::new(
                self.id(),
                ViolationScope::Book,
                message,
                Severity::Warning,
            ));
        }

        Ok(violations)
    }
}

/// Group cells by their associated value (workbook name or URL)
fn group_cells_by_value(cells: Vec<(u32, u32, String)>) -> Vec<(String, Vec<(u32, u32)>)> {
    use std::collections::HashMap;

    let mut grouped: HashMap<String, Vec<(u32, u32)>> = HashMap::new();
    for (row, col, value) in cells {
        grouped
            .entry(value)
            .or_insert_with(Vec::new)
            .push((row, col));
    }

    grouped.into_iter().collect()
}

/// Format a single contiguous range
fn format_single_range(cells: &[(u32, u32)]) -> String {
    use crate::violation::CellReference;

    if cells.is_empty() {
        return String::new();
    }

    if cells.len() == 1 {
        return CellReference::new(cells[0].0, cells[0].1).to_string();
    }

    let min_row = cells.iter().map(|(r, _)| r).min().unwrap();
    let max_row = cells.iter().map(|(r, _)| r).max().unwrap();
    let min_col = cells.iter().map(|(_, c)| c).min().unwrap();
    let max_col = cells.iter().map(|(_, c)| c).max().unwrap();

    let start = CellReference::new(*min_row, *min_col);
    let end = CellReference::new(*max_row, *max_col);

    format!("{}:{}", start, end)
}

/// Find contiguous ranges from a list of cells
fn find_contiguous_ranges(cells: &[(u32, u32)]) -> Vec<Vec<(u32, u32)>> {
    use std::collections::{HashSet, VecDeque};

    let cell_set: HashSet<(u32, u32)> = cells.iter().copied().collect();
    let mut visited: HashSet<(u32, u32)> = HashSet::new();
    let mut ranges: Vec<Vec<(u32, u32)>> = Vec::new();

    for &cell in cells {
        if visited.contains(&cell) {
            continue;
        }

        // BFS to find all connected cells
        let mut range = Vec::new();
        let mut queue = VecDeque::new();
        queue.push_back(cell);
        visited.insert(cell);

        while let Some((row, col)) = queue.pop_front() {
            range.push((row, col));

            // Check all 4 adjacent cells (up, down, left, right)
            let neighbors = [
                (row.wrapping_sub(1), col),
                (row + 1, col),
                (row, col.wrapping_sub(1)),
                (row, col + 1),
            ];

            for neighbor in neighbors {
                if cell_set.contains(&neighbor) && !visited.contains(&neighbor) {
                    visited.insert(neighbor);
                    queue.push_back(neighbor);
                }
            }
        }

        ranges.push(range);
    }

    ranges
}

/// Extract external workbook names from formula, returns empty list if no external reference
fn extract_external_workbook_names(formula: &str) -> Vec<String> {
    let mut names = Vec::new();
    let mut chars = formula.chars().peekable();
    let mut in_string = false;
    let mut current_name = String::new();
    let mut collecting_bracket = false;

    while let Some(c) = chars.next() {
        if c == '"' {
            if in_string {
                // Check if it's an escaped double quote ("")
                if let Some(&next_c) = chars.peek() {
                    if next_c == '"' {
                        chars.next(); // Consume the escaped quote
                        continue;
                    }
                }
                in_string = false;
            } else {
                in_string = true;
            }
            continue;
        }

        if in_string {
            continue;
        }

        if c == '[' {
            collecting_bracket = true;
            current_name.clear();
            current_name.push('[');
        } else if c == ']' {
            if collecting_bracket {
                current_name.push(']');
                // Only add if it's not empty brackets (unlikely but safe)
                if current_name.len() > 2 {
                    // Ignore ODS internal references which look like [.A1]
                    // Ignore ODS error references like [#REF!]
                    // Ignore ODS internal sheet references like [$Sheet1.A1]
                    if !current_name.starts_with("[.")
                        && !current_name.contains("#REF!")
                        && !current_name.starts_with("[$")
                    {
                        names.push(current_name.clone());
                    }
                }
                collecting_bracket = false;
            }
        } else if collecting_bracket {
            current_name.push(c);
        }
    }

    names
}
/// Tries to resolve the workbook path relative to the current workbook
fn check_workbook_exists(current_workbook_path: &std::path::Path, workbook_ref: &str) -> bool {
    // Extract filename from reference (remove brackets)
    let filename = workbook_ref.trim_start_matches('[').trim_end_matches(']');

    // Try to resolve relative to current workbook directory
    if let Some(parent) = current_workbook_path.parent() {
        let workbook_path = parent.join(filename);
        if workbook_path.exists() {
            return true;
        }
    }

    // Try as absolute path
    std::path::Path::new(filename).exists()
}

/// Check if a URL is accessible (returns true if accessible, false otherwise)
/// Uses HTTP HEAD request with proper status code checking
#[cfg(feature = "link-validation")]
fn check_url_status(url: &str, timeout_secs: u64) -> bool {
    use reqwest::blocking::Client;
    use std::time::Duration;

    let client = match Client::builder()
        .timeout(Duration::from_secs(timeout_secs))
        .build()
    {
        Ok(c) => c,
        Err(_) => return false,
    };

    // Use HEAD request to avoid downloading content
    match client.head(url).send() {
        Ok(response) => response.status().is_success(),
        Err(_) => false,
    }
}

#[cfg(not(feature = "link-validation"))]
fn check_url_status(_url: &str, _timeout_secs: u64) -> bool {
    // Fallback: assume valid if feature not enabled
    true
}

/// Check if text is an external URL
fn is_external_url(text: &str) -> bool {
    text.starts_with("http://")
        || text.starts_with("https://")
        || text.starts_with("ftp://")
        || text.starts_with("file://")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_external_links_in_formula() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=[Book1.xlsx]Sheet1!A1".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = ExternalLinksRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "SEC001");
        assert!(violations[0].message.contains("[Book1.xlsx]"));
        assert!(violations[0].message.contains("range"));
    }

    #[test]
    fn test_external_url_in_text() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text("https://example.com".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        // Use ALL mode to detect URLs (default is WORKBOOK only)
        let rule = ExternalLinksRule {
            mode: LinkDetectionMode::All,
            status: LinkStatus::All,
            timeout_secs: 5,
        };
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert!(violations[0].message.contains("https://example.com"));
        assert!(violations[0].message.contains("range"));
    }
    #[test]
    fn test_ignore_brackets_in_strings() {
        let mut cells = HashMap::new();
        // Formula with brackets in string: ="Amount: " & TEXT(A1, "[$$-en-US]0.00")
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula(
                    "=\"Amount: \" & TEXT(A1, \"[$$-en-US]0.00\")".to_string(),
                ),
            },
        );
        // Valid external link
        cells.insert(
            (0, 1),
            Cell {
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::formula("=[Book1.xlsx]Sheet1!A1".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = ExternalLinksRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert!(violations[0].message.contains("[Book1.xlsx]"));
    }
    #[test]
    fn test_ignore_ods_internal_references() {
        let mut cells = HashMap::new();
        // ODS style internal reference: of:=[.A1] + [.B2]
        // This should be ignored.
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::formula("=of:=[.A1]+[.B2]".to_string()),
            },
        );
        // Mixed: valid external link + ODS internal + REF error + Internal Sheet Ref
        cells.insert(
            (0, 1),
            Cell {
                num_fmt: None,
                row: 0,
                col: 1,
                value: CellValue::formula(
                    "=[Book1.xlsx]Sheet1!A1 + [.C3] + [#REF!] + [$Sheet1.A1]".to_string(),
                ),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
            hidden_columns: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            sheet_path: None,
            formula_parsing_error: None,
            conditional_formatting_count: 0,
            conditional_formatting_ranges: Vec::new(),
        };

        let workbook = Workbook {
            path: PathBuf::from("test.ods"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: Vec::new(),
        };

        let rule = ExternalLinksRule::default();
        let violations = rule.check(&workbook).unwrap();

        // Should only find Book1.xlsx, NOT .A1 or .B2 or .C3
        // Currently, without fix, it will likely find 4 violations or 1 violation with multiple ranges if grouped incorrectly?
        // Actually, it groups by workbook name.
        // So we expect violations for: Book1.xlsx, .A1, .B2, .C3

        // We want ONLY Book1.xlsx

        let messages: Vec<String> = violations.iter().map(|v| v.message.clone()).collect();
        println!("Violations found: {:?}", messages);

        // To verify reproduction, we ASSERT that we DO NOT have false positives.
        // This test will FAIL before the fix.

        let has_book1 = messages.iter().any(|m| m.contains("[Book1.xlsx]"));
        let has_a1 = messages.iter().any(|m| m.contains("[.A1]"));
        let has_ref = messages.iter().any(|m| m.contains("[#REF!]"));
        let has_sheet_ref = messages.iter().any(|m| m.contains("[$Sheet1.A1]"));

        assert!(has_book1, "Should detect actual external link");
        assert!(!has_a1, "Should NOT detect ODS internal reference [.A1]");
        assert!(!has_ref, "Should NOT detect REF error [#REF!]");
        assert!(
            !has_sheet_ref,
            "Should NOT detect ODS internal sheet reference [$Sheet1.A1]"
        );
    }
    #[test]
    fn test_external_link_in_metadata() {
        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells: HashMap::new(),
            used_range: Some((0, 0)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            defined_names: HashMap::new(),
            hidden_sheets: Vec::new(),
            has_macros: false,
            external_links: vec!["external_workbook.xlsx".to_string()],
        };

        let config = LinterConfig::default();
        let rule = ExternalLinksRule::new(&config);
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].scope, ViolationScope::Book);
        assert!(violations[0].message.contains("external_workbook.xlsx"));
        assert!(violations[0].message.contains("metadata"));
    }
}
