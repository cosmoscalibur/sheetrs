//! SEC005: Web URL links in cell values

use super::{LinterRule, RuleCategory};
use crate::config::LinterConfig;
use crate::reader::Workbook;
use crate::violation::{Severity, Violation, ViolationScope};
use anyhow::Result;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkScope {
    Book,  // Only report book-level violations
    Sheet, // Only report sheet-level violations
}

impl LinkScope {
    fn from_str(s: &str) -> Self {
        match s.to_uppercase().as_str() {
            "BOOK" => LinkScope::Book,
            "SHEET" => LinkScope::Sheet,
            _ => LinkScope::Book, // Default: BOOK
        }
    }
}

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

pub struct WebUrlsRule {
    scope: LinkScope,
    status: LinkStatus,
    timeout_secs: u64,
}

impl WebUrlsRule {
    pub fn new(config: &LinterConfig) -> Self {
        let scope = config
            .get_param_str("url_links_scope", None)
            .map(LinkScope::from_str)
            .unwrap_or(LinkScope::Book);

        let status = config
            .get_param_str("url_links_status", None)
            .map(LinkStatus::from_str)
            .unwrap_or(LinkStatus::All);

        let timeout_secs = config
            .get_param_int("url_timeout_seconds", None)
            .unwrap_or(5) as u64;

        Self {
            scope,
            status,
            timeout_secs,
        }
    }
}

impl Default for WebUrlsRule {
    fn default() -> Self {
        Self {
            scope: LinkScope::Book,
            status: LinkStatus::All,
            timeout_secs: 5,
        }
    }
}

impl LinterRule for WebUrlsRule {
    fn id(&self) -> &str {
        "SEC005"
    }

    fn name(&self) -> &str {
        "Web URL links"
    }

    fn category(&self) -> RuleCategory {
        RuleCategory::SecurityAndPrivacy
    }

    fn check(&self, workbook: &Workbook) -> Result<Vec<Violation>> {
        let mut violations = Vec::new();
        let mut seen_urls = std::collections::HashSet::new();

        // Collect URLs from all sheets
        for sheet in &workbook.sheets {
            let mut url_cells: Vec<(u32, u32, String)> = Vec::new();

            for cell in sheet.all_cells() {
                if let crate::reader::workbook::CellValue::Text(text) = &cell.value {
                    let urls = extract_urls(text);
                    for url in urls {
                        seen_urls.insert(url.clone());
                        url_cells.push((cell.row, cell.col, url));
                    }
                }
            }

            // SHEET scope: create range-based violations per sheet
            if matches!(self.scope, LinkScope::Sheet) && !url_cells.is_empty() {
                let grouped = group_cells_by_value(url_cells);
                for (url, cells) in grouped {
                    // Validate URL status if status is INVALID
                    if matches!(self.status, LinkStatus::Invalid)
                        && check_url_status(&url, self.timeout_secs)
                    {
                        continue; // Skip valid URLs
                    }

                    let ranges = find_contiguous_ranges(&cells);
                    for range in ranges {
                        let message = if matches!(self.status, LinkStatus::Invalid) {
                            format!(
                                "Invalid external URL '{}' (not accessible) in range: {}",
                                url,
                                format_single_range(&range)
                            )
                        } else {
                            format!(
                                "External URL '{}' found in range: {}",
                                url,
                                format_single_range(&range)
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
        }

        // BOOK scope: report unique URLs across entire workbook
        if matches!(self.scope, LinkScope::Book) && !seen_urls.is_empty() {
            for url in seen_urls {
                // Validate URL status if status is INVALID
                if matches!(self.status, LinkStatus::Invalid)
                    && check_url_status(&url, self.timeout_secs)
                {
                    continue; // Skip valid URLs
                }

                let message = if matches!(self.status, LinkStatus::Invalid) {
                    format!(
                        "Invalid external URL '{}' (not accessible) found in workbook.",
                        url
                    )
                } else {
                    format!("External URL '{}' found in workbook.", url)
                };
                violations.push(Violation::new(
                    self.id(),
                    ViolationScope::Book,
                    message,
                    Severity::Warning,
                ));
            }
        }

        Ok(violations)
    }
}

/// Extract URLs from text using regex
fn extract_urls(text: &str) -> Vec<String> {
    use regex::Regex;
    use std::sync::OnceLock;

    static URL_PATTERN: OnceLock<Regex> = OnceLock::new();
    let re = URL_PATTERN.get_or_init(|| Regex::new(r"(https?://|ftp://|file://)[^\s]+").unwrap());

    re.find_iter(text).map(|m| m.as_str().to_string()).collect()
}

/// Group cells by their URL value
fn group_cells_by_value(cells: Vec<(u32, u32, String)>) -> Vec<(String, Vec<(u32, u32)>)> {
    use std::collections::HashMap;

    let mut grouped: HashMap<String, Vec<(u32, u32)>> = HashMap::new();
    for (row, col, value) in cells {
        grouped.entry(value).or_default().push((row, col));
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

/// Check if a URL is accessible (returns true if accessible, false otherwise)
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::workbook::{Cell, CellValue, Sheet};
    use std::collections::HashMap;
    use std::path::PathBuf;

    #[test]
    fn test_url_in_text_cell() {
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
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        // Test SHEET scope
        let rule = WebUrlsRule {
            scope: LinkScope::Sheet,
            status: LinkStatus::All,
            timeout_secs: 5,
        };
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].rule_id, "SEC005");
        assert!(violations[0].message.contains("https://example.com"));
        assert!(violations[0].message.contains("range"));
    }

    #[test]
    fn test_url_book_scope() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text("Visit https://example.com for more".to_string()),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        // Test BOOK scope (default)
        let rule = WebUrlsRule::default();
        let violations = rule.check(&workbook).unwrap();

        assert_eq!(violations.len(), 1);
        assert_eq!(violations[0].scope, ViolationScope::Book);
        assert!(violations[0].message.contains("https://example.com"));
    }

    #[test]
    fn test_multiple_urls_in_cell() {
        let mut cells = HashMap::new();
        cells.insert(
            (0, 0),
            Cell {
                num_fmt: None,
                row: 0,
                col: 0,
                value: CellValue::Text(
                    "Check https://example.com and https://test.org".to_string(),
                ),
            },
        );

        let sheet = Sheet {
            name: "Sheet1".to_string(),
            cells,
            used_range: Some((1, 1)),
            ..Default::default()
        };

        let workbook = Workbook {
            path: PathBuf::from("test.xlsx"),
            sheets: vec![sheet],
            ..Default::default()
        };

        let rule = WebUrlsRule {
            scope: LinkScope::Book,
            status: LinkStatus::All,
            timeout_secs: 5,
        };
        let violations = rule.check(&workbook).unwrap();

        // Should detect both URLs
        assert_eq!(violations.len(), 2);
    }
}
