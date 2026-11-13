//! What the linter says of the integrated workbook `tests/lint_book.{xlsx,ods}`
//! (made by `tests/make_lint_book.py`, which tells where each rule is triggered).
//!
//! The expected violations of each file are stated for that file. The values the files cache are what
//! `LibreOffice` computes (`sheetrs-query/tests/libreoffice_oracle.rs` compares them cell by cell).

use sheetrs_io::Book;
use sheetrs_linter::{FormatContext, Linter, LinterConfig, RuleId, ViolationScope};
use std::collections::BTreeSet;

const XLSX: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/lint_book.xlsx");
const ODS: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/lint_book.ods");

fn column_name(mut col: u32) -> String {
    let mut name = String::new();
    col += 1;
    while col > 0 {
        let rest = (col - 1) % 26;
        name.insert(0, char::from(b'A' + rest as u8));
        col = (col - 1) / 26;
    }
    name
}

/// `RULE | location | message`, sorted. The age of FILE1002 is replaced by a placeholder: it changes
/// with the day the test runs.
fn lint(path: &str, config: LinterConfig) -> Vec<String> {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let ctx = FormatContext {
        workbook: &workbook,
    };
    let mut lines: Vec<String> = Linter::with_config(config)
        .lint_workbook(&workbook)
        .iter()
        .map(|v| {
            let at = match &v.scope {
                ViolationScope::Book => "Book".to_string(),
                ViolationScope::Sheet(i) => {
                    format!("Sheet {}", workbook.sheets[usize::from(*i)].name)
                }
                ViolationScope::Cell(i, cell) => format!(
                    "{}!{}{}",
                    workbook.sheets[usize::from(*i)].name,
                    column_name(cell.col),
                    cell.row + 1
                ),
            };
            let mut message = v.format_message(&ctx);
            if v.rule_id == RuleId::File1002 {
                let days: i64 = message
                    .strip_prefix("File is ")
                    .and_then(|m| m.strip_suffix(" days old"))
                    .unwrap()
                    .parse()
                    .unwrap();
                assert!(days > 365, "the file is dated 2020-01-15, so it is old");
                message = "File is <days> days old".to_string();
            }
            format!("{} | {at} | {message}", v.rule_id)
        })
        .collect();
    lines.sort();
    lines
}

fn expected(common: &[&str], own: &[&str]) -> Vec<String> {
    let mut all: Vec<String> = common.iter().chain(own).map(|s| (*s).to_string()).collect();
    all.sort();
    all
}

fn rules(lines: &[String]) -> BTreeSet<String> {
    lines
        .iter()
        .map(|l| l.split(" | ").next().unwrap().to_string())
        .collect()
}

/// Violations that both files have.
const COMMON: &[&str] = &[
    "CALC201 | Calc!B2 | Hardcoded values found in formula: 1.19",
    "CALC201 | Calc!B3 | Hardcoded values found in formula: 3.14",
    "CALC201 | Complexity!B1 | Hardcoded values found in formula: 6",
    "CALC201 | Vulnerable!E6 | Hardcoded values found in formula: 6",
    "CALC202 | Loops!A1 | Circular reference detected: Loops!A1 -> Loops!A2 -> Loops!A1",
    "CALC203 | Calc!B4 | Double operator \"++\" found in formula. This may indicate a typo.",
    "CALC203 | Calc!B5 | Double operator \"+-\" found in formula. This may indicate a typo.",
    "CALC204 | Calc!B6 | VLOOKUP uses approximate match (default). Add FALSE as 4th argument for exact match.",
    "CALC204 | Calc!B7 | MATCH uses approximate match (default). Set 3rd argument to 0 for exact match.",
    "CALC204 | Vulnerable!D3 | LOOKUP always uses approximate match. Use XLOOKUP or INDEX/MATCH for exact match.",
    "CALC205 | Calc!B8 | Overlapping ranges in SUM(): INPUTS!B2:B6 and INPUTS!B4:B6 overlap, causing double-counting.",
    "CPX502 | Sheet Complexity | Merged cells in range: A30:C30",
    "CPX502 | Sheet Complexity | Merged cells in range: A32:B33",
    "CPX503 | Sheet Complexity | Sheet has 6 conditional formatting rules. Ranges: F1:F10",
    "CPX504 | Complexity!B1 | Deeply nested IF statements (6 levels). Consider lookup tables or IFS.",
    "CPX505 | Complexity!B2 | Formula with deep nesting (6 levels). Consider simplifying.",
    "CPX506 | Complexity!B3 | Formula with many operations (9).",
    "CPX507 | Complexity!B4 | Formula references 3 distinct sheets.",
    "CPX507 | Summary!B2 | Formula references 3 distinct sheets.",
    "CPX507 | Summary!B6 | Formula references 5 distinct sheets.",
    "CPX507 | Summary!B7 | Formula references 4 distinct sheets.",
    "CPX508 | Complexity!B5 | Formula with many references (11).",
    "CPX509 | Complexity!B6 | Long formula (265 characters).",
    "DATA701 | Sheet Sheet 2 | Non-descriptive sheet name 'Sheet 2'",
    "DATA701 | Sheet Sheet2 | Non-descriptive sheet name 'Sheet2'",
    "DATA702 | Sheet Data | Numeric data stored as text in range: B2:B4",
    "DATA703 | Vulnerable!B1 | Date format 'mm/dd/yy hh:mm AM/PM' does not match required format",
    "DATA704 | Sheet Complexity | Long text cells in range: B6",
    "DATA704 | Sheet Data | Long text cells in range: D2",
    "DATA705 | Data!D2 | Cell contains unnecessary whitespace",
    "DATA705 | Data!E2 | Cell contains unnecessary whitespace",
    "DATA705 | Data!E3 | Cell contains unnecessary whitespace",
    "DATA706 | Data!F2 | Formula references numeric text cell: Data!B2. Convert text to number.",
    "DATA706 | Data!F3 | Formula references numeric text cell: Data!B2. Convert text to number.",
    "ERR101 | Book | Named range 'Gone' has broken reference: #REF!",
    "ERR102 | Errors!A2 | Cell contains error value: #N/A",
    "ERR102 | Errors!A4 | Cell contains error value: #DIV/0!",
    "ERR102 | Errors!A5 | Cell contains error value: #NUM!",
    "ERR102 | Summary!B5 | Cell contains error value: #DIV/0!",
    "ERR102 | Vulnerable!E4 | Cell contains error value: #VALUE!",
    "ERR103 | Errors!A3 | Formula references cell with error: Errors!A1",
    "ERR103 | Errors!A4 | Formula references cell with error: Errors!A3",
    "ERR103 | Summary!B5 | Formula references cell with error: Errors!A1",
    "EXT801 | Sheet Links | External workbook reference prices.ods found in range: B1:B3",
    "EXT802 | Sheet Links | External URL 'ftp://files.example.com/archive' found in range: A4",
    "EXT802 | Sheet Links | External URL 'https://example.com/data' found in range: A1:A2",
    "FILE1002 | Book | File is <days> days old",
    "HID901 | Sheet Hidden | Hidden sheet: Hidden",
    "HID902 | Sheet Hide | Hidden columns: B:D",
    "HID902 | Sheet Hide | Hidden columns: F",
    "HID902 | Sheet Hide | Hidden rows: 3:7",
    "INT401 | Seq!D6 | Value cell interrupts a formula sequence in column D (positions 2–19).",
    "INT402 | Seq!D11 | Empty cell interrupts a formula sequence in column D (positions 2–19).",
    "INT403 | Seq!D16 | Different formula interrupts a formula sequence in column D (positions 2–19).",
    "REF301 | Book | Named range 'Gone' is defined but never used",
    "REF301 | Book | Named range 'TaxRate' is defined but never used",
    "REF302 | Book | Confusingly similar sheet names: Sheet2, Sheet 2",
    "REF303 | Sheet Empty | Sheet 'Empty' is completely empty",
    "REF304 | Sheet Blanks | Used range extends beyond data: last used cell H16, last data/formula cell E11 (threshold: 2/2 rows/cols)",
    "REF305 | Sheet Blanks | Blank columns within used range: A-C. Consider removing or filling these columns.",
    "REF305 | Sheet Blanks | Blank columns within used range: F-H. Consider removing or filling these columns.",
    "REF305 | Sheet Blanks | Blank rows within used range: 1-3. Consider removing or filling these rows.",
    "REF305 | Sheet Blanks | Blank rows within used range: 6-10, 12-16. Consider removing or filling these rows.",
    "REF305 | Sheet Complexity | Blank rows within used range: 14-29. Consider removing or filling these rows.",
    "REF306 | Sheet Empty | Sheet 'Empty' is not connected with other sheets: no other sheet refers to it and it refers to none",
    "REF306 | Sheet Unlinked | Sheet 'Unlinked' is not connected with other sheets: no other sheet refers to it and it refers to none",
    "REF307 | Sheet Refs | Whole-column reference (e.g., A:A) found in range: B1. Use bounded ranges for better performance.",
    "REF307 | Sheet Refs | Whole-row reference (e.g., 1:1) found in range: B2. Use bounded ranges for better performance.",
    "REF309 | Refs!B3 | Formula references empty cell: Refs!D50",
    "REF310 | Refs!B4 | Range A2:A5 may be too short: the block continues in Refs!A6:A8 (3 rows). Suggested range: A2:A8.",
    "REF310 | Vulnerable!H1 | Range A1:A3 may be too short: the block continues in Vulnerable!A4:A5 (2 rows). Suggested range: A1:A5.",
    "REF310 | Vulnerable!H2 | Range A1:A3 may be too short: the block continues in Vulnerable!A4:A5 (2 rows). Suggested range: A1:A5.",
    "REF310 | Vulnerable!H3 | Range A1:A3 may be too short: the block continues in Vulnerable!A4:A5 (2 rows). Suggested range: A1:A5.",
    "REF310 | Vulnerable!J1 | Range A1:A3 may be too short: the block continues in Vulnerable!A4:A5 (2 rows). Suggested range: A1:A5.",
    "REF310 | Vulnerable!K1 | Range A1:A3 may be too short: the block continues in Vulnerable!A4:A5 (2 rows). Suggested range: A1:A5.",
    "REF310 | Vulnerable!L1 | Range A1:A3 may be too short: the block continues in Vulnerable!A4:A5 (2 rows). Suggested range: A1:A5.",
    "REF311 | Book | Defined name 'Local' is defined more than once: in the workbook, sheet 'Calc'",
    "VBA1101 | Book | Workbook contains macros or scripts. Review for security concerns.",
    "VUL601 | Sheet Vulnerable | Formula 'A1:A3*2' is duplicated 3 times in ranges: H1:H3. Consider using named ranges or helper cells.",
    "VUL601 | Sheet Vulnerable | Formula 'Inputs!$B$2*Inputs!$C$2' is duplicated 3 times in ranges: G1:G3. Consider using named ranges or helper cells.",
    "VUL601 | Sheet Vulnerable | Formula 'TRANSPOSE(A1:A3)' is duplicated 3 times in ranges: J1:L1. Consider using named ranges or helper cells.",
    "VUL602 | Sheet Vulnerable | Volatile function CELL() found in range: B5. Consider alternatives for better performance.",
    "VUL602 | Sheet Vulnerable | Volatile function INDIRECT() found in range: B4. Consider alternatives for better performance.",
    "VUL602 | Sheet Vulnerable | Volatile function NOW() found in range: B1. Consider alternatives for better performance.",
    "VUL602 | Sheet Vulnerable | Volatile function OFFSET() found in range: B3. Consider alternatives for better performance.",
    "VUL602 | Sheet Vulnerable | Volatile function RAND() found in range: B2. Consider alternatives for better performance.",
    "VUL603 | Sheet Escapes | Empty string test (=\"\" or LEN()=0) found in range: A3. Consider using ISBLANK() for better readability.",
    "VUL603 | Sheet Vulnerable | Empty string test (=\"\" or LEN()=0) found in range: C1:C2. Consider using ISBLANK() for better readability.",
    "VUL604 | Sheet Calc | Error-prone function VLOOKUP() found in range: B6. Use XLOOKUP / INDEX+MATCH instead.",
    "VUL604 | Sheet Vulnerable | Error-prone function HLOOKUP() found in range: D2. Use XLOOKUP / INDEX+MATCH instead.",
    "VUL604 | Sheet Vulnerable | Error-prone function LOOKUP() found in range: D3. Use XLOOKUP / INDEX+MATCH instead.",
    "VUL604 | Sheet Vulnerable | Error-prone function VLOOKUP() found in range: D1. Use XLOOKUP / INDEX+MATCH instead.",
    "VUL605 | Sheet Vulnerable | Legacy array formula (CSE) found in range: H1:H3. Convert to dynamic array.",
    "VUL605 | Sheet Vulnerable | Legacy array formula (CSE) found in range: J1:L1. Convert to dynamic array.",
    "VUL606 | Sheet Vulnerable | Deprecated function CONCATENATE() found in range: E1. Use CONCAT / TEXTJOIN() instead.",
    "VUL606 | Sheet Vulnerable | Deprecated function FORECAST() found in range: E6. Use FORECAST.LINEAR() instead.",
    "VUL606 | Sheet Vulnerable | Deprecated function LOOKUP() found in range: D3. Use XLOOKUP / INDEX+MATCH() instead.",
    "VUL606 | Sheet Vulnerable | Deprecated function MODE() found in range: E4. Use MODE.SNGL() instead.",
    "VUL606 | Sheet Vulnerable | Deprecated function PERCENTILE() found in range: E5. Use PERCENTILE.INC() instead.",
    "VUL606 | Sheet Vulnerable | Deprecated function RANK() found in range: E3. Use RANK.EQ() instead.",
    "VUL606 | Sheet Vulnerable | Deprecated function STDEV() found in range: E2. Use STDEV.S() instead.",
];

/// The XLSX stores a date format with the language prefix (`[$-409]dd/mm/yyyy`), which the rule does
/// not count in the comparison: the required format (`C3`) is accepted, as in the ODS. A sheet cannot be very
/// hidden in an ODS. `CEILING` and `FLOOR` are written to the XLSX as the `.MATH` functions that
/// replace them, so VUL606 has nothing to say of them there.
const XLSX_ONLY: &[&str] = &[
    "DATA703 | Data!C2 | Date format '[$-409]dd/mm/yyyy' does not match required format",
    "DATA703 | Data!C4 | Date format '[$-409]yyyy-mm-dd' does not match required format",
    "DATA703 | Data!C5 | Date format '[$-409]dd-mmm-yy' does not match required format",
    "HID901 | Sheet Vault | Very hidden sheet: Vault (it does not appear in the list of sheets to unhide)",
];

/// The ODS keeps the formats as the user typed them (`dd/mm/yyyy`) and the required one is accepted.
/// `Vault` is an ordinary hidden sheet there.
const ODS_ONLY: &[&str] = &[
    "DATA703 | Data!C2 | Date format 'dd/mm/yyyy' does not match required format",
    "DATA703 | Data!C4 | Date format 'yyyy-mm-dd' does not match required format",
    "DATA703 | Data!C5 | Date format 'dd-mmm-yy' does not match required format",
    "HID901 | Sheet Vault | Hidden sheet: Vault",
    "VUL606 | Sheet Vulnerable | Deprecated function CEILING() found in range: E7. Use CEILING.MATH() instead.",
    "VUL606 | Sheet Vulnerable | Deprecated function FLOOR() found in range: E8. Use FLOOR.MATH() instead.",
];

#[test]
fn linting_the_lint_book_xlsx_gives_exactly_the_violations_it_was_made_to_have() {
    assert_eq!(
        lint(XLSX, LinterConfig::default()),
        expected(COMMON, XLSX_ONLY)
    );
}

#[test]
fn linting_the_lint_book_ods_gives_exactly_the_violations_it_was_made_to_have() {
    assert_eq!(
        lint(ODS, LinterConfig::default()),
        expected(COMMON, ODS_ONLY)
    );
}

/// CPX501 and FILE1001 need a book of more than 50 sheets and a file of more than 10 MB with the
/// defaults: they are triggered with the limits lowered.
fn small_limits() -> LinterConfig {
    LinterConfig::from_toml("[global]\nmax_sheets = 10\nmax_file_size_mb = 0\n").unwrap()
}

#[test]
fn the_limits_that_no_small_file_exceeds_are_reported_when_lowered() {
    for path in [XLSX, ODS] {
        let lines = lint(path, small_limits());
        let cpx501: Vec<&String> = lines.iter().filter(|l| l.starts_with("CPX501")).collect();
        assert_eq!(cpx501.len(), 1, "{path}");
        assert_eq!(
            cpx501[0], "CPX501 | Book | Workbook has 20 sheets",
            "{path}"
        );
        assert_eq!(
            lines.iter().filter(|l| l.starts_with("FILE1001")).count(),
            1,
            "{path}"
        );
    }
}

/// Every rule of the linter is triggered by the books, each format on its own. FILE1003 (the 1904
/// date system) is the exception: a book would change the serial of every date for one setting, so
/// the rule is tested on a read-result model (`file1003_date_system_1904.rs`) and the readers on
/// small workbooks made in memory (`sheetrs-io/tests/date_system_tests.rs`). The books, in the 1900
/// system, must not report it.
#[test]
fn every_rule_is_triggered_by_the_book_of_each_format() {
    let mut all: BTreeSet<String> =
        sheetrs_linter::rules::registry::rule_metadata(&LinterConfig::default())
            .iter()
            .map(|m| m.id.to_string())
            .collect();
    assert!(all.remove("FILE1003"));
    for book in [XLSX, ODS] {
        let mut seen = rules(&lint(book, LinterConfig::default()));
        seen.extend(rules(&lint(book, small_limits())));
        assert!(!seen.contains("FILE1003"), "{book}");
        let missing: Vec<&String> = all.difference(&seen).collect();
        assert!(missing.is_empty(), "{book}: no violation of {missing:?}");
    }
}
