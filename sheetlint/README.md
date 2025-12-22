# sheetlint

**sheetlint** is a fast, configurable linter for Excel (XLSX) and ODS spreadsheets. It helps maintain data quality, security, and performance by enforcing a customizable set of rules.

## Installation

```bash
cargo install --path .
```

## Usage

```bash
sheetlint <FILE> [OPTIONS]
```

### Options

- `-c, --config <FILE>`: Path to configuration file (default: `sheetlint.toml`).
- `-f, --format <FORMAT>`: Output format: `text` (default) or `json`.

## Configuration

Create a `sheetlint.toml` file:

```toml
[global]
enabled_rules = ["ALL"]
disabled_rules = ["PERF004"] 
date_format = "dd-mm-yy"

[sheets."RawData"]
disabled_rules = ["UX", "SM"]
```

## Rule Reference

### Error Rules (ERR)

| ID | Description | Default Active | Params |
|----|-------------|----------------|--------|
| **ERR001** | Error cell values (#DIV/0!, #REF!, etc.) | Yes | None |
| **ERR002** | Broken named ranges | Yes | None |
| **ERR003** | Circular references | Yes | `expand_ranges_in_dependencies` (bool, default: false) |

### Security Rules (SEC)

| ID | Description | Default Active | Params |
|----|-------------|----------------|--------|
| **SEC001** | External links in formulas, URLs, or metadata | Yes | `external_links_type` (string: "URL"\|"WORKBOOK"\|"ALL", default: "WORKBOOK"), `external_links_status` (string: "INVALID"\|"ALL", default: "ALL"), `url_timeout_seconds` (int, default: 5) |
| **SEC002** | Hidden sheets | No | None |
| **SEC003** | Hidden columns or rows | No | None |
| **SEC004** | Macros and scripts detection (VBA, ODS Basic/Scripts) | No | None |

### Performance Rules (PERF)

| ID | Description | Default Active | Params |
|----|-------------|----------------|--------|
| **PERF001** | Unused named ranges | Yes | None |
| **PERF002** | Unused sheets (filled with content but unreferenced) | Yes | None |
| **PERF005** | Empty unused sheets (no content, no formulas, unreferenced) | Yes | None |
| **PERF003** | Large used range (empty cells beyond data) | Yes | `max_extra_row` (int, default 2), `max_extra_column` (int, default 2) |
| **PERF004** | Excessive conditional formatting (:warning: Not tested) | No | `max_conditional_formatting` (int, default 5) |

### Usability Rules (UX)

| ID | Description | Default Active | Params |
|----|-------------|----------------|--------|
| **UX001** | Inconsistent number formatting | Yes | None |
| **UX002** | Inconsistent date formatting | No | `date_format` (string, default: "mm/dd/yyyy") |
| **UX003** | Blank rows/columns in used range | No | `max_blank_row` (int, default 2), `max_blank_column` (int, default 2) |

### Maintainability Rules (SM)

| ID | Description | Default Active | Params |
|----|-------------|----------------|--------|
| **SM001** | Excessive sheet counts | Yes | `max_sheets` (int, default 50) |
| **SM002** | Confusingly similar sheet names (normalized: lowercase, alphanumeric only) | Yes | None |
| **SM003** | Long text cells | No | `max_text_length` (int, default 255) |
| **SM004** | Merged cells | No | None |
| **SM005** | Non-descriptive sheet names | Yes | `avoid_sheet_names` (`list<string>`, default: ["sheet", "copy"]) |

### Formula Rules (FORM)

| ID | Description | Default Active | Params |
|----|-------------|----------------|--------|
| **FORM001** | Long formulas | No | `max_formula_length` (int, default 255) |
| **FORM002** | Volatile functions (NOW, RAND, etc.) | Yes | None |
| **FORM003** | Duplicate formulas | Yes | None |
| **FORM004** | Whole column/row references (A:A, 1:1) | Yes | None |
| **FORM005** | Empty string logic tests (=A1="") | Yes | None |
| **FORM006** | Deep formula nesting | No | `max_formula_nesting` (int, default 5) |
| **FORM007** | Deep IF statement nesting | No | `max_if_nesting` (int, default 5) |
| **FORM008** | Hardcoded numeric values in formulas | Yes | `ignore_hardcoded_int_values` (bool, default true), `ignore_hardcoded_power_of_ten` (bool, default true), `ignore_hardcoded_num_values` (`list<string>`, default []) |
| **FORM009** | Usage of VLOOKUP/HLOOKUP (recommend XLOOKUP or INDEX/MATCH) | Yes | None |
