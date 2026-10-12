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

See the [`sheetlint.toml`](../sheetlint.toml) example at the workspace root.
Rule parameters are listed in the Params column of the
[Rule Reference](#rule-reference).

## Rule Reference

**sheetlint** uses a hierarchical rule system (1xx to 11xx) across 11 categories.
Every rule listed is implemented and active by default. Scope, dependencies
between rules and parity with PerfectXL and ExceLint are in the
[comparative report](comparative_report.md). Parameters are global
unless marked † (overridable per sheet under `[sheets."<name>"]`); defaults are
shown in [`sheetlint.toml`](../sheetlint.toml).

### 1. Excel Errors (ERR1xx)

| ID | Description | Params |
|----|-------------|--------|
| **ERR101** | Broken defined name | None |
| **ERR102** | Excel error | None |
| **ERR103** | Reference to error | None |

### 2. Unreliable Calculations (CALC2xx)

| ID | Description | Params |
|----|-------------|--------|
| **CALC201** | Hardcoded numbers in formulas | `ignore_hardcoded_num_values`†, `ignore_hardcoded_int_values`†, `ignore_hardcoded_power_of_ten`† |
| **CALC202** | Circular reference | None |
| **CALC203** | Double operator | `calc203_allow_double_negative` |
| **CALC204** | Approximate lookup | None |
| **CALC205** | Double count | None |

### 3. Reference Issues (REF3xx)

| ID | Description | Params |
|----|-------------|--------|
| **REF301** | Unused defined name | None |
| **REF302** | Duplicate sheet name | None |
| **REF303** | Empty sheet | None |
| **REF304** | Large used range | None |
| **REF305** | Blank row or column | None |
| **REF306** | Unused sheet | None |
| **REF307** | Whole column or row reference | None |
| **REF309** | Reference to empty cell | None |
| **REF310** | Longer cell reference expected | `ref310_min_adjacent_data` |

### 4. Formula Interruptions (INT4xx)

| ID | Description | Params |
|----|-------------|--------|
| **INT401** | Formula interrupted by data | `int_min_formula_sequence` |
| **INT402** | Formula interrupted by empty cell | `int_min_formula_sequence` |
| **INT403** | Formula interrupted by other formula | `int_min_formula_sequence` |

### 5. Complexity (CPX5xx)

CPX504, CPX505, CPX506, CPX508 and CPX509 form a cascade: CPX504, CPX505 and
CPX506 report at most one per formula cell, in that order; CPX508 and CPX509
are checked only when none of them fires.

| ID | Description | Params |
|----|-------------|--------|
| **CPX501** | Excessive sheet count | `max_sheets` |
| **CPX502** | Merged cells | None |
| **CPX503** | Excessive conditional formatting | `max_conditional_formatting`† |
| **CPX504** | Deep IF nesting | `max_if_nesting` |
| **CPX505** | Many nested functions | `max_formula_nesting` |
| **CPX506** | Many operations | `max_operations` |
| **CPX507** | Multiple sheet references | `max_sheet_refs` |
| **CPX508** | Many references | `max_references` |
| **CPX509** | Long formula | `max_formula_length` |

### 6. Vulnerable Formulas (VUL6xx)

| ID | Description | Params |
|----|-------------|--------|
| **VUL601** | Duplicate formula | None |
| **VUL602** | Volatile function | None |
| **VUL603** | Empty string test (`=""`) | None |
| **VUL604** | Error-prone functions | None |
| **VUL605** | Legacy array formula | None |
| **VUL606** | Deprecated function | None |

### 7. Data Issues (DATA7xx)

| ID | Description | Params |
|----|-------------|--------|
| **DATA701** | Generic sheet name | `avoid_sheet_names`† |
| **DATA702** | Number stored as text | None |
| **DATA703** | Inconsistent date format | `date_format`† |
| **DATA704** | Long text cell | `max_text_length`† |
| **DATA705** | Unnecessary space | None |
| **DATA706** | Numeric text calculation | None |

### 8. External References (EXT8xx)

| ID | Description | Params |
|----|-------------|--------|
| **EXT801** | External workbook reference | None |
| **EXT802** | Web URLs | `url_links_status`, `url_timeout_seconds` |

### 9. Hidden Information (HID9xx)

| ID | Description | Params |
|----|-------------|--------|
| **HID901** | Hidden worksheet (including very hidden) | None |
| **HID902** | Hidden rows or columns | None |

### 10. Files & Settings (FILE10xx)

| ID | Description | Params |
|----|-------------|--------|
| **FILE1001** | Large file size | `max_file_size_mb` |
| **FILE1002** | Old spreadsheet | `max_age_days` |
| **FILE1003** | 1904 date system | None |

### 11. VBA Issues (VBA11xx)

| ID | Description | Params |
|----|-------------|--------|
| **VBA1101** | Workbook has macros | None |
