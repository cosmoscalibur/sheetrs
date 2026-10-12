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

### 1. Excel Errors (ERR1xx)

| ID | Description | Status |
|----|-------------|--------|
| **ERR101** | Broken defined name | ✅ |
| **ERR102** | Excel error | ✅ |
| **ERR103** | Reference to error | ❌ |

### 2. Unreliable Calculations (CALC2xx)

| ID | Description | Status | Params |
|----|-------------|--------|--------|
| **CALC201** | Hardcoded numbers in formulas | ✅ | `ignore_hardcoded_num_values`, `ignore_hardcoded_int_values`, `ignore_hardcoded_power_of_ten` |
| **CALC202** | Circular reference | ✅ | None |
| **CALC203** | Double operator | ❌ | None |
| **CALC204** | Approximate lookup | ✅ | None |
| **CALC205** | Double count | ❌ | None |

### 3. Reference Issues (REF3xx)

| ID | Description | Status | Params |
|----|-------------|--------|--------|
| **REF301** | Unused defined name | ✅ | None |
| **REF302** | Duplicate sheet name | ✅ | None |
| **REF303** | Empty sheet | ✅ | None |
| **REF304** | Large used range | ✅ | None |
| **REF305** | Blank row or column | ✅ | None |
| **REF306** | Unused sheet | ✅ | None |
| **REF307** | Whole column or row reference | ✅ | None |
| **REF309** | Reference to empty cell | ✅ | None |
| **REF310** | Longer cell reference expected | ✅ | `ref310_min_adjacent_data` |

### 4. Formula Interruptions (INT4xx)

| ID | Description | Status |
|----|-------------|--------|
| **INT401** | Formula interrupted by data | ❌ |
| **INT402** | Interrupted by empty | ❌ |
| **INT403** | Formula interrupted by other formula | ❌ |

### 5. Complexity (CPX5xx)

| ID | Description | Status | Params |
|----|-------------|--------|--------|
| **CPX501** | Excessive sheet counts | ✅ | `max_sheets` |
| **CPX502** | Merged cells | ✅ | None |
| **CPX503** | Excessive cond formatting | ✅ | `max_conditional_formatting` |
| **CPX504** | Deep IF nesting | ✅ | `max_if_nesting` |
| **CPX505** | Many nested functions | ✅ | `max_formula_nesting` |
| **CPX506** | Many operations | ❌ | None |
| **CPX507** | Multiple sheet ref | ❌ | None |
| **CPX508** | Many references | ❌ | None |
| **CPX509** | Long formulas | ✅ | `max_formula_length` |

### 6. Vulnerable Formulas (VUL6xx)

| ID | Description | Status |
|----|-------------|--------|
| **VUL601** | Duplicate formula | ✅ |
| **VUL602** | Volatile function | ✅ |
| **VUL603** | Empty string test (="") | ✅ |
| **VUL604** | Error Prone Functions | 🚧 |
| **VUL605** | Legacy array | ✅ |
| **VUL606** | Deprecated func | ✅ |

### 7. Data Issues (DATA7xx)

| ID | Description | Status | Params |
|----|-------------|--------|--------|
| **DATA701** | Generic sheet name | ✅ | `avoid_sheet_names` |
| **DATA702** | Number stored as text | ✅ | None |
| **DATA703** | Inconsistent date format | ✅ | `date_format` |
| **DATA704** | Long text cells | ✅ | `max_text_length` |
| **DATA705** | Unnecessary space | ❌ | None |
| **DATA706** | Numeric text calculation | ✅ | None |
| **DATA708** | Sensitive data | ❌ | None |

### 8. External References (EXT8xx)

| ID | Description | Status |
|----|-------------|--------|
| **EXT801** | Name ext ref | ❌ |
| **EXT802** | External workbook reference | ✅ |
| **EXT803** | Web URLs | ✅ |
| **EXT804** | Pivot ext ref | ❌ |
| **EXT805** | Chart ext ref | ❌ |

### 9. Hidden Information (HID9xx)

| ID | Description | Status |
|----|-------------|--------|
| **HID901** | Hidden defined name | ❌ |
| **HID902** | Very Hidden Worksheet | ❌ |
| **HID903** | Hidden Worksheet | ✅ |
| **HID904** | Hidden Rows or Columns | ✅ |
| **HID905** | Hidden formula | ❌ |
| **HID906** | Invisible cell value | ❌ |

### 10. Files & Settings (FILE10xx)

| ID | Description | Status |
|----|-------------|--------|
| **FILE1001** | Large file size | ❌ |
| **FILE1002** | Old spreadsheet | ❌ |
| **FILE1003** | Date system 1904 | ❌ |

### 11. VBA Issues (VBA11xx)

| ID | Description | Status |
|----|-------------|--------|
| **VBA1101** | Workbook has macros | ✅ |
