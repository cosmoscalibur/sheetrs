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
Every rule listed is implemented and active by default. Parameters are global
unless marked † (overridable per sheet under `[sheets."<name>"]`); defaults are
shown in [`sheetlint.toml`](../sheetlint.toml).

### Parity

The Parity column compares each rule with the matching inspection of the
[PerfectXL Risk Finder](https://www.perfectxl.com/learning-center/risk-finder-library/)
and, for formula interruptions, with
[ExceLint](https://doi.org/10.1145/3276518):

- **=** same behavior as the PerfectXL inspection.
- **≈** same concept; the differences are listed.
- **?** not compared: no public definition of the PerfectXL inspection was
  available.

PerfectXL grades findings as observation, weakness or risk, sometimes by how
far a threshold is exceeded. sheetlint assigns each rule a fixed severity.

### 1. Excel Errors (ERR1xx)

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **ERR101** | Broken defined name | ? | None |
| **ERR102** | Excel error | = | None |
| **ERR103** | Reference to error | = | None |

### 2. Unreliable Calculations (CALC2xx)

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **CALC201** | Hardcoded numbers in formulas | ≈ PerfectXL ignores common constants such as 10, 100 and 1000. Here powers of ten and a configurable list of numbers are ignored. | `ignore_hardcoded_num_values`†, `ignore_hardcoded_int_values`†, `ignore_hardcoded_power_of_ten`† |
| **CALC202** | Circular reference | ? | None |
| **CALC203** | Double operator | ≈ PerfectXL flags pairs such as `++`, `+-`, `-+` and `%%`, and accepts a minus after `*`. Here `+` or `-` followed by any operator, and pairs of `*`, `/`, `^`, are flagged; `%%` is not. `--` is flagged unless allowed by config. | `calc203_allow_double_negative` |
| **CALC204** | Approximate lookup | ≈ PerfectXL checks VLOOKUP-style functions whose last argument is TRUE or omitted. Here MATCH without exact match type and LOOKUP are also flagged. | None |
| **CALC205** | Double count | ≈ PerfectXL flags any cell counted more than once in a formula, directly or through referenced cells, and lowers the level at three or more counts. Here only overlapping ranges inside one SUM-, COUNT- or AVERAGE-family call are flagged. | None |

### 3. Reference Issues (REF3xx)

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **REF301** | Unused defined name | = | None |
| **REF302** | Duplicate sheet name | ? | None |
| **REF303** | Empty sheet | ? | None |
| **REF304** | Large used range | ? | None |
| **REF305** | Blank row or column | ? | None |
| **REF306** | Unused sheet | ? | None |
| **REF307** | Whole column or row reference | ? | None |
| **REF309** | Reference to empty cell | ? | None |
| **REF310** | Longer cell reference expected | ≈ PerfectXL lowers the level when the reference covers less than half of the data block. Here there is no size ratio: the range must be followed by at least `ref310_min_adjacent_data` cells of the same value type. | `ref310_min_adjacent_data` |

### 4. Formula Interruptions (INT4xx)

sheetlint scans rows and columns for runs of at least
`int_min_formula_sequence` formulas with the same R1C1 pattern, bridging
one-cell gaps. ExceLint instead groups formulas into rectangular regions and
reports statistical outliers, so it has no per-kind rules or minimum run
length.

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **INT401** | Formula interrupted by data | ≈ PerfectXL limits how much of the range the data cells may cover. Here there is no coverage limit. | `int_min_formula_sequence` |
| **INT402** | Formula interrupted by empty cell | ≈ PerfectXL ignores gaps at a corner or along a full top or bottom edge. Here a single empty cell between formulas of a run is reported; two or more end the run. | `int_min_formula_sequence` |
| **INT403** | Formula interrupted by other formula | ≈ PerfectXL checks only along the formula direction and only when the other formulas cover at most 20% of the range. Here both rows and columns are checked, with no coverage limit. | `int_min_formula_sequence` |

### 5. Complexity (CPX5xx)

CPX504, CPX505, CPX506, CPX508 and CPX509 form a cascade: CPX504, CPX505 and
CPX506 report at most one per formula cell, in that order; CPX508 and CPX509
are checked only when none of them fires.

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **CPX501** | Excessive sheet count | ? | `max_sheets` |
| **CPX502** | Merged cells | ? | None |
| **CPX503** | Excessive conditional formatting | ? | `max_conditional_formatting`† |
| **CPX504** | Deep IF nesting | ? | `max_if_nesting` |
| **CPX505** | Many nested functions | ≈ PerfectXL counts a chain of at least 5 nested function calls. Here parenthesis depth above 5 is counted. | `max_formula_nesting` |
| **CPX506** | Many operations | ≈ PerfectXL counts at least 8 functions and operators. Here only operators are counted, above 8. | `max_operations` |
| **CPX507** | Multiple sheet references | ? | `max_sheet_refs` |
| **CPX508** | Many references | ≈ PerfectXL triggers at 10 references; here above 10. | `max_references` |
| **CPX509** | Long formula | ≈ PerfectXL triggers at 200 characters and not when Many References fires. Here above 255, and it can fire together with CPX508. | `max_formula_length` |

### 6. Vulnerable Formulas (VUL6xx)

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **VUL601** | Duplicate formula | ? | None |
| **VUL602** | Volatile function | ≈ PerfectXL lists NOW, RAND, RANDARRAY, RANDBETWEEN and TODAY, and raises a weakness only for NOW or TODAY under manual calculation. Here OFFSET, INDIRECT, INFO and CELL are also flagged, RANDARRAY is not, and every use is reported. | None |
| **VUL603** | Empty string test (`=""`) | ? | None |
| **VUL604** | Error-prone functions | ≈ PerfectXL includes CELL. Here VLOOKUP, HLOOKUP and LOOKUP are flagged; CELL is reported by VUL602. | None |
| **VUL605** | Legacy array formula | ? | None |
| **VUL606** | Deprecated function | ? | None |

### 7. Data Issues (DATA7xx)

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **DATA701** | Generic sheet name | ? | `avoid_sheet_names`† |
| **DATA702** | Number stored as text | ? | None |
| **DATA703** | Inconsistent date format | ? | `date_format`† |
| **DATA704** | Long text cell | ? | `max_text_length`† |
| **DATA705** | Unnecessary space | ? | None |
| **DATA706** | Numeric text calculation | ≈ PerfectXL also flags text that looks like a date or time. Here only text that parses as a number. | None |

### 8. External References (EXT8xx)

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **EXT801** | External workbook reference | ? | None |
| **EXT802** | Web URLs | ? | `url_links_status`, `url_timeout_seconds` |

### 9. Hidden Information (HID9xx)

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **HID901** | Hidden worksheet (including very hidden) | ? | None |
| **HID902** | Hidden rows or columns | ? | None |

### 10. Files & Settings (FILE10xx)

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **FILE1001** | Large file size | ? | `max_file_size_mb` |
| **FILE1002** | Old spreadsheet | ? | `max_age_days` |
| **FILE1003** | 1904 date system | = | None |

### 11. VBA Issues (VBA11xx)

| ID | Description | Parity | Params |
|----|-------------|--------|--------|
| **VBA1101** | Workbook has macros | ? | None |
