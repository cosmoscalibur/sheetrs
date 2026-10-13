# Comparative Analysis: sheetrs vs PerfectXL Risk Finder and ExceLint

This report compares `sheetrs` linter rules with the
[PerfectXL Risk Finder inspections](https://www.perfectxl.com/learning-center/risk-finder-library/)
and with [ExceLint](https://doi.org/10.1145/3276518) structural analysis.

Every rule below is implemented as a walker rule and active by default, so the
report no longer tracks support or migration status. Configuration parameters
and their defaults are in the [rule reference](README.md#rule-reference) and
[`sheetlint.toml`](../sheetlint.toml).

> **TODO:** merge this report into the rule reference in
> [`README.md`](README.md) once parity with the reference tools is settled and
> every `?` below is resolved.

## Legend

**Ordering:** by category, then by rule ID. IDs follow the scope hierarchy
**Book** (lowest ID) > **Sheet** > **Cell**.

**Scope:** `<granularity>::<inspected data>`. `(inter)` marks analysis across
cells or sheets rather than of a single cell.

**Dependencies:**

- `X` — rule `X` emits the rule's violations or collects the data it reads.
  `X` runs whenever the rule is enabled; with `X` disabled, only the
  violations of `X` are dropped.
- `!X` — the rule is not reported where rule `X` fires, whether or not `X` is
  enabled.

**Parity** (with the PerfectXL inspection of the same name unless stated):

- **=** same behavior.
- **≈** same concept; the differences are listed.
- **?** not compared: no public definition of the PerfectXL inspection was
  available.

PerfectXL grades findings as observation, weakness or risk, sometimes by how
far a threshold is exceeded. sheetrs assigns each rule a fixed severity.

---

## 1. Excel Errors (ERR1xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **ERR101** (Broken Defined Name) | `book::defined_names` | `RefersTo: =#REF!` | - | ? | Defined names whose reference contains `#REF!`. |
| **ERR102** (Excel Error) | `cell::value` | `#DIV/0!`, `#REF!` | `!CALC202`, `!ERR103` | = | Cells holding an error value, or a formula with an error literal. An error cell referenced by a formula is reported as ERR103 instead. |
| **ERR103** (Reference to Error) | `cell::formula (inter)` | `=A1` (A1 is `#N/A`) | `ERR102`, `CALC202`, `!CALC202` | = | Formulas that reference an error cell. Emitted by ERR102 from the CALC202 dependency graph. |

## 2. Unreliable Calculations (CALC2xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **CALC201** (Hardcoded Number) | `cell::formula` | `=A1*1.19` | - | ≈ PerfectXL ignores common constants such as 10, 100 and 1000. Here powers of ten and a configurable list of numbers are ignored. | Numeric constants embedded in formulas. |
| **CALC202** (Circular Reference) | `cell::formula (inter)` | `A1: =B1`, `B1: =A1` | - | ? | Dependency cycles across cells and sheets. Builds the dependency graph and referenced-sheet set used by other rules. |
| **CALC203** (Double Operator) | `cell::formula` | `=A1++B1` | - | ≈ PerfectXL flags pairs such as `++`, `+-`, `-+` and `%%`, and accepts a minus after `*`. Here `+` or `-` followed by any operator, and pairs of `*`, `/`, `^`, are flagged; `%%` is not. `--` is flagged unless allowed by config. | Consecutive operators outside string literals. |
| **CALC204** (Approximate Lookup) | `cell::formula` | `=VLOOKUP(A1,B:C,2)` | - | ≈ PerfectXL checks VLOOKUP-style functions whose last argument is TRUE or omitted. Here MATCH without exact match type and LOOKUP are also flagged. | Lookups using approximate match. |
| **CALC205** (Double Count) | `cell::formula` | `=SUM(A1:A10,A5:A15)` | - | ≈ PerfectXL flags any cell counted more than once in a formula, directly or through referenced cells, and lowers the level at three or more counts. Here only overlapping ranges inside one SUM-, COUNT- or AVERAGE-family call are flagged. | Overlapping range arguments in one aggregation call. |

## 3. Reference Issues (REF3xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **REF301** (Unused Defined Name) | `book::defined_names (inter)` | Name defined, never used | - | = | Defined names not referenced by any formula. |
| **REF302** (Duplicate Sheet Name) | `book::sheet_names` | `Tax` / `tax ` | - | ? | Sheet names equal after lowercasing and removing non-alphanumeric characters. |
| **REF303** (Empty Sheet) | `sheet::content` | Sheet with no cells | - | ? | Sheets with no cells. |
| **REF304** (Large Used Range) | `sheet::layout` | Used range ends 3+ rows past the data | - | ? | Used range extending more than 2 rows or columns beyond the last data or formula cell. |
| **REF305** (Blank Row or Column) | `sheet::layout` | Empty rows inside the data area | - | ? | Fully blank rows or columns inside the used range. |
| **REF306** (Unused Sheet) | `sheet::references (inter)` | No formula refers to the sheet | `CALC202` | ? | Sheets that no formula references, including formulas on the sheet itself. Not reported for single-sheet workbooks. |
| **REF307** (Whole Column or Row Reference) | `cell::formula` | `=SUM(A:A)`, `=SUM(1:1)` | - | ? | References to entire columns or rows. |
| **REF309** (Reference to Empty Cell) | `cell::formula (inter)` | `=A1` (A1 is blank) | `CALC202` | ? | Formulas referencing cells with no value or formula. Skips sheets with no content. |
| **REF310** (Longer Cell Reference Expected) | `cell::formula (inter)` | `=SUM(A1:A5)` with data in A6 | `CALC202` | ≈ PerfectXL lowers the level when the reference covers less than half of the data block. Here there is no size ratio: the range must be followed by at least `ref310_min_adjacent_data` cells of the same value type. | Ranges that stop before adjacent data of the same type. |

## 4. Formula Interruptions (INT4xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **INT401** (Formula Interrupted by Data) | `cell::formula (inter)` | `F`, `F`, `F`, `42`, `F` | - | ≈ PerfectXL limits how much of the range the data cells may cover. Here there is no coverage limit. | A value cell inside a run of matching formulas. |
| **INT402** (Interrupted by Empty) | `cell::formula (inter)` | `F`, `F`, `F`, *(blank)*, `F` | - | ≈ PerfectXL ignores gaps at a corner or along a full top or bottom edge. Here a single empty cell between formulas of a run is reported; two or more end the run. | An empty cell inside a run of matching formulas. |
| **INT403** (Formula Interrupted by Other Formula) | `cell::formula (inter)` | `=A1+B1`, `=A2+B2`, `=A3+B3`, `=A4*B4` | - | ≈ PerfectXL checks only along the formula direction and only when the other formulas cover at most 20% of the range. Here both rows and columns are checked, with no coverage limit. | A formula with a different R1C1 pattern inside a run. |

A run is a sequence of formulas with the same R1C1 pattern along a row or
column; a single empty cell does not end it. An interruption is reported only
when at least `int_min_formula_sequence` consecutive matching formulas sit
directly before or after it. See [ExceLint](#12-structural-chapter-excelint)
for how this differs from region-based outlier detection.

## 5. Complexity (CPX5xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **CPX501** (Sheet Counts) | `book::sheets` | 51 sheets | - | ? | Workbooks with more sheets than `max_sheets`. |
| **CPX502** (Merged Cells) | `sheet::layout` | Merged `A1:B2` | - | ? | Merged cell ranges. |
| **CPX503** (Excessive Conditional Formatting) | `sheet::style` | 6 conditional formatting rules | - | ? | Sheets with more conditional formatting rules than `max_conditional_formatting`. |
| **CPX504** (Deep IF Nesting) | `cell::formula` | 6 nested `IF` | - | ? | IF nesting deeper than `max_if_nesting`. Emits the CPX505/506/508/509 cascade. |
| **CPX505** (Many Nested Functions) | `cell::formula` | Parenthesis depth 6 | `CPX504`, `!CPX504` | ≈ PerfectXL counts a chain of at least 5 nested function calls. Here parenthesis depth above 5 is counted. | Parenthesis nesting deeper than `max_formula_nesting`. |
| **CPX506** (Many Operations) | `cell::formula` | 9 operators | `CPX504`, `!CPX504`, `!CPX505` | ≈ PerfectXL counts at least 8 functions and operators. Here only operators are counted, above 8. | More operators than `max_operations`. |
| **CPX507** (Multiple Sheet Ref) | `cell::formula` | `=S1!A1+S2!A1+S3!A1` | - | ? | Formulas referencing at least `max_sheet_refs` distinct sheets. Independent of the cascade. |
| **CPX508** (Many References) | `cell::formula` | 11 cell references | `CPX504`, `!CPX504`, `!CPX505`, `!CPX506` | ≈ PerfectXL triggers at 10 references; here above 10. | More cell references than `max_references`. |
| **CPX509** (Long Formula) | `cell::formula` | 256 bytes | `CPX504`, `!CPX504`, `!CPX505`, `!CPX506` | ≈ PerfectXL triggers at 200 characters and not when Many References fires. Here above 255 UTF-8 bytes, so non-ASCII formulas trigger with fewer characters, and it can fire together with CPX508. | Formulas longer than `max_formula_length` bytes. |

## 6. Vulnerable Formulas (VUL6xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **VUL601** (Duplicate Formula) | `cell::formula (inter)` | `=SUM(B1:B9)` in A1 and D7 | - | ? | Identical formulas in non-adjacent cells of a sheet. |
| **VUL602** (Volatile Function) | `cell::formula` | `=TODAY()`, `=INDIRECT(A1)` | - | ≈ PerfectXL lists NOW, RAND, RANDARRAY, RANDBETWEEN and TODAY, and raises a weakness only for NOW or TODAY under manual calculation. Here OFFSET, INDIRECT, INFO and CELL are also flagged, RANDARRAY is not, and every use is reported. | Functions recalculated on every change. |
| **VUL603** (Empty String Test) | `cell::formula` | `=IF(A1="",…)`, `LEN(A1)=0` | - | ? | Empty-string or zero-length tests; recommends `ISBLANK`. |
| **VUL604** (Error Prone Functions) | `cell::formula` | `=VLOOKUP(A1,B:C,2,FALSE)` | - | ≈ PerfectXL includes CELL. Here VLOOKUP, HLOOKUP and LOOKUP are flagged; CELL is reported by VUL602. | Every use of VLOOKUP, HLOOKUP or LOOKUP; recommends XLOOKUP or INDEX/MATCH. |
| **VUL605** (Legacy Array) | `cell::formula` | `{=SUM(A1:A3*B1:B3)}` | - | ? | CSE array formulas. |
| **VUL606** (Deprecated Func) | `cell::formula` | `=CONCATENATE(A1,B1)` | - | ? | Functions with modern replacements, such as CONCATENATE, STDEV, RANK or LOOKUP. |

LOOKUP is reported by CALC204, VUL604 and VUL606.

## 7. Data Issues (DATA7xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **DATA701** (Generic Sheet Name) | `sheet::name` | `Sheet1`, `Copy of Data` | - | ? | Sheet names containing a pattern from `avoid_sheet_names`. |
| **DATA702** (Number Stored as Text) | `cell::value` | `"123"` stored as text | - | ? | Text cells whose content parses as a number. |
| **DATA703** (Inconsistent Date Format) | `cell::format` | `dd/mm/yyyy` when `mm/dd/yyyy` is expected | - | ? | Date-formatted cells whose format differs from `date_format`. |
| **DATA704** (Long Text Cell) | `cell::value` | 300-byte text | - | ? | Text longer than `max_text_length` UTF-8 bytes. |
| **DATA705** (Unnecessary Space) | `cell::value` | `" Total "` | - | ? | Leading or trailing whitespace in text cells. |
| **DATA706** (Numeric Text Calculation) | `cell::formula (inter)` | `=A1+1` (A1 is the text `"10"`) | `CALC202` | ≈ PerfectXL also flags text that looks like a date or time. Here only text that parses as a number. | Formulas referencing numeric-looking text cells. |

## 8. External References (EXT8xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **EXT801** (External Workbook Reference) | `cell::formula` | `=[1]Sheet1!A1` | - | ? | Formulas linking to other spreadsheet files. |
| **EXT802** (Web URLs) | `cell::value` | `https://example.com` | - | ? | URLs in cell values; optionally only unreachable ones (`url_links_status`). |

## 9. Hidden Information (HID9xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **HID901** (Hidden Worksheet) | `sheet::visibility` | `hidden` / `veryHidden` | - | ? | Hidden sheets, including very hidden ones, reported under one ID. |
| **HID902** (Hidden Rows or Columns) | `sheet::layout` | Rows `5:7` hidden | - | ? | Hidden rows or columns, merged into ranges. |

## 10. Files & Settings (FILE10xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **FILE1001** (Large File Size) | `book::file` | 12 MB file | - | ? | Files larger than `max_file_size_mb`. |
| **FILE1002** (Old Spreadsheet) | `book::metadata` | Last modified 2 years ago | - | ? | Modification date older than `max_age_days`. |
| **FILE1003** (1904 Date System) | `book::settings` | `date1904: true` | - | = | Workbooks using the 1904 date system. |

## 11. VBA Issues (VBA11xx)

| SheetRS Rule | Scope | Example | Dependencies | Parity | Concept / Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **VBA1101** (Workbook has Macros) | `book::file` | `vbaProject.bin`, ODS `Basic/` | - | ? | Presence of VBA or ODS Basic/Scripts content. |

---

## 12. Structural Chapter: ExceLint

ExceLint takes a different approach from the rule-based systems of `sheetrs`
and PerfectXL. Its closest `sheetrs` counterparts are INT401–INT403.

### The Rectangles Principle

ExceLint groups cells into **rectangular regions** of homogeneous formulas. It
converts formulas to relative (R1C1) notation to find structural similarities
regardless of their position on the grid. `sheetrs` uses the same R1C1
normalization, but scans one-dimensional runs along rows and columns.

### Statistical Outlier Detection

Rather than checking a list of best practices, ExceLint uses information theory
(the Minimum Description Length principle) to flag cells that are "highly
surprising" relative to their neighbors in a rectangle. `sheetrs` reports every
interruption next to at least `int_min_formula_sequence` consecutive matching
formulas, with no statistical ranking and with separate rule IDs per interruption kind.

---

## 13. Out of Scope & Technical Constraints

Some PerfectXL risk categories and auditing techniques are excluded from this
comparison because of the limits of static analysis of OpenXML and ODS files.

### Legacy Format Limitations

- **Old File Type (.xls)**: `sheetrs` reads **OpenXML (.xlsx, .xlsm)** and
  **ODS**. Binary legacy formats are out of scope.

### VBA Code Auditing

- **Procedure length, recorded macros**: auditing VBA logic is outside the
  scope of formula and data integrity analysis.

---

## Sources

- [PerfectXL Risk Finder Library](https://www.perfectxl.com/learning-center/risk-finder-library/)
- [PerfectXL Release Notes](https://www.perfectxl.com/resources/downloads-support/release-notes/)
- [PerfectXL: Detect hardcoded numbers in formulas](https://www.perfectxl.com/resources/use-cases/detect-hardcoded-numbers-in-formulas/)
- [PerfectXL: XLOOKUP](https://perfectxl.com/excel-tools/perfectxl-risk-eliminator/xlookup)
- Barowy, Berger, Zorn. [ExceLint: Automatically Finding Spreadsheet Formula Errors](https://doi.org/10.1145/3276518). OOPSLA 2018.
