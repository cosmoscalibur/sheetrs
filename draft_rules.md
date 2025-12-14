# Unresolved errors

- **ERR001**: `ExternalLinksRule` - Detects external links in the workbook.
- **ERR002**: `BrokenNamedRangesRule` - Detects broken named ranges.
- **ERR003**: `CircularReferenceRule` - Detects circular references.

# Security and privacy

- **SEC001**: `ExternalLinksRule` - Detects external links in the workbook.
- **SEC002**: `HiddenSheetsRule` - Detects hidden sheets.
- **SEC003**: `HiddenColumnsRowsRule` - Detects hidden columns or rows.
- **SEC004**: `MacrosVbaRule` - Detects if the workbook contains macros or VBA code.
- **SEC005**: Possible corruption file in formula parser.

## Formatting and Usability (UX)

- **UX001**: `InconsistentNumberFormatRule` - Detects if possible number value is formatted as text instead of number.
- **UX002**: `InconsistentDateFormatRule` - Detect if date not formatted as a specific format setup (`date_format="MM/DD/YYYY"`)
- **UX003**: `BlankRowsColumnsRule` - Detects blank rows or columns within the used range.
  - Configuration:
    - `max_blank_row` (default: 2): Maximum allowed contiguous blank rows before reporting.
    - `max_blank_column` (default: 2): Maximum allowed contiguous blank columns before reporting

# Structural and maintainibility

- **SM001**: `ExcessiveSheetCountsRule` - Detects excessive sheet counts (default: 50).
- **SM002**: `DuplicateSheetNamesRule` - Detects duplicate sheet names (case insensitive convention).
- **SM003**: `LongTextCellRule` - Detects long text cell (default: 255).
- **SM004**: `MergedCellsRule` - Detects merged cells.
- **SM005**: `NonDescriptiveSheetNameRule` - Detects non-descriptive sheet names (avoid default names patterns as "sheet" or "copy", configurable and normalized as lower text when compare).

# Performance

- **PERF001**: `UnusedNamedRangesRule` - Detects unused named ranges.
- **PERF002**: `UnusedSheetsRule` - Detects unused sheets.
- **PERF003**: `LargeUsedRangeRule` - Detects large used range detection (default: last used more than 2 rows or columns beyond last cell with formula or value).
- **PERF004**: `ExcessiveConditionalFormattingRule` - Detects excessive conditional formatting (default: 20 or full column/row in sheet).

# Formula Best Practices (FORM)
- **FORM001**: `LongFormulaRule` - Detects long formula (default: 8192 characters).
- **FORM002**: `VolatileFunctionsRule` - Detects volatile functions (default: NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT).
- **FORM003**: `DuplicateFormulasRule` - Detects identical formulas repeated in close proximity.
- **FORM004**: `WholeColumnRowRefsRule` - Detects whole-column or whole-row references (e.g. `A:A`).
- **FORM005**: `EmptyStringTestRule` - Detects comparisons with empty string `""` instead of `ISBLANK()`.
- **FORM006**: `DeepFormulaNestingRule` - Detects deep formula nesting (max_formula_nesting: 5).
- **FORM007**: `DeepIfNestingRule` - Detects deeply nested if statements (max_if_nesting: 5).
- **FORM008**: `HardcodedValuesInFormulasRule` - Detects hardcoded values in formulas.
  - `ignore_hardcoded_num_values=[1.5]`
  - `ignore_hardcoded_int_values=true`
  - `ignore_hardcoded_power_of_ten=true`.
- **FORM009**: `VLookupHLookupUsageRule` - Detects VLOOKUP and HLOOKUP usage -> (XLOOKUP, INDEX/MATCH).
