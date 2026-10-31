# Test Assets

This directory contains test files used across all sheetrs project components (sheetrs-core, sheetlint, sheetedit, sheetstats).

## Files

### minimal_test.xlsx
- **Created with**: LibreOffice Calc
- **Purpose**: Minimal test file containing various spreadsheet features for testing parsers and linters
- **Contains**: External workbook references, formulas, various cell types, formatting

### minimal_test.ods
- **Created from**: Exported from minimal_test.xlsx using LibreOffice Calc
- **Purpose**: ODS equivalent of the XLSX file for format parity testing
- **Goal**: Should produce identical linting results as the XLSX version

## Current Limitations

Due to LibreOffice Calc limitations, these test files currently do **not** include:
- VBA macros (not supported in LibreOffice)
- Broken/invalid external references (LibreOffice auto-corrects these)

## Future Enhancements

To create more comprehensive test files covering all violation scenarios:
- Use Microsoft Excel on Windows to create test files with VBA macros
- Include broken external references and other edge cases
- Ensure both ODS and XLSX versions remain equivalent

## Usage in Tests

### Unit Tests
Use `include_bytes!` macro to embed files and avoid I/O overhead:
```rust
const TEST_XLSX: &[u8] = include_bytes!("../../tests/minimal_test.xlsx");
const TEST_ODS: &[u8] = include_bytes!("../../tests/minimal_test.ods");
```

### Integration Tests
Use relative paths from workspace root:
```rust
let xlsx_path = "tests/minimal_test.xlsx";
let ods_path = "tests/minimal_test.ods";
```

### Caching Parsed Workbooks
Consider a `std::sync::LazyLock` to cache parsed workbooks across tests to improve performance.

## Structure fixtures

`sheetrs-core/tests/assets/structure/book.{xlsx,ods}` is one workbook that LibreOffice 26.8 wrote
(`sheetrs-core/scripts/gen_structure_fixtures.py`): sheet `T` with numbers and formulas, sheet `F` with
one formula per reference form that reads `T`, defined names, a merged area, notes, manual page
breaks and a shape. The tests of rows
and columns deleted and of sheets and names removed (`sheetrs-core/tests/structure_*`) start from it,
and `sheetrs-core/scripts/compare_structure_with_libreoffice.py` checks `sheetedit` against LibreOffice
itself on the same file.
