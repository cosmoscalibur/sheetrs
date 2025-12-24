# Test Assets

This directory contains test files used across all sheetrs project components (sheetrs lib, sheetlint, sheetcli, sheetstats).

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
Consider using `lazy_static` or `once_cell` to cache parsed workbooks across tests to improve performance.
