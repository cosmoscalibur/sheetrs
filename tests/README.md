# Test Assets

This directory contains test files used across all sheetrs project components (sheetrs-io, sheetrs-linter, sheetrs-edit, sheetrs-stats).

## Files

### lint_book.xlsx, lint_book.ods
- **Made by**: `make_lint_book.py` (this folder), with LibreOffice 26.8 listening on a socket
  (`scripts/libreoffice-server.sh`). Both files are written from the same document in memory after
  `calculateAll()`, so every value they cache is what LibreOffice computes.
- **Purpose**: one integrated workbook with at least one violation of every linter rule a file can
  trigger, plus clean sheets, hidden and very hidden sheets, names (used, unused, broken, defined twice),
  merged cells, conditional formats, a validation, a note, a chart, an external workbook and a matrix
  formula. It is the shared input of the linter, statistics, reader, query and formula tests (and of
  the Python tests). The docstring of `make_lint_book.py` says where each rule is triggered.
- **Beyond LibreOffice's own output**: the generator writes into the saved archives what a document in
  memory cannot say: the sheet `Vault` is `veryHidden` in the XLSX, a `xl/vbaProject.bin` part is added
  to the XLSX (the ODS gets a Basic library through UNO), and the last-modified date is 2020-01-15
  (FILE1002 is a matter of age; the ODS `dc:date` has no time zone and a fraction of a second, as LibreOffice writes it).
- **Goal**: each file is judged by what LibreOffice does with it. What the tests expect of the linter is stated for each file
  (`sheetrs-linter/tests/lint_book_tests.rs`), and the oracle (see below) records what LibreOffice says
  of each cell.
- **Not in it**: CPX501 and FILE1001 (a book of more than 50 sheets, a file of more than 10 MB: the tests
  lower `max_sheets` and `max_file_size_mb`) and FILE1003 (the 1904 date system changes the serial of every date: the rule has unit tests over a read-result model, and `sheetrs-io/tests/date_system_tests.rs` makes small workbooks in memory to test that the readers state it).

## Usage in Tests

### Unit Tests
Use `include_bytes!` macro to embed files and avoid I/O overhead:
```rust
const TEST_XLSX: &[u8] = include_bytes!("../../tests/lint_book.xlsx");
const TEST_ODS: &[u8] = include_bytes!("../../tests/lint_book.ods");
```

### Integration Tests
Use relative paths from workspace root:
```rust
let xlsx_path = "tests/lint_book.xlsx";
let ods_path = "tests/lint_book.ods";
```

### Caching Parsed Workbooks
Consider a `std::sync::LazyLock` to cache parsed workbooks across tests to improve performance.

## What LibreOffice says of each file (the oracle)

`sheetrs-io/tests/assets/oracle/<name>.<ext>.oracle.json` records, for each fixture file and each
non-empty cell, what LibreOffice itself reports when it opens that very file: the value, the text it
shows, the format code, the formula in its notation, the error code, and the sheets of the workbook.
`make_oracle.py` (same folder) writes them for the fixtures it lists (`cells`, `inherit`, `formulas`, `formats`,
`errors`, `lint_book`, `book`, `input`). `inherit.{xlsx,ods}` (made by `make_inherit.py`) probes the format a
formula cell takes from its result; `sheetrs-query/tests/assets/find_formatted.json` (made by
`make_find_oracle.py`) records what `findAll` of LibreOffice finds with and without `SearchFormatted`:

```bash
soffice --headless --invisible --norestore --accept="socket,host=localhost,port=2002;urp;" &
/opt/libreoffice26.8/program/python sheetrs-io/tests/assets/oracle/make_oracle.py
```

`locale_es-CO.{xlsx,ods}` (made by `make_locale.py`) are saved by a LibreOffice with the `es-CO`
regional setting, and their oracle is recorded by one (`LO_PORT=2003 .../python make_oracle.py
locale_es-CO` against a server started with that setting: see the docstring of `make_oracle.py`):
the shown text of numbers, dates, booleans and the text a formula makes of a number, in a locale other
than `en-US`. `sheetrs-compiler/tests/assets/semantics_<tag>.*` (`es-ES`, `es-CO`, `es-MX`, `en-GB`,
`de-DE`; made by `gen_semantics_fixture.py`) are the oracle of the engine under each locale, among them the
rows of issues #84 and #106 (`TEXT(350*49799,"[$$-en-US]#.###.##0")`, a number joined to a text).

A test states, for each file, what is expected of it, and takes the expectation from the oracle
when LibreOffice can give it (`sheetrs-query/tests/libreoffice_oracle.rs` compares every cell; the
readers, the statistics and the editor compare what they count or keep). A difference found and not approved is listed in
the `PENDING` list of the test, which fails when the difference disappears. A decision of SheetRS that
LibreOffice cannot answer (what a linter rule reports) is stated for each file instead.

## Structure fixtures

`sheetrs-edit/tests/assets/structure/book.{xlsx,ods}` is one workbook that LibreOffice 26.8 wrote
(`sheetrs-edit/scripts/gen_structure_fixtures.py`): sheet `T` with numbers and formulas, sheet `F` with
one formula per reference form that reads `T`, defined names, a merged area, notes, manual page
breaks and a shape. The tests of rows
and columns deleted and of sheets and names removed (`sheetrs-edit/tests/structure_*`) start from it,
and `sheetrs-edit/scripts/compare_structure_with_libreoffice.py` checks `sheetrs edit` against LibreOffice
itself on the same file.

`sheetrs-edit/tests/assets/volatile/volatile.{xlsx,ods}` (`sheetrs-edit/scripts/gen_volatile_fixtures.py`, the
standard library only, no application) holds `V!A1 =NOW()`, `V!B1 =TODAY()`, `V!C1 =RAND()` and a few rows
to delete, with no cached result: the tests of `now` and `seed` of `calculate` (Rust, `sheetrs edit`,
Python, WebAssembly) read what the engine writes into it. `sheetrs-compiler/tests/assets/volatile.{xlsx,ods}` are
copies of the same two files (the tests of `sheetc` flags and of the `Book` in memory read them).

`sheetrs-edit/tests/assets/use_cache/*.{xlsx,ods}` (`make_fixtures.py`: openpyxl builds the XLSX, LibreOffice
makes the ODS, `fixture_determinism.normalize` fixes the bytes; run it twice and compare) are the workbooks of the
tests of `use_cache` (`sheetrs-edit/tests/use_cache.rs`, `sheetrs-cli/tests/edit_use_cache.rs`). Their formulas cache
a result the engine would not give, so that a formula that kept its cache is told from one calculated again:
`book`, `indirect`, `repro` and `array` are the cases of the feature; `volatile` (`TODAY()`, `NOW()`, `RAND()`),
`rules` (a whole row, `ROW()`, `SHEET()`, a defined name, a sheet to delete), `quarter` (40 formulas, the quarter
above which the cache is not used) and `runs` (the rectangles of 65 separate and of 70 adjacent replaced cells)
each pin a rule that decides which formulas an edit reaches (a test fails when the rule is off).
