### Changed

- The format of a workbook (XLSX, XLSM or ODS) is read from its content and is no longer an argument in any layer: `Source::from_bytes(data)` and `Source::from_file(path)` detect it (the `mimetype` entry of an ODS, the content type of the workbook part of an OOXML package) and fail with an unsupported-format error naming what the file is when it is none of the three; the extension is never consulted and the content wins. The readers (`read_workbook_from_reader`, `read_analysis_from_reader_into` and the rest) lost their `extension` parameter; the `Format` type no longer parses from a name or a path. `sheetrs.Book(source)` loses `format=`, and the WebAssembly functions lose their `ext` argument (`lint_workbook(bytes, config_toml)`, `edit_workbook(bytes, operations)`, `explain_edit(operations)`, ...). `Source::from_file` now reports a file that cannot be opened when it is made.

### Removed

- WebAssembly: `verify_workbook`. The check against a reference stays in `sheetc verify`.
