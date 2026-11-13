### Added

- `sheetc`: a workbook with no settings sheet (or one that must not change) declares its inputs,
  outputs and metadata in a JSON file, named by the new key `settings_file` of `sheetc.toml` (a
  path relative to the file) or by `--settings-file`: `{"inputs": ["Data!A1", {"name": "years",
  "range": "Data!A2"}], "outputs": [{"name": "result", "range": "Data!B1:B3"}], "meta": [...]}`.
  `inputs` and `outputs` are arrays of ranges, or of objects that give a range an alias; the checks
  are the sheet's. A row of the settings sheet with no `key` is also only a range. It replaces the
  settings sheet; naming both is an error. `sheetrs_compiler::settings::read_json` is the library
  side.
- `sheetc calculate` is the in-memory mode (compile and run, no executable).
