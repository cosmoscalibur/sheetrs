### Added

- `sheetc`: a workbook with no settings sheet (or one that must not change) declares its inputs,
  outputs and metadata in a JSON file, named by the new key `settings_file` of `sheetc.toml` (a
  path relative to the file) or by `--settings-file`: `{"inputs": {"rate": "Data!A1"}, "outputs":
  {"result": "Data!B1:B3"}, "meta": [{"key": "version", "value": "1"}]}`. The checks are the
  sheet's (ranges that resolve, keys that are not repeated nor a sheet or a name). Existing
  workbooks compile and calculate unmodified, with their inputs changed. It replaces the settings
  sheet; naming both is an error. `sheetrs_compiler::settings::read_json` is the library side.
- Documented convention: building or running with `--outputs any` and no inputs recalculates a
  workbook as it is (nothing is an input, any range can be asked for). `sheetc run` is the
  in-memory mode (compile and run, no executable).
