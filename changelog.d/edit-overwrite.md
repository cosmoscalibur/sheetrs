### Added

- `sheetrs-edit`: `Editor::write_in_place` and `sheetrs edit --in-place` replace the file being edited with
  the result. The result is written to a temporary file in the same folder and renamed over the
  original when it is complete: a failed run leaves the original as it was and no file behind, and
  the permissions of the original are kept. `write(path)` still requires a path. A workbook given as bytes
  has no file to replace (`Error::NoFile`).
