### Added

- Add `reader::Selection`, `Window`, `Properties` and `SheetInfo`, and `Source::read_selection_into`
  (also `WorkbookReader::read_selection_into` and `read_selection_from_reader_into`): a read of
  what is asked and nothing else, in XLSX and ODS. A sheet that is not selected is not parsed (an
  XLSX does not open its part of the archive), the reading of a sheet stops after the last row a
  window reaches (an ODS ends with the last table selected), the cells outside the windows are
  skipped, and the properties that are not asked for (`value`, `formula`, `number_format`) are not
  built; a selective read does not say whether a formula is an array formula (`is_array` is
  `false`). On a sheet of 4 million cells, ten rows read in under 10 ms and
  13 MB, in either format. Rows are expected in ascending order, as Excel requires: in an XLSX
  that has them out of order the cells after the stop are not read (a full read is not affected).
- `ods`: the names of a document now include `table:named-expression` (constants and formulas such
  as `K = 42` or `F = Sheet1!$A$1*2`), which only the XLSX reader gave before.
- Add `Workbook::name_conflicts` (`NameConflict`, `NameScope`) and the rule **REF311** (Defined name
  defined more than once, `Warning`): a name that the document defines for the workbook and for a
  sheet, or for several sheets, means a different thing depending on the sheet of the formula, and
  `Workbook::defined_names` (a map from a name to one definition) cannot say which. The map keeps
  one of them and the workbook now says there are several; the rule reports each such name with its
  scopes. It reads `localSheetId` in an XLSX and the position of the `named-expressions` in an ODS.
- `HID901` tells a very hidden sheet from a hidden one: the message of the first says that it does
  not appear in the list of sheets to unhide.

### Changed

- **Breaking:** `Sheet::visible: bool` is `Sheet::state: SheetState` (`Visible`, `Hidden`,
  `VeryHidden`). An XLSX has the three states (`state` of `<sheet>`); an ODS has the first two. The
  default sheet is `Visible` (the default of the boolean was "not visible"). `reader::
  extract_sheet_states_from_xlsx` gives the state of every sheet.
- `CellSink` receivers get the same cells as before from a full read; nothing changes for them.

### Fixed

- Speed up the XLSX reader by 20 % in instructions (3.68 to 2.96 G for 500,000 cells, 60 % of
  them formulas; 3.20 to 2.48 G with shared strings): it allocated about seven times per cell, and
  now it does not. The attributes of a cell are borrowed and not copied, the buffers of a cell are
  kept between cells, and `quick-xml` no longer keeps a list of the attributes seen to look for
  duplicates in the elements of a cell. The cells read are identical.
- Speed up the ODS reader by 36 % in instructions (7.42 to 4.76 G for 500,000 cells): the names
  and the macros of the document were read in two more passes over `content.xml`, and are now read
  in the pass of the cells. The document is read once.
