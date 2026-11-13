### Changed

- `calculate` is a step of the edit plan and nothing else (Rust `Editor::calculate`, Python `Book.calculate`, the `calculate` operation of `edit_workbook`): the engine writes the inputs and calculates every formula as the operations before it leave the workbook; the workbook is saved with fresh results and without the mark to recalculate it when it is opened. It stops being a query that returns values in Python: the outputs are read with `Book.to_dict()` (in memory) or `Book.save(..., json=path)`, and in Rust with `Editor::keep_outputs` and `Report::write_json`.

### Added

- Where the engine cannot calculate the workbook, the plan still writes it, marked to be recalculated when it is opened, and warns that the recalculation will be done by the spreadsheet program that opens it: `Report::warnings` (`Warning::NotCalculated`) in Rust, `sheetrs.RecalculationWarning` in Python, `warnings` in the result of `edit_workbook`, and a line on the standard error in `sheetrs edit`. With the outputs asked for it is an error.
- The `calculate` steps run with the engine of `sheetrs-compiler`, which `sheetrs-edit` depends on: nothing is handed in.
- Python: `RecalculationWarning`.

### Removed

- Python: macOS and Windows wheels leave the CI and release matrices (issue #99); Linux x86_64 and aarch64 stay.
