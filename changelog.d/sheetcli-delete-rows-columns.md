### Added

- `sheetedit --delete-rows` and `--delete-cols` delete rows and columns of an XLSX or an ODS, in the coordinates of the input file, and adjust the formulas of every sheet, the defined names, merged areas, conditional formats, validations, hyperlinks, filters, notes, manual page breaks, the database and print ranges of an ODS and the shapes and chart series of an XLSX as LibreOffice 26.8 does: a reference to a deleted cell is `#REF!`, a range shrinks around the deletion and a 3D reference (`Sheet1:Sheet3!A1`) follows the rules that LibreOffice gives it, and a shape whose cell is deleted stays anchored to the cell with the same index. `--recalc-on-open` makes the application that opens the result compute every formula (see `sheetedit/README.md`).

### Fixed

- `sheetedit --remove-sheets` turns the references to the removed sheets into `#REF!` (they stayed as text of a sheet that no longer exists), and a name that used it keeps a `#REF!` definition. `--remove-ranges` leaves the name unknown in the other names that used it.
