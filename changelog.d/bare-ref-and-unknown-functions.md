### Changed

- A bare `#REF!` of a formula reads as LibreOffice reads it: in an XLSX it is the unknown name `#ref!` and the cell is `#NAME?` (whatever error the file caches), in an ODS the `#ref!` LibreOffice writes is the error `#REF!`; one inside a reference (`Sheet!#REF!`, `[#REF!]`) is unchanged. `sheetrs edit` converts such a formula to `#NAME?` in an XLSX. The analysis form of an ODS writes the error in capitals. New `sheetrs_formula::calc_read::has_bare_ref_error`.
- The name of a function LibreOffice does not have is shown in lowercase from an XLSX (`casa()`, `_xlfn.casa()`), as LibreOffice shows it; an ODS keeps what the file holds.
- The oracle of the query tests no longer compares `minimal_test`: its cases are covered by purpose-built fixtures (`errors.xlsx`, `errors.ods`, `lint_book`), and an ODS that embeds Basic macros shows `#MACRO?` or `#NAME?` for an unknown function according to how LibreOffice runs, which is not a stable reference.
