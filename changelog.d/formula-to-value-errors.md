### Fixed

- `formula_to_value` leaves an empty text where the formula gave an error, in XLSX and in ODS, as LibreOffice's *Formula to Value* does (it kept the error). The oracle sequences of `sheetrs-edit/tests/recalculation.rs` hold the real expectation, with a new one (`value_only`) for the errors a file caches.
