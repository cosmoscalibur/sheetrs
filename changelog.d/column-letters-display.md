### Changed

- **Breaking:** Replace `parser_utils::column_index_to_letters` and `CellReference::col_to_letter` with `parser_utils::column_letters`, an allocation-free `impl Display` built with `fmt::from_fn`. `CellReference` display, XLSX shared-formula translation, INT401–403, HID902 and REF305 messages write column letters without intermediate `String`s. Production benchmark: about 1.3% (XLSX) and 0.9% (ODS) fewer user-space instructions, same peak RAM.

### Removed

- **Breaking:** Remove `CellReference::to_excel_ref`; use its `Display` implementation.
