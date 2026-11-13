//! Tests of the writers against real workbooks. They call `modify_workbook` directly: the fixtures
//! they read are the ones of the editor oracle (`sheetrs-edit/tests/assets`), made with
//! `LibreOffice` by the generators in `sheetrs-edit/scripts`.

mod structure_libreoffice;
mod structure_xlsx;
mod writer_basic;
