//! Tests of the editor against real workbooks. They use the writer directly (`modify_workbook`) as
//! the reference that the editor's passes are compared with, and so they live inside the crate; the
//! tests of the writers themselves are in `sheetrs-io` (feature `write`).

mod plan_equivalence;
