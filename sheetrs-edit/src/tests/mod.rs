//! Tests of the writer and of the editor against real workbooks. They use the writer directly
//! (`modify_workbook`) as the reference that the editor's passes are compared with, and so they
//! live inside the crate: the writer is not part of its public API.

mod plan_equivalence;
mod structure_parity;
mod structure_xlsx;
mod writer_basic;
