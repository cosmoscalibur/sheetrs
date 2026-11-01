//! sheetrs-stats: workbook and per-sheet statistics.
//!
//! The library behind the `sheetstats` executable, also used by the language bindings. The
//! statistics of a workbook are one function of its [`Source`](sheetrs_core::reader::Source):
//!
//! ```no_run
//! use sheetrs_core::reader::Source;
//!
//! let source = Source::from_file("tests/minimal_test.xlsx").unwrap();
//! let stats = sheetrs_stats::collect(&source).unwrap();
//! println!("{} formulas", stats.total_formulas);
//! ```

mod stats;

pub use stats::{SheetStats, WorkbookStats, collect, percentage};
