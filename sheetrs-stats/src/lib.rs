//! sheetrs-stats: workbook and per-sheet statistics.
//!
//! The library behind `sheetrs stats`, also used by the language bindings. The
//! statistics of a workbook are what [`BookStats::stats`] gives for a [`Book`](sheetrs_io::Book):
//!
//! ```no_run
//! use sheetrs_io::Book;
//! use sheetrs_stats::BookStats;
//!
//! let book = Book::from_file("tests/lint_book.xlsx").unwrap();
//! let stats = book.stats().unwrap();
//! println!("{} formulas", stats.total_formulas);
//! ```

mod stats;

#[cfg(feature = "cli")]
pub mod cli;

pub use stats::{BookStats, SheetStats, WorkbookStats, percentage};
