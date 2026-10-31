//! sheetrs-stats: workbook and per-sheet statistics.
//!
//! The library behind the `sheetstats` executable, also used by the language bindings
//! (`sheetrs-py`).
//!
//! ```no_run
//! use sheetrs_stats::WorkbookStats;
//!
//! let stats = WorkbookStats::from_path("tests/minimal_test.xlsx").unwrap();
//! println!("{} formulas", stats.total_formulas);
//! ```

mod stats;

pub use stats::{SheetStats, StatsSink, WorkbookStats, percentage};
