//! Differential test harness: cases are applied to `LibreOffice`, used as a
//! black-box oracle, and the observed values become the expected output.

pub mod case;
pub mod compare;
pub mod crafted;
pub mod oracle;
