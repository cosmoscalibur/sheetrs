//! The public types can be moved to another thread and shared between them, so a program can lint,
//! query or edit several workbooks at once.

use sheetrs_linter::{LintReport, Linter, LinterConfig, Violation};

fn is_send_and_sync<T: Send + Sync>() {}

#[test]
fn public_types_are_send_and_sync() {
    is_send_and_sync::<LintReport>();
    is_send_and_sync::<Linter>();
    is_send_and_sync::<LinterConfig>();
    is_send_and_sync::<Violation>();
}
