//! The public types can be moved to another thread and shared between them, so a program can lint,
//! query or edit several workbooks at once.

use sheetrs_protocol::{
    CellRecord, Diagnostic, Find, Query, QueryResponse, Replace, Request, ValueRanges,
};

fn is_send_and_sync<T: Send + Sync>() {}

#[test]
fn public_types_are_send_and_sync() {
    is_send_and_sync::<CellRecord>();
    is_send_and_sync::<Diagnostic>();
    is_send_and_sync::<Find>();
    is_send_and_sync::<ValueRanges>();
    is_send_and_sync::<Query>();
    is_send_and_sync::<QueryResponse>();
    is_send_and_sync::<Replace>();
    is_send_and_sync::<Request>();
}
