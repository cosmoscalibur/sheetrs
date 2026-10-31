//! The tests that compare the engine with LibreOffice's values (the oracle and the probes) and
//! the compiler's behavior on small workbooks. One test crate, so the helpers of `common` are
//! shared by its modules.

mod common;
mod engine;
mod locale;
mod oracle;
mod pipeline;
mod probe;
mod unsupported;
