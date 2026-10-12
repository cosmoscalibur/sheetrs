### Fixed

- Collapse nested match conditionals in the XLSX and ODS parsers to satisfy `clippy::collapsible_match`, restoring a clean `cargo clippy -- -D warnings` on current stable. Parser behavior is unchanged.
