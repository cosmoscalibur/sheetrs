### Changed

- Clippy runs with `clippy::all` and `clippy::pedantic` on every crate, with `-D warnings` in CI.
  The pedantic lints that stay allowed (numeric casts bounded by the format limits, exact float
  comparison as Excel does it, mathematical names, long flat `match`es, local helper items,
  `assert!(x.is_empty())`) are listed with their reasons in the root `Cargo.toml`. Public
  functions that return a `Result` or can panic document it (`# Errors`, `# Panics`).
