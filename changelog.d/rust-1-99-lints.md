### Changed

- Raise the minimum supported Rust version to 1.99 and pin it in `rust-toolchain.toml`.
- Enforce workspace clippy lints (`manual_let_else`, `uninlined_format_args`, `collection_is_never_read`, `significant_drop_tightening`, `cast_possible_wrap`, `format_push_string`); `just lint` now checks all targets.
- `WalkerRule::name` and `RuleCategory::as_str` return `&'static str`.
- An ODS external reference with a quoted empty URI (`''`) yields an empty workbook path instead of the literal quotes.

### Removed

- Remove the unused public functions `extract_formats_from_xlsx`, `Sheet::cells_in_column`, `Sheet::cells_in_row` and `Sheet::last_data_cell`.
