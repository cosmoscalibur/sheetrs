### Removed

- Build-time constant folding is no longer configurable in any interface (issue #106): `--no-fold` of `sheetc` and the `fold` key of `sheetc.toml` (and of the `options` of the bindings) are gone, and the key is an error like any unknown one. Folding is always on; `Spec::fold` stays as an option of the Rust library for the tests that compare a folded and an unfolded program and for low-level use.
