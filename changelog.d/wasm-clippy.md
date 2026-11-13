### Fixed

- The writer of a native executable (`bundle::write_executable`) no longer calls `drop` on a file that has no destructor on WebAssembly (a clippy error for that target); the file is written through `File::create` and `write_all`.
