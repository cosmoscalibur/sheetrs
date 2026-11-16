### Changed

- A reference that went is no longer told from an error written in the formula: there is no "reference that a deletion removed" in the engine: deleted cells and a deleted sheet leave the same text, and `sheetrs_formula::edit::Removal`, `sheetrs_compiler::Removed` and `RemovedAt`, `Spec::removed`, `Spec::removed_at` and `sheetrs_io::write::RemovedRefs` are gone.
