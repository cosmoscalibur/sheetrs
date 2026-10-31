### Changed

- CALC202 reads references from the formula AST (`sheetrs-formula`) instead of a regular expression (a formula that does not parse keeps the old reading) and has two levels: an error when a cell of the loop holds an error in the cache (a circular result as the application stored it), a warning when only the dependency graph closes the loop. A loop whose cells hold no error no longer suppresses ERR102/ERR103 on them.
