### Changed

- The generators of the binary test fixtures write the same bytes on every run (`scripts/fixture_determinism.py`: zip entry dates, document properties, date fields and the order of the master pages of an ODS become constants), so that regenerating a fixture no longer leaves modified binary files in the working tree; `scripts/fixture_determinism.py --check` lists the tracked workbooks that are not.
