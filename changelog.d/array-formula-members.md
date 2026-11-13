### Fixed

- A cell that is a member of an array formula (legacy or dynamic) shows its own element of the result after a calculation or an edit, in an XLSX and in an ODS, instead of the top-left value (issue #127): the readers tell the compiler how far each member sits from the top-left cell (`CellSink::array_member`) and the engine's array wrapper gives that element. A result of one row or one column is repeated over the block and a position past the result is `#N/A`, as in LibreOffice.
