### Fixed

- An XLSX written in one pass that converts formulas to values and deletes rows or columns keeps the references of the formulas it does not convert moving with the deletion when the engine writes the result it calculated for them (a formula `NOW()+A2` that moves to row 1 now reads `NOW()+A1`, and so do shared groups and arrays): the formulas that got a calculated result bypassed the rewrite, kept their old text and showed the result of another cell. An ODS was not affected.
