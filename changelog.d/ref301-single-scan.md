### Changed

- Speed up REF301: each formula is scanned once for all defined names with an Aho-Corasick automaton, instead of searching every distinct formula for each name at the end of the workbook. Reported violations are unchanged.
