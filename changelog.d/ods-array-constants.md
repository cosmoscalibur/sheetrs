### Fixed

- ODS: the row separator `|` of an array constant (`{1;2|3;4}`) is read as `;`, as Excel writes it, so formulas with array constants parse.
