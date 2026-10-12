### Added

Implemented five must-be audit rules:

- **CALC203** (Double Operator): Detects redundant consecutive arithmetic operators in formulas, with configurable double-negative tolerance.
- **INT401** (Interrupted by Data): Detects value cells breaking a consistent formula sequence.
- **INT402** (Interrupted by Empty): Detects empty cells breaking a consistent formula sequence.
- **INT403** (Interrupted by Other): Detects different formulas breaking a consistent formula sequence.
- **VUL606** (Deprecated Function): Identifies deprecated Excel functions and suggests modern replacements.

The three INT rules share a single data collection pass via `Arc`-backed shared state, avoiding 3× redundant work.
