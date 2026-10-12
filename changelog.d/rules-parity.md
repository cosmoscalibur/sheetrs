### Added

- Implement CALC205 (Double Count) and REF310 (Longer Ref Expected) rules. Refactor VUL604 to aggregation pattern and add LOOKUP detection. Extract shared cell-reference parsing into `helpers` module, removing duplication across calc202, calc205, and ref310.

### Fixed

- Fix VUL604 false positive on XLOOKUP via word-boundary checking. Fix CALC205 false positive on cross-sheet references by tracking sheet qualifiers.
