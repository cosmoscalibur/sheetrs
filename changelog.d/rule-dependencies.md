### Changed

- A disabled provider still runs, and takes its usual time, whenever a rule
  that depends on it is enabled.

### Removed

- Remove the sheetlint warning that rules depending on a disabled CALC202
  produce no violations.

### Fixed

- Disabling CALC202, CPX504 or ERR102 no longer changes the results of the rules
  that depend on them (REF306, REF309, REF310, DATA706, ERR102, ERR103,
  CPX505/506/508/509). The linter runs the providers of the enabled rules and
  drops the violations of the providers that are not enabled.
