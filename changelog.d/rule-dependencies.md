### Fixed

- Disabling CALC202, CPX504 or ERR102 no longer changes the results of the rules
  that depend on them (REF306, REF309, REF310, DATA706, ERR102, ERR103,
  CPX505/506/508/509). The linter runs the providers of the enabled rules and
  drops the violations of the providers that are not enabled.
