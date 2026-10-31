### Changed

- DATA704 measures the text in characters, not bytes. DATA701 normalizes the configured patterns like the sheet name (case and surrounding spaces). DATA702 no longer reports formulas. FILE1002 is named "Old Spreadsheet" (it measures age, not format).
- VBA1101 finds ODS macros in `Basic/` and `Scripts/` libraries listed by the manifest and in macros bound to events (`vnd.sun.star.script:`).
- REF306 treats a sheet as unused when it has no reference to or from another sheet (defined names resolved to the sheet they name).
- REF307, CPX504, CPX505 and VUL602 read the formula without its string literals and quoted sheet names; CPX504 counts only open `IF(` calls (not `SUMIF(` or `COUNTIF(`).
- REF310 reports a range only when the cells after it continue the block it covers: constants after constants, the same formula pattern after formulas (a total or an input next to the range does not count). The default `ref310_min_adjacent_data` stays 1 (PerfectXL parity); the parameter must be an integer of at least 1, otherwise the configuration is refused.
- CALC205 reports only `SUM`, `COUNT`, `COUNTA` and `AVERAGE`: the conditional functions (a criterion inside the range, a sum range equal to the range) and `SUMPRODUCT` are not double counts.
