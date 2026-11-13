### Added

- An executable built with `--embed-template` (and `sheetc calculate FILE --embed-template`) answers
  JSON as well as the workbook, decided when it runs: `--output OUT` writes the workbook (nothing on
  the standard output), no `--output` gives the JSON of any executable (`--range`, `--pretty`,
  `--serve`), and `--output OUT --json` gives both from one calculation. Supporting JSON adds a negligible
  amount of code to the executable (its base, `sheetc`, has it already) and no time to the workbook answer
  (measured with `dist` builds), so it is not a build option. `--meta` and `--info` say which
  answers an executable has (`answers`).
- The settings (sheet or file) of an embedded workbook are honored instead of discarded: the JSON
  answer gives the declared outputs, only the declared inputs are accepted (what does not depend on
  them is calculated once, at build time) and `--meta` lists them. The workbook answer writes the inputs
  given and the outputs that are formulas, and the formulas outside the program keep the result the
  template stored. A recalculated workbook (`--recalculate`) still ignores them.
- `sheetrs-engine`: `Program::handle_with` (the answer to a request and, from the same calculation, the
  values of some rectangles) and `Program::read_rects`. `sheetrs-answer`: `fill_template_with_answer`.
