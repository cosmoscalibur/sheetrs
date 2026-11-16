### Fixed

- A build of the embedded mode (`sheetc build --embed-template`, a template module of WebAssembly) keeps its own rules again: the outputs the
  workbook declares are written, and the formulas no output holds keep the result the template stored (the rule that every embedded build
  answers with the whole workbook calculated is gone). It is `sheetc calculate` that tells its build that the workbook is the answer
  (`Spec::whole_workbook`), and only then. The WebAssembly smoke test checks a module against the contract of a build.
