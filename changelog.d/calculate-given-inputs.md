### Fixed

- `sheetc calculate` (and the `calculate` of the bindings) with the workbook as the answer no longer calculates and writes the cone of every
  input the workbook declares when the request gives none: the roots of a calculation are the inputs it is given (`Spec::given_inputs`) and the cells
  that are volatile or have no result. On the 2.4 M formulas of the benchmark `calculate -o` with nothing given went from 15.8 s to 5.9 s and writes
  the template as it is; with the input given its cone is calculated and written, as before. A request read from the standard input is not known when the
  program is built and keeps the cone of every declared input.
- A calculation that reaches no formula writes the workbook as it is instead of failing with "there is no formula or input to write".
