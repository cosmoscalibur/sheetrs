### Changed

- Free the text of the formulas of the model as soon as `sheetc build` and `sheetc calculate` have lowered them: `compile` accepts the model by value (`compile(model, &spec)`, besides `compile(&model, &spec)`, which leaves the model as it is). On the generated workbooks of `docs/performance.md`, `formulas.xlsx` peaks at 1152 MB instead of 1309 MB; time and the program are unchanged.
