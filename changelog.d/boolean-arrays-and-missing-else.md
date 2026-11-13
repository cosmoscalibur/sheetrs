### Fixed

- A cell whose number is not 0 or 1 and has a boolean format (`IF({0;1},TRUE(),5)`) is that number in the engine and in the ODS reader (`boolean-value="5"` was read as TRUE), as in the XLSX of the same cell; the engine matches LibreOffice 26.8 on opening each file (`-TRUE()` is -1).
- The element that a missing `else` gives in an array (`INDEX(IF({0;1},1),1)`) is an empty text that `COUNTA` does not count (`LEN` 0, `TYPE` 2, `+0` is `#VALUE!`), shown as `FALSE` by the cell that holds the array; `AND` and `OR` of an array with a text in it are `Err:502`.
