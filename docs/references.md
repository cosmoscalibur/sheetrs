# Reference projects

Projects that inspire SheetRS. They are not compared with it: they are looked at for new features and for
the shape of an API (the behavior reference is LibreOffice, with Excel for what LibreOffice does not
say; the names come from the LibreOffice interface, UNO, and then from the tools below).

| Project | What it is | What to look at |
| ------- | ---------- | --------------- |
| [calamine](https://github.com/tafia/calamine) (Rust) | Reader of XLSX, ODS, XLS and XLSB | Reading ranges and ways to iterate; which formats and properties it covers |
| [IronCalc](https://github.com/ironcalc/IronCalc) (Rust) | Spreadsheet engine with Python and WebAssembly bindings | Formula evaluation, import and export, the API of its bindings |
| [gridcraft](https://github.com/storytold/gridcraft) (Rust) | Open-source spreadsheet application (Apache 2.0, alpha) with crates for cells and values, number formats, a formula parser and a function library, XLSX, CSV and TSV, a command-line interface and an MCP server | The split of the function library and the number formatting; its CLI and the MCP server as a way to expose a spreadsheet to an agent |
| [fastexcel](https://github.com/ToucanToco/fastexcel) (Python) | Excel reader on top of calamine | Selective reading and tabular output |
| [xlwings](https://www.xlwings.org/) (Python) | Excel automation from Python | Names of objects (`Book`, `range`) and operations |
| [openpyxl](https://openpyxl.readthedocs.io/) (Python) | Reading and writing XLSX | Edit operations and their names |
| [SheetJS](https://sheetjs.com/) (JavaScript) | Reading and writing spreadsheets in JavaScript | `read`/`write`, options objects, typed cells, dates as `Date` |
| [Handsontable](https://handsontable.com/) (JavaScript) | Data grid with a formula engine (HyperFormula) | How formulas and recalculation are exposed to a web interface |
| LibreOffice (UNO and the Calc interface) and Excel (VBA) | The reference | Names of actions, behavior and dialogs |
