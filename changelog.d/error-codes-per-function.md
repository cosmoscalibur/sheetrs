### Fixed

- The engine gives the error code LibreOffice 26.8 gives, and the oracle compares the code in every row (it used to accept any code the function gives somewhere):
  - an error written in an array (`{1;#N/A}`) makes the formula `Err:539` when it is compiled, wherever it is;
  - a plain value where a range goes is refused with the code of the function (`INDEX(5,1)` is `#REF!`, `XMATCH`, `MATCH`, `VLOOKUP`, `SORT`, `UNIQUE`, `SUMIF`, `COUNTIF`, `OFFSET`, `ROW`… are `Err:504`); sets of different sizes are `Err:502` for `CORREL`, `SLOPE`, `FORECAST`… and `#VALUE!` for `SUMPRODUCT` and `SUMX2MY2`;
  - the number of arguments is a table probed in LibreOffice for every function (`functions/arity.rs`, `scripts/probe_arity.py`): the code of each refused count (`Err:504`, `508`, `509`, `510`, `511`, `518`), and the counts it takes that the engine now runs: `SUM()`, `AVERAGE()`, `ROUND(x)` and the optional flags and parameters of `NORMDIST`, `POISSON`, `GAMMADIST`, `F.DIST`, `LOGINV`, `LOGNORMDIST`, `HYPGEOMDIST` and `BETADIST`; the ones it takes that the engine does not run stop the build;
  - the database functions tell a cell from a range of one cell in the field and the criteria (`T1:T1` is a block with no condition, `T1` is no block) and ignore blank criteria rows;
  - `FOURIER` gives `Err:502` for too many columns or rows and `#VALUE!` for a blank or a text, alone or inside another function;
  - `ODDFPRICE` and `ODDFYIELD` refuse their invalid arguments with `Err:502`, `COUPNUM` and the other coupon functions read a text basis and frequency as LibreOffice does, `FINDB` and `SEARCH` give `Err:502` for a start that does not fit;
  - `F.INV(1, d1, d2)` gives LibreOffice's answer (`5 * 2^268` for 5 and 8 degrees of freedom, `Err:523` when its search does not converge) instead of stopping the build.
- `sheetrs-compiler/scripts/lo_probe.py` asks LibreOffice what a list of formulas gives; `probe_arity.py` and `probe_scalar_args.py` build on it.
