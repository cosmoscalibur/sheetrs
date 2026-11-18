# XLSX formulas

Formula elements read by `sheetrs-xlsx`: types, cached results, shared groups and array formulas.

## Observations

| Behaviour | Test |
| --- | --- |
| Of the cached results of formula cells only a plain formula's non-zero number is trusted on load; strings, booleans, errors, dates, zero, shared and array formulas, and missing caches are recalculated, and a data table is an empty cell. | `only_a_plain_non_zero_number_cache_is_trusted_over_xlsx` in `sheetrs-xlsx` |
| The `t` of a formula is `normal`, `shared`, `array` or `dataTable`, written exactly; any other value, and a shared group below zero, drops the formula and keeps the cell value as a constant. The group number is an integer prefix and a shared formula without a group is an ordinary one. | `formula_types_and_shared_groups_over_xlsx` in `sheetrs-xlsx` |
| An array formula fills every cell of its `ref` with its text, as cells to recalculate: later cells that hold a formula replace it, later constants do not, and earlier cells in the area are replaced. The `ref` has to start at the cell that holds the formula, in the grammar of the shared `ref`, or the cell is dropped with its cached value; a range whose end is a column alone or an incomplete row fills only its first row, and a reversed range counts as its normalized form. An empty text fills the area too. Areas are cut at the sheet. | `array_formulas_fill_their_area_over_xlsx` in `sheetrs-xlsx` |
| A shared formula group is kept only when one of its anchors (a cell with formula text or a `ref`) has a `ref` that covers the anchor itself; otherwise the anchor and every dependent of the group are empty cells, a dependent whose group has no anchor included. One valid anchor keeps the whole group, whichever order the anchors come in. A member of a dropped group writes nothing: where it shared an address with another cell, in either order, that cell stays. | `formula_types_and_shared_groups_over_xlsx` in `sheetrs-xlsx` |
| A `ref` is a cell, or two corners (cells, columns or rows) in either order, followed by spaces only (a tab, line feed or no-break space, written as a character reference, rejects it); letters in any case, `$` before a column or a row, a `+` before a row, leading zeros, and rows beyond the sheet are accepted in the end corner. The end corner is lenient: with a column valid in the sheet followed by a digit or a sign and a row that is zero, negative or beyond the sheet, anything after the column is ignored and the range runs to the last row; a row inside the sheet must be followed by nothing. The end corner may be a column alone, with or without `$`; a column alone or a row alone is not a cell reference on its own. | `formula_types_and_shared_groups_over_xlsx` in `sheetrs-xlsx` |

## Open defects

- A plain formula with a non-zero number cached is flagged as trusted, but LibreOffice
  recalculates it when its text does not compile (`<f>  </f>` is `Err:520`), when it is volatile
  (`NOW()`) and when it is circular or evaluates to an error.
- An array formula over the whole sheet (`A1:XFD1048576`) makes LibreOffice fail to load the
  file, and the reader refuses it; see [storage limits](../../model/storage.md).
