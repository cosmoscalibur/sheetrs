# Formula to Value

Operation observed through `sheetrs-testkit`.

## Observations

| Behaviour | Test |
| --- | --- |
| Formula to Value turns a formula that evaluates to an error into an empty text cell. | `converting_an_error_formula_to_a_value_leaves_an_empty_string` in `sheetrs-testkit` |
| Formula to Value converts the cached results, so the cells must be calculated first. | `oracle.py`, operation `to_values` |

## Open defects

None.
