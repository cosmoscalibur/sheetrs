# sheetrs-wasm

WebAssembly build of the SheetRS linter and statistics for the browser (`--target web`
ES module). Reads XLSX and ODS workbooks from bytes.

```js
import init, {
  init_panic_hook,
  lint_workbook_records,
  workbook_stats,
  get_rules_definition,
} from "sheetrs-wasm";

await init();            // loads the .wasm next to the module
init_panic_hook();

const bytes = new Uint8Array(await file.arrayBuffer());
const violations = lint_workbook_records(bytes, "xlsx", ""); // third argument: sheetlint.toml text
const stats = workbook_stats(bytes, "xlsx");
```

| Function | Returns |
| -------- | ------- |
| `lint_workbook_records(bytes, ext, config_toml)` | array of `{rule_id, rule_name, severity, message, scope, sheet, cell, row, col}` |
| `workbook_stats(bytes, ext)` | object with the fields of `sheetstats --format json` (`total_file_size` is 0) |
| `get_rules_definition()` | array of `{id, name, category, is_default}` |
| `lint_workbook(bytes, ext, config_toml)` | the human-readable report, as text |
| `get_workbook_stats(bytes, ext)` | the human-readable statistics, as text |

`ext` is `"xlsx"`, `"xlsm"` or `"ods"`. Errors are thrown as strings.

Build the package with `scripts/package-wasm.sh` (see the repository README).
