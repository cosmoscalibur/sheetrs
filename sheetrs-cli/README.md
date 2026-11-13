# sheetrs-cli

The `sheetrs` executable: one entry point with the subcommands `query`, `edit`, `lint` and `stats` over
XLSX, XLSM and ODS workbooks. The subcommands are the libraries `sheetrs-query`, `sheetrs-edit`,
`sheetrs-linter` and `sheetrs-stats` behind their `cli` features; the flags, their order of
application and the exit codes are described in the README of each. `sheetc` (build, calculate,
inspect, verify, evaluate) is the executable of `sheetrs-compiler`.

```console
$ sheetrs lint -i book.xlsx -f sarif
$ sheetrs query -i book.xlsx --range 'Data!A1:C9' -p value,formula
$ sheetrs edit -i book.xlsx --delete-rows 'Data!3:10' -o out.xlsx
$ sheetrs stats -i book.xlsx --format json
```

The API of every layer is experimental: names and flags may change.
