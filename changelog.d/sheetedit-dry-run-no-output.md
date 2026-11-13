### Fixed

- `sheetrs edit --dry-run` needs no `--output` or `--in-place`: it shows the passes and writes nothing; given one of them it also says where the result would go. A command that changes files still needs an output.
- The web demo offers the linter's JSON (the text of `sheetrs lint --format json`, downloadable) next to the report, the table and SARIF.
- The figures of the WebAssembly target and of the build at scale are no longer copied into the READMEs and `DESIGN.md`: they point to the Comparisons of `scripts/README.md`, where each has the commit it was measured at.
