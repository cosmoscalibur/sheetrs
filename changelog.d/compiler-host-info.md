### Changed

- `INFO("system")`, `INFO("osversion")`, `INFO("release")` and `CELL("filename")` are answered from the machine and the executable's directory when the program runs, and can be fixed when it is built or run (`--system`, `--osversion`, `--release`, `--directory`, `--filename`, `[host]` in `sheetc.toml`), as the clock is. `INFO("directory")` is `#N/A`, as in LibreOffice. Before, they stopped the build.
