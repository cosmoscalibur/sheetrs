### Added

- `--meta` of an executable built by `sheetc` lists the properties the document declares (`meta_type` `document`: `title`, `creator`, `created`, `modified`, `generator`… named alike in an XLSX and an ODS) and, when it was built with `outputs = "named"` or `"any"`, the defined names it can read (`meta_type` `name`). The readers give the workbook its `properties`.

- `sheetc --version` and `--version` of an executable print the versions of the compiler and the runtime, and `--info` of an executable has a `build` block: compiler, platform, build options, the workbook (file name, format and SHA-256) and the UTC time of the build (`SOURCE_DATE_EPOCH` fixes it).

### Changed

- A settings-sheet row with a `key` and no `meta_type` is metadata of the category `extra` instead of a build error.
