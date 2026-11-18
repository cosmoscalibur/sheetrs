# XLSX macros

## Observations

| Behaviour | Test |
| --- | --- |
| The project of the relationship of type `vbaProject` is a compound file whose `VBA/dir` stream lists the modules; each is a Basic module of a library called as the project, `Standard` when the project has no name. The `PROJECT` stream, the stream of each module and the code are not what decides: a module the `dir` stream lists without a stream of its own is listed, and one listed only in `PROJECT` is not. A module listed twice is one module, and names that differ only in case are two. Reference records in the `dir` stream, and a `dir` stream long enough to take sectors of its own and several chunks, do not change the list. | `vba_projects_over_xlsx` in `sheetrs-xlsx` |
| A file that is not a compound file, and a project without modules, list none. | `vba_projects_over_xlsx` in `sheetrs-xlsx` |

## Open defects

- The projects of the tests are built from the public layouts of the compound file and of the
  project, because LibreOffice saves no VBA project. Protected and signed projects, and projects
  with a performance cache, have not been compared.
- The reader does not look at a project of more than 64 MiB, at a `dir` stream of more than
  16 MiB of text, or at a compound file whose header lies about its sectors; it lists no module
  for them. The limits of LibreOffice are not known.
- XLM macro sheets are read as sheets and not listed.
