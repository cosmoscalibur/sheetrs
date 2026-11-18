# ODS macros

## Observations

| Behaviour | Test |
| --- | --- |
| The libraries of the package are those `Basic/script-lc.xml` lists, and the modules of each are those its `Basic/<library>/script-lb.xml` lists. A library holds no module when one of its modules has no file; a name listed twice, as a library or as a module, leaves the package without any library; names that differ only in case make the package unreadable. A listing that is not well-formed XML lists nothing, and the package loads without it; the other libraries stay. | `basic_libraries_over_ods` in `sheetrs-ods` |
| The libraries that `LibreOffice` saves in a document are the ones it holds when it loads it again. | `basic_libraries_saved_by_the_oracle_over_ods` in `sheetrs-ods` |

## Open defects

- Scripts in other languages (`Scripts/`) and the scripts that events call in `content.xml` are
  not listed; the oracle reports Basic libraries only.
