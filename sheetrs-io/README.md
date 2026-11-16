# sheetrs-io

The input and output of the SheetRS suite. `Book` is the one object: a workbook as a file or as bytes,
whose format (XLSX, XLSM or ODS) is detected from the content. It reads the workbook model, streams its
cells, and, with the feature `write`, keeps the operations pending on it and writes the result; the
other crates add their verbs to it (`BookQuery`, `BookEdit`, `BookCompile`, `BookLint`, `BookStats`).

```rust,no_run
use sheetrs_io::Book;

let book = Book::from_file("book.xlsx")?;
println!("{} {:?}", book.format().extension(), book.date_system()?);
# Ok::<(), sheetrs_io::Error>(())
```

The locale is a property of the `Book`: `Book::with_locale` gives it (a `locale::LocaleTag`, which parses
from `"es-CO"`), and every tool that takes the book follows it (`TEXT` and the conversion of a number to
text when a workbook is calculated, the `display` of a query). Without it the book has the locale its
system gives (`Book::default_locale`, `locale::system`: `LC_ALL`, `LC_CTYPE` or `LANG`, as `LibreOffice`
reads them on Linux), for an XLSX and an ODS alike; `en-US` is what there is when the system has none and in
WebAssembly. With the feature `config`, `locale::from_config`
reads the `locale` of `[global]` in a `sheetrs.toml`.

The API is experimental: names and shapes may change.
