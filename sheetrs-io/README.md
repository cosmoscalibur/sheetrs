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

The API is experimental: names and shapes may change.
