### Changed

- The writers of XLSX and ODS compress the part they rewrite in another thread while the next bytes are
  produced (a few chunks of memory, none on WebAssembly). A total recalculation of 2.4 M formulas written to a
  workbook takes 16 % less in an XLSX and 11 % less in an ODS. The entries of the archive have the same
  content and order; the compressed bytes can differ by a few from before, and are the same on every run
  and on every target.
