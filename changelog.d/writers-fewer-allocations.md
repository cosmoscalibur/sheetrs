### Changed

- The XLSX and ODS writers allocate less per cell: the content of a cell is read into buffers kept between cells, the
  attributes of an ODS cell are read in one pass, and a cell that does not change is copied as it is. Rewriting a
  workbook is 15-25 % faster (14 MB peak, unchanged) and the output is byte for byte the same.
