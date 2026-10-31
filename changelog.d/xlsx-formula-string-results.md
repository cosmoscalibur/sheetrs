### Fixed

- XLSX: the cached string result of a formula (`t="str"`) that reads as a number, such as `"007"` or `"1010"`, is now read as text instead of a number.
