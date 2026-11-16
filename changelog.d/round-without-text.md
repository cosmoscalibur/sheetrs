### Changed

- `ROUND`, `ROUNDUP`, `ROUNDDOWN` and the roundings of number formats (`TEXT`, `DOLLAR`, `BAHTTEXT`) compute the
  15-significant-digit rounding in integer arithmetic instead of writing the number as text: 20 times faster per call
  (295 → 15 ns), a request of a workbook with a `ROUND` per row takes 9-20 % less, and the results are the same bits
  (compared with the text definition over 10^8 values; values out of the range of the integer path still use it).
