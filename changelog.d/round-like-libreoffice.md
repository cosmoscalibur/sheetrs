### Fixed

- `ROUND`, `ROUNDUP` and `ROUNDDOWN` give what LibreOffice gives (the digits are cut to a whole number with a 15-digit
  tolerance; `ROUNDUP`/`ROUNDDOWN` of a value with a fractional part and fewer than 12 digits first bring it to 12
  significant digits, as in `ROUNDUP(8.06-8; 2)` = `0.06`; then the number is scaled by `10^digits`, the product is brought
  to 15 digits, the floor or the ceiling is taken and it is scaled back; digits are cut to the fractional bits of the
  number and whole numbers are left alone). A negative number that rounds to nothing is `-0`, and a count of digits
  beyond 32 767 is clamped instead of an error. Over 420 000 values with the three functions (-14 to 14 digits), a probe of
  the forms of the call and 6 669 results of extreme values there is no difference with LibreOffice; before there
  were 75 in `ROUND`, 1 843 in `ROUNDDOWN` and 4 882 in `ROUNDUP` in 127 000 cases.
