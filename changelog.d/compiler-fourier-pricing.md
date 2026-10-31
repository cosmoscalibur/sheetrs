### Added

- `sheetc` implements `FOURIER`, `OPT_BARRIER`, `OPT_TOUCH`, `OPT_PROB_HIT`, `OPT_PROB_INMONEY`, `COLOR`, `CURRENT` and `EUROCONVERT` with oracle rows computed by LibreOffice 26.8 (see `FIDELITY.md`). `FOURIER` by rows of an array written in the formula stops the build: LibreOffice transposes that array in the formula at each calculation.
- An ODS name of the pricing add-in (`COM.SUN.STAR.SHEET.ADDIN.PRICINGFUNCTIONS.GETOPT…`) resolves to the function.

### Changed

- `ERF` and `ERFC` follow the algorithm of the C library (glibc), which LibreOffice calls, instead of a series: the results differ in the last digit.
