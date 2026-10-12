### Fixed

- Replace O(n) year-by-year loop in `parse_ods_date()` with O(1) closed-form formula. Use correct Gregorian counting instead of replicating the Excel 1900 leap year bug, since ODS/LibreOffice does not have this bug. Parity test updated to compare calendar dates rather than raw serial numbers.
