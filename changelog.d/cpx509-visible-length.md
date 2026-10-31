### Fixed

- CPX509 counts a formula's characters without the namespace prefix of its functions (`_xlfn.CONCAT` and `COM.MICROSOFT.CONCAT` count as `CONCAT`), so the same formula has the same length in an XLSX and an ODS.
