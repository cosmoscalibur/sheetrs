---
trigger: always_on
---

Specific ODS and XLSX logic are part of `<format>_parser.rs` module.
Common general logic should be located in extra files according context: `parser_utils.rs`, `workbook.rs` or new file (require approve).
oC (Separation of Concerns): Keep rules format-agnostic. The rule should receive a "clean" string, and the parser should handle the "dirt" of the file format.