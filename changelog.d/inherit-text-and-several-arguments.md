### Fixed

- The format a formula cell takes from its result: the text format (`@`) is never taken from a cell (`=A1` and `SUM(A1,A2)` over cells formatted as text are `General`, as in LibreOffice), and an aggregate with several arguments takes the format of the first argument that holds a number (`SUM(A4,A6)` over a text and a date is a date), passing over the arguments that hold none. New fixture `inherit_text.{xlsx,ods}`, made by `make_inherit_text.py`, in the oracle.
