"""The linter (``sheetlint``): ``Linter(...).lint(source)``."""

from ._sheetrs import Linter, Violation

__all__ = ["Linter", "Violation"]
