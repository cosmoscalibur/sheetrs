"""The linter (`sheetlint`): `lint` runs the rules, `rules` lists them."""

import os
from typing import Any, Literal, TypedDict

Source = str | os.PathLike[str] | bytes

class Violation:
    """One finding of the linter."""

    rule_id: str
    rule_name: str
    severity: Literal["error", "warning", "info"]
    message: str
    scope: Literal["book", "sheet", "cell"]
    sheet: str | None
    cell: str | None
    row: int | None
    col: int | None
    def to_dict(self) -> dict[str, Any]: ...

class Rule(TypedDict):
    id: str
    name: str
    category: str
    default: bool

def lint(
    source: Source,
    *,
    config: str | None = None,
    config_path: str | os.PathLike[str] | None = None,
    extension: str | None = None,
) -> list[Violation]:
    """Lint a workbook given by path or as bytes (then `extension` is required)."""

def rules() -> list[Rule]:
    """The linter's rules, ordered by id."""
