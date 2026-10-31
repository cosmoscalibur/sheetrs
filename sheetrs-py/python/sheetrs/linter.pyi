"""The linter (`sheetlint`)."""

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
    enabled: bool

class Linter:
    """The linter with a configuration validated at construction.

    `config` is the TOML text of a `sheetlint.toml`, `config_path` the path of one (not both);
    an unknown rule, category or parameter raises `ValueError`.
    """

    def __init__(
        self,
        *,
        config: str | None = None,
        config_path: str | os.PathLike[str] | None = None,
    ) -> None: ...
    def lint(self, source: Source, *, extension: str | None = None) -> list[Violation]:
        """Lint a workbook given by path or as bytes (then `extension` is required)."""
    def rules(self) -> list[Rule]:
        """The rules ordered by id; `enabled` tells whether this configuration runs each one."""

def lint(
    source: Source,
    *,
    config: str | None = None,
    config_path: str | os.PathLike[str] | None = None,
    extension: str | None = None,
) -> list[Violation]:
    """Lint a workbook given by path or as bytes (then `extension` is required)."""

