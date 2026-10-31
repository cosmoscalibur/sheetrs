"""Workbook statistics (`sheetstats`)."""

import os
from typing import TypedDict

class SheetStats(TypedDict):
    name: str
    used_cells: int
    formulas: int
    errors: int
    empty_cells: int
    end_of_sheet: str | None
    comments: int
    charts: int
    tables: int
    xml_size: int

class WorkbookStats(TypedDict):
    total_sheets: int
    total_named_ranges: int
    total_file_size: int
    total_used_cells: int
    total_formulas: int
    total_errors: int
    total_empty_cells: int
    total_comments: int
    total_charts: int
    total_tables: int
    total_xml_size: int
    sheets: list[SheetStats]

def compute(
    source: str | os.PathLike[str] | bytes, *, extension: str | None = None
) -> WorkbookStats:
    """Statistics of a workbook given by path or as bytes (`total_file_size` is 0 for bytes)."""
