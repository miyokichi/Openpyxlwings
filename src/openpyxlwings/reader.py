"""Backward-compatible read imports.

Prefer :class:`openpyxlwings.ExcelWorkbook` for new code.
"""

from openpyxlwings.workbook import ExcelWorkbook as ExcelReader
from openpyxlwings.workbook import read_cell_at, read_range, read_range_at, read_sheet, sheet_names

__all__ = [
    "ExcelReader",
    "read_cell_at",
    "read_range",
    "read_range_at",
    "read_sheet",
    "sheet_names",
]
