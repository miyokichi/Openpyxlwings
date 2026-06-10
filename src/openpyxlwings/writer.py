"""Backward-compatible write imports.

Prefer :class:`openpyxlwings.ExcelWorkbook` for new code.
"""

from openpyxlwings.workbook import ExcelWorkbook as ExcelWriter
from openpyxlwings.workbook import write_range, write_range_at, write_values, write_values_at

__all__ = [
    "ExcelWriter",
    "write_range",
    "write_range_at",
    "write_values",
    "write_values_at",
]
