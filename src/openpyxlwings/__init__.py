"""Fast Excel reads with openpyxl and safer writes through xlwings."""

from openpyxlwings.border_table import BorderTable
from openpyxlwings.format import ExcelFormat, ExtractedMatch, Placeholder, TablePattern
from openpyxlwings.plan import WritePlan
from openpyxlwings.selected_columns import SelectedColumnsTable
from openpyxlwings.workbook import (
    ExcelWorkbook,
    read_cell_at,
    read_range,
    read_range_at,
    read_sheet,
    sheet_names,
    write_range,
    write_range_at,
    write_values,
    write_values_at,
)

ExcelReader = ExcelWorkbook
ExcelWriter = ExcelWorkbook

__all__ = [
    "BorderTable",
    "ExcelFormat",
    "ExcelWorkbook",
    "ExcelReader",
    "ExcelWriter",
    "ExtractedMatch",
    "Placeholder",
    "SelectedColumnsTable",
    "TablePattern",
    "WritePlan",
    "read_cell_at",
    "read_range",
    "read_range_at",
    "read_sheet",
    "sheet_names",
    "write_range",
    "write_range_at",
    "write_values",
    "write_values_at",
]

__version__ = "0.1.0"
