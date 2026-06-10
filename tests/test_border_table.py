from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import Border, Side

from openpyxlwings import BorderTable, ExcelWorkbook
from openpyxlwings.exceptions import BorderTableShapeError


def make_bordered_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"

    values = [
        ["", "2026", "2027"],
        ["Region", "Sales", "Sales"],
        ["East", 10, 20],
        ["West", 30, 40],
    ]
    thin = Side(style="thin")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    for row_offset, row in enumerate(values, start=2):
        for column_offset, value in enumerate(row, start=2):
            cell = sheet.cell(row=row_offset, column=column_offset)
            cell.value = value
            cell.border = border

    workbook.save(path)


def test_get_bordered_table_detects_range_and_headers(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    make_bordered_workbook(path)

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table(
            "Report",
            3,
            3,
            header_rows=2,
            header_columns=1,
        )

    assert table.range == "B2:D5"
    assert table.values == [
        [None, "2026", "2027"],
        ["Region", "Sales", "Sales"],
        ["East", 10, 20],
        ["West", 30, 40],
    ]
    assert table.column_headers == [["2026", "2027"], ["Sales", "Sales"]]
    assert table.row_headers == [["East"], ["West"]]
    assert table.data == [[10, 20], [30, 40]]


def test_border_table_edit_methods_update_values(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    make_bordered_workbook(path)

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table("Report", 4, 4, header_rows=2, header_columns=1)

    table.set_value(1, 2, "FY2026")
    table.set_body_value(2, 2, 99)
    table.add_row([50, 60], row_headers=["North"])
    table.add_column([70, 80, 90], column_headers=["2028", "Sales"])
    table.add_header_row(["Area", "Actual", "Actual", "Plan"])
    table.add_header_column(["Group", "Metric", "Metric", "A", "B", "C"])

    assert table.header_rows == 3
    assert table.header_columns == 2
    assert table.values == [
        [None, "Group", "FY2026", "2027", "2028"],
        ["Region", "Metric", "Sales", "Sales", "Sales"],
        ["Area", "Metric", "Actual", "Actual", "Plan"],
        ["East", "A", 10, 20, 70],
        ["West", "B", 30, 99, 80],
        ["North", "C", 50, 60, 90],
    ]


def test_border_table_rejects_bad_shapes(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    make_bordered_workbook(path)

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table("Report", 4, 4, header_rows=2, header_columns=1)

    with pytest.raises(BorderTableShapeError):
        table.set_body_value(10, 1, "outside")

    with pytest.raises(BorderTableShapeError):
        table.add_row([1])

    with pytest.raises(BorderTableShapeError):
        table.add_column([1])


def test_get_bordered_table_rejects_missing_internal_border(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    make_bordered_workbook(path)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Broken"
    thin = Side(style="thin")
    full = Border(top=thin, bottom=thin, left=thin, right=thin)
    missing_right = Border(top=thin, bottom=thin, left=thin)
    for row in range(1, 3):
        for column in range(1, 3):
            sheet.cell(row=row, column=column).border = full
    sheet.cell(row=2, column=1).border = missing_right
    sheet.cell(row=2, column=2).border = Border(top=thin, bottom=thin, right=thin)
    workbook.save(path)

    with ExcelWorkbook(path) as reader:
        with pytest.raises(BorderTableShapeError):
            reader.get_bordered_table("Broken", 1, 1)


def test_border_table_save_delegates_to_workbook() -> None:
    class FakeWorkbook:
        def __init__(self) -> None:
            self.saved_table = None

        def _save_bordered_table(self, table: BorderTable) -> None:
            self.saved_table = table

    workbook = FakeWorkbook()
    table = BorderTable(
        workbook=workbook,  # type: ignore[arg-type]
        sheet="Report",
        start_row=1,
        start_column=1,
        values=[["Name", "Score"], ["Alice", 10]],
    )
    table.add_row(["Bob", 20])

    table.save()

    assert workbook.saved_table is table
    assert table._insertions == []
