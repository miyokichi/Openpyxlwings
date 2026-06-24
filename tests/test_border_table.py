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


def make_amount_header_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Amount"

    values = [
        ["header1", "header2", "amount", "amount"],
        ["header_col1", "header2_col1", 100, 200],
        ["header_col2", "header2_col2", 300, 400],
        ["header_col3", "header2_col3", 500, 600],
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


def test_get_bordered_table_by_header_detects_variable_amount_columns(tmp_path: Path) -> None:
    path = tmp_path / "amount.xlsx"
    make_amount_header_workbook(path)

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table_by_header(
            "Amount",
            ["header1", "header2"],
            value_header_contains="amount",
        )

    assert table.range == "B2:E5"
    assert table.header_rows == 1
    assert table.header_columns == 2
    assert table.column_headers == [["amount", "amount"]]
    assert table.row_headers == [
        ["header_col1", "header2_col1"],
        ["header_col2", "header2_col2"],
        ["header_col3", "header2_col3"],
    ]
    assert table.data == [[100, 200], [300, 400], [500, 600]]


def test_get_bordered_table_by_header_is_case_insensitive(tmp_path: Path) -> None:
    path = tmp_path / "amount.xlsx"
    make_amount_header_workbook(path)

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table_by_header(
            "Amount",
            ["HEADER1", "HEADER2"],
            value_header_contains="AMOUNT",
        )

    assert table.range == "B2:E5"


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


def test_repeated_bordered_table_reads_reuse_single_file_open(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = tmp_path / "book.xlsx"
    make_bordered_workbook(path)

    import openpyxlwings.workbook as workbook_module

    styled_opens = 0
    real_load_workbook = workbook_module.load_workbook

    def counting_load_workbook(*args, **kwargs):
        nonlocal styled_opens
        if kwargs.get("read_only") is False:
            styled_opens += 1
        return real_load_workbook(*args, **kwargs)

    monkeypatch.setattr(workbook_module, "load_workbook", counting_load_workbook)

    with ExcelWorkbook(path) as workbook:
        first = workbook.get_bordered_table("Report", 3, 3, header_rows=2, header_columns=1)
        second = workbook.get_bordered_table("Report", 3, 3, header_rows=2, header_columns=1)

    assert first.range == second.range == "B2:D5"
    assert styled_opens == 1
