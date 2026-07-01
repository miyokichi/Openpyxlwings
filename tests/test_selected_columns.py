from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import Border, Side

from openpyxlwings import ExcelWorkbook, SelectedColumnsTable, WritePlan
from openpyxlwings.exceptions import BorderTableNotFoundError, BorderTableShapeError
from openpyxlwings.plan import _SelectedColumnsTableOp
from openpyxlwings.selected_columns import _SelectedColumn

THIN = Side(style="thin")
FULL = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def _write_grid(path: Path, title: str, values) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = title
    for row_offset, row in enumerate(values, start=2):
        for column_offset, value in enumerate(row, start=2):
            cell = sheet.cell(row=row_offset, column=column_offset)
            cell.value = value
            cell.border = FULL
    workbook.save(path)


def make_amount_workbook(path: Path) -> None:
    _write_grid(
        path,
        "Amount",
        [
            ["header1", "header2", "amount", "amount"],
            ["col1a", "col1b", 100, 200],
            ["col2a", "col2b", 300, 400],
            ["col3a", "col3b", 500, 600],
        ],
    )


def test_selects_only_requested_header_column(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    make_amount_workbook(path)

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table_by_columns("Amount", ["header1"])

    assert table.column_headers == ["header1"]
    assert table.data == [["col1a"], ["col2a"], ["col3a"]]
    assert table.row_count == 3


def test_value_header_contains_selects_every_amount_column(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    make_amount_workbook(path)

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table_by_columns(
            "Amount",
            ["header1"],
            value_header_contains="amount",
        )

    # header2 is not requested, so it is excluded; both amount columns appear.
    assert table.column_headers == ["header1", "amount", "amount"]
    assert table.data == [
        ["col1a", 100, 200],
        ["col2a", 300, 400],
        ["col3a", 500, 600],
    ]
    # Each selected column remembers its source Excel column (B, C, D).
    assert [column.source_column for column in table.columns] == [2, 4, 5]


def test_continuation_includes_bordered_empty_cell(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _write_grid(
        path,
        "Gap",
        [
            ["key", "amount"],
            ["a", 100],
            ["b", None],
            ["c", 300],
        ],
    )

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table_by_columns(
            "Gap",
            ["key"],
            value_header_contains="amount",
        )

    assert table.data == [["a", 100], ["b", None], ["c", 300]]


def test_ragged_columns_are_padded_with_none() -> None:
    # Columns read with different lengths are squared up into a rectangle.
    table = SelectedColumnsTable(
        workbook=None,  # type: ignore[arg-type]
        sheet="S",
        start_row=2,
        start_column=2,
        end_row=4,
        end_column=3,
        header_row=2,
        columns=[
            _SelectedColumn("key", 2, ["a", "b", "c"]),
            _SelectedColumn("amount", 3, [100]),
        ],
    )

    assert table.row_count == 3
    assert table.data == [["a", 100], ["b", None], ["c", None]]


def test_missing_header_raises(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    make_amount_workbook(path)

    with ExcelWorkbook(path) as workbook:
        with pytest.raises(BorderTableNotFoundError):
            workbook.get_bordered_table_by_columns("Amount", ["nope"])


def test_add_row_and_add_column_update_virtual_table(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    make_amount_workbook(path)

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table_by_columns(
            "Amount",
            ["header1"],
            value_header_contains="amount",
        )

    table.add_row(["col4a", 700, 800])
    assert table.row_count == 4
    assert table.added_rows == 1
    assert table.data[-1] == ["col4a", 700, 800]

    table.add_column([1, 2, 3, 4], header="ratio")
    assert table.column_headers == ["header1", "amount", "amount", "ratio"]
    assert table.columns[-1].source_column is None
    assert table.data[0] == ["col1a", 100, 200, 1]


def test_add_row_rejects_wrong_width(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    make_amount_workbook(path)

    with ExcelWorkbook(path) as workbook:
        table = workbook.get_bordered_table_by_columns("Amount", ["header1"])

    with pytest.raises(BorderTableShapeError):
        table.add_row(["too", "many"])


def test_save_delegates_to_workbook() -> None:
    class FakeWorkbook:
        def __init__(self) -> None:
            self.saved_table = None

        def _save_selected_columns_table(self, table: SelectedColumnsTable) -> None:
            self.saved_table = table

    workbook = FakeWorkbook()
    table = SelectedColumnsTable(
        workbook=workbook,  # type: ignore[arg-type]
        sheet="Amount",
        start_row=2,
        start_column=2,
        end_row=4,
        end_column=3,
        header_row=2,
        columns=[
            _SelectedColumn("key", 2, ["a", "b"]),
            _SelectedColumn("amount", 3, [100, 200]),
        ],
    )
    table.add_row(["c", 300])
    table.save()

    assert workbook.saved_table is table
    assert table.added_rows == 0  # save() rebaselines the original row count


def test_write_plan_snapshots_selected_columns_table() -> None:
    table = SelectedColumnsTable(
        workbook=None,  # type: ignore[arg-type]
        sheet="Amount",
        start_row=2,
        start_column=2,
        end_row=4,
        end_column=3,
        header_row=2,
        columns=[
            _SelectedColumn("key", 2, ["a", "b"]),
            _SelectedColumn("amount", 3, [100, 200]),
        ],
    )

    plan = WritePlan()
    plan.add_selected_columns_table(table)

    # Mutating the table afterwards must not change the queued operation.
    table.add_row(["c", 300])
    table.add_column([9, 9, 9], header="extra")

    assert len(plan) == 1
    op = next(iter(plan))
    assert isinstance(op, _SelectedColumnsTableOp)
    assert op.added_rows == 0
    assert op.columns == (
        (2, "key", ("a", "b")),
        (3, "amount", (100, 200)),
    )
