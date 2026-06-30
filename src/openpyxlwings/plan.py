"""Detached, cacheable write instructions for :class:`ExcelWorkbook`.

A :class:`WritePlan` collects write operations without touching Excel or
xlwings, so it can be built and cached outside any workbook context. The
operations are executed later by :meth:`ExcelWorkbook.apply`, which opens the
Excel session only at that point.
"""

from __future__ import annotations

from collections.abc import Iterator
from copy import deepcopy
from dataclasses import dataclass
from typing import TYPE_CHECKING

from openpyxlwings.workbook import (
    CellValue,
    RangeValue,
    Table,
    _cell_address,
    _range_address,
)

if TYPE_CHECKING:
    from openpyxlwings.border_table import BorderTable
    from openpyxlwings.selected_columns import SelectedColumnsTable
    from openpyxlwings.workbook import _XlwingsWriteSession


@dataclass(frozen=True)
class _WriteValuesOp:
    sheet: str | None
    cell: str
    values: RangeValue
    expand: bool = False

    def apply(self, writer: _XlwingsWriteSession) -> None:
        writer.write_values(self.sheet, self.cell, self.values, expand=self.expand)


@dataclass(frozen=True)
class _ClearContentsOp:
    sheet: str | None
    address: str

    def apply(self, writer: _XlwingsWriteSession) -> None:
        writer.clear_contents(self.sheet, self.address)


@dataclass(frozen=True)
class _BorderedTableOp:
    sheet: str | None
    start_row: int
    start_column: int
    values: Table
    end_row: int
    end_column: int
    insertions: tuple[tuple[str, int], ...]

    def apply(self, writer: _XlwingsWriteSession) -> None:
        writer.apply_bordered_table(
            self.sheet,
            self.start_row,
            self.start_column,
            self.values,
            self.end_row,
            self.end_column,
            self.insertions,
        )


@dataclass(frozen=True)
class _SelectedColumnsTableOp:
    sheet: str | None
    start_row: int
    start_column: int
    end_row: int
    end_column: int
    header_row: int
    added_rows: int
    columns: tuple[tuple[int | None, CellValue, tuple[CellValue, ...]], ...]

    def apply(self, writer: _XlwingsWriteSession) -> None:
        writer.apply_selected_columns_table(
            self.sheet,
            self.start_row,
            self.start_column,
            self.end_row,
            self.end_column,
            self.header_row,
            self.added_rows,
            [(source, header, list(values)) for source, header, values in self.columns],
        )


_WriteOp = _WriteValuesOp | _ClearContentsOp | _BorderedTableOp | _SelectedColumnsTableOp


class WritePlan:
    """An ordered, reusable collection of deferred write operations.

    Builder methods mirror the write API of :class:`ExcelWorkbook`, validate
    their arguments immediately, and return ``self`` so calls can be chained.
    The plan holds no Excel resources and is not consumed by being applied, so
    one plan can be applied to several workbooks.
    """

    def __init__(self) -> None:
        self._ops: list[_WriteOp] = []

    def write_values(
        self,
        sheet: str | None,
        cell: str,
        values: RangeValue,
        *,
        expand: bool = False,
    ) -> WritePlan:
        """Queue a write of a scalar, row, column, or 2-D values at ``cell``.

        ``values`` is snapshotted now, so the queued write reflects the value at
        this point even if the original object is mutated or rebound later.
        """

        self._ops.append(_WriteValuesOp(sheet, cell, deepcopy(values), expand))
        return self

    def write_values_at(
        self,
        sheet: str | None,
        row: int,
        column: int,
        values: RangeValue,
        *,
        expand: bool = False,
    ) -> WritePlan:
        """Queue a write by 1-based row and column numbers."""

        return self.write_values(sheet, _cell_address(row, column), values, expand=expand)

    def clear_contents(self, sheet: str | None, address: str) -> WritePlan:
        """Queue clearing of values/formulas from a range such as ``A1:D10``."""

        self._ops.append(_ClearContentsOp(sheet, address))
        return self

    def clear_contents_at(
        self,
        sheet: str | None,
        start_row: int,
        start_column: int,
        end_row: int,
        end_column: int,
    ) -> WritePlan:
        """Queue clearing of a range by 1-based row and column numbers."""

        return self.clear_contents(
            sheet,
            _range_address(start_row, start_column, end_row, end_column),
        )

    def add_bordered_table(self, table: BorderTable) -> WritePlan:
        """Queue an edited bordered table to be written back to its position.

        A snapshot of the table's current values, bounds, and pending row/column
        insertions is taken now, so later edits to ``table`` do not affect what
        this plan writes. Unlike :meth:`BorderTable.save`, this neither writes to
        Excel nor saves until :meth:`ExcelWorkbook.apply` runs.
        """

        self._ops.append(
            _BorderedTableOp(
                sheet=table.sheet,
                start_row=table.start_row,
                start_column=table.start_column,
                values=[list(row) for row in table.values],
                end_row=table.end_row,
                end_column=table.end_column,
                insertions=tuple(
                    (insertion.axis, insertion.index) for insertion in table._insertions
                ),
            )
        )
        return self

    def add_selected_columns_table(self, table: SelectedColumnsTable) -> WritePlan:
        """Queue an edited column-selected virtual table to be written back.

        A snapshot of the table's current columns, values, bounds, and appended
        row count is taken now, so later edits to ``table`` do not affect what
        this plan writes. Unlike :meth:`SelectedColumnsTable.save`, this neither
        writes to Excel nor saves until :meth:`ExcelWorkbook.apply` runs.
        """

        self._ops.append(
            _SelectedColumnsTableOp(
                sheet=table.sheet,
                start_row=table.start_row,
                start_column=table.start_column,
                end_row=table.end_row,
                end_column=table.end_column,
                header_row=table.header_row,
                added_rows=table.added_rows,
                columns=tuple(
                    (column.source_column, column.header, tuple(column.values))
                    for column in table.columns
                ),
            )
        )
        return self

    def clear(self) -> None:
        """Discard every queued operation."""

        self._ops.clear()

    def __iter__(self) -> Iterator[_WriteOp]:
        return iter(self._ops)

    def __len__(self) -> int:
        return len(self._ops)

    def __repr__(self) -> str:
        return f"{type(self).__name__}({len(self._ops)} ops)"
