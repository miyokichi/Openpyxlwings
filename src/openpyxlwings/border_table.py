"""Border-based table detection and editing."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from openpyxlwings.exceptions import (
    BorderTableNotFoundError,
    BorderTableShapeError,
)

if TYPE_CHECKING:
    from openpyxlwings.workbook import CellValue, ExcelWorkbook, Table


@dataclass
class _Insertion:
    axis: str
    index: int


@dataclass
class BorderTable:
    """Editable model for a plain Excel range divided by borders."""

    workbook: ExcelWorkbook
    sheet: str | None
    start_row: int
    start_column: int
    values: Table
    header_rows: int = 1
    header_columns: int = 0
    _original_rows: int = field(init=False, repr=False)
    _original_columns: int = field(init=False, repr=False)
    _insertions: list[_Insertion] = field(default_factory=list, init=False, repr=False)

    def __post_init__(self) -> None:
        self._validate_shape()
        self._original_rows = len(self.values)
        self._original_columns = self.column_count

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def column_count(self) -> int:
        return len(self.values[0]) if self.values else 0

    @property
    def end_row(self) -> int:
        return self.start_row + self.row_count - 1

    @property
    def end_column(self) -> int:
        return self.start_column + self.column_count - 1

    @property
    def range(self) -> str:
        return f"{_cell_address(self.start_row, self.start_column)}:{_cell_address(self.end_row, self.end_column)}"

    @property
    def data(self) -> Table:
        return [row[self.header_columns :] for row in self.values[self.header_rows :]]

    @property
    def row_headers(self) -> Table:
        if self.header_columns == 0:
            return []
        return [row[: self.header_columns] for row in self.values[self.header_rows :]]

    @property
    def column_headers(self) -> Table:
        return [row[self.header_columns :] for row in self.values[: self.header_rows]]

    def set_value(self, row: int, column: int, value: CellValue) -> None:
        """Set a value in the full table using 1-based table coordinates."""

        self._validate_table_position(row, column)
        self.values[row - 1][column - 1] = value

    def set_body_value(self, row: int, column: int, value: CellValue) -> None:
        """Set a value in the body area using 1-based body coordinates."""

        body_rows = self.row_count - self.header_rows
        body_columns = self.column_count - self.header_columns
        _validate_position(row, column, max_row=body_rows, max_column=body_columns)
        self.values[self.header_rows + row - 1][self.header_columns + column - 1] = value

    def add_row(
        self,
        values: list[CellValue],
        *,
        row_headers: list[CellValue] | None = None,
        position: int | None = None,
    ) -> None:
        """Add a body row. ``position`` is 1-based inside the body."""

        body_rows = self.row_count - self.header_rows
        body_columns = self.column_count - self.header_columns
        insert_body_index = body_rows + 1 if position is None else position
        if insert_body_index < 1 or insert_body_index > body_rows + 1:
            raise BorderTableShapeError("row position is outside the body range.")
        if len(values) != body_columns:
            raise BorderTableShapeError("row values length does not match table body width.")

        headers = [] if row_headers is None else list(row_headers)
        if len(headers) != self.header_columns:
            raise BorderTableShapeError("row_headers length does not match header_columns.")

        table_index = self.header_rows + insert_body_index - 1
        self.values.insert(table_index, headers + list(values))
        self._insertions.append(_Insertion("row", self.start_row + table_index))

    def add_column(
        self,
        values: list[CellValue],
        *,
        column_headers: list[CellValue] | None = None,
        position: int | None = None,
    ) -> None:
        """Add a body column. ``position`` is 1-based inside the body."""

        body_rows = self.row_count - self.header_rows
        body_columns = self.column_count - self.header_columns
        insert_body_index = body_columns + 1 if position is None else position
        if insert_body_index < 1 or insert_body_index > body_columns + 1:
            raise BorderTableShapeError("column position is outside the body range.")
        if len(values) != body_rows:
            raise BorderTableShapeError("column values length does not match table body height.")

        headers = [] if column_headers is None else list(column_headers)
        if len(headers) != self.header_rows:
            raise BorderTableShapeError("column_headers length does not match header_rows.")

        table_column_index = self.header_columns + insert_body_index - 1
        for row_index, row in enumerate(self.values):
            value = headers[row_index] if row_index < self.header_rows else values[row_index - self.header_rows]
            row.insert(table_column_index, value)
        self._insertions.append(_Insertion("column", self.start_column + table_column_index))

    def add_header_row(
        self,
        values: list[CellValue],
        *,
        position: int | None = None,
    ) -> None:
        """Add a column-header row. ``position`` is 1-based inside headers."""

        insert_index = self.header_rows + 1 if position is None else position
        if insert_index < 1 or insert_index > self.header_rows + 1:
            raise BorderTableShapeError("header row position is outside the header range.")
        if len(values) != self.column_count:
            raise BorderTableShapeError("header row values length does not match table width.")

        table_index = insert_index - 1
        self.values.insert(table_index, list(values))
        self.header_rows += 1
        self._insertions.append(_Insertion("row", self.start_row + table_index))

    def add_header_column(
        self,
        values: list[CellValue],
        *,
        position: int | None = None,
    ) -> None:
        """Add a row-header column. ``position`` is 1-based inside headers."""

        insert_index = self.header_columns + 1 if position is None else position
        if insert_index < 1 or insert_index > self.header_columns + 1:
            raise BorderTableShapeError("header column position is outside the header range.")
        if len(values) != self.row_count:
            raise BorderTableShapeError("header column values length does not match table height.")

        table_column_index = insert_index - 1
        for row, value in zip(self.values, values, strict=True):
            row.insert(table_column_index, value)
        self.header_columns += 1
        self._insertions.append(_Insertion("column", self.start_column + table_column_index))

    def save(self) -> None:
        """Write the edited table back to its original workbook position."""

        self.workbook._save_bordered_table(self)
        self._original_rows = self.row_count
        self._original_columns = self.column_count
        self._insertions.clear()

    def _validate_shape(self) -> None:
        if not self.values:
            raise BorderTableShapeError("table values cannot be empty.")
        width = len(self.values[0])
        if width == 0:
            raise BorderTableShapeError("table width cannot be zero.")
        if any(len(row) != width for row in self.values):
            raise BorderTableShapeError("table values must be rectangular.")
        if self.header_rows < 0 or self.header_columns < 0:
            raise BorderTableShapeError("header counts cannot be negative.")
        if self.header_rows >= len(self.values):
            raise BorderTableShapeError("header_rows must leave at least one body row.")
        if self.header_columns >= width:
            raise BorderTableShapeError("header_columns must leave at least one body column.")

    def _validate_table_position(self, row: int, column: int) -> None:
        _validate_position(row, column, max_row=self.row_count, max_column=self.column_count)


def detect_bordered_table(
    workbook: ExcelWorkbook,
    worksheet: Worksheet,
    sheet: str | None,
    row: int,
    column: int,
    *,
    header_rows: int = 1,
    header_columns: int = 0,
) -> BorderTable:
    """Detect a rectangular bordered table containing ``row``/``column``."""

    _validate_position(row, column)
    _reject_merged_cell(worksheet, row, column)

    start_row = row
    end_row = row
    start_column = column
    end_column = column

    while (
        start_row > 1
        and _has_horizontal_boundary(worksheet, start_row - 1, column)
        and _cell_has_any_border(worksheet, start_row - 1, column)
    ):
        start_row -= 1
    while (
        end_row < worksheet.max_row
        and _has_horizontal_boundary(worksheet, end_row, column)
        and _cell_has_any_border(worksheet, end_row + 1, column)
    ):
        end_row += 1
    while (
        start_column > 1
        and _has_vertical_boundary(worksheet, row, start_column - 1)
        and _cell_has_any_border(worksheet, row, start_column - 1)
    ):
        start_column -= 1
    while (
        end_column < worksheet.max_column
        and _has_vertical_boundary(worksheet, row, end_column)
        and _cell_has_any_border(worksheet, row, end_column + 1)
    ):
        end_column += 1

    if start_row == end_row or start_column == end_column:
        raise BorderTableNotFoundError("A bordered table was not found from the start cell.")

    _validate_bordered_rectangle(worksheet, start_row, start_column, end_row, end_column)
    values = [
        [cell.value for cell in row_cells]
        for row_cells in worksheet.iter_rows(
            min_row=start_row,
            max_row=end_row,
            min_col=start_column,
            max_col=end_column,
        )
    ]
    return BorderTable(
        workbook=workbook,
        sheet=sheet,
        start_row=start_row,
        start_column=start_column,
        values=values,
        header_rows=header_rows,
        header_columns=header_columns,
    )


def detect_bordered_table_by_header(
    workbook: ExcelWorkbook,
    worksheet: Worksheet,
    sheet: str | None,
    header_values: list[CellValue],
    *,
    value_header_contains: str,
    header_row: int = 1,
    match_case: bool = False,
) -> BorderTable:
    """Detect a bordered table by fixed header values and value-column text."""

    if not header_values:
        raise BorderTableShapeError("header_values must contain at least one value.")
    if header_row < 1:
        raise BorderTableShapeError("header_row must be 1 or greater.")
    if not value_header_contains:
        raise BorderTableShapeError("value_header_contains cannot be empty.")

    header_width = len(header_values)
    for row in range(1, worksheet.max_row + 1):
        for column in range(1, worksheet.max_column - header_width + 2):
            if not _header_values_match(
                worksheet,
                row,
                column,
                header_values,
                match_case=match_case,
            ):
                continue

            table = detect_bordered_table(
                workbook,
                worksheet,
                sheet,
                row,
                column,
                header_rows=header_row,
                header_columns=header_width,
            )
            table_header_row_index = row - table.start_row + 1
            if table_header_row_index != header_row:
                continue

            relative_header_row = table.values[header_row - 1]
            first_value_column = _find_first_value_header_column(
                relative_header_row,
                value_header_contains,
                match_case=match_case,
            )
            if first_value_column is None:
                continue
            if first_value_column <= header_width:
                raise BorderTableShapeError(
                    "value-header columns must be to the right of header_values."
                )
            if first_value_column != header_width + 1:
                raise BorderTableShapeError(
                    "header_values must cover every row-header column before the value area."
                )

            table.header_rows = header_row
            table.header_columns = first_value_column - 1
            table._validate_shape()
            return table

    raise BorderTableNotFoundError("A bordered table matching the header values was not found.")


def _validate_bordered_rectangle(
    worksheet: Worksheet,
    start_row: int,
    start_column: int,
    end_row: int,
    end_column: int,
) -> None:
    for row in range(start_row, end_row + 1):
        for column in range(start_column, end_column + 1):
            _reject_merged_cell(worksheet, row, column)

    for column in range(start_column, end_column + 1):
        if not _has_top_border(worksheet, start_row, column):
            raise BorderTableShapeError("table top border is incomplete.")
        if not _has_bottom_border(worksheet, end_row, column):
            raise BorderTableShapeError("table bottom border is incomplete.")

    for row in range(start_row, end_row + 1):
        if not _has_left_border(worksheet, row, start_column):
            raise BorderTableShapeError("table left border is incomplete.")
        if not _has_right_border(worksheet, row, end_column):
            raise BorderTableShapeError("table right border is incomplete.")

    for row in range(start_row, end_row):
        for column in range(start_column, end_column + 1):
            if not _has_horizontal_boundary(worksheet, row, column):
                raise BorderTableShapeError("table has a missing internal horizontal border.")

    for column in range(start_column, end_column):
        for row in range(start_row, end_row + 1):
            if not _has_vertical_boundary(worksheet, row, column):
                raise BorderTableShapeError("table has a missing internal vertical border.")


def _reject_merged_cell(worksheet: Worksheet, row: int, column: int) -> None:
    cell_ref = _cell_address(row, column)
    for merged_range in worksheet.merged_cells.ranges:
        if cell_ref in merged_range:
            raise BorderTableShapeError("merged cells are not supported in bordered tables.")


def _has_vertical_boundary(worksheet: Worksheet, row: int, left_column: int) -> bool:
    return _has_right_border(worksheet, row, left_column) or _has_left_border(
        worksheet,
        row,
        left_column + 1,
    )


def _has_horizontal_boundary(worksheet: Worksheet, top_row: int, column: int) -> bool:
    return _has_bottom_border(worksheet, top_row, column) or _has_top_border(
        worksheet,
        top_row + 1,
        column,
    )


def _has_top_border(worksheet: Worksheet, row: int, column: int) -> bool:
    return _has_side(worksheet.cell(row=row, column=column).border.top)


def _has_bottom_border(worksheet: Worksheet, row: int, column: int) -> bool:
    return _has_side(worksheet.cell(row=row, column=column).border.bottom)


def _has_left_border(worksheet: Worksheet, row: int, column: int) -> bool:
    return _has_side(worksheet.cell(row=row, column=column).border.left)


def _has_right_border(worksheet: Worksheet, row: int, column: int) -> bool:
    return _has_side(worksheet.cell(row=row, column=column).border.right)


def _has_side(side: object) -> bool:
    return bool(getattr(side, "style", None))


def _cell_has_any_border(worksheet: Worksheet, row: int, column: int) -> bool:
    border = worksheet.cell(row=row, column=column).border
    return any(
        _has_side(side)
        for side in (border.top, border.bottom, border.left, border.right)
    )


def _header_values_match(
    worksheet: Worksheet,
    row: int,
    column: int,
    expected_values: list[CellValue],
    *,
    match_case: bool,
) -> bool:
    for offset, expected in enumerate(expected_values):
        actual = worksheet.cell(row=row, column=column + offset).value
        if _normalize_value(actual, match_case=match_case) != _normalize_value(
            expected,
            match_case=match_case,
        ):
            return False
    return True


def _find_first_value_header_column(
    values: list[CellValue],
    text: str,
    *,
    match_case: bool,
) -> int | None:
    needle = text if match_case else text.casefold()
    for index, value in enumerate(values, start=1):
        haystack = "" if value is None else str(value)
        if not match_case:
            haystack = haystack.casefold()
        if needle in haystack:
            return index
    return None


def _normalize_value(value: CellValue, *, match_case: bool) -> str:
    normalized = "" if value is None else str(value).strip()
    return normalized if match_case else normalized.casefold()


def _cell_address(row: int, column: int) -> str:
    return f"{get_column_letter(column)}{row}"


def _validate_position(
    row: int,
    column: int,
    *,
    max_row: int | None = None,
    max_column: int | None = None,
) -> None:
    if row < 1:
        raise BorderTableShapeError("row must be 1 or greater.")
    if column < 1:
        raise BorderTableShapeError("column must be 1 or greater.")
    if max_row is not None and row > max_row:
        raise BorderTableShapeError("row is outside the table range.")
    if max_column is not None and column > max_column:
        raise BorderTableShapeError("column is outside the table range.")
