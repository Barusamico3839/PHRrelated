"""Thin helpers for interacting with Excel exclusively through COM."""

from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, List, Sequence, Tuple

import pythoncom  # type: ignore
import win32com.client  # type: ignore


@contextmanager
def open_workbook(path: Path | str, *, read_only: bool = False, visible: bool = False):
    """Open an Excel workbook via COM and always tear it down safely."""

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    workbook = excel.Workbooks.Open(
        str(Path(path)),
        UpdateLinks=False,
        ReadOnly=read_only,
    )
    try:
        yield workbook
    finally:
        try:
            workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def get_used_range_bounds(sheet) -> Tuple[int, int, int, int]:
    """Return (row_start, row_end, col_start, col_end) for a worksheet."""

    try:
        used = sheet.UsedRange
    except Exception:
        return 1, 0, 1, 0
    if used is None:
        return 1, 0, 1, 0
    start_row = int(used.Row)
    start_col = int(used.Column)
    row_count = int(getattr(used.Rows, "Count", 0) or 0)
    col_count = int(getattr(used.Columns, "Count", 0) or 0)
    end_row = start_row + row_count - 1 if row_count else start_row - 1
    end_col = start_col + col_count - 1 if col_count else start_col - 1
    return start_row, end_row, start_col, end_col


def iter_rows(
    sheet,
    *,
    start_row: int = 1,
    start_col: int = 1,
    end_col: int | None = None,
) -> Iterator[tuple[int, List]]:
    """Yield (row_index, row_values) across the used range."""

    row_start, row_end, col_start, col_end = get_used_range_bounds(sheet)
    if row_end < row_start:
        return
    start_row = max(start_row, row_start)
    if end_col is None:
        end_col = col_end
    col_start = max(start_col, col_start)
    end_col = max(col_start, end_col)
    for row_idx in range(start_row, row_end + 1):
        yield row_idx, read_row(sheet, row_idx, col_start, end_col)


def read_row(sheet, row_idx: int, start_col: int, end_col: int) -> List:
    """Read values for a single row between the provided columns."""

    rng = sheet.Range(sheet.Cells(row_idx, start_col), sheet.Cells(row_idx, end_col))
    values = rng.Value
    if start_col == end_col:
        return [values]
    if isinstance(values, tuple):
        return list(values)
    return [values]


def write_row(sheet, row_idx: int, values: Sequence, start_col: int = 1) -> None:
    """Write a contiguous block of values into a row."""

    if not values:
        return
    end_col = start_col + len(values) - 1
    rng = sheet.Range(sheet.Cells(row_idx, start_col), sheet.Cells(row_idx, end_col))
    if len(values) == 1:
        rng.Value = values[0]
    else:
        rng.Value = tuple(values)
