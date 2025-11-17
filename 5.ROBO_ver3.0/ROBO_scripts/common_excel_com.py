"""Lightweight helpers for interacting with Excel via COM (pywin32)."""

from __future__ import annotations

import contextlib
import time
from pathlib import Path
from typing import Any, Iterable, Optional

try:
    import win32com.client  # type: ignore
    from win32com.client import constants  # type: ignore
except Exception as exc:  # pragma: no cover - pywin32 must be installed on Windows
    win32com = None  # type: ignore
    _IMPORT_ERROR = exc
else:
    _IMPORT_ERROR = None


class ExcelCOM:
    """Context manager to control a single Excel Application instance via COM."""

    def __init__(self, visible: bool = False) -> None:
        if win32com is None:
            raise RuntimeError("pywin32 (win32com) が見つからないため Excel COM を利用できません。") from _IMPORT_ERROR
        self.visible = visible
        self.app = None
        self.workbooks: list[Any] = []

    def __enter__(self) -> "ExcelCOM":
        self.app = win32com.DispatchEx("Excel.Application")
        self.app.Visible = self.visible
        self.app.DisplayAlerts = False
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        for wb in list(self.workbooks):
            with contextlib.suppress(Exception):
                wb.Close(SaveChanges=False)
        with contextlib.suppress(Exception):
            if self.app is not None:
                self.app.Quit()
        self.app = None
        self.workbooks.clear()

    def open(self, path: Path | str, read_only: bool = False) -> Any:
        if self.app is None:
            raise RuntimeError("Excel application is not initialised.")
        wb = self.app.Workbooks.Open(str(path), UpdateLinks=False, ReadOnly=read_only)
        self.workbooks.append(wb)
        return wb


def read_range(workbook_path: Path | str, sheet_name: str, cell_refs: Iterable[str]) -> dict[str, Any]:
    """Return a dict of cell_ref -> value using COM."""

    workbook_path = Path(workbook_path)
    results: dict[str, Any] = {}
    with ExcelCOM(visible=False) as excel:
        wb = excel.open(workbook_path, read_only=True)
        sheet = wb.Worksheets(sheet_name)
        for ref in cell_refs:
            results[ref] = sheet.Range(ref).Value
    return results


def read_rows(
    workbook_path: Path | str,
    sheet_name: str,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> list[list[Any]]:
    """Read a rectangular region via COM and return as a nested list."""

    workbook_path = Path(workbook_path)
    with ExcelCOM(visible=False) as excel:
        wb = excel.open(workbook_path, read_only=True)
        sheet = wb.Worksheets(sheet_name)
        rng = sheet.Range(
            sheet.Cells(start_row, start_col),
            sheet.Cells(end_row, end_col),
        )
        values = rng.Value
    rows: list[list[Any]] = []
    if isinstance(values, tuple):
        for row in values:
            rows.append(list(row))
    else:
        rows.append([values])
    return rows


def write_cells(
    workbook_path: Path | str,
    sheet_name: str,
    assignments: dict[str, Any],
    save: bool = True,
    visible: bool = False,
) -> None:
    """Write multiple cell references via COM."""

    workbook_path = Path(workbook_path)
    with ExcelCOM(visible=visible) as excel:
        wb = excel.open(workbook_path, read_only=False)
        sheet = wb.Worksheets(sheet_name)
        for ref, value in assignments.items():
            sheet.Range(ref).Value = value
        if save:
            wb.Save()


def wait_for_file(path: Path | str, timeout: int = 10) -> None:
    """Ensure the target workbook exists before COM operations."""

    path = Path(path)
    deadline = time.time() + timeout
    while not path.exists() and time.time() < deadline:
        time.sleep(0.2)
    if not path.exists():
        raise FileNotFoundError(f"Excel ファイルが見つかりません: {path}")
