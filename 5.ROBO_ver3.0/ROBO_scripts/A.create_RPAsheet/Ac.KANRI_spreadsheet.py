"""Step Ac: read company data from the 管理表 workbook."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

from excel_com import get_used_range_bounds, open_workbook

if TYPE_CHECKING:  # pragma: no cover - typing only
    from main import ChoujiRobo


def _normalize_tehai_value(value: object) -> str:
    text = "" if value is None else str(value).strip()
    digits = "".join(ch for ch in text if ch.isdigit())
    return digits or text


def _find_row_via_find(sheet, search_values: list) -> int | None:
    for candidate in search_values:
        if candidate is None or candidate == "":
            continue
        try:
            cell = sheet.Columns(2).Find(
                What=candidate,
                LookIn=-4163,  # xlValues
                LookAt=1,  # xlWhole
                SearchOrder=1,  # xlByRows
                SearchDirection=1,  # xlNext
                MatchCase=False,
            )
        except Exception:
            cell = None
        if cell is not None:
            try:
                return int(cell.Row)
            except Exception:
                pass
    return None


def _find_row_via_scan(sheet, target_normalized: str, robot: "ChoujiRobo") -> int | None:
    row_start, row_end, _, _ = get_used_range_bounds(sheet)
    for row_idx in range(max(2, row_start), row_end + 1):
        raw_value = sheet.Cells(row_idx, 2).Value
        cell_normalized = _normalize_tehai_value(raw_value)
        matched = False
        if target_normalized and cell_normalized:
            matched = cell_normalized == target_normalized or cell_normalized.startswith(
                target_normalized
            )
        if not matched:
            matched = robot._safe_str(raw_value) == robot._safe_str(robot.state.tehai_number)
        if matched:
            return row_idx
    return None


def run(robot: "ChoujiRobo") -> None:
    robot.current_phase = "Ac.KANRI_spreadsheet"
    logger = logging.getLogger("chouji_robo.excel")
    logger.info("Ac.KANRI_spreadsheet: 管理表データを取得")

    source_book = robot.paths.kanri_report_book
    logger.debug("管理表パス: %s", source_book)
    if not source_book.exists():
        raise FileNotFoundError(f"管理表が見つかりません: {source_book}")

    target_row_index: int | None = None
    source_sheet_name: str | None = None
    target_normalized = _normalize_tehai_value(robot.state.tehai_number)
    search_values: list = []
    safe_tehai = robot._safe_str(robot.state.tehai_number)
    if safe_tehai:
        search_values.append(safe_tehai)
    if target_normalized and target_normalized not in search_values:
        search_values.append(target_normalized)
    if target_normalized.isdigit():
        try:
            search_values.append(int(target_normalized))
        except Exception:
            pass

    try:
        with open_workbook(source_book, read_only=True) as workbook:
            for sheet in workbook.Worksheets:
                target_row_index = _find_row_via_find(sheet, search_values)
                if target_row_index is None:
                    target_row_index = _find_row_via_scan(sheet, target_normalized, robot)
                if target_row_index is not None:
                    source_sheet_name = str(getattr(sheet, "Name", ""))
                    robot.state.company_name = robot._safe_str(sheet.Cells(target_row_index, 6).Value)
                    robot.state.pin = robot._safe_str(sheet.Cells(target_row_index, 7).Value)
                    raw_mail_time = sheet.Cells(target_row_index, 11).Value
                    robot.state.mail_time = robot._parse_excel_datetime(raw_mail_time)
                    break
    except Exception as exc:
        raise PermissionError(
            f"管理表をExcel COMで開けませんでした。ファイルが開かれていないか確認してください: {source_book}"
        ) from exc

    if target_row_index is None:
        raise ValueError(f"管理NO.{robot.state.tehai_number} の行が見つかりませんでした。")

    if source_sheet_name is not None:
        logger.debug("管理表シート: %s で行を取得", source_sheet_name)

    logger.info(
        "管理表取得結果: company_name=%s, pin=%s, mail_time=%s",
        robot.state.company_name,
        robot.state.pin,
        robot.state.mail_time,
    )
    print(f"[INFO] company_name: {robot.state.company_name}")
    print(f"[INFO] pin: {robot.state.pin}")
    print(f"[INFO] mail_time: {robot.state.mail_time}")
