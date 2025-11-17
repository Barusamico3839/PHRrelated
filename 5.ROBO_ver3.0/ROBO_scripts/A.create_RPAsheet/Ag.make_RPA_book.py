"""Step Ag: produce the final RPA workbook."""

from __future__ import annotations

import logging
import shutil
import time
from typing import TYPE_CHECKING, Optional

import pythoncom
import win32com.client

if TYPE_CHECKING:  # pragma: no cover - typing only
    from main import ChoujiRobo


def run(robot: "ChoujiRobo") -> None:
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
    except Exception:
        pass
    robot.current_phase = "Ag.make_RPA_book"
    logger = logging.getLogger("chouji_robo.excel")
    logger.info("Ag.make_RPA_book: RPAブックを作成します")

    robot._ensure_directory(robot.paths.rpa_book_destination.parent)
    shutil.copy2(robot.paths.rpa_local_book, robot.paths.rpa_book_destination)
    logger.debug("RPAブックを %s にコピーしました。", robot.paths.rpa_book_destination)

    excel = None
    wb_rpa = None
    previous_settings = {}
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        previous_settings = _capture_excel_settings(excel)
        _apply_fast_excel_settings(excel)

        wb_rpa = excel.Workbooks.Open(str(robot.paths.rpa_book_destination))

        sheet_name_candidates: list[str] = []
        if robot.state.company_name:
            sheet_name_candidates.append(f"{robot.state.company_name}弔事連絡票")
        sheet_name_candidates.append("弔事連絡票")

        sheet_names_snapshot = _collect_sheet_names(wb_rpa)
        logger.debug("シート一覧: %s", sheet_names_snapshot)

        rpa_sheet = _find_sheet_by_name(wb_rpa, "RPAシート")
        if rpa_sheet is None:
            rpa_sheet = _find_sheet_by_keywords(wb_rpa, ["rpa", "シート"])
        company_sheet = _find_first_existing_sheet(wb_rpa, sheet_name_candidates)
        if company_sheet is None:
            company_sheet = _find_sheet_by_keywords(wb_rpa, ["弔事連絡票"])

        def log_and_set(cell_address: str, value: Optional[str], description: str, condition: bool = True) -> None:
            if rpa_sheet is None:
                return
            value_to_set = value or ""
            print(f"[INFO] RPA sheet {description} = {value_to_set or '(empty)'}")
            if not condition:
                return
            if not _write_cell_with_retry(rpa_sheet, cell_address, value_to_set, logger):
                logger.error("RPA sheet %s write failed after retries.", cell_address)

        if rpa_sheet is not None:
            try:
                current_d3 = robot._safe_str(rpa_sheet.Range("D3").Value).lower()
            except Exception:
                current_d3 = ""
            if current_d3 not in ("excel", ""):
                logger.debug("D3 の既存値 (=%s) は更新対象外かもしれません。", current_d3)
            log_and_set("D3", robot.state.mail_sender, "D3")
            log_and_set("D9", robot.state.mail_cc, "D9")
            log_and_set("D11", robot.state.mail_bcc, "D11")
            log_and_set("D102", robot.state.reply_email_body, "D102")
        else:
            logger.error("RPA�V�[�g��������Ȃ��������� D��̏������݂͍s���܂���B")
        _trim_processing_sheet(robot, wb_rpa, logger)

        wb_rpa.Save()
    finally:
        if wb_rpa is not None:
            try:
                wb_rpa.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                _restore_excel_settings(excel, previous_settings)
            except Exception:
                pass
            try:
                excel.Quit()
            except Exception:
                pass
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _trim_processing_sheet(robot: "ChoujiRobo", workbook, logger: logging.Logger) -> None:
    company_name = robot._safe_str(getattr(robot.state, "company_name", "")).strip()
    if not company_name:
        logger.debug("company_name が未設定のため RPAシート下処理2 の調整をスキップします。")
        return

    sheet = _find_sheet_by_name(workbook, "RPAシート下処理2")
    if sheet is None:
        sheet = _find_sheet_by_keywords(workbook, ["RPAシート下処理2"])
    if sheet is None:
        logger.debug("RPAシート下処理2 が見つからなかったため行削除をスキップします。")
        return

    target_row = None
    for row in range(1, 601):
        try:
            value = robot._safe_str(sheet.Cells(row, 2).Value)
        except Exception:
            value = ""
        if value == company_name:
            target_row = row
            break

    if target_row is None:
        logger.debug("RPAシート下処理2 の B 列に %s が見つからなかったため行削除をスキップします。", company_name)
        return

    start_row = max(1, target_row - 1)
    end_row = min(600, target_row + 98)
    logger.debug("RPAシート下処理2 行抽出: keep=%d-%d", start_row, end_row)

    for row in range(600, 0, -1):
        if row < start_row or row > end_row:
            try:
                sheet.Rows(row).Delete()
            except Exception as exc:
                logger.debug("RPAシート下処理2 行削除失敗 row=%d: %s", row, exc)


def _capture_excel_settings(excel) -> dict:
    settings = {}
    for attr in ("Visible", "DisplayAlerts", "ScreenUpdating"):
        try:
            settings[attr] = getattr(excel, attr)
        except Exception:
            settings[attr] = None
    for attr in ("EnableEvents", "Calculation"):
        try:
            settings[attr] = getattr(excel, attr)
        except Exception:
            settings[attr] = None
    return settings


def _apply_fast_excel_settings(excel) -> None:
    try:
        excel.Visible = False
    except Exception:
        pass
    try:
        excel.DisplayAlerts = False
    except Exception:
        pass
    try:
        excel.ScreenUpdating = False
    except Exception:
        pass
    try:
        excel.EnableEvents = False
    except Exception:
        pass
    try:
        excel.Calculation = -4135  # xlCalculationManual
    except Exception:
        pass


def _restore_excel_settings(excel, settings: dict) -> None:
    for attr, value in settings.items():
        if value is None:
            continue
        try:
            setattr(excel, attr, value)
        except Exception:
            pass


def _find_sheet_by_name(wb, name: str) -> Optional[object]:
    logger = logging.getLogger("chouji_robo.excel")
    name_clean = _normalize_sheet_token(name)
    sheet_names = []
    for sheet in wb.Worksheets:
        sheet_name = getattr(sheet, "Name", "")
        sheet_names.append(sheet_name)
        if _sheet_name_matches(sheet_name, name_clean):
            return sheet
    logger.debug("%s シートは見つかりませんでした (シート一覧=%s)", name, sheet_names)
    return None


def _find_first_existing_sheet(wb, candidates: list[str]) -> Optional[object]:
    logger = logging.getLogger("chouji_robo.excel")
    for name in candidates:
        if not name:
            continue
        name_clean = _normalize_sheet_token(name)
        for sheet in wb.Worksheets:
            sheet_name = getattr(sheet, "Name", "")
            if _sheet_name_matches(sheet_name, name_clean):
                return sheet
        logger.debug("%s シートは見つかりませんでした。", name)
    return None


def _find_sheet_by_keywords(wb, keywords: list[str]) -> Optional[object]:
    logger = logging.getLogger("chouji_robo.excel")
    normalized_keywords = [_normalize_sheet_token(keyword) for keyword in keywords if keyword]
    for sheet in wb.Worksheets:
        sheet_name = getattr(sheet, "Name", "")
        sheet_clean = _normalize_sheet_token(sheet_name)
        if all(keyword in sheet_clean for keyword in normalized_keywords):
            return sheet
    logger.debug("キーワード %s に一致するシートは見つかりませんでした。", keywords)
    return None


def _collect_sheet_names(wb) -> list[str]:
    names = []
    for sheet in wb.Worksheets:
        try:
            names.append(str(sheet.Name))
        except Exception:
            names.append("(unknown)")
    return names


def _sheet_name_matches(sheet_name: str, target_clean: str) -> bool:
    sheet_clean = _normalize_sheet_token(sheet_name)
    if sheet_clean == target_clean:
        return True
    if target_clean and target_clean in sheet_clean:
        return True
    return False


def _normalize_sheet_token(name: str) -> str:
    stripped = name.replace(" ", "").replace("　", "")
    return stripped.strip().casefold()


def _write_cell_with_retry(sheet, address: str, value: str, logger: logging.Logger, max_attempts: int = 20, delay: float = 0.5) -> bool:
    for attempt in range(1, max_attempts + 1):
        try:
            sheet.Range(address).Value = value
            logger.debug('RPA sheet %s write succeeded on attempt %d.', address, attempt)
            return True
        except Exception as exc:
            logger.debug("RPA sheet %s write failed (%d/%d): %s", address, attempt, max_attempts, exc)
            time.sleep(delay)
    return False

