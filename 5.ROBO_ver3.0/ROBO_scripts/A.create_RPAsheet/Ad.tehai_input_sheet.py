"""Step Ad: prepare 手配入力シート for the current company."""

from __future__ import annotations

import logging
import shutil
import time
from typing import TYPE_CHECKING

import pythoncom
import win32com.client

if TYPE_CHECKING:  # pragma: no cover - typing only
    from main import ChoujiRobo


def run(robot: "ChoujiRobo") -> None:
    robot.current_phase = "Ad.tehai_input_sheet"
    logger = logging.getLogger("chouji_robo.excel")
    logger.info("Ad.tehai_input_sheet: 手配入力シートを準備中…")

    company_name = robot.state.company_name
    if not company_name:
        raise RuntimeError("会社名が設定されていません。")

    source_rpa_book = robot.paths.panasonic_rpa_book()
    target_rpa_book = robot.paths.rpa_local_book
    robot._ensure_directory(target_rpa_book.parent)

    logger.debug("RPA元ブック: %s", source_rpa_book)
    logger.debug("RPA先ブック: %s", target_rpa_book)
    shutil.copy2(source_rpa_book, target_rpa_book)
    logger.info("RPAブックをローカルへコピーしました。")

    company_input_path = robot.paths.company_input_sheet(company_name)
    if not company_input_path.exists():
        raise FileNotFoundError(f"会社別手配入力シートが見つかりません: {company_input_path}")
    logger.debug("会社別手配入力シート: %s", company_input_path)

    temp_company_path = robot.paths.temp_input_sheet_dir
    robot._ensure_directory(temp_company_path.parent)
    shutil.copy2(company_input_path, temp_company_path)
    logger.info("会社別手配入力シートをテンポラリへコピーしました: %s", temp_company_path)

    excel = None
    wb_company = None
    wb_rpa = None
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception as exc:
            logger.debug("Excel 計算モード設定をスキップしました: %s", exc)

        wb_company = excel.Workbooks.Open(str(temp_company_path))
        sheet_input = wb_company.Worksheets("入力欄")

        pin_column = None
        used_columns = sheet_input.UsedRange.Columns.Count
        for col in range(1, used_columns + 10):
            header = robot._safe_str(sheet_input.Cells(3, col).Value)
            if header == "PIN":
                pin_column = col
                break
        if pin_column is None:
            raise ValueError("入力欄シートに PIN 見出しが見つかりません。")

        sheet_input.Cells(4, pin_column).Value = robot.state.pin
        logger.debug("PIN セルを書き込みました: (4, %s) => %s", pin_column, robot.state.pin)
        wb_company.Save()

        if company_name.upper() == "PID":
            target_range = sheet_input.Range("E4")
            display_name = ""
            max_attempts = 10
            for attempt in range(max_attempts):
                wait_time = 0.2 if attempt == 0 else 2.0
                time.sleep(wait_time)
                try:
                    excel.CalculateUntilAsyncQueriesDone()
                except Exception:
                    pass
                try:
                    excel.CalculateFull()
                except Exception:
                    pass

                value = target_range.Value
                formula = str(target_range.Formula or "")
                display_name = robot._safe_str(value)
                if formula.startswith("=") and (display_name.startswith("=") or display_name == ""):
                    if attempt + 1 < max_attempts:
                        logger.debug("name_katakana が未計算のため再取得します (%d/%d)", attempt + 1, max_attempts)
                        continue
                    raise RuntimeError("name_katakana が数式のまま取得されました。時間をおいて再実行してください。")
                break

            robot.state.name_katakana = display_name
            logger.info("PID なので name_katakana を取得: %s", robot.state.name_katakana)
            print(f"[INFO] name_katakana={robot.state.name_katakana}")

        wb_rpa = excel.Workbooks.Open(str(target_rpa_book))
        _log_sheet_names(wb_rpa, logger, "before_tehai_copy")
        generic_sheet_name = "手配入力シート"
        company_sheet_name = f"{company_name}手配入力シート"
        _safe_delete_sheet(wb_rpa, generic_sheet_name, logger)

        sheet_input.Copy(Before=wb_rpa.Worksheets(1))
        new_sheet = wb_rpa.Worksheets(1)
        new_sheet.Name = generic_sheet_name

        for sheet in wb_rpa.Worksheets:
            sheet.Cells.Replace(What=company_sheet_name, Replacement=generic_sheet_name, LookAt=2, SearchOrder=1, MatchCase=False)

        _safe_delete_sheet(wb_rpa, company_sheet_name, logger)

        _log_sheet_names(wb_rpa, logger, "after_tehai_copy")

        wb_rpa.Save()
        wb_company.Save()

    finally:
        if wb_rpa is not None:
            try:
                wb_rpa.Close(SaveChanges=True)
            except Exception:
                pass
        if wb_company is not None:
            try:
                wb_company.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _safe_delete_sheet(workbook, sheet_name: str, logger: logging.Logger) -> bool:
    try:
        sheet = workbook.Worksheets(sheet_name)
    except Exception:
        return False
    actual_name = str(getattr(sheet, "Name", ""))
    if "RPAシート" in actual_name:
        logger.debug("Skipping deletion for sheet '%s' to preserve RPAシート.", actual_name)
        return False
    try:
        sheet.Delete()
        logger.debug("Deleted sheet '%s'.", actual_name)
        return True
    except Exception as exc:
        logger.debug("Failed to delete sheet '%s': %s", actual_name, exc)
        return False


def _log_sheet_names(workbook, logger: logging.Logger, label: str) -> None:
    try:
        names = [str(getattr(sheet, "Name", "")) for sheet in workbook.Worksheets]
        logger.debug("Workbook sheets (%s): %s", label, names)
    except Exception as exc:
        logger.debug("Unable to enumerate sheets (%s): %s", label, exc)

