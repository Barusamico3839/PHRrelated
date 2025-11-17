"""Step Af: update 弔事連絡票 sheet in the RPA workbook."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

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

    robot.current_phase = "Af.chouji_renraku_hyou"
    logger = logging.getLogger("chouji_robo.excel")
    logger.info("Af.chouji_renraku_hyou: 弔事連絡票シートを更新中…")

    if not robot.paths.temp_forms_book.exists():
        raise FileNotFoundError("temp_弔事連絡票.xlsx が存在しません。")

    excel = None
    wb_source = None
    wb_rpa = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb_source = excel.Workbooks.Open(str(robot.paths.temp_forms_book))
        source_sheet = None
        for sheet in wb_source.Worksheets:
            if "弔事連絡票" in robot._safe_str(sheet.Name):
                source_sheet = sheet
                break
        if source_sheet is None:
            raise ValueError("弔事連絡票を含むシートが temp ブック内に見つかりません。")

        wb_rpa = excel.Workbooks.Open(str(robot.paths.rpa_local_book))
        _log_sheet_names(wb_rpa, logger, "before_chouji_copy")
        generic_sheet_name = "弔事連絡票"
        company_sheet_name = f"{robot.state.company_name}弔事連絡票"
        _safe_delete_sheet(wb_rpa, generic_sheet_name, logger)

        source_sheet.Copy(Before=wb_rpa.Worksheets(1))
        new_sheet = wb_rpa.Worksheets(1)
        new_sheet.Name = generic_sheet_name

        for sheet in wb_rpa.Worksheets:
            sheet.Cells.Replace(What=company_sheet_name, Replacement=generic_sheet_name, LookAt=2, SearchOrder=1, MatchCase=False)

        _safe_delete_sheet(wb_rpa, company_sheet_name, logger)

        _log_sheet_names(wb_rpa, logger, "after_chouji_copy")

        wb_rpa.Save()

    finally:
        if wb_rpa is not None:
            try:
                wb_rpa.Close(SaveChanges=True)
            except Exception:
                pass
        if wb_source is not None:
            try:
                wb_source.Close(SaveChanges=False)
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

