#!/usr/bin/env python3
"""Generate PDF/Excel exports for the 弔事連絡票 sheet."""

from __future__ import annotations

import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional

from excel_com import open_workbook


LOGGER = logging.getLogger("chouji_robo.email")


def _sanitize_filename(value: str) -> str:
    normalized = value.strip()
    normalized = re.sub(r'[\\/:*?"<>|]', "_", normalized)
    if not normalized:
        normalized = datetime.now().strftime("弔事連絡票_%Y%m%d_%H%M%S")
    return normalized


def _ensure_target_sheet(workbook, name: str):
    try:
        return workbook.Worksheets(name)
    except Exception as exc:  # pragma: no cover - COM guard
        raise RuntimeError(f"'{name}' シートが見つかりません。") from exc


def _read_cell(sheet, address: str) -> str:
    try:
        value = sheet.Range(address).Value
    except Exception:
        value = ""
    return "" if value is None else str(value).strip()


def run(robot) -> None:
    robot.current_phase = "Ea.create_excel_PDF"
    LOGGER.info("Ea.create_excel_PDF: PDF/Excel の生成を開始します。")

    output_dir = robot.paths.rpa_book_dir
    robot._ensure_directory(output_dir)

    rpa_book = robot.paths.rpa_book_destination
    if not rpa_book.exists():
        raise FileNotFoundError(f"RPAブックが見つかりません: {rpa_book}")

    # 既存ファイルの整理
    for item in output_dir.iterdir():
        try:
            if item.resolve() == rpa_book.resolve():
                continue
        except Exception:
            if item.name == rpa_book.name:
                continue
        if item.is_dir():
            shutil.rmtree(item, ignore_errors=True)
        else:
            try:
                item.unlink()
            except Exception:
                LOGGER.warning("ファイル %s の削除に失敗しました。", item, exc_info=True)

    with open_workbook(rpa_book) as workbook:
        rpa_sheet = _ensure_target_sheet(workbook, "RPAシート")
        company_sheet = _ensure_target_sheet(workbook, "弔事連絡票")

        base_name = _sanitize_filename(_read_cell(rpa_sheet, "D13"))
        pdf_path = output_dir / f"{base_name}.pdf"
        xlsx_path = output_dir / f"{base_name}.xlsx"

        LOGGER.debug("PDF を %s に書き出します。", pdf_path)
        company_sheet.ExportAsFixedFormat(0, str(pdf_path))

        LOGGER.debug("Excel を %s に書き出します。", xlsx_path)
        excel = workbook.Application
        company_sheet.Copy()  # 新しいブックとしてコピー
        copied_book = excel.ActiveWorkbook
        try:
            copied_book.SaveAs(str(xlsx_path), FileFormat=51)  # xlOpenXMLWorkbook
        finally:
            copied_book.Close(SaveChanges=False)

    robot.state.generated_pdf_path = str(pdf_path)
    robot.state.generated_excel_path = str(xlsx_path)
    LOGGER.info("PDF/Excel の生成が完了しました: %s, %s", pdf_path, xlsx_path)
