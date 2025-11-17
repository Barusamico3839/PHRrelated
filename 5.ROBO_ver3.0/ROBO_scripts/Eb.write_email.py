#!/usr/bin/env python3
"""Create an Outlook draft populated with data from the RPA sheet."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict

import win32com.client  # type: ignore

import pythoncom
from excel_com import open_workbook


LOGGER = logging.getLogger("chouji_robo.email")


def _read_rpa_values(robot, addresses: Dict[str, str]) -> Dict[str, str]:
    result: Dict[str, str] = {}
    with open_workbook(robot.paths.rpa_book_destination, read_only=True) as workbook:
        sheet = workbook.Worksheets("RPAシート")
        for key, addr in addresses.items():
            try:
                value = sheet.Range(addr).Value
            except Exception:
                value = ""
            result[key] = robot._safe_str(value)
    return result


def run(robot) -> None:
    robot.current_phase = "Eb.write_email"
    LOGGER.info("Eb.write_email: Outlook 下書きを作成します。")

    pdf_path = robot.state.generated_pdf_path
    if not pdf_path or not Path(pdf_path).exists():
        raise FileNotFoundError("PDF が見つかりません。Ea.create_excel_PDF を先に実行してください。")

    data = _read_rpa_values(
        robot,
        {
            "to": "D3",
            "subject": "D12",
            "cc": "D25",
            "bcc": "D26",
            "body": "D27",
        },
    )

    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
    except Exception:
        pass

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = data["to"]
        mail.Subject = data["subject"]
        if data["cc"]:
            mail.CC = data["cc"]
        if data["bcc"]:
            mail.BCC = data["bcc"]
        mail.Body = data["body"]
        mail.Attachments.Add(Source=str(pdf_path))
        mail.Save()

        robot.state.outlook_draft_entry_id = mail.EntryID or ""
        robot.state.outlook_draft_store_id = getattr(mail, "StoreID", "") or ""
        LOGGER.info("Outlook 下書きを保存しました。EntryID=%s", robot.state.outlook_draft_entry_id)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
