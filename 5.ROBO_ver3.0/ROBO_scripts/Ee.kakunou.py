#!/usr/bin/env python3
"""Optionally archive generated PDF/Excel files."""

from __future__ import annotations

import logging
import shutil
from pathlib import Path

from tkinter import messagebox


LOGGER = logging.getLogger("chouji_robo.email")


def run(robot) -> None:
    robot.current_phase = "Ee.kakunou"
    if messagebox is None:
        LOGGER.debug("tkinter.messagebox が利用できないため格納確認をスキップします。")
        return

    pdf_path = robot.state.generated_pdf_path
    xlsx_path = robot.state.generated_excel_path
    if not pdf_path or not Path(pdf_path).exists():
        LOGGER.debug("格納する PDF が見つからないためスキップします。")
        return

    answer = messagebox.askyesno("格納確認", "今回の弔事連絡票(PDFファイル)を指定場所に格納しますか？", parent=robot.root)
    if not answer:
        return

    company_name = robot._safe_str(robot.state.company_name)
    if not company_name:
        messagebox.showwarning("会社名未設定", "会社名が未設定のため格納先を決定できません。")
        return

    target_dir = robot.paths.company_archive_dir(company_name)
    robot._ensure_directory(target_dir)
    LOGGER.info("PDF/Excel を %s に格納します。", target_dir)

    for source in (pdf_path, xlsx_path):
        if not source:
            continue
        source_path = Path(source)
        if not source_path.exists():
            continue
        destination = target_dir / source_path.name
        shutil.copy2(source_path, destination)
        LOGGER.debug("%s を %s にコピーしました。", source_path, destination)
