"""Top-level orchestrator for step A (create_RPAsheet)."""

from __future__ import annotations

import logging

from common import PathRegistry
from excel_com import open_workbook
from module_loader import load_helper


def _hide_non_target_sheets(logger: logging.Logger) -> None:
    """Hide every sheet except 'RPAシート' and '奉行メール' in the final workbook."""

    workbook_path = PathRegistry().rpa_book_destination
    if not workbook_path.exists():
        logger.warning("RPAブックが見つからないためシートの非表示処理をスキップします: %s", workbook_path)
        return

    keep_titles = {"RPAシート", "奉行メール", "弔事連絡票"}
    changed = False

    with open_workbook(workbook_path) as workbook:
        for sheet in workbook.Worksheets:
            title = str(getattr(sheet, "Name", "") or "")
            is_target = title in keep_titles
            is_hidden = int(getattr(sheet, "Visible", 1) or 1) == 0
            if is_target and is_hidden:
                sheet.Visible = -1  # xlSheetVisible
                changed = True
                logger.debug("シート '%s' を再表示しました。", title)
            elif not is_target and not is_hidden:
                sheet.Visible = 0  # xlSheetHidden
                changed = True
                logger.debug("シート '%s' を非表示に設定しました。", title)

        if not changed:
            logger.info("非表示にする追加シートはありませんでした。")
            return

        workbook.Save()
        logger.info("RPAブックで対象外シートを非表示にしました: %s", workbook_path)


def run(robot) -> None:
    robot.current_phase = "A.create_RPAsheet"
    logger = logging.getLogger("chouji_robo")
    logger.info("A.create_RPAsheet を開始しました。")

    load_helper("Ac.KANRI_spreadsheet").run(robot)
    load_helper("Ad.tehai_input_sheet").run(robot)
    load_helper("Ae.get_mail").run(robot)
    load_helper("Af.chouji_renraku_hyou").run(robot)
    load_helper("Ag.make_RPA_book").run(robot)

    _hide_non_target_sheets(logger)

    logger.info("A.create_RPAsheet のすべてのサブステップが完了しました。")
