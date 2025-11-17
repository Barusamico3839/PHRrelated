#!/usr/bin/env python3
"""Run the email creation workflow (step E)."""

from __future__ import annotations

import logging
from datetime import datetime

from module_loader import load_helper


LOGGER = logging.getLogger("chouji_robo.email")


def run(robot) -> bool:
    robot.current_phase = "E.create_email"
    LOGGER.info("E.create_email: メール作成フェーズを開始します。")

    if robot.state.workflow_started_at is None:
        robot.state.workflow_started_at = datetime.now()

    success = False
    error_message = ""

    try:
        while True:
            load_helper("Ea.create_excel_PDF").run(robot)
            load_helper("Eb.write_email").run(robot)
            action = load_helper("Ec.show_mail_and_PDF").run(robot)
            if action == "redo":
                LOGGER.info("PDFの変更が選択されたため、再度 PDF/メールを生成します。")
                continue
            break
        success = True
    except Exception as exc:  # pragma: no cover - defensive
        error_message = str(exc)
        LOGGER.exception("E ステップでエラーが発生しました: %s", exc)
    finally:
        robot.state.workflow_finished_at = datetime.now()
        robot.state.workflow_error = error_message

        try:
            load_helper("Ed.error_log").run(robot, success=success, error_message=error_message)
        except Exception:
            LOGGER.exception("Ed.error_log の実行に失敗しました。")

        if success:
            try:
                load_helper("Ee.kakunou").run(robot)
            except Exception:
                LOGGER.exception("Ee.kakunou の実行に失敗しました。")
            load_helper("Ef.last_dialog").run(robot, success=True, error_message="")
        else:
            load_helper("Ef.last_dialog").run(robot, success=False, error_message=error_message or "原因不明のエラーです。")

    return success
