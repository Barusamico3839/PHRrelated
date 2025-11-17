#!/usr/bin/env python3
"""Show the final dialog depending on workflow success."""

from __future__ import annotations

import tkinter as tk


def _build_dialog(robot, title: str, lines: list[tuple[str, str]]) -> None:
    window = tk.Toplevel(robot.root)
    window.title(title)
    window.geometry("520x260")
    window.attributes("-topmost", True)

    for idx, (text, font) in enumerate(lines):
        label = tk.Label(window, text=text, font=font, wraplength=480, justify="center")
        pady = 18 if idx == 0 else 8
        label.pack(pady=(pady, 0))

    done = tk.BooleanVar(value=False)

    def _close() -> None:
        done.set(True)
        window.destroy()

    button = tk.Button(window, text="OK", width=12, command=_close)
    button.pack(pady=20)

    window.grab_set()
    window.wait_variable(done)


def run(robot, *, success: bool, error_message: str) -> None:
    robot.current_phase = "Ef.last_dialog"
    if success:
        lines = [
            ("手配が完了しました！", "HGS創英角ﾎﾟｯﾌﾟ体 24"),
            ("ロボをご利用ありがとうございました。", "游ゴシック 16"),
            (error_message or "エラーは発生していません。", "游ゴシック 12"),
        ]
        _build_dialog(robot, "完了", lines)
    else:
        lines = [
            ("申し訳ありません、ロボはエラーで終わりました。", "游ゴシック 18"),
            ("以下がエラー内容です。", "游ゴシック 12"),
            (error_message or "原因不明のエラーです。", "游ゴシック 12"),
        ]
        _build_dialog(robot, "エラー", lines)
