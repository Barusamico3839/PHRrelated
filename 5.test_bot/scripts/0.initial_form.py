# -*- coding: utf-8 -*-
import sys
import os
import re
import datetime as _dt
from typing import List, Tuple

try:
    import tkinter as tk
    from tkinter import messagebox
    from tkinter import scrolledtext
except Exception as e:  # pragma: no cover
    print(f"[0.initial_form] Tkinter import error: {e}")
    raise


def _rainbow_colors(n: int) -> List[str]:
    import colorsys
    return [
        "#%02x%02x%02x" % tuple(int(c * 255) for c in colorsys.hsv_to_rgb(i / max(n, 1), 0.75, 0.95))
        for i in range(n)
    ]


def _parse_tehai_numbers(text: str) -> List[int]:
    parts = [p for p in re.split(r"\s+", text.strip()) if p]
    nums: List[int] = []
    for p in parts:
        if not p.isdigit():
            raise ValueError(f"手配番号に数字以外が含まれています: {p}")
        if len(p) != 4:
            raise ValueError(f"手配番号は4桁である必要があります: {p}")
        nums.append(int(p))
    if not nums:
        raise ValueError("手配番号が入力されていません")
    return nums


def run() -> Tuple[str, List[int], str]:
    root = tk.Tk()
    root.title("ロボットテスター")
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    # ESCで終了
    def _esc(_evt=None):
        print("[0.initial_form] ESCが押されました。プログラムを終了します。")
        root.destroy()
        sys.exit(0)
    root.bind("<Escape>", _esc)

    # Top frame for rainbow title
    title_frame = tk.Frame(root)
    title_frame.pack(padx=12, pady=(12, 4))
    title_text = "ロボットテスター"
    colors = _rainbow_colors(len(title_text))
    for ch, color in zip(title_text, colors):
        lbl = tk.Label(title_frame, text=ch, fg=color, font=("Meiryo", 24, "bold"))
        lbl.pack(side=tk.LEFT)

    # Flow name
    tk.Label(root, text="テストするフローの名前をコピペしてください", font=("Meiryo", 16)).pack(anchor="w", padx=12, pady=(8, 2))
    flow_name_var = tk.StringVar()
    flow_name_entry = tk.Entry(root, textvariable=flow_name_var, width=50, font=("Meiryo", 12))
    flow_name_entry.pack(fill="x", padx=12)

    # Tehai numbers
    tk.Label(
        root,
        text=(
            "テストする手配番号を半角スペース区切りで記入してください。すべて半角で入力してください\n"
            "ex. 4785 4886 4897 4667,,,,"
        ),
        font=("Meiryo", 16),
        justify="left",
    ).pack(anchor="w", padx=12, pady=(12, 2))

    tehai_text = scrolledtext.ScrolledText(root, width=60, height=6, wrap=tk.WORD, font=("Meiryo", 12))
    tehai_text.pack(fill="both", expand=False, padx=12)

    result: Tuple[str, List[int], str] = ("", [], "")

    def on_close():
        print("[0.initial_form] ユーザーが閉じるを選択しました。プログラムを終了します。")
        root.destroy()
        sys.exit(0)

    def on_confirm():
        nonlocal result
        try:
            raw_name = flow_name_var.get().strip()
            if not raw_name:
                raise ValueError("フロー名が未入力です")
            suffix = " - Power Automate.url"
            flow_name = raw_name if raw_name.endswith(suffix) else (raw_name + suffix)
            tehai_numbers = _parse_tehai_numbers(tehai_text.get("1.0", tk.END))
            msg = f"{len(tehai_numbers)}個のテストを開始します。"
            try:
                messagebox.showinfo("確認", msg, parent=root)
            except TypeError:
                messagebox.showinfo("確認", msg)
            ts = _dt.datetime.now().strftime("%m%d_%H%M")
            result = (flow_name, tehai_numbers, ts)
            print(f"[0.initial_form] 入力受理 flow_name={flow_name}, 件数={len(tehai_numbers)}, 開始時刻={ts}")
            root.destroy()
        except Exception as e:
            print(f"[0.initial_form] 入力エラー: {e}")
            try:
                messagebox.showerror("入力エラー", str(e), parent=root)
            except TypeError:
                messagebox.showerror("入力エラー", str(e))

    btn_frame = tk.Frame(root)
    btn_frame.pack(fill="x", padx=12, pady=12)
    tk.Button(btn_frame, text="確定", command=on_confirm, width=12, font=("Meiryo", 12)).pack(side=tk.LEFT, expand=True)
    tk.Button(btn_frame, text="閉じる", command=on_close, width=12, font=("Meiryo", 12)).pack(side=tk.RIGHT, expand=True)

    flow_name_entry.focus_set()
    root.mainloop()

    if not result[0]:
        raise SystemExit(0)
    return result


if __name__ == "__main__":  # manual run
    try:
        fn, nums, ts = run()
        print(f"flow_name={fn}")
        print(f"tehai_numbers={nums}")
        print(f"timestamp={ts}")
    except Exception as e:
        print(f"[0.initial_form] 例外: {e}")
        sys.exit(1)
