"""Step Aa: initial form UI creation."""

from __future__ import annotations

try:
    import tkinter as tk
    from tkinter import messagebox, ttk
except Exception:  # pragma: no cover - tkinter should exist on Windows
    tk = None  # type: ignore
    messagebox = None  # type: ignore
    ttk = None  # type: ignore


def build(robot) -> None:
    if tk is None:
        raise RuntimeError("tkinter is required to run chouji_ROBO.")

    robot.current_phase = "Aa.create_initial_form"
    logger = robot.ui_logger
    logger.info("Aa.create_initial_form を起動しました。")

    form = tk.Toplevel(robot.root)
    form.title("弔事ロボット")
    form.geometry("480x260")
    form.resizable(False, False)
    form.configure(bg="#f7f3ff")
    form.protocol("WM_DELETE_WINDOW", robot._shutdown)

    title_label = tk.Label(
        form,
        text="弔事ロボット",
        font=("HGS創英角ﾎﾟｯﾌﾟ体", 20),
        fg="#111111",
        bg="#f7f3ff",
    )
    title_label.pack(pady=(20, 8))

    prompt_label = tk.Label(
        form,
        text="対応したいメールの管理番号を入力してください",
        font=("游ゴシック Medium", 16),
        fg="#666666",
        bg="#f7f3ff",
    )
    prompt_label.pack(pady=(0, 20))

    input_var = tk.StringVar()
    entry = ttk.Entry(form, textvariable=input_var, font=("游ゴシック Medium", 16), width=10, justify="center")
    entry.pack(pady=(0, 12))
    entry.focus_set()

    button_frame = ttk.Frame(form)
    button_frame.pack(pady=(12, 8))

    confirm_btn = ttk.Button(
        button_frame,
        text="確定",
        command=lambda: robot._handle_tehai_submit(form, input_var.get().strip()),
        width=12,
    )
    confirm_btn.grid(row=0, column=0, padx=6)

    cancel_btn = ttk.Button(button_frame, text="終了", command=robot._shutdown, width=12)
    cancel_btn.grid(row=0, column=1, padx=6)

    form.bind("<Return>", lambda event: confirm_btn.invoke())
