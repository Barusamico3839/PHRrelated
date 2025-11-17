#!/usr/bin/env python3
"""Display the Outlook draft and generated PDF, prompting the user for confirmation."""

from __future__ import annotations

import ctypes
import logging
import os
import subprocess
from pathlib import Path
from typing import Literal, Optional

import pythoncom
import win32com.client  # type: ignore
import win32con  # type: ignore
import win32gui  # type: ignore
import win32process  # type: ignore
import tkinter as tk
from tkinter import messagebox


LOGGER = logging.getLogger("chouji_robo.email")


def _show_main_dialog(robot) -> Literal["next", "redo"]:
    result = tk.StringVar(value="next")

    window = tk.Toplevel(robot.root)
    window.title("メール確認")
    window.geometry("520x200")
    window.attributes("-topmost", True)

    label = tk.Label(
        window,
        text="メール作成が完了しました！内容に間違いがなければ手動でメールの送信を行ってください。\n"
        "添付ファイルを変更したい場合は「PDF変更」のボタンを押してください",
        font=("游ゴシック", 14),
        wraplength=480,
        justify="left",
    )
    label.pack(padx=20, pady=20)

    button_frame = tk.Frame(window)
    button_frame.pack(pady=10)

    def _set_and_close(value: str) -> None:
        result.set(value)
        window.destroy()

    next_btn = tk.Button(button_frame, text="次に進む", width=14, command=lambda: _set_and_close("next"))
    next_btn.grid(row=0, column=0, padx=8)

    redo_btn = tk.Button(button_frame, text="PDF変更", width=14, command=lambda: _set_and_close("redo"))
    redo_btn.grid(row=0, column=1, padx=8)

    window.grab_set()
    window.wait_variable(result)
    return result.get()  # type: ignore[return-value]


def _find_edge_executable() -> Optional[Path]:
    candidates = []
    program_files_x86 = os.environ.get("PROGRAMFILES(X86)")
    program_files = os.environ.get("PROGRAMFILES")
    local_app_data = os.environ.get("LOCALAPPDATA")
    if program_files_x86:
        candidates.append(Path(program_files_x86) / "Microsoft" / "Edge" / "Application" / "msedge.exe")
    if program_files:
        candidates.append(Path(program_files) / "Microsoft" / "Edge" / "Application" / "msedge.exe")
    if local_app_data:
        candidates.append(Path(local_app_data) / "Microsoft" / "Edge" / "Application" / "msedge.exe")
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _launch_pdf_with_edge(pdf_path: Path, robot) -> None:
    edge = _find_edge_executable()
    if not edge:
        LOGGER.warning("Microsoft Edge の実行ファイルが見つからないため、既定アプリで PDF を開きます。")
        os.startfile(pdf_path)  # type: ignore[attr-defined]
        robot.state.edge_process_pid = None
        return
    try:
        proc = subprocess.Popen([str(edge), str(pdf_path)], creationflags=subprocess.CREATE_NO_WINDOW)
        robot.state.edge_process_pid = proc.pid
    except Exception:
        robot.state.edge_process_pid = None
        LOGGER.exception("Edge で PDF を開けませんでした。既定アプリを試します。")
        os.startfile(pdf_path)  # type: ignore[attr-defined]


def _close_edge(robot) -> None:
    pid = robot.state.edge_process_pid
    if pid:
        subprocess.run(["taskkill", "/PID", str(pid), "/F", "/T"], check=False, capture_output=True)
        robot.state.edge_process_pid = None
    else:
        subprocess.run(["taskkill", "/IM", "msedge.exe", "/F", "/T"], check=False, capture_output=True)


def _set_process_window_rect(pid: int, left: int, top: int, width: int, height: int) -> None:
    target_hwnd = None

    def _callback(hwnd, extra):
        nonlocal target_hwnd
        if not win32gui.IsWindowVisible(hwnd):
            return True
        _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
        if window_pid == pid:
            target_hwnd = hwnd
            return False
        return True

    win32gui.EnumWindows(_callback, None)
    if target_hwnd:
        win32gui.SetWindowPos(
            target_hwnd,
            None,
            left,
            top,
            width,
            height,
            win32con.SWP_NOZORDER | win32con.SWP_SHOWWINDOW,
        )


def _screen_size() -> tuple[int, int]:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass
    user32 = ctypes.windll.user32
    return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)


def _position_outlook_window(draft) -> None:
    try:
        inspector = draft.GetInspector()
        hwnd = inspector.WindowHandle
    except Exception:
        return
    width, height = _screen_size()
    win32gui.SetWindowPos(
        hwnd,
        None,
        0,
        0,
        width // 2,
        height,
        win32con.SWP_NOZORDER | win32con.SWP_SHOWWINDOW,
    )


def _position_edge_window(robot) -> None:
    pid = robot.state.edge_process_pid
    if not pid:
        return
    width, height = _screen_size()
    _set_process_window_rect(pid, width // 2, 0, width // 2, height)


def _prompt_pdf_edit(robot) -> None:
    LOGGER.info("PDF変更が選択されたため、弔事連絡票シートを表示します。")
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
    except Exception:
        pass

    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        workbook = excel.Workbooks.Open(str(robot.paths.rpa_book_destination))
        try:
            sheet = workbook.Worksheets("弔事連絡票")
        except Exception:
            sheet = None
        excel.Visible = True
        if sheet is not None:
            try:
                sheet.Activate()
            except Exception:
                pass

        prompt = tk.Toplevel(robot.root)
        prompt.title("PDF変更")
        prompt.geometry("380x160")
        prompt.attributes("-topmost", True)

        label = tk.Label(
            prompt,
            text="内容を修正してください。\n修正が完了したら「次に進む」を押してください。",
            font=("游ゴシック", 14),
            wraplength=340,
            justify="left",
        )
        label.pack(padx=20, pady=20)

        done = tk.BooleanVar(value=False)

        def _finish() -> None:
            done.set(True)
            prompt.destroy()

        button = tk.Button(prompt, text="次に進む", width=14, command=_finish)
        button.pack(pady=10)

        prompt.grab_set()
        prompt.wait_variable(done)

        workbook.Save()
        workbook.Close(SaveChanges=True)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def run(robot) -> Literal["next", "redo"]:
    robot.current_phase = "Ec.show_mail_and_PDF"
    LOGGER.info("Ec.show_mail_and_PDF: メールとPDFを表示します。")

    entry_id = robot.state.outlook_draft_entry_id
    if not entry_id:
        raise RuntimeError("Outlook 下書きが見つかりません。")

    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
    except Exception:
        pass

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        session = outlook.GetNamespace("MAPI")
        store_id = robot.state.outlook_draft_store_id or None
        if store_id:
            draft = session.GetItemFromID(entry_id, store_id)
        else:
            draft = session.GetItemFromID(entry_id)
        draft.Display()
        _position_outlook_window(draft)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    pdf_path = robot.state.generated_pdf_path
    if pdf_path and Path(pdf_path).exists():
        try:
            _launch_pdf_with_edge(Path(pdf_path), robot)
            _position_edge_window(robot)
        except OSError:
            LOGGER.warning("PDF %s を既定アプリで開けませんでした。", pdf_path, exc_info=True)
    else:
        LOGGER.warning("PDF パスが無効のため閲覧をスキップします: %s", pdf_path)

    action = _show_main_dialog(robot)
    if action == "redo":
        _close_edge(robot)
        _prompt_pdf_edit(robot)
        return "redo"
    return "next"
