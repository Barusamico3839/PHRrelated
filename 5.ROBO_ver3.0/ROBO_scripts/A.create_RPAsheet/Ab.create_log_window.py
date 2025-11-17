"""Log window utilities for chouji_ROBO."""

from __future__ import annotations

import logging
import queue
import threading
from typing import Optional

try:
    import tkinter as tk
    from tkinter import ttk
except Exception:  # pragma: no cover - tkinter should exist on Windows
    tk = None  # type: ignore
    ttk = None  # type: ignore

from common import LOG_QUEUE_POLL_MS


class TkLogHandler(logging.Handler):
    """Pushes log records into a Tk text widget."""

    def __init__(
        self,
        text_widget: "tk.Text",
        stop_event: threading.Event,
        poll_interval_ms: int = LOG_QUEUE_POLL_MS,
    ) -> None:
        super().__init__()
        self._text_widget = text_widget
        self._stop_event = stop_event
        self._queue: "queue.Queue[str]" = queue.Queue()
        self._poll_interval_ms = poll_interval_ms
        self._text_widget.after(self._poll_interval_ms, self._drain_queue)

    def emit(self, record: logging.LogRecord) -> None:
        message = self.format(record)
        self._queue.put(message)

    def _drain_queue(self) -> None:
        if self._stop_event.is_set():
            return
        try:
            while True:
                message = self._queue.get_nowait()
                self._text_widget.configure(state="normal")
                self._text_widget.insert("end", message + "\n")
                self._text_widget.see("end")
                self._text_widget.configure(state="disabled")
        except queue.Empty:
            pass
        finally:
            if not self._stop_event.is_set():
                self._text_widget.after(self._poll_interval_ms, self._drain_queue)


class LogWindowManager:
    """Creates and governs the translucent log window."""

    def __init__(self, root: "tk.Tk", stop_event: threading.Event) -> None:
        if tk is None:
            raise RuntimeError("tkinter is required to render the log window.")
        self._root = root
        self._stop_event = stop_event
        self._force_stop_event = threading.Event()
        self._window: Optional["tk.Toplevel"] = None
        self._text_widget: Optional["tk.Text"] = None
        self._handler: Optional[TkLogHandler] = None

    @property
    def handler(self) -> TkLogHandler:
        if self._handler is None:
            raise RuntimeError("Log handler not initialised yet.")
        return self._handler

    @property
    def force_stop_event(self) -> threading.Event:
        return self._force_stop_event

    def create(self) -> None:
        window = tk.Toplevel(self._root)
        window.title("chouji_ROBO ログモニター")
        window.configure(bg="#444444")
        window.attributes("-alpha", 0.6)
        window.attributes("-topmost", 0)
        window.geometry(self._initial_geometry(window))
        window.resizable(False, False)
        window.protocol("WM_DELETE_WINDOW", self._handle_force_close)
        self._window = window

        frame = ttk.Frame(window)
        frame.pack(fill="both", expand=True, padx=4, pady=4)

        button = ttk.Button(frame, text="ロボ強制終了", command=self._handle_force_close)
        button.pack(anchor="ne", padx=2, pady=(2, 4))

        text = tk.Text(
            frame,
            height=20,
            width=60,
            bg="#555555",
            fg="#ffffff",
            font=("Yu Gothic UI", 10),
            relief="flat",
        )
        text.pack(fill="both", expand=True)
        text.configure(state="disabled")
        text.lower()
        self._text_widget = text
        self._handler = TkLogHandler(text, self._stop_event)
        self._schedule_lowering()

    def _initial_geometry(self, window: "tk.Toplevel") -> str:
        window.update_idletasks()
        width = max(int(window.winfo_screenwidth() / 4), 320)
        height = max(int(window.winfo_screenheight() / 4), 200)
        return f"{width}x{height}+0+0"

    def _schedule_lowering(self) -> None:
        if self._window is None or self._stop_event.is_set():
            return
        self._window.lower()
        self._window.after(2000, self._schedule_lowering)

    def _handle_force_close(self) -> None:
        logging.getLogger("chouji_robo.ui").warning("ロボ強制終了ボタンが押されました。シャットダウンします。")
        self._force_stop_event.set()


def create_log_window(root: "tk.Tk", stop_event: threading.Event) -> LogWindowManager:
    manager = LogWindowManager(root, stop_event)
    manager.create()
    return manager
