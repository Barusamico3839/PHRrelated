import time
from typing import Optional


def _import_uia():
    try:
        import uiautomation as auto  # type: ignore
        return auto
    except Exception:
        return None


def _find_power_automate_hwnd() -> int:
    auto = _import_uia()
    try:
        if auto:
            root = auto.GetRootControl()
            for w in getattr(root, 'GetChildren', lambda: [])():
                try:
                    if getattr(w, 'ControlTypeName', '') == 'WindowControl':
                        nm = (getattr(w, 'Name', '') or '')
                        aid = (getattr(w, 'AutomationId', '') or '')
                        cls = (getattr(w, 'ClassName', '') or '')
                        if ('Power Automate' in nm) or (aid == 'ConsoleMainWindow') or (cls == 'WinAutomationWindow'):
                            hwnd = int(getattr(w, 'NativeWindowHandle', 0) or 0)
                            if hwnd:
                                return hwnd
                except Exception:
                    continue
    except Exception:
        pass
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        GetWindowTextW = user32.GetWindowTextW
        GetWindowTextLengthW = user32.GetWindowTextLengthW
        GetClassNameW = user32.GetClassNameW
        EnumWindows = user32.EnumWindows
        IsWindowVisible = user32.IsWindowVisible
        GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]  # type: ignore
        GetWindowTextLengthW.argtypes = [wintypes.HWND]  # type: ignore
        GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]  # type: ignore

        targets = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)  # type: ignore
        def _enum_proc(hwnd, _lparam):
            try:
                if not IsWindowVisible(hwnd):
                    return True
                length = GetWindowTextLengthW(hwnd)
                title = ctypes.create_unicode_buffer(length + 1)
                GetWindowTextW(hwnd, title, length + 1)
                clsbuf = ctypes.create_unicode_buffer(256)
                GetClassNameW(hwnd, clsbuf, 256)
                title_s = title.value or ''
                cls_s = clsbuf.value or ''
                if ('Power Automate' in title_s) or (cls_s == 'WinAutomationWindow'):
                    targets.append(hwnd)
            except Exception:
                pass
            return True

        EnumWindows(_enum_proc, 0)
        if targets:
            return int(targets[0])
    except Exception:
        pass
    return 0


def hide_pad_console() -> bool:
    """Hide the Power Automate main console window (non-intrusive)."""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        hwnd = _find_power_automate_hwnd()
        if not hwnd:
            return False
        SW_HIDE = 0
        SWP_NOSIZE = 0x0001
        SWP_NOMOVE = 0x0002
        SWP_NOACTIVATE = 0x0010
        SWP_HIDEWINDOW = 0x0080
        try:
            user32.ShowWindow(hwnd, SW_HIDE)
        except Exception:
            pass
        try:
            user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE | SWP_HIDEWINDOW)
        except Exception:
            pass
        return True
    except Exception:
        return False


def close_pad_console() -> bool:
    """Request close of the Power Automate main console window."""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        hwnd = _find_power_automate_hwnd()
        if not hwnd:
            return False
        try:
            WM_CLOSE = 0x0010
            user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
            return True
        except Exception:
            pass
        # UIA fallback: Alt+F4
        try:
            auto = _import_uia()
            if auto:
                try:
                    user32.SetForegroundWindow(hwnd)
                except Exception:
                    pass
                auto.SendKeys('%{F4}')
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False

