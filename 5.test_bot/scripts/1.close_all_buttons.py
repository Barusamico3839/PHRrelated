# -*- coding: utf-8 -*-
import time
import sys
import os
from typing import Optional, Dict, Any


def _import_uia() -> Optional[object]:
    try:
        import uiautomation as auto
        return auto
    except Exception as e:
        print(f"[1.close_all_buttons] uiautomation 未利用: {e}")
        return None


def _find_win(auto, name=None, automation_id=None, class_name=None, timeout=3):
    """BFS from UIA root across all top-level windows (multi-monitor safe)."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            root = auto.GetRootControl()
            q = list(getattr(root, 'GetChildren', lambda: [])())
            seen = set()
            while q:
                c = q.pop(0)
                if id(c) in seen:
                    continue
                seen.add(id(c))
                try:
                    if name is not None and (getattr(c, 'Name', None) != name):
                        pass
                    else:
                        if automation_id is not None and (getattr(c, 'AutomationId', None) != automation_id):
                            pass
                        else:
                            if class_name is not None and (getattr(c, 'ClassName', None) != class_name):
                                pass
                            else:
                                return c
                except Exception:
                    pass
                try:
                    for ch in c.GetChildren():
                        if id(ch) not in seen:
                            q.append(ch)
                except Exception:
                    pass
        except Exception:
            pass
        time.sleep(0.2)
    return None

def _dismiss_toast_fast(auto, rounds: int = 2) -> int:
    """Quickly dismiss toast notifications by clicking DismissButton at coordinates.
    Returns the number of clicks performed (up to 3).
    """
    clicked = 0
    if not auto:
        return 0
    if _show_desktop(auto):
        print('[1.close_all_buttons] Minimized open windows with Win+D.')
    else:
        print('[1.close_all_buttons] Could not send Win+D; continuing without minimising.')
    for _ in range(max(1, int(rounds))):
        try:
            root = auto.GetRootControl()
            q = list(getattr(root, 'GetChildren', lambda: [])())
            seen = set()
            while q:
                c = q.pop(0)
                if id(c) in seen:
                    continue
                seen.add(id(c))
                try:
                    aid = getattr(c, 'AutomationId', '') or ''
                    cls = getattr(c, 'ClassName', '') or ''
                    if aid == 'NormalToastView' or cls == 'FlexibleToastView':
                        try:
                            btn = c.ButtonControl(AutomationId='DismissButton')
                        except Exception:
                            btn = None
                        if btn and btn.Exists(0):
                            try:
                                print('burnt toast出現！撃退します')
                            except Exception:
                                pass
                            rect = getattr(btn, 'BoundingRectangle', None)
                            cx = cy = None
                            try:
                                if isinstance(rect, (tuple, list)) and len(rect) >= 4:
                                    l, t, r, b = rect[0], rect[1], rect[2], rect[3]
                                    cx = int((int(l) + int(r)) / 2)
                                    cy = int((int(t) + int(b)) / 2)
                                else:
                                    l = getattr(rect, 'left', None); t = getattr(rect, 'top', None)
                                    r = getattr(rect, 'right', None); b = getattr(rect, 'bottom', None)
                                    if None not in (l, t, r, b):
                                        cx = int((int(l) + int(r)) / 2)
                                        cy = int((int(t) + int(b)) / 2)
                            except Exception:
                                cx = cy = None
                            if cx is not None and cy is not None:
                                try:
                                    import ctypes, time as _t
                                    user32 = ctypes.windll.user32
                                    try:
                                        user32.SetProcessDPIAware()
                                    except Exception:
                                        pass
                                    user32.SetCursorPos(int(cx), int(cy))
                                    user32.mouse_event(0x0002, 0, 0, 0, 0)
                                    _t.sleep(0.015)
                                    user32.mouse_event(0x0004, 0, 0, 0, 0)
                                    clicked += 1
                                    if clicked >= 3:
                                        return clicked
                                    continue
                                except Exception:
                                    try:
                                        import pyautogui
                                        pyautogui.FAILSAFE = False
                                        pyautogui.click(cx, cy)
                                        clicked += 1
                                        if clicked >= 3:
                                            return clicked
                                        continue
                                    except Exception:
                                        pass
                except Exception:
                    pass
                try:
                    for ch in c.GetChildren():
                        if id(ch) not in seen:
                            q.append(ch)
                except Exception:
                    pass
        except Exception:
            pass
    return clicked


def _control_center(ctrl) -> tuple:
    """Return (cx, cy) integer tuple for control center if bounding rect is available."""
    rect = getattr(ctrl, 'BoundingRectangle', None)
    try:
        if isinstance(rect, (tuple, list)) and len(rect) >= 4:
            l, t, r, b = rect[0], rect[1], rect[2], rect[3]
        else:
            l = getattr(rect, 'left', None)
            t = getattr(rect, 'top', None)
            r = getattr(rect, 'right', None)
            b = getattr(rect, 'bottom', None)
        if None in (l, t, r, b):
            return ()
        cx = int((int(l) + int(r)) / 2)
        cy = int((int(t) + int(b)) / 2)
        return (cx, cy)
    except Exception:
        return ()


def _click_center(ctrl) -> bool:
    """Click the center of a control; falls back to UI Automation click."""
    pt = _control_center(ctrl)
    if pt:
        x, y = pt
        try:
            import ctypes, time as _t
            user32 = ctypes.windll.user32
            try:
                user32.SetProcessDPIAware()
            except Exception:
                pass
            user32.SetCursorPos(int(x), int(y))
            user32.mouse_event(0x0002, 0, 0, 0, 0)
            _t.sleep(0.02)
            user32.mouse_event(0x0004, 0, 0, 0, 0)
            return True
        except Exception:
            try:
                import pyautogui
                pyautogui.FAILSAFE = False
                pyautogui.click(int(x), int(y))
                return True
            except Exception:
                pass
    try:
        ctrl.Click()
        return True
    except Exception:
        return False


def _close_window_by_hwnd(hwnd: int) -> bool:
    """Send WM_CLOSE to the specified native window handle."""
    try:
        if not hwnd:
            return False
        import ctypes
        user32 = ctypes.windll.user32
        WM_CLOSE = 0x0010
        user32.PostMessageW(int(hwnd), WM_CLOSE, 0, 0)
        # Allow the target window to process the close request
        time.sleep(0.1)
        return True
    except Exception:
        return False


def _find_window_hwnd(ctrl) -> int:
    """Ascend the UIA tree to locate the owning window handle."""
    try:
        current = ctrl
        steps = 0
        last_hwnd = 0
        window_prefix = 'WindowsForms10.Window'
        while current and steps < 10:
            try:
                hwnd = getattr(current, 'NativeWindowHandle', None) or 0
            except Exception:
                hwnd = 0
            try:
                cls = getattr(current, 'ClassName', '') or ''
            except Exception:
                cls = ''
            if hwnd:
                last_hwnd = int(hwnd)
                if cls.startswith(window_prefix):
                    return last_hwnd
            try:
                current = current.GetParentControl()
            except Exception:
                current = None
            steps += 1
        return last_hwnd
    except Exception:
        return 0


def _force_close_shadan_window_native() -> bool:
    """Last-resort close for 遮断予告/メッセージ通知 windows via Win32 API."""
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        targets = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        def _enum_proc(hwnd, _lparam):
            try:
                if not user32.IsWindowVisible(hwnd):
                    return True
                length = user32.GetWindowTextLengthW(hwnd)
                if length <= 0:
                    return True
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value or ''
                if any(key in title for key in ('遮断予告', 'メッセージ通知')):
                    targets.append(hwnd)
            except Exception:
                pass
            return True

        user32.EnumWindows(_enum_proc, 0)
        for hwnd in targets:
            if _close_window_by_hwnd(hwnd):
                try:
                    print('[1.close_all_buttons] 遮断予告をWM_CLOSEで閉じました。')
                except Exception:
                    pass
                return True
    except Exception:
        pass
    return False


def _click_shadan_notice(auto) -> bool:
    """Find and click the '閉じる' button on 遮断予告 notification windows."""
    if not auto:
        return False

    candidates = []
    try:
        btn = auto.ButtonControl(AutomationId='1903760')
    except Exception:
        btn = None
    if btn:
        candidates.append(btn)
    try:
        by_name = auto.ButtonControl(Name='閉じる')
    except Exception:
        by_name = None
    if by_name and by_name not in candidates:
        candidates.append(by_name)

    for candidate in candidates:
        try:
            if not candidate or not candidate.Exists(0.1):
                continue
            parent = None
            try:
                parent = candidate.GetParentControl()
            except Exception:
                parent = None
            try:
                window_hwnd = _find_window_hwnd(parent or candidate)
            except Exception:
                window_hwnd = 0
            if parent:
                try:
                    parent.SetActive()
                except Exception:
                    pass
            if _click_center(candidate):
                try:
                    print('[1.close_all_buttons] Clicked ShadanYokoku close button.')
                except Exception:
                    pass
                return True
            try:
                invoke = candidate.GetInvokePattern()
                if invoke:
                    invoke.Invoke()
                    print('[1.close_all_buttons] Invoked ShadanYokoku close button.')
                    return True
            except Exception:
                pass
            esc_sent = False
            try:
                if parent:
                    try:
                        parent.SetFocus()
                    except Exception:
                        pass
                auto.SendKeys('{ESC}', interval=0.0, waitTime=0.0)
                time.sleep(0.1)
                esc_sent = True
            except Exception:
                esc_sent = False
            if esc_sent:
                try:
                    print('[1.close_all_buttons] ESCで遮断予告を閉じる試行を行いました。')
                except Exception:
                    pass
                return True
            if window_hwnd and _close_window_by_hwnd(window_hwnd):
                try:
                    print('[1.close_all_buttons] 遮断予告をWM_CLOSEで閉じる試行を行いました。')
                except Exception:
                    pass
                return True
        except Exception:
            pass

    target_name = '閉じる'
    target_aid = '1903760'
    target_class = 'WindowsForms10.BUTTON.app.0.d7ec25_r22_ad1'
    try:
        root = auto.GetRootControl()
    except Exception:
        return False

    queue = list(getattr(root, 'GetChildren', lambda: [])())
    seen = set()
    while queue:
        ctrl = queue.pop(0)
        if id(ctrl) in seen:
            continue
        seen.add(id(ctrl))
        try:
            name = getattr(ctrl, 'Name', '') or ''
            aid = getattr(ctrl, 'AutomationId', '') or ''
            cls = getattr(ctrl, 'ClassName', '') or ''
            if (aid == target_aid or name == target_name) and (not target_class or cls == target_class):
                if not ctrl.Exists(0.1):
                    continue
                parent = None
                try:
                    parent = ctrl.GetParentControl()
                except Exception:
                    parent = None
                try:
                    window_hwnd = _find_window_hwnd(parent or ctrl)
                except Exception:
                    window_hwnd = 0
                if parent:
                    try:
                        parent.SetActive()
                    except Exception:
                        pass
                if _click_center(ctrl):
                    try:
                        print('[1.close_all_buttons] Clicked ShadanYokoku close button (BFS).')
                    except Exception:
                        pass
                    return True
                try:
                    invoke = ctrl.GetInvokePattern()
                    if invoke:
                        invoke.Invoke()
                        print('[1.close_all_buttons] Invoked ShadanYokoku close button (BFS).')
                        return True
                except Exception:
                    pass
                esc_sent = False
                try:
                    if parent:
                        try:
                            parent.SetFocus()
                        except Exception:
                            pass
                    auto.SendKeys('{ESC}', interval=0.0, waitTime=0.0)
                    time.sleep(0.1)
                    esc_sent = True
                except Exception:
                    esc_sent = False
                if esc_sent:
                    try:
                        print('[1.close_all_buttons] ESCで遮断予告を閉じる試行を行いました。(BFS)')
                    except Exception:
                        pass
                    return True
                if window_hwnd and _close_window_by_hwnd(window_hwnd):
                    try:
                        print('[1.close_all_buttons] 遮断予告をWM_CLOSEで閉じる試行を行いました。(BFS)')
                    except Exception:
                        pass
                    return True
        except Exception:
            pass
        try:
            for child in ctrl.GetChildren():
                if id(child) not in seen:
                    queue.append(child)
        except Exception:
            pass
    if _force_close_shadan_window_native():
        return True
    return False


def _show_desktop(auto) -> bool:
    """Try to minimise all windows using Win+D."""
    try:
        if auto:
            auto.SendKeys('{Win}d', interval=0.0, waitTime=0.0)
            time.sleep(0.2)
            return True
    except Exception:
        pass
    try:
        import ctypes, time as _t
        user32 = ctypes.windll.user32
        KEYEVENTF_KEYUP = 0x0002
        VK_LWIN = 0x5B
        VK_D = 0x44
        user32.keybd_event(VK_LWIN, 0, 0, 0)
        user32.keybd_event(VK_D, 0, 0, 0)
        _t.sleep(0.03)
        user32.keybd_event(VK_D, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
        return True
    except Exception:
        return False


def _click_ok_stronger(auto, parent, name: str, automation_id: str) -> bool:
    """Even more aggressive OK click: dismiss toasts, coordinate click, retry.
    """
    try:
        for _ in range(5):
            # Dismiss any toast quickly
            try:
                _dismiss_toast_fast(auto, rounds=2)
            except Exception:
                pass
            # Re-acquire button each try
            btn = None
            try:
                btn = parent.ButtonControl(AutomationId=automation_id)
            except Exception:
                btn = None
            if not btn or not btn.Exists(0):
                try:
                    btn = parent.ButtonControl(Name=name)
                except Exception:
                    btn = None
            try:
                parent.SetActive()
            except Exception:
                pass
            if btn and btn.Exists(0):
                # Prefer coordinate click
                rect = getattr(btn, 'BoundingRectangle', None)
                cx = cy = None
                try:
                    if isinstance(rect, (tuple, list)) and len(rect) >= 4:
                        l, t, r, b = rect[0], rect[1], rect[2], rect[3]
                        cx = int((int(l) + int(r)) / 2)
                        cy = int((int(t) + int(b)) / 2)
                    else:
                        l = getattr(rect, 'left', None); t = getattr(rect, 'top', None)
                        r = getattr(rect, 'right', None); b = getattr(rect, 'bottom', None)
                        if None not in (l, t, r, b):
                            cx = int((int(l) + int(r)) / 2)
                            cy = int((int(t) + int(b)) / 2)
                except Exception:
                    cx = cy = None
                if cx is not None and cy is not None:
                    try:
                        import ctypes, time as _t
                        user32 = ctypes.windll.user32
                        try:
                            user32.SetProcessDPIAware()
                        except Exception:
                            pass
                        user32.SetCursorPos(int(cx), int(cy))
                        user32.mouse_event(0x0002, 0, 0, 0, 0)
                        _t.sleep(0.02)
                        user32.mouse_event(0x0004, 0, 0, 0, 0)
                        print("[1.close_all_buttons] 'OK' ボタンをクリック")
                        return True
                    except Exception:
                        try:
                            import pyautogui
                            pyautogui.FAILSAFE = False
                            pyautogui.click(cx, cy)
                            print("[1.close_all_buttons] 'OK' ボタンをクリック")
                            return True
                        except Exception:
                            pass
                # Try UIA Click
                try:
                    btn.Click()
                    print("[1.close_all_buttons] 'OK' ボタンをクリック")
                    return True
                except Exception:
                    pass
            try:
                time.sleep(0.08)
            except Exception:
                pass
        # final fallback: press Enter a few times
        for _ in range(3):
            try:
                try:
                    parent.SetFocus()
                except Exception:
                    pass
                auto.SendKeys('{ENTER}', interval=0.0, waitTime=0.0)
                print("[1.close_all_buttons] 'OK' ボタンをクリック")
                return True
            except Exception:
                try:
                    time.sleep(0.05)
                except Exception:
                    pass
    except Exception:
        pass
    return False


def _click_ok(auto, parent, name: str, automation_id: str):
    try:
        btn = None
        try:
            btn = parent.ButtonControl(AutomationId=automation_id)
        except Exception:
            pass
        if not btn or not btn.Exists(0):
            try:
                btn = parent.ButtonControl(Name=name)
            except Exception:
                btn = None
        if btn and btn.Exists(0):
            btn.Click()
            print("[1.close_all_buttons] 'OK' ボタンをクリック")
            return True
    except Exception as e:
        print(f"[1.close_all_buttons] OKクリック例外: {e}")
    return False


def _click_ok_robust(auto, parent, name: str, automation_id: str) -> bool:
    """Robust OK click that ignores overlays and prefers coordinate click.
    Tries UIA Click, then coordinate click, then keyboard Enter.
    """
    try:
        btn = None
        try:
            btn = parent.ButtonControl(AutomationId=automation_id)
        except Exception:
            btn = None
        if not btn or not btn.Exists(0):
            try:
                btn = parent.ButtonControl(Name=name)
            except Exception:
                btn = None
        try:
            parent.SetActive()
        except Exception:
            pass
        if btn and btn.Exists(0):
            # Try UIA Click
            try:
                btn.Click()
                print("[1.close_all_buttons] 'OK' ボタンをクリック")
                return True
            except Exception:
                pass
            # Fallback: coordinate click at center of button
            rect = getattr(btn, 'BoundingRectangle', None)
            cx = cy = None
            if isinstance(rect, (tuple, list)) and len(rect) >= 4:
                l, t, r, b = rect[0], rect[1], rect[2], rect[3]
                cx = int((int(l) + int(r)) / 2)
                cy = int((int(t) + int(b)) / 2)
            else:
                l = getattr(rect, 'left', None); t = getattr(rect, 'top', None)
                r = getattr(rect, 'right', None); b = getattr(rect, 'bottom', None)
                if None not in (l, t, r, b):
                    cx = int((int(l) + int(r)) / 2)
                    cy = int((int(t) + int(b)) / 2)
            if cx is not None and cy is not None:
                try:
                    import ctypes, time as _t
                    user32 = ctypes.windll.user32
                    try:
                        user32.SetProcessDPIAware()
                    except Exception:
                        pass
                    user32.SetCursorPos(int(cx), int(cy))
                    user32.mouse_event(0x0002, 0, 0, 0, 0)
                    _t.sleep(0.02)
                    user32.mouse_event(0x0004, 0, 0, 0, 0)
                    print('[1.close_all_buttons] OK clicked')
                    return True
                except Exception:
                    try:
                        import pyautogui
                        pyautogui.FAILSAFE = False
                        pyautogui.click(cx, cy)
                        print('[1.close_all_buttons] OK clicked')
                        return True
                    except Exception:
                        pass
        # Keyboard fallback: send Enter to the dialog
        try:
            try:
                parent.SetFocus()
            except Exception:
                pass
            auto.SendKeys('{ENTER}', interval=0.0, waitTime=0.0)
            print("[1.close_all_buttons] 'OK' ボタンをクリック")
            return True
        except Exception:
            pass
    except Exception as e:
        try:
            print(f"[1.close_all_buttons] OKクリック例外: {e}")
        except Exception:
            pass
    return False


def _get_dialog_message(auto, dlg) -> Optional[str]:
    try:
        doc = None
        try:
            doc = dlg.DocumentControl(AutomationId='rtbMessage')
        except Exception:
            pass
        if not doc or not doc.Exists(0):
            try:
                doc = dlg.DocumentControl()
            except Exception:
                doc = None
        if not doc or not doc.Exists(0):
            return None
        try:
            vp = doc.GetValuePattern()
            if vp:
                return vp.Value
        except Exception:
            pass
        try:
            return doc.Name
        except Exception:
            return None
    except Exception as e:
        print(f"[1.close_all_buttons] ダイアログ本文取得失敗: {e}")
        return None


def _terminate_power_automate() -> None:
    names = [
        "PAD.Desktop.exe",
        "PAD.Console.Host.exe",
        "PowerAutomateDesktop.NativeBridge.exe",
        "PAD.Agent.exe",
    ]
    for n in names:
        try:
            import subprocess
            subprocess.run(["taskkill", "/IM", n, "/T", "/F"], capture_output=True, text=True)
        except Exception:
            pass


def run(tehai_number: int) -> Dict[str, Any]:
    print(f"[1.close_all_buttons] 開始 tehai_number={tehai_number}")
    try:
        src_path = os.path.join(
            os.path.expanduser("~"),
            "Desktop",
            "【全社標準】弔事対応フォルダ",
            "2. RPAブック",
            "RPAブック.xlsx",
        )
        if os.path.exists(src_path):
            os.remove(src_path)
            print(f"[1.close_all_buttons] 削除しました: {src_path}")
    except Exception as e:
        print(f"[1.close_all_buttons] 削除時の例外: {e}")

    auto = _import_uia()
    message_dialog: Optional[str] = None
    _step1_start_ts = time.time()
    _step1_min_end_ts = _step1_start_ts + 60.0
    if not auto:
        print("[1.close_all_buttons] UI操作不可のため次工程へ")
        return {"message_dialog": None, "next": "step2"}

    deadline = time.time() + 600  # 10 min
    while time.time() < deadline:
        try:
            if _click_shadan_notice(auto):
                continue
        except Exception:
            pass
        # 対応するメール → 入力 → OK
        try:
            mail_win = _find_win(auto, name="対応するメール", automation_id="FormInputDialog", class_name="WindowsForms10.Window.8.app.0.6e7d48_r7_ad1", timeout=1)
            if mail_win:
                try:
                    edit = mail_win.EditControl(AutomationId='txtUserInput')
                except Exception:
                    edit = None

                if edit and edit.Exists(0):
                    try:
                        edit.SetValue(str(tehai_number))
                        print('[1.close_all_buttons] 入力欄に手配番号を設定: {}'.format(tehai_number))
                    except Exception:
                        pass
                    # verify and fallback to SendKeys if needed
                    set_ok = False
                    try:
                        vp = edit.GetValuePattern()
                        if vp and vp.Value == str(tehai_number):
                            set_ok = True
                    except Exception:
                        pass
                    if not set_ok:
                        try:
                            try:
                                # Bring the dialog to foreground to avoid stray keystrokes
                                mail_win.SetActive()
                            except Exception:
                                pass
                            edit.SetFocus()
                            # Use uiautomation's explicit modifier syntax and send in one shot
                            auto.SendKeys('{Ctrl}a{Delete}' + str(tehai_number), interval=0.005, waitTime=0.0)
                            # verify again
                            try:
                                vp2 = edit.GetValuePattern()
                                if vp2 and vp2.Value == str(tehai_number):
                                    set_ok = True
                            except Exception:
                                # if ValuePattern not available, assume success
                                set_ok = True
                        except Exception:
                            pass
                    if set_ok:
                        # Stronger OK: dismiss toasts and click by coordinates
                        if not _click_ok_stronger(auto, mail_win, name='OK', automation_id='btnOk'):
                            _click_ok_robust(auto, mail_win, name='OK', automation_id='btnOk')
                    else:
                        print('[1.close_all_buttons] 入力確認ができないため、OKは押しません')

        except Exception:
            pass

        # 実行する範囲 → OK
        try:
            range_win = _find_win(auto, name="実行する範囲", automation_id="FormSelectDialog", class_name="WindowsForms10.Window.8.app.0.6e7d48_r7_ad1", timeout=1)
            if range_win:
                print("[1.close_all_buttons] '実行する範囲' 検出 → OK")
                if not _click_ok_stronger(auto, range_win, name="OK", automation_id="btnOk"):
                    _click_ok_robust(auto, range_win, name="OK", automation_id="btnOk")
        except Exception:
            pass

        # ダイアログ → メッセージ判定
        try:
            dlg = _find_win(auto, name="ダイアログ", automation_id="FormMessageBox", class_name="WindowsForms10.Window.208.app.0.6e7d48_r7_ad1", timeout=1)
            if dlg:
                message_dialog = _get_dialog_message(auto, dlg)
                if message_dialog:
                    print(f"[1.close_all_buttons] ダイアログ: {message_dialog}")
                    if "エラー" in message_dialog:
                        _terminate_power_automate()
                        return {"message_dialog": message_dialog, "next": "step4"}
                    if "メール内容の確認をしてください。" in message_dialog:
                        _terminate_power_automate()
                        # Ensure step1 takes at least 60 seconds before finishing
                        try:
                            remain = _step1_min_end_ts - time.time()
                            if remain > 0:
                                time.sleep(remain)
                        except Exception:
                            pass
                        return {"message_dialog": message_dialog, "next": "step2"}
                if not _click_ok_stronger(auto, dlg, name="OK", automation_id="4392456"):
                    _click_ok_robust(auto, dlg, name="OK", automation_id="4392456")
        except Exception:
            pass

        # フロー完了トースト
        # トースト（通知）が出たら DismissButton を座標優先でクリックして閉じる
        try:
            root = auto.GetRootControl()
            q = list(getattr(root, "GetChildren", lambda: [])())
            seen = set()
            while q:
                c = q.pop(0)
                if id(c) in seen:
                    continue
                seen.add(id(c))
                try:
                    aid = getattr(c, "AutomationId", "") or ""
                    cls = getattr(c, "ClassName", "") or ""
                    if aid == "NormalToastView" or cls == "FlexibleToastView":
                        try:
                            btn = c.ButtonControl(AutomationId="DismissButton")
                        except Exception:
                            btn = None
                        if btn and btn.Exists(0):
                            try:
                                print("burnt toast出現！撃退します")
                            except Exception:
                                pass
                            rect = getattr(btn, "BoundingRectangle", None)
                            cx = cy = None
                            try:
                                if isinstance(rect, (tuple, list)) and len(rect) >= 4:
                                    l, t, r, b = rect[0], rect[1], rect[2], rect[3]
                                    cx = int((int(l) + int(r)) / 2)
                                    cy = int((int(t) + int(b)) / 2)
                                else:
                                    l = getattr(rect, "left", None); t = getattr(rect, "top", None)
                                    r = getattr(rect, "right", None); b = getattr(rect, "bottom", None)
                                    if None not in (l, t, r, b):
                                        cx = int((int(l) + int(r)) / 2)
                                        cy = int((int(t) + int(b)) / 2)
                            except Exception:
                                cx = cy = None
                            clicked = False
                            if cx is not None and cy is not None:
                                try:
                                    import ctypes, time as _t
                                    user32 = ctypes.windll.user32
                                    try:
                                        user32.SetProcessDPIAware()
                                    except Exception:
                                        pass
                                    user32.SetCursorPos(int(cx), int(cy))
                                    user32.mouse_event(0x0002, 0, 0, 0, 0)
                                    _t.sleep(0.02)
                                    user32.mouse_event(0x0004, 0, 0, 0, 0)
                                    clicked = True
                                except Exception:
                                    try:
                                        import pyautogui
                                        pyautogui.FAILSAFE = False
                                        pyautogui.click(cx, cy)
                                        clicked = True
                                    except Exception:
                                        pass
                            if not clicked:
                                try:
                                    btn.Click()
                                except Exception:
                                    pass
                except Exception:
                    pass
                try:
                    for ch in c.GetChildren():
                        if id(ch) not in seen:
                            q.append(ch)
                except Exception:
                    pass
        except Exception:
            pass

        time.sleep(0.5)

    # 10-minute watch finished
    print('[1.close_all_buttons] 10-minute watch finished, no popups detected')
    return {"message_dialog": message_dialog, "next": "error"}


if __name__ == "__main__":
    try:
        run(1234)
    except Exception as e:
        print(f"[1.close_all_buttons] 例外: {e}")
        sys.exit(1)
