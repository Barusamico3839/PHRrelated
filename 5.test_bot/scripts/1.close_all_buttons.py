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

    deadline = time.time() + 300  # 5分
    while time.time() < deadline:
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

    # 5分経過
    print('[1.close_all_buttons] 5分待機中にポップアップがありませんでした')
    return {"message_dialog": message_dialog, "next": "error"}


if __name__ == "__main__":
    try:
        run(1234)
    except Exception as e:
        print(f"[1.close_all_buttons] 例外: {e}")
        sys.exit(1)
