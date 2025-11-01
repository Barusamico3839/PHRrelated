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
                        _click_ok(auto, mail_win, name='OK', automation_id='btnOk')
                    else:
                        print('[1.close_all_buttons] 入力確認ができないため、OKは押しません')

        except Exception:
            pass

        # 実行する範囲 → OK
        try:
            range_win = _find_win(auto, name="実行する範囲", automation_id="FormSelectDialog", class_name="WindowsForms10.Window.8.app.0.6e7d48_r7_ad1", timeout=1)
            if range_win:
                print("[1.close_all_buttons] '実行する範囲' 検出 → OK")
                _click_ok(auto, range_win, name="OK", automation_id="btnOk")
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
                        return {"message_dialog": message_dialog, "next": "step2"}
                _click_ok(auto, dlg, name="OK", automation_id="4392456")
        except Exception:
            pass

        # フロー完了トースト
        try:
            root = auto.GetRootControl()
            for w in root.GetChildren():
                try:
                    nm = (getattr(w, 'Name', '') or '')
                    aid = getattr(w, 'AutomationId', '') or ''
                    cls = getattr(w, 'ClassName', '') or ''
                    if aid == 'NormalToastView' or cls == 'FlexibleToastView':
                        if 'Power Automate' in nm and '正常に完了' in nm:
                            print('[1.close_all_buttons] 完了トースト検出')
                            _terminate_power_automate()
                            return {"message_dialog": message_dialog, "next": "step4"}
                except Exception:
                    continue
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
