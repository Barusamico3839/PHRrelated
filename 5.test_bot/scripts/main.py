# -*- coding: utf-8 -*-
import os
import sys
import time
import importlib.util
import subprocess
from typing import Optional

# 画像クリックの座標キャリブレーション（仮想スクリーン座標系）
# 毎回一定のズレがあるため、検出座標にこの補正を加えてクリックします。
CALIB_OFFSET_X = 90
CALIB_OFFSET_Y = -124


def _module_from(path: str, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"モジュールをロードできません: {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


def _scripts_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _launch_flow_shortcut(flow_name: str) -> None:
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    candidates = [
        os.path.join(desktop, flow_name + ".lnk"),
        os.path.join(desktop, flow_name),
    ]
    for p in candidates:
        if os.path.exists(p):
            print(f"[main] ショートカット実行: {p}")
            os.startfile(p)  # type: ignore[attr-defined]
            return
    raise FileNotFoundError(f"デスクトップにショートカットが見つかりません: {candidates}")


def _terminate_power_automate() -> None:
    names = [
        "PAD.Desktop.exe",
        "PAD.Console.Host.exe",
        "PowerAutomateDesktop.NativeBridge.exe",
        "PAD.Agent.exe",
    ]
    try:
        for n in names:
            try:
                subprocess.run(["taskkill", "/IM", n, "/T", "/F"], capture_output=True, text=True)
            except Exception:
                pass
        print("[main] Power Automate 関連プロセスを強制終了しました（存在する場合）")
    except Exception as e:
        print(f"[main] プロセス終了時の警告: {e}")


def _import_uia():
    try:
        import uiautomation as auto
        return auto
    except Exception as e:
        print(f"[main] uiautomation 利用不可: {e}")
        return None


def _safe_getattr(obj, name, default=""):
    try:
        return getattr(obj, name, default) or default
    except Exception:
        return default


def _debug_log_candidates(auto) -> None:
    try:
        root = auto.GetRootControl()
        tops = list(getattr(root, 'GetChildren', lambda: [])())
        print(f"[debug] top windows: {len(tops)}")
        for w in tops:
            nm = getattr(w, "Name", "") or ""
            ctn = getattr(w, "ControlTypeName", "") or ""
            aid = getattr(w, "AutomationId", "") or ""
            cls = getattr(w, "ClassName", "") or ""
            print(f"    - {ctn} Name=\"{nm}\" Aid=\"{aid}\" Cls=\"{cls}\"")
    except Exception as e:
        print(f"[debug] candidates error: {e}")

def _resolve_image_path(base_name: str) -> Optional[str]:
    """Find an image file by trying common extensions in scripts dir."""
    exts = [".png", ".jpg", ".jpeg", ".bmp"]
    base = os.path.join(_scripts_dir(), base_name)
    if os.path.isfile(base):
        return base
    for ext in exts:
        p = base + ext
        if os.path.isfile(p):
            return p
    return None


def _click_continue_by_image(timeout: float = 45.0, image_name: str = "zokkou_botton") -> bool:
    start_ts = time.time()
    last_report = -1
    img_path = _resolve_image_path(image_name)
    if not img_path:
        print("[debug] log restored")
    print("[debug] log restored")
    try:
        import pyautogui
        from PIL import Image
    except Exception as e:
        print("[debug] log restored")

    def locate_all_monitors() -> Optional[tuple]:
        # mss を使って全モニタ撮影 → PIL画像でテンプレ一致
        try:
            import mss
            with mss.mss() as sct:
                monitors = sct.monitors[1:]
                for mon in monitors:
                    try:
                        shot = sct.grab(mon)
                        screen_img = Image.frombytes('RGB', shot.size, shot.rgb)
                        box = None
                        try:
                            box = pyautogui.locate(img_path, screen_img, confidence=0.8, grayscale=True)
                        except Exception:
                            box = pyautogui.locate(img_path, screen_img)
                        if box:
                            # 絶対座標（仮想スクリーン）に統一し、中心は四捨五入で算出
                            cx = int(round(mon['left'] + box.left + (box.width / 2.0)))
                            cy = int(round(mon['top'] + box.top + (box.height / 2.0)))
                            return (cx, cy)
                    except Exception:
                        continue
        except Exception:
            # フォールバック: プライマリのみ
            try:
                try:
                    box2 = pyautogui.locateOnScreen(img_path, confidence=0.8, grayscale=True)
                except Exception:
                    box2 = pyautogui.locateOnScreen(img_path)
                if box2:
                    center = pyautogui.center(box2)
                    return (int(center.x), int(center.y))
            except Exception:
                return None
        return None

    while time.time() - start_ts < timeout:
        pos = locate_all_monitors()
        sec = int(time.time() - start_ts)
        if pos:
            try:
                if _move_mouse_and_click(pos[0], pos[1]):
                    print(f"[main] 画像クリック成功: {pos} (t={sec}s)")
                    return True
                else:
                    print(f"[main] 画像クリック失敗: low-levelクリックも失敗 (pos={pos})")
                    return False
            except Exception as e:
                print(f"[main] 画像クリック処理例外: {e} (pos={pos})")
                return False
        if sec != last_report:
            print("[debug] log restored")
        time.sleep(1.0)
    print("[main] 画像による『続行』が見つかりませんでした（タイムアウト）")
    return False


def _move_mouse_and_click_strict(x: int, y: int) -> bool:
    try:
        return _move_mouse_and_click_strict_v2(x, y)
    except Exception as e:
        print(f"[debug] strict click wrapper error: {e}")
        return False


def _move_mouse_and_click_strict_v2(x: int, y: int) -> bool:
    try:
        import ctypes, time as _t
        user32 = ctypes.windll.user32
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass
        user32.SetCursorPos(int(x), int(y))
        MOUSEEVENTF_LEFTDOWN = 0x0002
        MOUSEEVENTF_LEFTUP = 0x0004
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        _t.sleep(0.04)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        return True
    except Exception:
        try:
            import pyautogui, time as _t
            pyautogui.FAILSAFE = False
            pyautogui.moveTo(int(x), int(y), duration=0.35)
            pyautogui.mouseDown(); _t.sleep(0.04); pyautogui.mouseUp()
            return True
        except Exception:
            return False



def _click_continue_by_image_v2(timeout: float = 45.0, image_name: str = "zokkou_botton") -> bool:
    """Image search across all monitors and click until the image disappears.
    Proceeds only when the image vanishes, logs every action.
    """
    start_ts = time.time()
    last_report = -1
    img_path = _resolve_image_path(image_name)
    if not img_path:
        print(f"[debug] 画像が見つかりません: {image_name}(.png/.jpg/.jpeg/.bmp)")
        return False
    print(f"[debug] 画像探索開始: '{img_path}' timeout={timeout}s (全モニター対応)")
    try:
        import pyautogui
        from PIL import Image
    except Exception as e:
        print(f"[debug] 画像探索に必要なライブラリ不足: {e}")
        return False

    def locate_all_monitors() -> Optional[tuple]:
        try:
            import mss
            with mss.mss() as sct:
                for mon in sct.monitors[1:]:
                    try:
                        shot = sct.grab(mon)
                        screen_img = Image.frombytes('RGB', shot.size, shot.rgb)
                        try:
                            box = pyautogui.locate(img_path, screen_img, confidence=0.8, grayscale=True)
                        except Exception:
                            box = pyautogui.locate(img_path, screen_img)
                        if box:
                            cx = mon['left'] + box.left + box.width // 2
                            cy = mon['top'] + box.top + box.height // 2
                            return (cx, cy)
                    except Exception:
                        continue
        except Exception:
            try:
                try:
                    box2 = pyautogui.locateOnScreen(img_path, confidence=0.8, grayscale=True)
                except Exception:
                    box2 = pyautogui.locateOnScreen(img_path)
                if box2:
                    center = pyautogui.center(box2)
                    return (center.x, center.y)
            except Exception:
                return None
        return None

    while time.time() - start_ts < timeout:
        pos = locate_all_monitors()
        sec = int(time.time() - start_ts)
        if pos:
            try:
                tx = int(pos[0])
                ty = int(pos[1])
                if _move_mouse_and_click_strict_v2(tx, ty):
                    print(f"[main] 画像クリック: detect=({pos[0]},{pos[1]}) -> click=({tx},{ty}) (t={sec}s)")
                    vanish_deadline = time.time() + 8.0
                    while time.time() < vanish_deadline:
                        time.sleep(0.4)
                        if not locate_all_monitors():
                            print("[debug] '続行'画像が消えました")
                            return True
                    print("[debug] '続行'画像が残存→再クリック試行")
                else:
                    print(f"[main] 画像クリック失敗: low-levelクリック失敗 (detect={pos})")
            except Exception as e:
                print(f"[main] 画像クリック例外: {e} (detect={pos})")
        if sec != last_report:
            print(f"[debug] 画像探索 t={sec}s 経過 画像='{os.path.basename(img_path)}'")
            last_report = sec
        time.sleep(1.0)
    print("[main] 画像による続行押下に失敗（タイムアウト）")
    return False
_click_continue_by_image = _click_continue_by_image_v2
def _move_mouse_and_click(x: int, y: int) -> bool:
    """Move mouse gradually to (x,y) and click with fallbacks.
    Returns True if any click path reports success.
    """
    # 最優先: 低レベルAPI（仮想スクリーン座標に対応・負座標もOK）
    try:
        import ctypes
        user32 = ctypes.windll.user32
        # DPI aware にして座標ズレを減らす
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass
        # SetCursorPos は仮想スクリーン座標を受け付ける（負座標可）
        if user32.SetCursorPos(int(x), int(y)) == 0:
            print("[debug] log restored")
        MOUSEEVENTF_LEFTDOWN = 0x0002
        MOUSEEVENTF_LEFTUP = 0x0004
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.02)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        return True
    except Exception as e:
        print("[debug] log restored")
        try:
            import pyautogui
            pyautogui.FAILSAFE = False
            cur_x, cur_y = (0, 0)
            try:
                cur_x, cur_y = pyautogui.position()
            except Exception:
                pass
            print("[debug] log restored")
            pyautogui.mouseDown(); time.sleep(0.04); pyautogui.mouseUp()
            return True
        except Exception as e2:
            print("[debug] log restored")


def _click_runflow_continue(timeout: float = 45.0) -> bool:
    auto = _import_uia()
    print("[debug] 実行チェック開始 UIAあり={}".format(bool(auto)))
    if not auto:
        start_ts = time.time()
        last_report = -1
        while time.time() - start_ts < timeout:
            sec = int(time.time() - start_ts)
            if sec != last_report:
                print("[debug] 実行チェック t={}s diag_found=False btn_found=False (uiautomation無)".format(sec))
                last_report = sec
            time.sleep(1.0)
        print("[main] '続行' ボタンが見つかりませんでした（UIAなし/タイムアウト）")
        return False
    start_ts = time.time()
    deadline = start_ts + timeout
    printed_10s = False
    printed_30s = False
    last_report_sec = -1
    while time.time() < deadline:
        diag_found = False
        btn_found = False
        try:
            root = auto.GetRootControl()
            # 以降の詳細検索ロジックは既存のまま
        except Exception:
            pass
        # 進捗ログ
        sec = int(time.time() - start_ts)
        if sec != last_report_sec:
            print("[debug] 実行チェック t={}s diag_found={} btn_found={}".format(sec, diag_found, btn_found))
            last_report_sec = sec
        time.sleep(0.3)
    print("[main] '続行' ボタンが見つかりませんでした（タイムアウト）")
    return False
def _minimize_power_automate_window():
    auto = _import_uia()
    if not auto:
        return
    try:
        root = auto.GetRootControl()
        for w in root.GetChildren():
            try:
                if getattr(w, 'ControlTypeName', '') == 'WindowControl':
                    nm = (getattr(w, 'Name', '') or '')
                    if 'Power Automate' in nm:
                        try:
                            w.SetFocus()
                            auto.SendKeys('%{SPACE}')
                            time.sleep(0.2)
                            auto.SendKeys('n')
                            print('[main] Power Automate ウインドウを最小化しました')
                            return
                        except Exception:
                            pass
            except Exception:
                continue
    except Exception:
        pass


def _find_power_automate_hwnd() -> int:
    """Locate the Power Automate main console window handle.
    Returns hwnd (>0) if found, else 0.
    """
    # Try via UI Automation first
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
    # Fallback to Win32 APIs
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        GetWindowTextW = user32.GetWindowTextW
        GetWindowTextLengthW = user32.GetWindowTextLengthW
        GetClassNameW = user32.GetClassNameW
        EnumWindows = user32.EnumWindows
        IsWindowVisible = user32.IsWindowVisible
        GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        GetWindowTextLengthW.argtypes = [wintypes.HWND]
        GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

        targets = []
        @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        def _enum_proc(hwnd, lparam):
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


def _close_power_automate_console(mode: str = 'hide') -> bool:
    """Hide or close the Power Automate main console window.
    mode: 'hide' to hide (non-intrusive), 'close' to request close.
    Returns True if any action was taken.
    """
    try:
        import ctypes, time as _t
        user32 = ctypes.windll.user32
        hwnd = _find_power_automate_hwnd()
        if not hwnd:
            print("[debug] log restored")
        if mode == 'hide':
            SW_HIDE = 0
            # Hide via ShowWindow, then ensure hidden via SetWindowPos
            try:
                user32.ShowWindow(hwnd, SW_HIDE)
                SWP_NOSIZE=0x0001; SWP_NOMOVE=0x0002; SWP_NOACTIVATE=0x0010; SWP_HIDEWINDOW=0x0080
                user32.SetWindowPos(hwnd, 0, 0,0,0,0, SWP_NOSIZE|SWP_NOMOVE|SWP_NOACTIVATE|SWP_HIDEWINDOW)
                print('[main] Power Automate コンソールを非表示にしました')
                return True
            except Exception:
                pass
        # close by posting WM_CLOSE or Alt+F4
        try:
            WM_CLOSE = 0x0010
            user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
            print('[main] Power Automate コンソールに WM_CLOSE を送信しました')
            _t.sleep(0.2)
            return True
        except Exception:
            pass
        # UIA fallback: focus and Alt+F4
        try:
            auto = _import_uia()
            if auto:
                # Set focus to window by clicking title bar approx (safe)
                try:
                    user32.SetForegroundWindow(hwnd)
                except Exception:
                    pass
                auto.SendKeys('%{F4}')
                print('[main] Power Automate コンソールを Alt+F4 で閉じる指示を送りました')
                return True
        except Exception:
            pass
    except Exception as e:
        print("[debug] log restored")

def _rainbow_colors(n: int):
    import colorsys
    return [
        "#%02x%02x%02x" % tuple(int(c * 255) for c in colorsys.hsv_to_rgb(i / max(n, 1), 0.75, 0.95))
        for i in range(n)
    ]


def _show_finished_banner():
    try:
        import tkinter as tk
    except Exception as e:
        print(f"[main] Tkinter が利用できないため完了バナーをスキップ: {e}")
        return

    root = tk.Tk()
    root.title("完了")
    frame = tk.Frame(root)
    frame.pack(padx=16, pady=16)
    text = "テストが終わりました！！"
    colors = _rainbow_colors(len(text))
    for ch, color in zip(text, colors):
        tk.Label(frame, text=ch, fg=color, font=("Meiryo", 24, "bold")).pack(side=tk.LEFT)
    btn = tk.Button(root, text="OK", font=("Meiryo", 12), command=root.destroy)
    btn.pack(pady=(12, 8))
    try:
        root.attributes("-topmost", True)
        root.lift()
        root.focus_force()
        root.bind("<Escape>", lambda _e=None: root.destroy())
    except Exception:
        pass
    root.mainloop()


def _show_failed_banner():
    try:
        import tkinter as tk
    except Exception as e:
        print("ごめんなさい、、テストに失敗しました、、")
        print(f"[main] Tkinter が利用できないためエラーバナーをスキップ: {e}")
        return

    root = tk.Tk()
    root.title("エラー")
    frame = tk.Frame(root)
    frame.pack(padx=16, pady=16)
    msg = "ごめんなさい、、テストに失敗しました、、"
    tk.Label(frame, text=msg, fg="#0000FF", font=("Meiryo", 30, "bold")).pack()
    btn = tk.Button(root, text="OK", font=("Meiryo", 12), command=root.destroy)
    btn.pack(pady=(12, 8))
    try:
        root.attributes("-topmost", True)
        root.lift()
        root.focus_force()
        root.bind("<Escape>", lambda _e=None: root.destroy())
    except Exception:
        pass
    root.mainloop()


class _LogWindow:
    def __init__(self):
        try:
            import tkinter as tk
        except Exception as e:
            self._tk = None
            print(f"[main] LogWindow不可: {e}")
            return
        self._tk = tk
        self.root = tk.Tk()
        self.root.title("ログ")
        try:
            # 通常のウインドウ（オーバーレイではない）
            self.root.attributes("-topmost", False)
            # 画面の1/2（左上）。プライマリモニター基準
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            ww = max(520, int(sw / 2))
            wh = max(320, int(sh / 2))
            self.root.geometry(f"{ww}x{wh}+0+0")
            # 最背面へ送る（作成直後）
            try:
                self.root.lower()
            except Exception:
                pass
        except Exception:
            pass
        bg = "#202020"
        self.root.configure(bg=bg)
        try:
            self.root.attributes("-alpha", 0.7)  # 透過率30%（不透明度70%）
        except Exception:
            pass
        self.text = tk.Text(self.root, width=120, height=40, bg=bg, fg="#FFFFFF", bd=0, highlightthickness=0)
        self.text.configure(font=("Meiryo", 10))
        self.text.place(x=16, y=16, anchor="nw")
        def _esc(_evt=None):
            print("[main] ESCが押されました。プログラムを終了します。")
            try:
                _shutdown_logging_overlay()
            except Exception:
                pass
            sys.exit(0)
        try:
            self.root.bind("<Escape>", _esc)
        except Exception:
            pass
        # ヘッダ＋本文（Text）構成に変更。ボタンは同一ウインドウ内に配置。
        try:
            header = tk.Frame(self.root, bg=bg)
            header.pack(fill="x", side=tk.TOP)
            btn = tk.Button(header, text="テスターを強制終了", font=("Meiryo", 10, "bold"),
                            command=self._force_exit)
            btn.pack(side=tk.LEFT, padx=8, pady=6)
        except Exception as e:
            print("[debug] log restored")
        # 本文テキスト
        content = tk.Frame(self.root, bg=bg)
        content.pack(fill="both", expand=True)
        try:
            import tkinter as tk
            self.text = tk.Text(content, width=120, height=40, bg=bg, fg="#FFFFFF", bd=0, highlightthickness=0)
            self.text.configure(font=("Meiryo", 10))
            self.text.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        except Exception:
            pass

        # 初期描画
        self._last_refresh = 0.0
        self.root.update_idletasks(); self.root.update()
        # 更新で前面に出てしまうことを防ぐため都度背面へ
        try:
            self.root.lower()
        except Exception:
            pass

    def write(self, s: str):
        if not getattr(self, 'root', None):
            return
        try:
            self.text.insert('end', s)
            self.text.see('end')
        except Exception:
            pass
        # 更新のスロットリング（最大約30fps）
        try:
            import time as _t
            now = _t.time()
            if now - getattr(self, '_last_refresh', 0.0) >= 0.033:
                self.root.update_idletasks(); self.root.update()
                # 表示更新後も背面へ保持
                try:
                    self.root.lower()
                except Exception:
                    pass
                self._last_refresh = now
        except Exception:
            pass

    def _force_exit(self):
        print("[main] 'テスターを強制終了' が押されました。プログラムを終了します。")
        try:
            _shutdown_logging_overlay()
        except Exception:
            pass
        import os; os._exit(0)

    def flush(self):
        pass


class _StdoutTee:
    def __init__(self, base, logwin: _LogWindow):
        self.base = base
        self.logwin = logwin
    def write(self, s):
        try:
            self.base.write(s)
        except Exception:
            pass
        try:
            self.logwin.write(s)
        except Exception:
            pass
    def flush(self):
        try:
            self.base.flush()
        except Exception:
            pass


def _shutdown_logging_overlay():
    try:
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
    except Exception:
        pass
    try:
        gw = globals().get("_GLOBAL_LOGWIN")
        if gw and getattr(gw, 'root', None):
            try:
                gw.root.destroy()
            except Exception:
                pass
        # 追加: 制御用ウインドウも破棄
        try:
            if gw and getattr(gw, 'ctrl', None):
                gw.ctrl.destroy()
        except Exception:
            pass
        globals().pop("_GLOBAL_LOGWIN", None)
    except Exception:
        pass


def main():
    sdir = _scripts_dir()
    initial_form = _module_from(os.path.join(sdir, "0.initial_form.py"), "initial_form")
    close_all = _module_from(os.path.join(sdir, "1.close_all_buttons.py"), "close_all")
    get_pad = _module_from(os.path.join(sdir, "2.get_pad_result.py"), "get_pad")
    get_mail = _module_from(os.path.join(sdir, "3.get_mail_result.py"), "get_mail")
    _ = _module_from(os.path.join(sdir, "4.save_results.py"), "save_results")
    compare = _module_from(os.path.join(sdir, "5.compare_results.py"), "compare")
    next_case = _module_from(os.path.join(sdir, "6.start_next_case.py"), "next_case")

    try:
        flow_name, tehai_numbers, timestamp = initial_form.run()
    except Exception as e:
        print(f"[main] 初期入力エラー: {e}")
        _show_failed_banner()
        sys.exit(1)

    logwin = _LogWindow()
    try:
        globals()["_GLOBAL_LOGWIN"] = logwin
    except Exception:
        pass
    sys.stdout = _StdoutTee(sys.__stdout__, logwin)
    sys.stderr = _StdoutTee(sys.__stderr__, logwin)
    print(f"[main] テスト開始: flow='{flow_name}', 件数={len(tehai_numbers)}, ts={timestamp}")

    i = 0
    total = len(tehai_numbers)
    while i < total:
        tehai_number = tehai_numbers[i]
        print(f"[main] ===== ケース {i+1}/{total} 手配番号={tehai_number} =====")
        # ステップ1開始前に、既存のPower Automate関連プロセスを強制終了
        try:
            _terminate_power_automate()
        except Exception as e:
            print(f"[main] 強制終了処理の警告: {e}")
        try:
            _launch_flow_shortcut(flow_name)
        except Exception as e:
            print(f"[main] ショートカット実行エラー: {e}")
            _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)

        # 起動直後に『続行』を画像で探索してクリック（UIAは使わない）
        try:
            print("[debug] ショートカット起動後 → 画像探索で続行ボタンを探します")
            if _click_continue_by_image(timeout=45.0, image_name="zokkou_botton"):
                try:
                    if not _close_power_automate_console(mode='hide'):
                        _minimize_power_automate_window()
                except Exception:
                    _minimize_power_automate_window()
            else:
                print("[main] 続行ボタンが消えないため終了します")
                _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)
        except Exception as e:
            print(f"[main] 続行ボタン処理中の例外: {e}")
            _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)
            _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)

        try:
            step1 = close_all.run(tehai_number)
        except Exception as e:
            print(f"[main] ステップ1エラー: {e}")
            _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)

        # 分岐: ステップ1の結果に応じて次へ
        try:
            next_action = None
            if isinstance(step1, dict):
                next_action = step1.get('next')
            if next_action == 'step4':
                try:
                    import importlib
                    save_mod = importlib.import_module('save_results')
                    if hasattr(save_mod, 'run'):
                        print('[main] ステップ4へ: save_results を実行します')
                        save_mod.run()
                except Exception as e:
                    print(f"[main] save_results 実行時の例外: {e}")
                # 次のケースへスキップ
                try:
                    msg = None
                    if isinstance(step1, dict):
                        msg = step1.get('message_dialog')
                    i = next_case.run(i, total, timestamp, tehai_number, msg or "")
                except Exception as e:
                    print(f"[main] ステップ6エラー: {e}")
                    _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)
                continue
            elif next_action == 'error':
                print('[main] ステップ1でエラー条件が満たされました。終了します。')
                _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)
            # next_action が step2 または None の場合は通常どおり続行
        except Exception:
            pass

        try:
            sheet_name = get_pad.run(tehai_number, timestamp)
        except Exception as e:
            print(f"[main] ステップ2エラー: {e}")
            _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)

        try:
            get_mail.run(tehai_number, timestamp, sheet_name)
        except Exception as e:
            print(f"[main] ステップ3エラー: {e}")
            _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)

        try:
            compare.run()
        except Exception as e:
            print(f"[main] ステップ5エラー: {e}")
            _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)

        try:
            message_dialog = None
            try:
                if isinstance(step1, dict):
                    message_dialog = step1.get("message_dialog")
            except Exception:
                message_dialog = None
            i = next_case.run(i, total, timestamp, tehai_number, message_dialog or "")
        except Exception as e:
            print(f"[main] ステップ6エラー: {e}")
            _shutdown_logging_overlay(); _show_failed_banner(); sys.exit(1)

    print("[main] 全ケース処理終了")
    try:
        _shutdown_logging_overlay()
    except Exception:
        pass
    _show_finished_banner()
    sys.exit(0)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[main] 致命的エラー: {e}")
        _show_failed_banner()
        sys.exit(1)




