# -*- coding: utf-8 -*-
import os
import sys
import time
from typing import Optional


def _import_openpyxl():
    try:
        import openpyxl  # type: ignore
        return openpyxl
    except Exception as e:
        print(f"[2.get_pad_result] openpyxl 未インストール: {e}")
        return None


def _import_uia() -> Optional[object]:
    try:
        import uiautomation as auto
        return auto
    except Exception as e:
        print(f"[2.get_pad_result] uiautomation 未インストールまたは利用不可: {e}")
        return None


def _ensure_results_book(dst_path: str):
    oxl = _import_openpyxl()
    if not oxl:
        raise RuntimeError("openpyxl が必要です")
    if os.path.exists(dst_path):
        return
    wb = oxl.Workbook()
    # 少なくとも1枚はシートを残す（空ブック保存対策）
    ws = wb.active
    ws.title = "Sheet1"
    wb.save(dst_path)
    print(f"[2.get_pad_result] 作成 {dst_path}")


def _copy_sheet_values(src_path: str, src_sheet: str, dst_path: str, new_sheet_title: str) -> str:
    oxl = _import_openpyxl()
    if not oxl:
        raise RuntimeError("openpyxl が必要です")
    if not os.path.exists(src_path):
        raise FileNotFoundError(f"ソースファイルが見つかりません: {src_path}")
    _ensure_results_book(dst_path)

    src_wb = oxl.load_workbook(src_path, data_only=True)
    if src_sheet not in src_wb.sheetnames:
        raise ValueError(f"シートが見つかりません: {src_sheet}")
    s = src_wb[src_sheet]

    dst_wb = oxl.load_workbook(dst_path)
    # シートが1枚も無い場合に備える（念のため）
    if not dst_wb.sheetnames:
        dst_wb.create_sheet("Sheet1")
    title = str(new_sheet_title)
    base_title = title
    suffix = 1
    while title in dst_wb.sheetnames:
        suffix += 1
        title = f"{base_title}_{suffix}"
    d = dst_wb.create_sheet(title)

    for row in s.iter_rows():
        for cell in row:
            d.cell(row=cell.row, column=cell.column, value=cell.value)

    dst_wb.save(dst_path)
    print(f"[2.get_pad_result] '{src_sheet}' を '{dst_path}' のシート '{title}' にコピーしました")
    return title


def _try_click_buttons():
    auto = _import_uia()
    if not auto:
        print("[2.get_pad_result] UI操作スキップ")
        return

    def find_button_by(name=None, automation_id=None):
        try:
            root = auto.WindowControl()
            for c in root.GetChildren():
                try:
                    btn = None
                    if automation_id:
                        btn = c.ButtonControl(AutomationId=automation_id)
                        if btn and btn.Exists(0):
                            return btn
                    if name:
                        btn = c.ButtonControl(Name=name)
                        if btn and btn.Exists(0):
                            return btn
                except Exception:
                    continue
        except Exception:
            pass
        return None

    btn = find_button_by(name="はい", automation_id="3672912")
    if btn:
        try:
            btn.Click()
            print("[2.get_pad_result] 'はい' をクリック")
        except Exception as e:
            print(f"[2.get_pad_result] 'はい' クリック失敗: {e}")

    deadline = time.time() + 6
    while time.time() < deadline:
        ok = find_button_by(name="OK", automation_id="8590578")
        if ok:
            try:
                ok.Click()
                print("[2.get_pad_result] 'OK' をクリック")
                time.sleep(0.4)
                continue
            except Exception as e:
                print(f"[2.get_pad_result] 'OK' クリック失敗: {e}")
                break
        time.sleep(0.5)


def run(tehai_number: int, timestamp: str) -> str:
    print(f"[2.get_pad_result] 開始 tehai_number={tehai_number}, ts={timestamp}")
    src_path = os.path.join(
        os.path.expanduser("~"),
        "Desktop",
        "【全社標準】弔事対応フォルダ",
        "2. RPAブック",
        "RPAブック.xlsx",
    )
    dst_dir = os.path.join(
        os.path.expanduser("~"),
        "Desktop",
        "【全社標準】弔事対応フォルダ",
        "5.test_bot",
    )
    os.makedirs(dst_dir, exist_ok=True)
    dst_path = os.path.join(dst_dir, f"results_{timestamp}.xlsx")

    try:
        new_sheet = _copy_sheet_values(src_path, "RPAシート", dst_path, str(tehai_number))
    except Exception as e:
        print(f"[2.get_pad_result] Excel処理エラー: {e}")
        raise

    _try_click_buttons()
    print("[2.get_pad_result] 完了")
    return new_sheet


if __name__ == "__main__":
    try:
        run(1234, "0101_1234")
    except Exception as e:
        print(f"[2.get_pad_result] 例外: {e}")
        sys.exit(1)
