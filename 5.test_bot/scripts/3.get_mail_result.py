# -*- coding: utf-8 -*-
import os
import sys
from typing import Optional
from difflib import SequenceMatcher


def _import_openpyxl():
    try:
        import openpyxl  # type: ignore
        return openpyxl
    except Exception as e:
        print(f"[3.get_mail_result] openpyxl 未インストール: {e}")
        return None


def _import_win32():
    try:
        import win32com.client  # type: ignore
        return win32com.client
    except Exception as e:
        print(f"[3.get_mail_result] pywin32 未インストール: {e}")
        return None


def _best_mail_match(sent_items, keyword: str, target: str):
    best = None
    best_score = -1.0
    count_checked = 0
    for item in sent_items:
        try:
            subj = str(getattr(item, 'Subject', '') or '')
            body = str(getattr(item, 'Body', '') or '')
            if keyword not in subj and keyword not in body:
                continue
            s1 = SequenceMatcher(None, subj, target).ratio()
            s2 = SequenceMatcher(None, body, target).ratio()
            score = max(s1, s2)
            if score > best_score:
                best = item
                best_score = score
            count_checked += 1
        except Exception:
            continue
    print(f"[3.get_mail_result] 候補メール数={count_checked}, ベストスコア={best_score:.3f}")
    return best


def run(tehai_number: int, timestamp: str, sheet_name: Optional[str] = None) -> None:
    print(f"[3.get_mail_result] 開始 tehai_number={tehai_number}, ts={timestamp}")
    oxl = _import_openpyxl()
    if not oxl:
        raise RuntimeError("openpyxl が必要です")
    w32 = _import_win32()
    if not w32:
        print("[3.get_mail_result] Outlook操作をスキップします (pywin32 未導入)")
        return

    dst_dir = os.path.join(
        os.path.expanduser("~"),
        "Desktop",
        "【全社標準】弔事対応フォルダ",
        "5.test_bot",
    )
    dst_path = os.path.join(dst_dir, f"results_{timestamp}.xlsx")
    if not os.path.exists(dst_path):
        raise FileNotFoundError(f"結果ブックが見つかりません: {dst_path}")

    wb = oxl.load_workbook(dst_path)
    target_sheet_name = sheet_name or str(tehai_number)
    if target_sheet_name not in wb.sheetnames:
        raise ValueError(f"シートが見つかりません: {target_sheet_name}")
    ws = wb[target_sheet_name]

    d15 = str(ws["D15"].value or "")
    d27 = str(ws["D27"].value or "")
    if not d15:
        raise ValueError("D15 が空です (検索キーワード)")
    if not d27:
        print("[3.get_mail_result] D27 が空です。キーワードのみで検索します")

    outlook = w32.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    sent = mapi.GetDefaultFolder(5)  # 送信済みアイテム
    items = sent.Items
    try:
        items.Sort("[SentOn]", True)
    except Exception:
        pass

    mail = _best_mail_match(items, d15, d27 or d15)
    if not mail:
        raise RuntimeError("条件に合う送信済みメールが見つかりませんでした")

    content = str(getattr(mail, 'Body', '') or '')
    ws["E3"].value = content

    max_line = max((len(line) for line in content.splitlines()), default=20)
    ws.column_dimensions['E'].width = max(20, min(max_line + 2, 120))

    wb.save(dst_path)
    print("[3.get_mail_result] メール内容を E3 に貼り付け、列幅を調整しました")


if __name__ == "__main__":
    try:
        run(1234, "0101_1234")
    except Exception as e:
        print(f"[3.get_mail_result] 例外: {e}")
        sys.exit(1)
