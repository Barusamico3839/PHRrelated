# -*- coding: utf-8 -*-
import os
import sys


def _import_win32():
    try:
        import win32com.client  # type: ignore
        return win32com.client
    except Exception as e:
        print(f"[6.start_next_case] pywin32 未インストール: {e}")
        return None


def _results_path(timestamp: str) -> str:
    return os.path.join(
        os.path.expanduser("~"),
        "Desktop",
        "【全社標準】弔事対応フォルダ",
        "5.test_bot",
        f"results_{timestamp}.xlsx",
    )


def _insert_message_shape(timestamp: str, tehai_number: int, message: str) -> None:
    if not message:
        print("[6.start_next_case] message_dialog が空のため図形挿入をスキップ")
        return
    w32 = _import_win32()
    if not w32:
        print("[6.start_next_case] Excel COM が利用できないため図形挿入をスキップ")
        return
    path = _results_path(timestamp)
    if not os.path.exists(path):
        print(f"[6.start_next_case] 結果ブックが見つかりません: {path}")
        return
    excel = None
    wb = None
    try:
        excel = w32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(path)
        try:
            ws = wb.Worksheets(str(tehai_number))
        except Exception:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = str(tehai_number)

        rng = ws.Range("H3")
        left = rng.Left
        top = rng.Top
        width = max(320, int(rng.Width * 6))
        height = max(80, int(rng.Height * 3))

        const = w32.constants
        try:
            shape = ws.Shapes.AddShape(const.msoShapeRoundedRectangle, left, top, width, height)
        except Exception:
            # fallback: 5 is msoShapeRoundedRectangle (typical)
            shape = ws.Shapes.AddShape(5, left, top, width, height)

        try:
            shape.TextFrame.Characters().Text = str(message)
            shape.TextFrame.Characters().Font.Size = 18
            shape.TextFrame.HorizontalAlignment = -4108  # xlCenter
            shape.TextFrame.VerticalAlignment = -4108    # xlCenter
        except Exception:
            try:
                # TextFrame2 alternative
                tr = shape.TextFrame2.TextRange
                tr.Characters.Text = str(message)
                tr.Characters.Font.Size = 18
            except Exception:
                pass

        wb.Save()
        print("[6.start_next_case] メッセージ図形を挿入しました (シート: {0}, H3)".format(tehai_number))
    except Exception as e:
        print(f"[6.start_next_case] 図形挿入時の例外: {e}")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=True)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


def run(current_index: int, total: int, timestamp: str, tehai_number: int, message_dialog: str) -> int:
    try:
        _insert_message_shape(timestamp, tehai_number, message_dialog or "")
    except Exception as e:
        print(f"[6.start_next_case] メッセージ挿入でエラー: {e}")
    nxt = current_index + 1
    print(f"[6.start_next_case] 次のケースへ {current_index} -> {nxt} / {total}")
    return nxt


if __name__ == "__main__":
    try:
        run(0, 10, "0101_1234", 1234, "テストメッセージ")
    except Exception as e:
        print(f"[6.start_next_case] 例外: {e}")
        sys.exit(1)
