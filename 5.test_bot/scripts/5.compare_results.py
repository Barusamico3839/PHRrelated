# -*- coding: utf-8 -*-
import os
import sys
import argparse
import datetime as _dt
import importlib.util
from typing import List, Optional


def _scripts_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _module_from(path: str, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"module load failed: {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


def run(*args, **kwargs):
    print("[5.compare_results] テスト用の比較ステップ（ダミー）を実行します。")


def _ensure_ts(ts: Optional[str]) -> str:
    return ts or _dt.datetime.now().strftime("%m%d_%H%M")


def main(argv: Optional[List[str]] = None) -> None:
    parser = argparse.ArgumentParser(
        prog="compare_step2_onward",
        description="step2以降（2→3→5）だけを手早くテスト実行します。",
    )
    parser.add_argument(
        "tehai_numbers",
        nargs="*",
        type=int,
        help="配列の手配番号（半角スペース区切り）。例: 4785 4886",
    )
    parser.add_argument(
        "--ts",
        "--timestamp",
        dest="timestamp",
        default=None,
        help="任意のタイムスタンプ（省略時は現在時刻 mmdd_HHMM）。",
    )
    parser.add_argument(
        "--skip-mail",
        action="store_true",
        help="step3（Outlook連携）をスキップします。",
    )

    ns = parser.parse_args(argv if argv is not None else sys.argv[1:])
    tehai_numbers: List[int] = ns.tehai_numbers or []
    ts = _ensure_ts(ns.timestamp)

    if not tehai_numbers:
        print("[5.compare_results] 手配番号が未指定のため、例: 1234 を試します。")
        tehai_numbers = [1234]

    sdir = _scripts_dir()
    get_pad = _module_from(os.path.join(sdir, "2.get_pad_result.py"), "get_pad")
    get_mail = _module_from(os.path.join(sdir, "3.get_mail_result.py"), "get_mail")

    print(f"[5.compare_results] step2以降テスト開始: 件数={len(tehai_numbers)}, ts={ts}")
    for i, num in enumerate(tehai_numbers, start=1):
        print(f"[5.compare_results] --- Case {i}/{len(tehai_numbers)}: 手配番号={num} ---")
        # step2: PAD結果取り込み（Excelコピー）
        try:
            sheet_name = get_pad.run(num, ts, i-1, None)
        except Exception as e:
            print(f"[5.compare_results] step2 エラー: {e}")
            sys.exit(1)

        # step3: 送信メールの本文をExcel(E3)へ（Outlook必須）
        if ns.skip_mail:
            print("[5.compare_results] step3 は --skip-mail 指定のためスキップします。")
        else:
            try:
                get_mail.run(num, ts, sheet_name, i-1)
            except Exception as e:
                print(f"[5.compare_results] step3 エラー: {e}")
                sys.exit(1)

        # step5: 結果比較（このファイルの run を呼ぶ）
        try:
            run()
        except Exception as e:
            print(f"[5.compare_results] step5 エラー: {e}")
            sys.exit(1)

    print("[5.compare_results] step2以降テストが完了しました。")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        print(f"[5.compare_results] 予期せぬエラー: {e}")
        sys.exit(1)
