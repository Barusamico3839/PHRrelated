# -*- coding: utf-8 -*-
import os
import sys
import argparse
import importlib.util
from typing import Optional, List


def _scripts_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _module_from(path: str, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"module load failed: {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


def main(argv: Optional[List[str]] = None) -> None:
    parser = argparse.ArgumentParser(
        prog="main_from_step2",
        description="step2以降（2→3→5）だけを実行してテストします。",
    )
    parser.add_argument(
        "tehai_numbers",
        nargs="*",
        type=int,
        help="手配番号（半角スペース区切り） 例: 4785 4886",
    )
    parser.add_argument(
        "--ts",
        "--timestamp",
        dest="timestamp",
        default=None,
        help="任意のタイムスタンプ（省略時は現在時刻 mmdd_HHMM を使用するのが一般的）。",
    )
    parser.add_argument(
        "--skip-mail",
        action="store_true",
        help="step3（Outlook 連携）をスキップ",
    )

    ns = parser.parse_args(argv if argv is not None else sys.argv[1:])
    tehai_numbers: List[int] = ns.tehai_numbers or []
    ts: Optional[str] = ns.timestamp

    sdir = _scripts_dir()
    get_pad = _module_from(os.path.join(sdir, "2.get_pad_result.py"), "get_pad")
    get_mail = _module_from(os.path.join(sdir, "3.get_mail_result.py"), "get_mail")
    compare = _module_from(os.path.join(sdir, "5.compare_results.py"), "compare")

    if not tehai_numbers:
        print("[main_from_step2] 手配番号が未指定のためサンプル 1234 を使用します。")
        tehai_numbers = [1234]

    if not ts:
        import datetime as _dt
        ts = _dt.datetime.now().strftime("%m%d_%H%M")

    print(f"[main_from_step2] start: 件数={len(tehai_numbers)}, ts={ts}")

    for idx, num in enumerate(tehai_numbers, start=1):
        wk = idx - 1  # 週目（0ベース）: 1件目=0→root_row=10
        print(f"[main_from_step2] === Case {idx}/{len(tehai_numbers)} 手配番号={num} 週目idx={wk} ===")

        # step2
        try:
            sheet_name = get_pad.run(num, ts, wk, None)
        except Exception as e:
            print(f"[main_from_step2] step2 エラー: {e}")
            sys.exit(1)

        # step3（任意）
        if ns.skip_mail:
            print("[main_from_step2] step3 は --skip-mail のためスキップします。")
        else:
            try:
                get_mail.run(num, ts, sheet_name, wk)
            except Exception as e:
                print(f"[main_from_step2] step3 エラー: {e}")
                sys.exit(1)

        # step5（比較・ダミー）
        try:
            compare.run()
        except Exception as e:
            print(f"[main_from_step2] step5 エラー: {e}")
            sys.exit(1)

    print("[main_from_step2] すべて完了")


if __name__ == "__main__":
    main()

