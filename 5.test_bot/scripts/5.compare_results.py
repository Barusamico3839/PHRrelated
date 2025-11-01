# -*- coding: utf-8 -*-
import sys


def run(*args, **kwargs):
    print("[5.compare_results] 後で開発します。現在はスキップします。")


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print(f"[5.compare_results] 例外: {e}")
        sys.exit(1)
