# -*- coding: utf-8 -*-
import sys


def run(*args, **kwargs):
    print("[4.save_results] このステップは使用しません")


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print(f"[4.save_results] 例外: {e}")
        sys.exit(1)
