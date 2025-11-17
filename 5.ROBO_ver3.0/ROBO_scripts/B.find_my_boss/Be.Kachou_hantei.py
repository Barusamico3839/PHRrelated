#!/usr/bin/env python3
"""Decide the CC routing rules based on boss job titles."""

from __future__ import annotations

import argparse
import json
import logging
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
import re
import unicodedata

SCRIPT_DIR = Path(__file__).resolve().parent
ROBO_SCRIPTS_ROOT = SCRIPT_DIR.parent
if str(ROBO_SCRIPTS_ROOT) not in sys.path:
    sys.path.append(str(ROBO_SCRIPTS_ROOT))

from common import PathRegistry
from excel_com import get_used_range_bounds, open_workbook

LOGGER = logging.getLogger("chouji_robo.kachou_hantei")
POSITIONS_CACHE = Path(__file__).with_name("positions_snapshot.json")

LABEL_COL = 8
NAME_COL = 9
EMAIL_COL = 10
DEPT_COL = 12
TITLE_COL = 13  # M列
MAIL_TARGET_COL = 14  # N列（タグ用）
FIRST_DATA_ROW = 5

CC_KEYWORDS = ("課長", "所長", "主幹")
HIGH_RANK_KEYWORDS = ("社長", "副社長", "常務", "監査役", "執行役員", "本部長")


def _normalise(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _normalize_sheet_title(value: str) -> str:
    if not value:
        return ""
    normalized = unicodedata.normalize("NFKC", value)
    normalized = re.sub(r"\s+", "", normalized).lower()
    return normalized


@dataclass
class PersonEntry:
    row: int
    label: str
    job_title: str


class RpaSheetAccessor:
    def __init__(self, book_path: Optional[Path] = None, sheet_name: str = "RPAシート") -> None:
        registry = PathRegistry()
        candidates: List[Path] = []
        if book_path:
            candidates.append(Path(book_path))
        candidates.append(Path(registry.rpa_book_destination))
        candidates.append(Path(registry.rpa_local_book))

        self.book_path: Optional[Path] = None
        for candidate in candidates:
            if candidate and candidate.exists():
                self.book_path = candidate
                LOGGER.info("[INFO] 使用するブックを検出しました: %s", candidate)
                break

        if self.book_path is None:
            raise FileNotFoundError(
                "RPAシートのブックが見つかりませんでした。--book でパスを指定してください。"
            )

        self.sheet_name = sheet_name
        self._cell_cache: Dict[Tuple[int, int], str] = {}
        self._pending_cells: Dict[Tuple[int, int], str] = {}
        LOGGER.info("[STEP] RPAシートの読み込みを開始します: book=%s sheet=%s", self.book_path, self.sheet_name)
        self._rows = self._load_rows()

    def _load_rows(self) -> List[PersonEntry]:
        people: List[PersonEntry] = []
        with open_workbook(self.book_path, read_only=True) as workbook:
            sheet = self._resolve_sheet(workbook)
            row_start, row_end, _, _ = get_used_range_bounds(sheet)
            LOGGER.info("[INFO] シートの使用範囲: rows=%s-%s", row_start, row_end)
            blank_run = 0
            for row in range(max(FIRST_DATA_ROW, row_start), row_end + 1):
                label = _normalise(sheet.Cells(row, LABEL_COL).Value)
                title = _normalise(sheet.Cells(row, TITLE_COL).Value)
                self._cell_cache[(row, TITLE_COL)] = title
                self._cell_cache[(row, MAIL_TARGET_COL)] = _normalise(
                    sheet.Cells(row, MAIL_TARGET_COL).Value
                )
                if not label and not title:
                    blank_run += 1
                    if blank_run >= 4:
                        break
                    continue
                blank_run = 0
                people.append(PersonEntry(row=row, label=label, job_title=title))
                LOGGER.debug("Row %s: label=%s title=%s", row, label, title)
        LOGGER.info("[STEP] RPAシートの読み込みが完了しました: 読み込み行=%s", len(people))
        return people

    def _resolve_sheet(self, workbook):
        target_normalized = _normalize_sheet_title(self.sheet_name)
        fallback = None
        for sheet in workbook.Worksheets:
            name = str(getattr(sheet, "Name", ""))
            if name == self.sheet_name:
                return sheet
            normalized = _normalize_sheet_title(name)
            if normalized == target_normalized:
                return sheet
            if fallback is None and "rpa" in normalized:
                fallback = sheet
        if fallback is not None:
            LOGGER.debug("指定シート '%s' が見つからないため '%s' を使用します。", self.sheet_name, getattr(fallback, "Name", ""))
            return fallback
        try:
            first_sheet = workbook.Worksheets(1)
            LOGGER.debug("指定シートが見つからないため先頭シート '%s' を使用します。", getattr(first_sheet, "Name", ""))
            return first_sheet
        except Exception:
            raise ValueError(f"シート '{self.sheet_name}' は存在しません。")

    def iter_people(self) -> List[PersonEntry]:
        return list(self._rows)

    def update_title(self, row: int, title: str) -> None:
        self._cell_cache[(row, TITLE_COL)] = title
        self._pending_cells[(row, TITLE_COL)] = title
        for person in self._rows:
            if person.row == row:
                person.job_title = title
                break

    def append_tag(self, row: int, column: int, tag: str) -> None:
        existing = self._cell_cache.get((row, column), "")
        if not existing:
            new_value = tag
        elif tag in existing:
            new_value = existing
        else:
            new_value = f"{existing} / {tag}"
        self._cell_cache[(row, column)] = new_value
        self._pending_cells[(row, column)] = new_value

    def get_cell_value(self, row: int, column: int) -> str:
        return self._cell_cache.get((row, column), "")

    def collect_positions(self, exclude_first: bool = True) -> List[str]:
        positions: List[str] = []
        for idx, person in enumerate(self._rows):
            if exclude_first and idx == 0:
                continue
            if person.job_title:
                positions.append(person.job_title)
        return positions

    def save(self) -> None:
        if not self._pending_cells:
            LOGGER.info("[INFO] Excel への変更が無いため、書き込みをスキップします。")
            return
        with open_workbook(self.book_path) as workbook:
            sheet = self._resolve_sheet(workbook)
            for (row, column), value in self._pending_cells.items():
                sheet.Cells(row, column).Value = value
            workbook.Save()
        LOGGER.info("[STEP] Excel へ %s 件のセル更新を書き込みました。", len(self._pending_cells))
        self._pending_cells.clear()


def _append_tag(sheet: RpaSheetAccessor, row: int, column: int, tag: str) -> None:
    sheet.append_tag(row, column, tag)


def _load_cached_positions(cache_path: Path = POSITIONS_CACHE) -> List[str]:
    if not cache_path.exists():
        LOGGER.info("[INFO] positions キャッシュが存在しません: %s", cache_path)
        return []
    try:
        LOGGER.info("[STEP] positions キャッシュの読み込みを開始します: %s", cache_path)
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
        positions = payload.get("positions") or []
        return [str(item) for item in positions if item]
    except Exception:
        LOGGER.warning("positions キャッシュの読み込みに失敗しました: %s", cache_path)
        return []


def _select_reference_title(idx: int, positions: Sequence[str], entry: PersonEntry) -> str:
    if idx < len(positions) and positions[idx]:
        return str(positions[idx])
    return entry.job_title


def apply_cc_logic(sheet: RpaSheetAccessor, positions: Sequence[str]) -> dict:
    people = sheet.iter_people()
    boss_entries = people[1:]
    if not positions:
        positions = [entry.job_title for entry in boss_entries]

    LOGGER.info("[STEP] CC 判定処理を開始します: 対象行=%s 既知positions=%s", len(boss_entries), len(positions))
    cc_rows: List[int] = []
    cc_assigned_row: Optional[int] = None

    for idx, entry in enumerate(boss_entries):
        title = _normalise(_select_reference_title(idx, positions, entry))
        if not title:
            LOGGER.debug("Row %s: 役職名が空のためスキップします。", entry.row)
            continue
        if any(keyword in title for keyword in CC_KEYWORDS):
            _append_tag(sheet, entry.row, MAIL_TARGET_COL, "ccに含む")
            cc_rows.append(entry.row)
            cc_assigned_row = entry.row
            LOGGER.debug("Row %s: ccに含む を付与しました (title=%s)", entry.row, title)
            break

    if cc_assigned_row is None and boss_entries:
        first_manager_row = boss_entries[0].row
        LOGGER.info("課長/所長/主幹が見付からないため、行 %s を cc 対象にします。", first_manager_row)
        _append_tag(sheet, first_manager_row, MAIL_TARGET_COL, "ccに含む")
        cc_rows.append(first_manager_row)
        cc_assigned_row = first_manager_row

    n16_value = _normalise(sheet.get_cell_value(16, MAIL_TARGET_COL))
    if n16_value == "※二次上司は必要" and cc_assigned_row is not None:
        second_row = cc_assigned_row + 1
        _append_tag(sheet, second_row, MAIL_TARGET_COL, "ccに含む")
        cc_rows.append(second_row)
        LOGGER.info("[INFO] 二次上司が必要なため、行 %s にも cc に含む を追加しました。", second_row)

    for entry in people:
        title = _normalise(entry.job_title)
        if not title:
            continue
        if any(keyword in title for keyword in HIGH_RANK_KEYWORDS):
            _append_tag(sheet, entry.row, MAIL_TARGET_COL, "上位役職")
            LOGGER.debug("Row %s: 上位役職タグを付与しました (title=%s)", entry.row, title)

    sheet.save()
    LOGGER.info("[STEP] CC 判定処理が完了しました。cc_rows=%s", cc_rows)
    return {
        "positions": list(positions),
        "cc_rows": cc_rows,
    }


def configure_logging(verbose: bool) -> None:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s :: %(message)s")
    handler.setFormatter(formatter)
    root = logging.getLogger()
    if not root.handlers:
        root.addHandler(handler)
    root.setLevel(logging.DEBUG if verbose else logging.INFO)
    LOGGER.setLevel(logging.DEBUG if verbose else logging.INFO)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="上司に対応する CC 設定を更新します。")
    parser.add_argument("--book", type=Path, help="RPAブックのパス。未指定時は既定値を使用します。")
    parser.add_argument("--sheet", default="RPAシート", help="対象シート名です。")
    parser.add_argument("--positions-cache", type=Path, default=POSITIONS_CACHE, help="positions キャッシュのパスです。")
    parser.add_argument("--verbose", action="store_true", help="詳細ログを出力します。")
    return parser.parse_args(argv)


def run(argv: Optional[Sequence[str]] = None) -> dict:
    args = parse_args(argv)
    configure_logging(args.verbose)
    LOGGER.info("[STEP] Be.Kachou_hantei を開始します。")
    sheet = RpaSheetAccessor(book_path=args.book, sheet_name=args.sheet)
    positions = _load_cached_positions(args.positions_cache)
    result = apply_cc_logic(sheet, positions)
    LOGGER.info("[STEP] Be.Kachou_hantei が完了しました。cc_rows=%s", result.get("cc_rows"))
    return result


def main(argv: Optional[Sequence[str]] = None) -> int:
    try:
        run(argv)
        return 0
    except KeyboardInterrupt:
        LOGGER.error("ユーザーによって処理が中断されました。")
        return 130
    except Exception as exc:
        LOGGER.exception("処理中にエラーが発生しました: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
