#!/usr/bin/env python3
"""Lookup job titles on PHONE APPLI and store them on the RPA sheet."""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Tuple, Set, Set
from urllib.parse import quote

try:
    from selenium import webdriver
    from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.edge.options import Options as EdgeOptions
    from selenium.webdriver.edge.service import Service as EdgeService
    from selenium.webdriver.remote.webelement import WebElement
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
except Exception as exc:  # pragma: no cover - selenium might be missing in CI
    webdriver = None  # type: ignore[assignment]
    SELENIUM_IMPORT_ERROR: Optional[Exception] = exc
else:  # pragma: no cover - import guard
    SELENIUM_IMPORT_ERROR = None

from excel_com import get_used_range_bounds, open_workbook

SCRIPT_DIR = Path(__file__).resolve().parent
ROBO_SCRIPTS_ROOT = SCRIPT_DIR.parent
DEFAULT_DRIVER_PATH = ROBO_SCRIPTS_ROOT.parent / "ROBO_tools" / "msedgedriver.exe"
if str(ROBO_SCRIPTS_ROOT) not in sys.path:
    sys.path.append(str(ROBO_SCRIPTS_ROOT))

try:
    from common import PathRegistry
except Exception as exc:  # pragma: no cover - running outside robot root
    raise RuntimeError("common.PathRegistry が読み込めません。実行ディレクトリを確認してください。") from exc

LOGGER = logging.getLogger("chouji_robo.find_job_title")
BASE_URL = "https://panasonic.phoneappli.net/front/login?returnTo=%2Ffront%2FinternalContacts%3FapiType%3Dsearch%26page%3D0%26size%3D30"
SEARCH_URL_TEMPLATE = (
    "https://panasonic.phoneappli.net/front/internalContacts"
    "?apiType=searchByFreeWord&divisionId=&freeWord={email}&freeWordSearchType=ALL&page=0&size=30"
)
POSITIONS_CACHE = Path(__file__).with_name("positions_snapshot.json")

LABEL_COL = 8
NAME_COL = 9
EMAIL_COL = 10
COMPANY_COL = 11
DEPT_COL = 12
TITLE_COL = 13
MAIL_TARGET_COL = 14
FIRST_DATA_ROW = 5
MAX_CONSECUTIVE_BLANKS = 4


@lru_cache(maxsize=1)
def _load_cache_payload() -> dict:
    if not POSITIONS_CACHE.exists():
        return {}
    try:
        payload = json.loads(POSITIONS_CACHE.read_text(encoding="utf-8"))
        if isinstance(payload, dict):
            return payload
    except Exception as exc:
        LOGGER.warning("positions_snapshot.json の読み込みに失敗しました: %s", exc)
    return {}


def _resolve_driver_path() -> Optional[Path]:
    override = os.getenv("CHOUJI_EDGE_DRIVER")
    if override:
        candidate = Path(override).expanduser()
        if candidate.exists():
            return candidate
        LOGGER.warning("CHOUJI_EDGE_DRIVER=%s が見つかりません。", override)

    payload = _load_cache_payload()
    json_path = payload.get("edge_driver_path")
    if isinstance(json_path, str) and json_path.strip():
        candidate = Path(json_path).expanduser()
        if candidate.exists():
            return candidate
        LOGGER.warning("positions_snapshot.json の edge_driver_path (%s) が見つかりません。", json_path)

    if DEFAULT_DRIVER_PATH.exists():
        return DEFAULT_DRIVER_PATH
    return None


def _normalise(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return text


@dataclass
class PersonEntry:
    """Represents a single row on the RPA sheet."""

    row: int
    label: str
    name: str
    email: str
    department: str
    job_title: str = ""


class RpaSheetAccessor:
    """Convenience wrapper for reading/writing the RPA sheet."""

    def __init__(self, book_path: Optional[Path] = None, sheet_name: str = "RPA�V�[�g") -> None:
        registry = PathRegistry()
        target = Path(book_path) if book_path else Path(registry.rpa_book_destination)
        if not target.exists():
            raise FileNotFoundError(f"RPA�V�[�g��������܂���: {target}")
        self.book_path = target
        self.sheet_name = sheet_name
        self._cell_cache: Dict[Tuple[int, int], str] = {}
        self._pending_cells: Dict[Tuple[int, int], str] = {}
        self._rows = list(self._load_rows())

    def _resolve_sheet(self, workbook):
        for sheet in workbook.Worksheets:
            name = str(getattr(sheet, "Name", ""))
            if name == self.sheet_name:
                return sheet
        raise ValueError(f"�V�[�g '{self.sheet_name}' ���u�b�N�ɑ��݂��܂���B")

    def _load_rows(self) -> List[PersonEntry]:
        people: List[PersonEntry] = []
        with open_workbook(self.book_path, read_only=True) as workbook:
            sheet = self._resolve_sheet(workbook)
            row_start, row_end, _, _ = get_used_range_bounds(sheet)
            blank_run = 0
            for row in range(max(FIRST_DATA_ROW, row_start), row_end + 1):
                label = _normalise(sheet.Cells(row, LABEL_COL).Value)
                name = _normalise(sheet.Cells(row, NAME_COL).Value)
                email = _normalise(sheet.Cells(row, EMAIL_COL).Value)
                department = _normalise(sheet.Cells(row, DEPT_COL).Value)
                title = _normalise(sheet.Cells(row, TITLE_COL).Value)
                self._cell_cache[(row, TITLE_COL)] = title
                if not any([label, name, email, department]):
                    blank_run += 1
                    if blank_run >= MAX_CONSECUTIVE_BLANKS:
                        break
                    continue
                blank_run = 0
                people.append(
                    PersonEntry(
                        row=row,
                        label=label,
                        name=name,
                        email=email,
                        department=department,
                        job_title=title,
                    )
                )
        return people

    def iter_people(self) -> Iterator[PersonEntry]:
        return iter(self._rows)

    def update_title(self, row: int, title: str) -> None:
        self._cell_cache[(row, TITLE_COL)] = title
        self._pending_cells[(row, TITLE_COL)] = title
        for person in self._rows:
            if person.row == row:
                person.job_title = title
                break

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
            return
        with open_workbook(self.book_path) as workbook:
            sheet = self._resolve_sheet(workbook)
            for (row, column), value in self._pending_cells.items():
                sheet.Cells(row, column).Value = value
            workbook.Save()
        self._pending_cells.clear()


@dataclass
class Candidate:
    """Represents a search result item inside PHONE APPLI."""

    container: WebElement
    description: str
    division_text: str
    position_preview: str = ""


class PhoneAppliClient:
    """Thin Selenium wrapper for PHONE APPLI PEOPLE."""

    def __init__(
        self,
        *,
        headless: bool = False,
        keep_browser_open: bool = False,
        base_url: str = BASE_URL,
        timeout: int = 25,
    ) -> None:
        if webdriver is None or SELENIUM_IMPORT_ERROR:
            raise ImportError("selenium がインストールされていません。pip install selenium を実行してください。") from SELENIUM_IMPORT_ERROR
        self.base_url = base_url
        self.keep_browser_open = keep_browser_open
        self.driver = self._create_driver(headless=headless)
        self.wait = WebDriverWait(self.driver, timeout)

    def _create_driver(self, *, headless: bool) -> webdriver.Edge:
        options = EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-features=msEdgeDataSharing")
        options.add_argument("--disable-blink-features=AutomationControlled")
        if headless:
            options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
        service = None
        candidate = _resolve_driver_path()
        if candidate:
            LOGGER.info("Using msedgedriver at %s", candidate)
            service = EdgeService(executable_path=str(candidate))
        try:
            if service:
                return webdriver.Edge(options=options, service=service)
            return webdriver.Edge(options=options)
        except WebDriverException as exc:
            raise RuntimeError("Edge WebDriver の起動に失敗しました。Edge がインストールされているか確認してください。") from exc

    def close(self) -> None:
        if self.keep_browser_open:
            return
        try:
            self.driver.quit()
        except Exception:
            pass

    def ensure_ready(self, login_wait: int = 120) -> None:
        LOGGER.info("PHONE APPLI にアクセスしています: %s", self.base_url)
        self.driver.get(self.base_url)
        try:
            login_button = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".o365-login-button, .normal_button.o365-login-button"))
            )
            login_button.click()
            LOGGER.info("Microsoft 365 ログインボタンをクリックしました。必要に応じて認証を完了してください。")
        except TimeoutException:
            LOGGER.debug("ログインボタンが見つかりませんでした。すでにサインイン済みと判断します。")
        self._wait_for_search_box(login_wait)

    def _wait_for_search_box(self, login_wait: int) -> None:
        LOGGER.info("検索画面の読み込みを待っています (最大 %s 秒)...", login_wait)
        WebDriverWait(self.driver, login_wait).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder*='キーワード']"))
        )

    def lookup_job_title(self, email: str, department: str) -> Optional[str]:
        if not email:
            return None
        LOGGER.info("URL検索: %s (department=%s)", email, department or "-")
        encoded = quote(email, safe="")
        target_url = SEARCH_URL_TEMPLATE.format(email=encoded)
        LOGGER.debug("検索URLへ遷移: %s", target_url)
        self.driver.get(target_url)
        time.sleep(0.5)
        candidates = self._collect_candidates()
        if not candidates:
            LOGGER.warning("検索結果が見つかりませんでした。")
            return None
        target = self._select_candidate(candidates, department)
        if target is None:
            LOGGER.warning("一致する候補を特定できませんでした。")
            return None
        title = self._open_candidate_and_read_position(target)
        LOGGER.info("取得した役職: %s", title or "(取得失敗)")
        return title

    def _collect_candidates(self) -> List[Candidate]:
        result_selector = "li.internal-item"

        def _has_items(driver: webdriver.Edge) -> bool:
            nodes = driver.find_elements(By.CSS_SELECTOR, result_selector)
            return any(node.text.strip() for node in nodes)

        try:
            self.wait.until(_has_items)
        except TimeoutException:
            return []

        nodes = self.driver.find_elements(By.CSS_SELECTOR, result_selector)
        candidates: List[Candidate] = []
        for node in nodes:
            text = node.text.strip()
            if not text:
                continue
            division_text = ""
            position_preview = ""
            try:
                division_texts: List[str] = []
                division_columns = node.find_elements(By.CSS_SELECTOR, "div._divisionColumn_ls35d_22")
                if not division_columns:
                    division_columns = node.find_elements(By.CSS_SELECTOR, "div.internal-item__column--division-position")
                for column in division_columns:
                    button_spans = column.find_elements(By.CSS_SELECTOR, "div._itemText_ls35d_61 span, button span, span")
                    for elem in button_spans:
                        value = elem.text.strip()
                        if value:
                            division_texts.append(value)
                seen_divisions: Set[str] = set()
                filtered = []
                for value in division_texts:
                    if value not in seen_divisions:
                        seen_divisions.add(value)
                        filtered.append(value)
                division_text = " / ".join(filtered)
            except Exception:
                division_text = ""

            try:
                position_values: List[str] = []
                position_nodes = node.find_elements(
                    By.CSS_SELECTOR,
                    "div._nameColumn_pij2g_22 span._positionPaddingLarge_pij2g_49 span, "
                    "div[class*='name'] span[class*='position'], "
                    "span[class*='positionPadding'] span, span[class*='positionPadding']",
                )
                for elem in position_nodes:
                    value = elem.text.strip()
                    if value:
                        position_values.append(value)
                seen_positions: Set[str] = set()
                filtered_positions = []
                for value in position_values:
                    if value not in seen_positions:
                        seen_positions.add(value)
                        filtered_positions.append(value)
                position_preview = " / ".join(filtered_positions)
            except Exception:
                position_preview = ""

            candidates.append(
                Candidate(container=node, description=text, division_text=division_text, position_preview=position_preview)
            )
        return candidates
        return candidates

    def _select_candidate(self, candidates: Sequence[Candidate], department: str) -> Optional[Candidate]:
        if not candidates:
            return None
        if len(candidates) == 1 or not department:
            return candidates[0]

        norm_department = department.strip()

        def _calc_score(text: str) -> float:
            if not text or not norm_department:
                return 0.0
            source = text.strip()
            return SequenceMatcher(None, source, norm_department).ratio()

        best_candidate = candidates[0]
        best_score = -1.0
        for candidate in candidates:
            division_score = _calc_score(candidate.division_text)
            description_score = _calc_score(candidate.description)
            score = division_score if division_score > 0 else description_score
            LOGGER.debug(
                "候補 '%s' division_score=%.3f description_score=%.3f",
                candidate.description.splitlines()[0] if candidate.description else "(no text)",
                division_score,
                description_score,
            )
            if score > best_score:
                best_score = score
                best_candidate = candidate
        LOGGER.info("部門一致率が最も高い候補を選択しました (score=%.3f)。", best_score)
        return best_candidate

    def _open_candidate_and_read_position(self, candidate: Candidate) -> Optional[str]:
        if candidate.position_preview:
            LOGGER.debug("一覧の役職を使用: %s", candidate.position_preview)
            return candidate.position_preview
        button = None
        button_selector = "button.internal-item__column--person, button"
        for attempt in range(10):
            try:
                button = candidate.container.find_element(By.CSS_SELECTOR, button_selector)
                break
            except NoSuchElementException:
                time.sleep(0.5)
        if button is None:
            LOGGER.error("候補にプロフィールボタンが見つかりません。")
            return None
        button.click()
        try:
            title = self._extract_position_text()
            if not title:
                LOGGER.error("役職情報の読み取りに失敗しました。")
            return title
        finally:
            self._close_profile_dialog()

    def _extract_position_text(self) -> Optional[str]:
        dialog = None
        try:
            dialog = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.profile-dialog, .profile-dialog__container"))
            )
        except TimeoutException:
            LOGGER.debug("プロフィールダイアログが表示されませんでした。再試行します。")
            dialog = None
        if dialog is None:
            for attempt in range(3):
                try:
                    time.sleep(1.0)
                    dialog = EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "div.profile-dialog, .profile-dialog__container")
                    )(self.driver)
                    if dialog is not None:
                        LOGGER.debug("プロフィールダイアログを再試行 %d 回目で検出しました。", attempt + 1)
                        break
                except Exception:
                    dialog = None
            else:
                LOGGER.debug("プロフィールダイアログを検出できませんでした。")
                return None

        selector_groups = [
            ".detail__position span",
            "[class*='detail__position'] span",
            "[data-v-cbd4949e] span",
            "[class*='position'] span",
        ]

        titles: List[str] = []
        for selector in selector_groups:
            try:
                for element in dialog.find_elements(By.CSS_SELECTOR, selector):
                    text = element.text.strip()
                    if text:
                        titles.append(text)
            except Exception:
                continue

        if not titles:
            try:
                for element in dialog.find_elements(By.TAG_NAME, "span"):
                    text = element.text.strip()
                    if text and "@" not in text and len(text) <= 20:
                        titles.append(text)
            except Exception:
                pass

        unique_titles: List[str] = []
        seen: Set[str] = set()
        for title in titles:
            if title not in seen:
                seen.add(title)
                unique_titles.append(title)

        if not unique_titles:
            return "役職未設定"

        LOGGER.debug("役職候補を検出: %s", unique_titles)
        return " / ".join(unique_titles)

    def _close_profile_dialog(self) -> None:
        dialog_selector = "div.profile-dialog"
        close_selectors = [
            "button.profile-dialog__close",
            "button[aria-label*='閉じる']",
            f"{dialog_selector} button[type='button']",
        ]
        for selector in close_selectors:
            try:
                close_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                close_button.click()
                WebDriverWait(self.driver, 5).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, dialog_selector)))
                return
            except TimeoutException:
                LOGGER.debug("Close selector '%s' が見つからなかったためリトライします。", selector)
                continue

        LOGGER.debug("close ボタンを検出できなかったため Esc キーでクローズを試みます。")
        try:
            self.driver.switch_to.active_element.send_keys(Keys.ESCAPE)
            WebDriverWait(self.driver, 5).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, dialog_selector)))
        except Exception:
            LOGGER.warning("Esc キーによるダイアログクローズも失敗しました。")


def persist_positions(positions: Sequence[str], destination: Path = POSITIONS_CACHE) -> None:
    payload = _load_cache_payload().copy()
    if "edge_driver_path" not in payload and DEFAULT_DRIVER_PATH.exists():
        payload["edge_driver_path"] = str(DEFAULT_DRIVER_PATH)
    payload["generated_at"] = datetime.now().isoformat()
    payload["positions"] = list(positions)
    destination.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    _load_cache_payload.cache_clear()
    LOGGER.info("positions を %s に保存しました。", destination)


def configure_logging(verbose: bool) -> None:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s :: %(message)s")
    handler.setFormatter(formatter)
    LOGGER.setLevel(logging.DEBUG if verbose else logging.INFO)
    root = logging.getLogger()
    if not root.handlers:
        root.setLevel(logging.DEBUG if verbose else logging.INFO)
        root.addHandler(handler)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PHONE APPLI から役職を取得して RPA シートを更新します。")
    parser.add_argument("--book", type=Path, help="RPAブックのパス (省略時は既定パス)。")
    parser.add_argument("--sheet", default="RPAシート", help="対象シート名。")
    parser.add_argument("--max-rows", type=int, help="処理する行数の上限。")
    parser.add_argument("--headless", action="store_true", help="Edge をヘッドレスで起動します。")
    parser.add_argument("--keep-browser-open", action="store_true", help="処理後もブラウザを閉じません。")
    parser.add_argument("--login-wait", type=int, default=120, help="ログイン完了待ちのタイムアウト秒数。")
    parser.add_argument("--verbose", action="store_true", help="詳細ログを有効にします。")
    return parser.parse_args(argv)


def run(argv: Optional[Sequence[str]] = None) -> dict:
    args = parse_args(argv)
    configure_logging(args.verbose)

    sheet = RpaSheetAccessor(book_path=args.book, sheet_name=args.sheet)
    people = list(sheet.iter_people())
    if not people:
        LOGGER.warning("処理対象が見つかりませんでした。")
        return {"positions": []}

    client = PhoneAppliClient(
        headless=args.headless,
        keep_browser_open=args.keep_browser_open,
    )

    processed = 0
    try:
        client.ensure_ready(login_wait=args.login_wait)
        for idx, person in enumerate(people):
            if idx == 0:
                LOGGER.debug("本人行はスキップします。")
                continue
            if args.max_rows and processed >= args.max_rows:
                LOGGER.info("max_rows=%s に達したため処理を終了します。", args.max_rows)
                break
            if not person.email:
                LOGGER.info("Row %s (%s) はメールアドレスが空のためスキップします。", person.row, person.label or person.name)
                continue
            try:
                title = client.lookup_job_title(person.email, person.department)
            except Exception as exc:
                LOGGER.exception("検索中にエラーが発生しました: %s", exc)
                title = None
            if not title:
                title = "上司検索失敗"
            sheet.update_title(person.row, title)
            processed += 1
            time.sleep(0.5)
    finally:
        client.close()

    sheet.save()
    positions = sheet.collect_positions(exclude_first=True)
    persist_positions(positions)
    return {"positions": positions, "processed": processed}


def main(argv: Optional[Sequence[str]] = None) -> int:  # pragma: no cover - CLI helper
    try:
        run(argv)
        return 0
    except KeyboardInterrupt:
        LOGGER.error("ユーザーによって中断されました。")
        return 130
    except Exception as exc:
        LOGGER.exception("処理中に致命的なエラーが発生しました: %s", exc)
        return 1


if __name__ == "__main__":  # pragma: no cover - CLI entrypoint
    sys.exit(main())
