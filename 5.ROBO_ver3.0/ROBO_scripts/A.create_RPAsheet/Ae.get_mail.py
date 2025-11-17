"""Step Ae: gather the relevant mail data from Outlook."""

from __future__ import annotations

import logging
import re
import time
from pathlib import Path
from datetime import datetime, timedelta
from typing import Any, List, Optional, TYPE_CHECKING
import pythoncom
SMTP_PROPERTY_URI = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

try:
    import tkinter as tk
    from tkinter import messagebox, simpledialog
except Exception:  # pragma: no cover - tkinter should exist on Windows
    tk = None  # type: ignore
    messagebox = None  # type: ignore
    simpledialog = None  # type: ignore

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover - handled gracefully at runtime
    win32com = None  # type: ignore

from excel_com import iter_rows, open_workbook

from common import MailEnvelope

if TYPE_CHECKING:  # pragma: no cover - typing only
    from main import ChoujiRobo


def run(robot: "ChoujiRobo") -> None:
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
    except Exception:
        pass

    try:
        _run_mail_session(robot)
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _run_mail_session(robot: "ChoujiRobo") -> None:
    robot.current_phase = "Ae.get_mail"
    logger = logging.getLogger("chouji_robo.mail")
    logger.info("Ae.get_mail: Outlookメールを検索します")

    if win32com is None:
        raise RuntimeError("pywin32 が見つからないため Outlook にアクセスできません。")

    robot.state.mail_sender = ""
    robot.state.mail_cc = ""
    robot.state.mail_bcc = ""
    robot.state.forms_row = []

    mail_time = robot.state.mail_time
    if mail_time is None:
        raise RuntimeError("mail_time が設定されていません。")

    namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    mail_sources = _gather_mail_sources(namespace, logger)
    if not mail_sources:
        raise RuntimeError("Outlook のメールフォルダを列挙できませんでした。")

    candidates = _collect_recent_messages(
        mail_sources,
        anchor=mail_time,
        seconds=180,
        logger=logger,
    )
    candidates = _sort_and_log_candidates(
        logger, mail_time, candidates, label="近傍スキャン(±180秒)"
    )

    if candidates:
        nearest = candidates[0]
        diff_seconds = _seconds_difference(nearest.received_at, mail_time)
        logger.info(
            "近傍スキャンで最寄りメールを決定: 受信=%s 件名=%s 差分=%.1f秒",
            nearest.received_at,
            nearest.subject,
            diff_seconds,
        )
        logger.info("メール本文プレビュー:\n%s", nearest.body)
        print(
            f"[INFO] Ae.get_mail: 管理時刻との差 {diff_seconds:.1f} 秒のメールを候補として使用します。"
        )
    else:
        nearest_overall = _find_nearest_message(mail_sources, mail_time, logger)
        if nearest_overall:
            diff_seconds = _seconds_difference(nearest_overall.received_at, mail_time)
            robot.mail_logger.error("指定時刻付近のメール取得ができません。最寄りのメール内容を出力します。")
            robot.mail_logger.error(
                "受信=%s 件名=%s 差分=%.1f秒",
                nearest_overall.received_at,
                nearest_overall.subject,
                diff_seconds,
            )
            robot.mail_logger.error("メール本文プレビュー:\n%s", nearest_overall.body)
            print(
                f"[WARN] Ae.get_mail: 管理時刻との差 {diff_seconds:.1f} 秒のメールのみ取得できました。"
            )
        raise RuntimeError(
            f"指定時刻のメールは見つかりませんでした。管理NO.[{robot.state.tehai_number}]を確認できますか？"
        )
    manual_phrase = f"弔事の発生した従業員：{robot.state.pin}"
    prioritized = [
        entry for entry in candidates if manual_phrase in robot._safe_str(entry.body)
    ]
    if prioritized:
        if _process_candidates_via_forms(robot, prioritized, logger):
            return
        logger.info("本文フレーズ一致メールでは必要情報を取得できませんでした。次の条件に進みます。")

    company_is_pid = robot._safe_str(robot.state.company_name).upper() == "PID"
    if company_is_pid:
        if _process_candidates_via_attachments(robot, candidates, logger, match_mode="name"):
            return
        logger.info("PID向け氏名検索でも情報を取得できませんでした。PIN検索にフォールバックします。")

    if _process_candidates_via_attachments(robot, candidates, logger, match_mode="pin"):
        return

    print("[WARN] Ae.get_mail: 自動解析で該当メールを確定できなかったため手動選択に切り替えます。")
    target_entry = _select_target_mail(robot, candidates)
    _process_manual_selection(robot, target_entry)


def _process_manual_selection(robot: "ChoujiRobo", entry: MailEnvelope) -> None:
    logger = logging.getLogger("chouji_robo.mail")
    body_text = robot._safe_str(entry.body)
    manual_phrase = "弔事の発生した従業員："

    if manual_phrase in body_text:
        logger.info(
            "手動選択: 本文に『弔事の発生した従業員：』があるためForms転記フローを実行します。"
        )
        print("[INFO] Ae.get_mail: 本文に対象フレーズがあるためSharePoint経由で取得します。")
        if _process_entry_via_forms(robot, entry):
            return
        logger.warning("手動選択: 本文フレーズはあるがForms転記でデータが取得できませんでした。添付を確認します。")

    logger.info(
        "手動選択: 本文に対象フレーズが無いため添付ファイルのみを確認します。"
    )
    print("[INFO] Ae.get_mail: 本文にフレーズなし -> 添付ファイル確認に切り替えます。")

    robot.state.selected_mail_entry = entry
    mail_item = getattr(entry, "raw_item", None)
    if mail_item is not None:
        _populate_mail_metadata(robot, mail_item, logger)

    company_is_pid = robot._safe_str(robot.state.company_name).upper() == "PID"
    attachment_modes = ["name", "pin"] if company_is_pid else ["pin"]
    if _process_entry_via_attachments(robot, entry, attachment_modes, logger):
        return

    logger.error("手動選択: 添付ファイルも無いため処理を続行できません。")
    raise RuntimeError("ロボに対応できるメールじゃありません")


def _process_candidates_via_forms(
    robot: "ChoujiRobo",
    entries: List[MailEnvelope],
    logger: logging.Logger,
) -> bool:
    for idx, entry in enumerate(entries, start=1):
        logger.info(
            "本文フレーズ優先: 候補 #%d をForms転記で確認します (受信=%s 件名=%s)",
            idx,
            entry.received_at,
            entry.subject,
        )
        if _process_entry_via_forms(robot, entry):
            return True
        logger.debug("本文フレーズ優先: 候補 #%d ではFormsデータを取得できませんでした。", idx)
    return False


def _process_candidates_via_attachments(
    robot: "ChoujiRobo",
    entries: List[MailEnvelope],
    logger: logging.Logger,
    match_mode: str,
) -> bool:
    for idx, entry in enumerate(entries, start=1):
        logger.info(
            "添付検索(%s): 候補 #%d を確認します (受信=%s 件名=%s)",
            match_mode,
            idx,
            entry.received_at,
            entry.subject,
        )
        if _process_entry_via_attachments(robot, entry, [match_mode], logger):
            return True
        logger.debug("添付検索(%s): 候補 #%d では一致しませんでした。", match_mode, idx)
    return False


def _process_entry_via_forms(robot: "ChoujiRobo", entry: MailEnvelope) -> bool:
    logger = logging.getLogger("chouji_robo.mail")
    robot.state.selected_mail_entry = entry
    mail_item = getattr(entry, "raw_item", None)
    if mail_item is not None:
        _populate_mail_metadata(robot, mail_item, logger)

    body = robot._safe_str(entry.body)
    url_match = _extract_first_url(body)
    if not url_match:
        logger.debug("Forms転記対象: URLが見つからないためスキップします。")
        return False

    logger.info("Forms転記対象: URL=%s を取得し、SharePointから行データを取得します。", url_match)
    robot.state.forms_row = robot._fetch_row_from_url(url_match, robot.state.pin)
    if not robot.state.forms_row:
        logger.warning("Forms転記対象: J列のPIN行が取得できませんでした。次の候補を確認します。")
        return False

    robot._write_forms_row_to_temp_book()
    print("[INFO] Ae.get_mail: Forms転記シート用の情報を取得しました。")
    return True


def _process_entry_via_attachments(
    robot: "ChoujiRobo",
    entry: MailEnvelope,
    modes: List[str],
    logger: logging.Logger,
) -> bool:
    robot.state.selected_mail_entry = entry
    mail_item = getattr(entry, "raw_item", None)
    if mail_item is not None:
        _populate_mail_metadata(robot, mail_item, logger)

    for mode in modes:
        if _handle_mail_attachments(robot, entry, match_mode=mode):
            logger.info("添付検索(%s): 必要情報を取得しました。", mode)
            print(f"[INFO] Ae.get_mail: 添付ファイル検索({mode})で必要情報を取得しました。")
            return True
    return False

def _populate_mail_metadata(robot: "ChoujiRobo", mail_item, logger: logging.Logger) -> None:
    sender = robot._safe_str(getattr(mail_item, "SenderEmailAddress", "")).strip()
    cc_addresses = _extract_recipient_addresses(robot, mail_item, recipient_type=2, logger=logger)
    bcc_addresses = _extract_recipient_addresses(robot, mail_item, recipient_type=3, logger=logger)
    reply_body = _build_reply_body(robot, mail_item, logger)

    robot.state.mail_sender = sender
    robot.state.mail_cc = "; ".join(cc_addresses)
    robot.state.mail_bcc = "; ".join(bcc_addresses)
    robot.state.reply_email_body = reply_body

    logger.info(
        "メールメタデータ: sender=%s cc_count=%d bcc_count=%d reply_len=%d",
        sender,
        len(cc_addresses),
        len(bcc_addresses),
        len(reply_body),
    )
    print(f"[INFO] mail_sender: {sender or '(empty)'}")
    print(f"[INFO] mail_cc: {robot.state.mail_cc or '(empty)'}")
    print(f"[INFO] mail_bcc: {robot.state.mail_bcc or '(empty)'}")
    print("[INFO] reply_email_body:")
    if reply_body:
        print(reply_body)
    else:
        print("(empty)")


def _get_primary_smtp_from_accessor(recipient, logger: logging.Logger) -> str:
    try:
        accessor = getattr(recipient, "PropertyAccessor", None)
        if accessor is None:
            return ""
        value = accessor.GetProperty(SMTP_PROPERTY_URI)
        value = str(value).strip() if value is not None else ""
        if value and "@" in value:
            return value
    except Exception as exc:
        logger.debug("PropertyAccessor から SMTP を取得できませんでした: %s", exc)
    return ""


def _extract_recipient_addresses(robot: "ChoujiRobo", mail_item, recipient_type: int, logger: logging.Logger) -> list[str]:
    addresses: list[str] = []
    recipients = getattr(mail_item, "Recipients", None)
    if recipients is None:
        return addresses
    try:
        count = recipients.Count
    except Exception:
        count = 0
    for index in range(1, count + 1):
        try:
            recipient = recipients.Item(index)
        except Exception as exc:
            logger.debug("Recipients.Item(%d) 取得に失敗: %s", index, exc)
            continue
        try:
            if getattr(recipient, "Type", None) != recipient_type:
                continue
        except Exception:
            continue
        email = _resolve_recipient_address(robot, recipient, logger)
        if email and "@" in email:
            addresses.append(email)
    logger.debug("recipient_type=%d resolved addresses: %s", recipient_type, addresses)
    return addresses


def _resolve_recipient_address(robot: "ChoujiRobo", recipient, logger: logging.Logger) -> str:
    email = _get_primary_smtp_from_accessor(recipient, logger)
    if email:
        return email

    address_entry = None
    try:
        address_entry = recipient.AddressEntry
    except Exception as exc:
        logger.debug("AddressEntry を取得できませんでした: %s", exc)

    if address_entry is not None:
        for attr in ("PrimarySmtpAddress", "SMTPAddress", "Address"):
            try:
                value = getattr(address_entry, attr, "")
                value = robot._safe_str(value).strip()
                if value and "@" in value:
                    email = value
                    break
            except Exception:
                continue
        if not email:
            try:
                exchange_user = address_entry.GetExchangeUser()
                if exchange_user is not None:
                    for attr in ("PrimarySmtpAddress", "SMTPAddress", "Address"):
                        primary = robot._safe_str(getattr(exchange_user, attr, "")).strip()
                        if primary and "@" in primary:
                            email = primary
                            break
            except Exception:
                pass
        if not email:
            try:
                distribution = address_entry.GetExchangeDistributionList()
                if distribution is not None:
                    for attr in ("PrimarySmtpAddress", "SMTPAddress", "Address"):
                        fallback = robot._safe_str(getattr(distribution, attr, "")).strip()
                        if fallback and "@" in fallback:
                            email = fallback
                            break
            except Exception:
                pass
        if not email:
            try:
                contact = address_entry.GetContact()
                if contact is not None:
                    for attr in (
                        "PrimarySmtpAddress",
                        "Email1Address",
                        "Email2Address",
                        "Email3Address",
                    ):
                        fallback = robot._safe_str(getattr(contact, attr, "")).strip()
                        if fallback and "@" in fallback:
                            email = fallback
                            break
            except Exception:
                pass

    if not email:
        email = _get_primary_smtp_from_accessor(recipient, logger)

    if not email:
        try:
            raw_address = robot._safe_str(getattr(recipient, "Address", "")).strip()
        except Exception:
            raw_address = ""
        if raw_address and "@" in raw_address:
            email = raw_address

    if email and "@" not in email:
        logger.debug("未解決アドレスを検出: %s", email)
        return ""

    return email


def _build_reply_body(robot: "ChoujiRobo", mail_item, logger: logging.Logger) -> str:
    reply_body = ""
    reply_item = None
    try:
        reply_item = mail_item.Reply()
        reply_body = robot._safe_str(getattr(reply_item, "Body", ""))
        if not reply_body.strip():
            reply_body = robot._safe_str(getattr(reply_item, "HTMLBody", ""))
    except Exception as exc:
        logger.debug("Reply draft generation failed, fallback to original body: %s", exc)
    finally:
        if reply_item is not None:
            try:
                reply_item.Close(0)  # 0 => olDiscard
            except Exception as close_exc:
                logger.debug("Reply draft close failed (ignored): %s", close_exc)

    if not reply_body.strip():
        try:
            reply_body = robot._safe_str(getattr(mail_item, "Body", ""))
        except Exception as exc:
            logger.debug("Original mail body fallback failed: %s", exc)
            reply_body = ""

    return reply_body

def _build_outlook_restriction(start: datetime, end: datetime) -> str:
    def fmt(value: datetime) -> str:
        return value.strftime("%m/%d/%Y %I:%M %p")

    return f"[ReceivedTime] >= '{fmt(start)}' AND [ReceivedTime] <= '{fmt(end)}'"


def _normalize_datetime(value: datetime) -> datetime:
    if value.tzinfo is not None:
        try:
            value = value.astimezone()
        except Exception:
            pass
        return value.replace(tzinfo=None)
    return value


def _seconds_difference(left: datetime, right: datetime) -> float:
    return abs((_normalize_datetime(left) - _normalize_datetime(right)).total_seconds())


def _collect_recent_messages(
    collections, anchor: datetime, seconds: int, logger: logging.Logger
) -> List[MailEnvelope]:
    window_start = anchor - timedelta(seconds=seconds)
    window_end = anchor + timedelta(seconds=seconds)
    restriction = _build_outlook_restriction(window_start, window_end)
    recent: List[MailEnvelope] = []
    for label, items in collections:
        try:
            scoped_items = items.Restrict(restriction)
        except Exception as exc:
            logger.debug("%s: フォールバックRestrictに失敗: %s", label, exc)
            continue
        for item in scoped_items:
            envelope = _build_envelope_from_item(
                item, logger, label=f"{label}フォールバック候補"
            )
            if envelope:
                recent.append(envelope)

    return recent


def _build_envelope_from_item(item, logger: logging.Logger, label: str) -> Optional[MailEnvelope]:
    if item is None:
        return None
    try:
        return MailEnvelope(
            entry_id=str(item.EntryID),
            subject=str(item.Subject),
            sender=str(item.SenderName),
            received_at=_convert_outlook_time(item.ReceivedTime),
            body=str(item.Body),
            raw_item=item,
        )
    except Exception as exc:
        logger.warning("%sの解析に失敗しました: %s", label, exc)
        return None


def _sort_and_log_candidates(
    logger: logging.Logger,
    anchor: datetime,
    candidates: List[MailEnvelope],
    label: str,
) -> List[MailEnvelope]:
    if not candidates:
        logger.info("%s: 候補メールはありません。", label)
        return []

    sorted_candidates = sorted(
        candidates,
        key=lambda entry: _seconds_difference(entry.received_at, anchor),
    )
    logger.info("%s: %d件の候補(管理時刻=%s)", label, len(sorted_candidates), anchor)
    for entry in sorted_candidates:
        logger.info("  受信=%s 件名=%s", entry.received_at, entry.subject)
    return sorted_candidates

def _find_nearest_message(
    collections, anchor: datetime, logger: logging.Logger, limit: int = 200
) -> Optional[MailEnvelope]:
    envelopes: List[MailEnvelope] = []
    count = 0
    for label, items in collections:
        try:
            items.IncludeRecurrences = True
            items.Sort("[ReceivedTime]", True)
        except Exception as exc:
            logger.debug("%s: 最寄りメール探索の準備に失敗: %s", label, exc)
        try:
            item = items.GetFirst()
        except Exception as exc:
            logger.debug("%s: 最寄りメール探索の初期化に失敗: %s", label, exc)
            continue
        while item is not None and count < limit:
            envelope = _build_envelope_from_item(
                item, logger, label=f"{label}近似候補"
            )
            if envelope:
                envelopes.append(envelope)
            try:
                item = items.GetNext()
            except Exception:
                break
            count += 1
        if count >= limit:
            break

    if not envelopes:
        return None

    return min(envelopes, key=lambda entry: _seconds_difference(entry.received_at, anchor))


def _gather_mail_sources(namespace, logger: logging.Logger):
    sources: List[tuple[str, Any]] = []
    visited: set[str] = set()

    def collect(folder, label_chain: List[str]) -> None:
        try:
            entry_id = str(getattr(folder, "EntryID", ""))
        except Exception:
            entry_id = ""
        if entry_id and entry_id in visited:
            return
        if entry_id:
            visited.add(entry_id)

        folder_name = str(getattr(folder, "Name", "Folder"))
        label = "/".join(label_chain + [folder_name]) if label_chain else folder_name

        default_item_type = getattr(folder, "DefaultItemType", None)
        default_message_class = str(getattr(folder, "DefaultMessageClass", "") or "")
        is_mail_folder = default_item_type == 0 or default_message_class.startswith("IPM.Note")

        try:
            items = folder.Items
        except Exception:
            items = None
        if items is not None and is_mail_folder:
            try:
                items.IncludeRecurrences = True
            except Exception:
                pass
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception as exc:
                logger.debug("%s: Items の並び替えに失敗: %s", label, exc)
            sources.append((label, items))
        elif items is not None:
            logger.debug("%s: メール以外のフォルダのため検索対象から除外します (DefaultItemType=%s DefaultMessageClass=%s)", label, default_item_type, default_message_class)

        try:
            subfolders = folder.Folders
        except Exception:
            return

        try:
            count = subfolders.Count
        except Exception:
            count = 0

        for idx in range(1, count + 1):
            try:
                sub = subfolders.Item(idx)
            except Exception as exc:
                logger.debug("%s: サブフォルダー取得に失敗: %s", label, exc)
                continue
            collect(sub, label_chain + [folder_name])

    try:
        stores = getattr(namespace, "Stores", None)
        if stores is not None:
            total = stores.Count
            for index in range(1, total + 1):
                store = stores.Item(index)
                store_name = str(
                    getattr(store, "DisplayName", getattr(store, "Name", f"Store{index}"))
                )
                try:
                    root = store.GetRootFolder()
                except Exception:
                    root = getattr(store, "Folders", None)
                if root is None:
                    continue
                collect(root, [store_name])
        else:
            folders = namespace.Folders
            for index in range(1, folders.Count + 1):
                folder = folders.Item(index)
                folder_name = str(getattr(folder, "Name", f"Folder{index}"))
                collect(folder, [folder_name])
    except Exception as exc:
        logger.debug("メールフォルダの列挙に失敗: %s", exc)

    if not sources:
        try:
            fallback = namespace.GetDefaultFolder(6)
            collect(fallback, ["DefaultInbox"])
        except Exception as exc:
            logger.debug("デフォルト受信トレイの確保に失敗: %s", exc)
    return sources


def _convert_outlook_time(value: object) -> datetime:
    if isinstance(value, datetime):
        return value
    return datetime.fromtimestamp(time.mktime(time.strptime(str(value)[:19], "%Y-%m-%d %H:%M:%S")))


def _select_target_mail(robot: "ChoujiRobo", candidates: List[MailEnvelope]) -> MailEnvelope:
    logger = logging.getLogger("chouji_robo.mail")
    pin_phrase = f"弔事の発生した従業員：{robot.state.pin}"
    for entry in candidates:
        if pin_phrase in entry.body:
            logger.debug("PIN フレーズを含むメールを自動選択しました: %s", entry.subject)
            return entry

    descriptions = [
        f"{idx+1}: {entry.received_at:%Y-%m-%d %H:%M} | {entry.subject} | {entry.sender}"
        for idx, entry in enumerate(candidates)
    ]
    previews = []
    for entry in candidates:
        body_preview = (entry.body or "").strip()
        if len(body_preview) > 2000:
            body_preview = body_preview[:2000] + "..."
        previews.append(body_preview or "(本文なし)")

    if tk is not None and robot.root is not None:
        choice = _prompt_mail_choice(robot, descriptions, previews)
    else:
        prompt = "\n".join(descriptions)
        choice = simpledialog.askinteger(
            "メールの選択",
            f"どのメールを対応したいですか？\n{prompt}",
            parent=robot.root,
            minvalue=1,
            maxvalue=len(candidates),
        )

    if choice is None:
        raise RuntimeError("メールが選択されなかったため中断します。")
    return candidates[choice - 1]


def _prompt_mail_choice(robot: "ChoujiRobo", descriptions: list[str], previews: list[str]) -> Optional[int]:
    if tk is None:
        return None

    result = {"value": None}
    top = tk.Toplevel(robot.root)
    top.title("メールの選択")
    top.grab_set()

    var = tk.IntVar(value=1)

    list_frame = tk.Frame(top)
    list_frame.pack(fill="both", expand=False, padx=10, pady=10)

    def update_preview(idx: int) -> None:
        preview_box.configure(state="normal")
        preview_box.delete("1.0", tk.END)
        preview_box.insert(tk.END, previews[idx - 1])
        preview_box.configure(state="disabled")

    for idx, label in enumerate(descriptions, start=1):
        rb = tk.Radiobutton(
            list_frame,
            text=label,
            variable=var,
            value=idx,
            justify="left",
            anchor="w",
            command=lambda i=idx: update_preview(i),
            wraplength=500,
        )
        rb.pack(fill="x", anchor="w")

    preview_box = tk.Text(top, height=12, width=80)
    preview_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    preview_box.insert(tk.END, previews[0])
    preview_box.configure(state="disabled")

    button_frame = tk.Frame(top)
    button_frame.pack(fill="x", pady=(0, 10))

    def on_ok() -> None:
        result["value"] = var.get()
        top.destroy()

    def on_cancel() -> None:
        result["value"] = None
        top.destroy()

    tk.Button(button_frame, text="決定", command=on_ok, width=10).pack(side="left", padx=5)
    tk.Button(button_frame, text="キャンセル", command=on_cancel, width=10).pack(side="right", padx=5)

    robot.root.wait_window(top)
    return result["value"]
def _extract_first_url(text: str) -> Optional[str]:
    pattern = re.compile(r"https?://[^\s<>\"']+")
    match = pattern.search(text)
    return match.group(0) if match else None


def _handle_mail_attachments(
    robot: "ChoujiRobo",
    entry: MailEnvelope,
    match_mode: str = "any",
) -> bool:
    logger = logging.getLogger("chouji_robo.mail")
    item = entry.raw_item
    attachments = getattr(item, "Attachments", None)
    if attachments is None or attachments.Count == 0:
        logger.debug("添付ファイルはありませんでした。")
        return False

    temp_dir = robot.paths.temp_forms_book.parent
    robot._ensure_directory(temp_dir)
    logger.debug("添付ファイルを順次確認します (count=%d)", attachments.Count)
    print(f"[INFO] Ae.get_mail: 添付ファイルチェック開始 (count={attachments.Count})")

    for index in range(1, attachments.Count + 1):
        attachment = attachments.Item(index)
        original_name = robot._safe_str(getattr(attachment, "FileName", "")) or f"attachment_{index}"
        suffix = Path(original_name).suffix.lower()
        if suffix not in {".xlsx", ".xlsm"}:
            logger.debug("添付ファイル %s は対象外の形式のためスキップします。", original_name)
            print(f"[INFO] Ae.get_mail: 添付 {original_name} は対象外 (suffix={suffix}) -> スキップ")
            continue

        temp_path = temp_dir / f"_attachment_{int(time.time())}_{index}{suffix}"
        attachment.SaveAsFile(str(temp_path))
        logger.info("添付ファイルを一時保存しました: %s", temp_path)
        print(f"[INFO] Ae.get_mail: 添付ファイルを一時保存しました -> {temp_path}")
        try:
            if _attachment_matches(robot, temp_path, match_mode=match_mode):
                logger.info("条件に該当する添付ファイルを temp_弔事連絡票.xlsx として保存しました。")
                print("[INFO] Ae.get_mail: 添付ファイルからPINまたは氏名を検出しました。")
                return True
        finally:
            temp_path.unlink(missing_ok=True)

    logger.debug("PIN/氏名に一致する添付ファイルは見つかりませんでした。")
    print("[WARN] Ae.get_mail: 添付ファイルから一致する情報は見つかりませんでした。")
    return False


def _attachment_matches(
    robot: "ChoujiRobo",
    file_path,
    match_mode: str = "any",
) -> bool:
    logger = logging.getLogger("chouji_robo.mail")
    normalized_name = robot._normalize_name(robot.state.name_katakana)
    normalized_pin = robot._normalize_name(robot.state.pin)
    use_name = match_mode in ("name", "any") and bool(normalized_name)
    use_pin = match_mode in ("pin", "any") and bool(normalized_pin)
    if not use_name and not use_pin:
        logger.debug(
            "�Y�t����(%s): ���p�ł���L�[������܂��� (name=%s pin=%s)",
            match_mode,
            normalized_name or "(none)",
            normalized_pin or "(none)",
        )
        return False

    logger.debug(
        "�Y�t����(%s) �L�[: name=%s pin=%s",
        match_mode,
        normalized_name or "(none)",
        normalized_pin or "(none)",
    )

    trace_lines: list[str] = []
    try:
        with open_workbook(file_path, read_only=True) as workbook:
            target_sheets = list(workbook.Worksheets)
            if not target_sheets:
                message = "添付ファイルにシートが見つかりませんでした。"
                logger.error(message)
                raise RuntimeError(message)

            sheet_names_snapshot = [
                robot._safe_str(getattr(sheet, "Name", "")) for sheet in target_sheets
            ]
            logger.debug("添付ブックのシート候補: %s", sheet_names_snapshot)
            for sheet in target_sheets:
                sheet_name = str(getattr(sheet, "Name", ""))
                logger.debug("�V�[�g %s ���������܂�", sheet_name)
                for row_index, row in iter_rows(sheet, start_row=1):
                    row_text = " ".join(robot._safe_str(value) for value in row)
                    trace_lines.append(f"[{sheet_name}:{row_index}] {row_text}")
                    normalized_row = robot._normalize_name(row_text)
                    if use_pin and normalized_pin and normalized_pin in normalized_row:
                        robot.state.forms_row = list(row)
                        robot._write_forms_row_to_temp_book(from_workbook=file_path)
                        logger.info("�Y�t�t�@�C������ PIN �s���擾���܂��� (sheet=%s)�B", sheet_name)
                        return True
                    if use_name and normalized_name and normalized_name in normalized_row:
                        robot.state.forms_row = list(row)
                        robot._write_forms_row_to_temp_book(from_workbook=file_path)
                        logger.info("�Y�t�t�@�C�����玁���s���擾���܂��� (sheet=%s)�B", sheet_name)
                        return True
    except Exception as exc:
        message = f"�Y�t�t�@�C���� Excel �Ƃ��ēǂݍ��߂܂���ł���: {exc}"
        logger.error(message)
        raise RuntimeError(message) from exc

    logger.error("�Y�t�t�@�C�������v����s�����ł��܂���ł����B")
    for line in trace_lines:
        logger.error("  %s", line)
    return False
