"""Common data structures and constants for chouji_ROBO step-A automation."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, List, Optional

HEARTBEAT_INTERVAL_MS = 5000
FORCE_STOP_POLL_MS = 500
LOG_QUEUE_POLL_MS = 120


@dataclass
class StepAState:
    """Shared state manipulated across the step-A workflow."""

    tehai_number: Optional[str] = None
    company_name: Optional[str] = None
    pin: Optional[str] = None
    mail_time: Optional[datetime] = None
    name_katakana: Optional[str] = None
    forms_row: List[Any] = field(default_factory=list)
    selected_mail_entry: Optional["MailEnvelope"] = None
    mail_sender: str = ""
    mail_cc: str = ""
    mail_bcc: str = ""
    reply_email_body: str = ""
    workflow_started_at: Optional[datetime] = None
    workflow_finished_at: Optional[datetime] = None
    workflow_error: str = ""
    workflow_feedback: str = ""
    generated_pdf_path: Optional[str] = None
    generated_excel_path: Optional[str] = None
    outlook_draft_entry_id: str = ""
    outlook_draft_store_id: str = ""
    edge_process_pid: Optional[int] = None


@dataclass
class MailEnvelope:
    """Lightweight representation of an Outlook mail item."""

    entry_id: str
    subject: str
    sender: str
    received_at: datetime
    body: str
    raw_item: Any = field(repr=False)


@dataclass
class PathRegistry:
    """Resolves filesystem locations while keeping the user root dynamic."""

    home: Path = field(default_factory=Path.home)

    @property
    def desktop_root(self) -> Path:
        return self.home / "Desktop" / "【全社標準】弔事対応フォルダ"

    @property
    def rpa_local_dir(self) -> Path:
        return self.desktop_root / "1. 下処理"

    @property
    def rpa_local_book(self) -> Path:
        return self.rpa_local_dir / "RPAブック下処理.xlsx"

    @property
    def temp_forms_book(self) -> Path:
        return self.rpa_local_dir / "temp_弔事連絡票.xlsx"

    @property
    def temp_input_sheet_dir(self) -> Path:
        return self.rpa_local_dir / "temp_手配入力シート.xlsx"

    @property
    def rpa_book_destination(self) -> Path:
        return self.desktop_root / "2. RPAブック" / "RPAブック.xlsx"

    @property
    def rpa_book_dir(self) -> Path:
        return self.rpa_book_destination.parent

    @property
    def panasonic_root(self) -> Path:
        return self.home / "Panasonic"

    @property
    def kanri_report_book(self) -> Path:
        return (
            self.panasonic_root
            / "TM.PHR_HR-PRIDEメンバー - 03_管理表"
            / "【管理表】 業務報告.xlsx"
        )

    def panasonic_rpa_book(self) -> Path:
        return (
            self.panasonic_root
            / "TM.PHR_HR-PRIDEメンバー - 標準化"
            / "1. 下処理"
            / "RPAブック下処理.xlsx"
        )

    def company_input_sheet(self, company_name: str) -> Path:
        return (
            self.panasonic_root
            / "TM.PHR_HR-PRIDEメンバー - 標準化"
            / "1. 下処理"
            / f"{company_name}手配入力シート.xlsx"
        )

    def company_forms_sheet(self, company_name: str) -> Path:
        return (
            self.panasonic_root
            / "TM.PHR_HR-PRIDEメンバー - 標準化"
            / "1. 下処理"
            / f"{company_name}_Forms転記シート.xlsx"
        )

    @property
    def error_log_book(self) -> Path:
        return (
            self.panasonic_root
            / "TM.PHR_HR-PRIDEメンバー - 標準化"
            / "3. 過去のデータ"
            / "error_log.xlsx"
        )

    def company_archive_dir(self, company_name: str) -> Path:
        return (
            self.panasonic_root
            / "TM.PHR_集中化 - General"
            / f"{company_name}弔事関連"
            / "99_ロボ弔事格納"
        )
