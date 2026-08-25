from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from string import Formatter
from threading import RLock
from typing import Any, Protocol
from uuid import uuid4

from gns_app.config import Settings
from gns_app.diagnostics import record_event, record_exception
from gns_app.runtime_commands import module_command
from gns_app.database import Database, utc_now
from gns_app.services.storage import (
    StorageError,
    ensure_within,
    sanitize_filename,
    save_pdf_stream,
)
from gns_app.services.windows_focus import focus_window_handle


OL_FOLDER_INBOX = 6
OL_FOLDER_DRAFTS = 16
OL_MAIL_ITEM = 0
OL_BY_VALUE = 1
OL_TEXT = 1
GNS_DRAFT_KEY_PROPERTY = "GNS_APP_DRAFT_KEY"
ACCOUNT_TYPE_LABELS = {
    0: "Exchange",
    1: "IMAP",
    2: "POP3",
    3: "HTTP",
    4: "Exchange ActiveSync",
    5: "Другой",
}
EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
PR_INTERNET_MESSAGE_ID = (
    "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
)
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"


class OutlookIntegrationError(RuntimeError):
    """Безопасная для показа пользователю ошибка интеграции Outlook."""


class OutlookUnsupportedPlatformError(OutlookIntegrationError):
    pass


class OutlookComponentMissingError(OutlookIntegrationError):
    pass


class OutlookConnectionError(OutlookIntegrationError):
    pass


class OutlookProbeTimeoutError(OutlookIntegrationError):
    pass


@dataclass(frozen=True, slots=True)
class OutlookScannedAttachment:
    attachment_index: int
    original_filename: str
    temporary_path: str
    size_bytes: int
    error_message: str = ""

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "OutlookScannedAttachment":
        return cls(
            attachment_index=int(payload.get("attachment_index") or 0),
            original_filename=str(payload.get("original_filename") or "document.pdf"),
            temporary_path=str(payload.get("temporary_path") or ""),
            size_bytes=int(payload.get("size_bytes") or 0),
            error_message=str(payload.get("error_message") or ""),
        )


@dataclass(frozen=True, slots=True)
class OutlookScannedMessage:
    source_key: str
    sender_smtp: str
    received_at: str
    attachment_count: int
    pdf_attachment_count: int
    attachments: tuple[OutlookScannedAttachment, ...]

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "OutlookScannedMessage":
        return cls(
            source_key=str(payload.get("source_key") or ""),
            sender_smtp=str(payload.get("sender_smtp") or ""),
            received_at=str(payload.get("received_at") or ""),
            attachment_count=int(payload.get("attachment_count") or 0),
            pdf_attachment_count=int(payload.get("pdf_attachment_count") or 0),
            attachments=tuple(
                OutlookScannedAttachment.from_dict(item)
                for item in payload.get("attachments", [])
                if isinstance(item, dict)
            ),
        )


@dataclass(frozen=True, slots=True)
class OutlookInboxScan:
    inspected_mail_count: int
    eligible_message_count: int
    known_message_count: int
    scan_error_count: int
    messages: tuple[OutlookScannedMessage, ...]

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "OutlookInboxScan":
        return cls(
            inspected_mail_count=int(payload.get("inspected_mail_count") or 0),
            eligible_message_count=int(payload.get("eligible_message_count") or 0),
            known_message_count=int(payload.get("known_message_count") or 0),
            scan_error_count=int(payload.get("scan_error_count") or 0),
            messages=tuple(
                OutlookScannedMessage.from_dict(item)
                for item in payload.get("messages", [])
                if isinstance(item, dict)
            ),
        )


@dataclass(frozen=True, slots=True)
class OutlookAccountInfo:
    display_name: str
    smtp_address: str
    account_type: str


@dataclass(frozen=True, slots=True)
class OutlookProbe:
    outlook_version: str
    profile_name: str
    current_user: str
    default_store: str
    inbox_name: str
    inbox_path: str
    inbox_item_count: int
    offline: bool
    sync_group_count: int
    accounts: tuple[OutlookAccountInfo, ...]
    sync_requested: bool = False

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "OutlookProbe":
        accounts = tuple(
            OutlookAccountInfo(
                display_name=str(item.get("display_name") or ""),
                smtp_address=str(item.get("smtp_address") or ""),
                account_type=str(item.get("account_type") or "Другой"),
            )
            for item in payload.get("accounts", [])
            if isinstance(item, dict)
        )
        return cls(
            outlook_version=str(payload.get("outlook_version") or ""),
            profile_name=str(payload.get("profile_name") or ""),
            current_user=str(payload.get("current_user") or ""),
            default_store=str(payload.get("default_store") or ""),
            inbox_name=str(payload.get("inbox_name") or ""),
            inbox_path=str(payload.get("inbox_path") or ""),
            inbox_item_count=int(payload.get("inbox_item_count") or 0),
            offline=bool(payload.get("offline", False)),
            sync_group_count=int(payload.get("sync_group_count") or 0),
            accounts=accounts,
            sync_requested=bool(payload.get("sync_requested", False)),
        )


@dataclass(frozen=True, slots=True)
class OutlookDiagnosticResult:
    state: str
    checked_at: str
    message: str
    probe: OutlookProbe | None = None

    @property
    def successful(self) -> bool:
        return self.state in {"ready", "sync_requested"}


@dataclass(frozen=True, slots=True)
class OutlookDraftResult:
    entry_id: str
    recipient_email: str
    subject: str
    attachment_name: str
    existing: bool = False

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "OutlookDraftResult":
        return cls(
            entry_id=str(payload.get("entry_id") or ""),
            recipient_email=str(payload.get("recipient_email") or ""),
            subject=str(payload.get("subject") or ""),
            attachment_name=str(payload.get("attachment_name") or ""),
            existing=bool(payload.get("existing", False)),
        )


class OutlookGateway(Protocol):
    def inspect(self, *, request_sync: bool = False) -> OutlookProbe: ...

    def create_draft(
        self,
        *,
        draft_key: str,
        recipient_email: str,
        subject: str,
        body: str,
        attachment_path: Path,
        attachment_name: str,
    ) -> OutlookDraftResult: ...


class OutlookWorkflow(Protocol):
    def get_active_employee(self) -> str: ...

    def create_upload(self, original_filename: str, stream: Any) -> str: ...

    def get_response_letter(self, letter_id: str) -> dict[str, Any] | None: ...

    def get_confirmed_signed_response_scan(
        self, letter_id: str
    ) -> dict[str, Any] | None: ...

    def match_gns_office(self, text: str) -> dict[str, Any] | None: ...


def _safe_text(value: Any) -> str:
    try:
        return str(value or "").strip()
    except Exception:
        return ""


def _safe_attribute(value: Any, name: str, default: Any = "") -> Any:
    try:
        return getattr(value, name)
    except Exception:
        return default


class PyWin32OutlookGateway:
    """Низкоуровневый Outlook COM-адаптер, запускаемый в дочернем процессе."""

    @staticmethod
    def _ensure_outlook_window(application: Any, inbox: Any) -> None:
        explorers = _safe_attribute(application, "Explorers", None)
        if explorers is None:
            return
        if int(_safe_attribute(explorers, "Count", 0) or 0) > 0:
            return
        try:
            inbox.Display()
        except Exception:
            # Видимое окно повышает устойчивость Outlook 2013, но невозможность
            # открыть Explorer не должна подменять результат чтения папки.
            return

    @staticmethod
    def _select_inbox(namespace: Any, mailbox: str) -> Any:
        mailbox_name = mailbox.strip().casefold()
        if not mailbox_name:
            return namespace.GetDefaultFolder(OL_FOLDER_INBOX)

        stores = _safe_attribute(namespace, "Stores", None)
        store_count = int(_safe_attribute(stores, "Count", 0) or 0)
        exact_matches: list[Any] = []
        partial_matches: list[Any] = []
        for index in range(1, store_count + 1):
            try:
                store = stores.Item(index)
                display_name = _safe_text(
                    _safe_attribute(store, "DisplayName")
                ).casefold()
                if display_name == mailbox_name:
                    exact_matches.append(store)
                elif mailbox_name in display_name:
                    partial_matches.append(store)
            except Exception:
                continue
        matches = exact_matches or partial_matches
        if len(matches) != 1:
            if not matches:
                raise OutlookConnectionError(
                    "Указанный почтовый ящик не найден в текущем профиле Outlook."
                )
            raise OutlookConnectionError(
                "Имя почтового ящика совпало с несколькими хранилищами Outlook. "
                "Укажите его точное имя."
            )
        try:
            return matches[0].GetDefaultFolder(OL_FOLDER_INBOX)
        except Exception as exc:
            raise OutlookConnectionError(
                "Outlook не открыл папку «Входящие» выбранного ящика."
            ) from exc

    def inspect(self, *, request_sync: bool = False) -> OutlookProbe:
        if sys.platform != "win32":
            raise OutlookUnsupportedPlatformError(
                "Интеграция Outlook доступна только в Windows."
            )
        try:
            import pythoncom  # type: ignore[import-not-found]
            from win32com import client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise OutlookComponentMissingError(
                "Не установлен локальный компонент связи с Outlook. "
                "Повторно запустите START.bat."
            ) from exc

        pythoncom.CoInitialize()
        try:
            application = client.Dispatch("Outlook.Application")
            namespace = application.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
            self._ensure_outlook_window(application, inbox)

            accounts: list[OutlookAccountInfo] = []
            account_collection = _safe_attribute(namespace, "Accounts", None)
            account_count = int(
                _safe_attribute(account_collection, "Count", 0) or 0
            )
            for index in range(1, account_count + 1):
                try:
                    account = account_collection.Item(index)
                    account_type = int(
                        _safe_attribute(account, "AccountType", 5) or 0
                    )
                    accounts.append(
                        OutlookAccountInfo(
                            display_name=_safe_text(
                                _safe_attribute(account, "DisplayName")
                            ),
                            smtp_address=_safe_text(
                                _safe_attribute(account, "SmtpAddress")
                            ),
                            account_type=ACCOUNT_TYPE_LABELS.get(
                                account_type,
                                f"Другой ({account_type})",
                            ),
                        )
                    )
                except Exception:
                    # Ошибка одного дополнительного аккаунта не должна скрывать
                    # доступность основной папки входящих.
                    continue

            if request_sync:
                # Outlook выполняет SendAndReceive асинхронно. На этом этапе
                # фиксируется только успешная передача команды приложению.
                namespace.SendAndReceive(False)

            current_user = _safe_attribute(namespace, "CurrentUser", None)
            default_store = _safe_attribute(namespace, "DefaultStore", None)
            sync_objects = _safe_attribute(namespace, "SyncObjects", None)
            items = _safe_attribute(inbox, "Items", None)
            return OutlookProbe(
                outlook_version=_safe_text(
                    _safe_attribute(application, "Version")
                ),
                profile_name=_safe_text(
                    _safe_attribute(namespace, "CurrentProfileName")
                ),
                current_user=_safe_text(
                    _safe_attribute(current_user, "Name")
                ),
                default_store=_safe_text(
                    _safe_attribute(default_store, "DisplayName")
                ),
                inbox_name=_safe_text(_safe_attribute(inbox, "Name")),
                inbox_path=_safe_text(_safe_attribute(inbox, "FolderPath")),
                inbox_item_count=int(_safe_attribute(items, "Count", 0) or 0),
                offline=bool(_safe_attribute(namespace, "Offline", False)),
                sync_group_count=int(
                    _safe_attribute(sync_objects, "Count", 0) or 0
                ),
                accounts=tuple(accounts),
                sync_requested=request_sync,
            )
        except OutlookIntegrationError:
            raise
        except Exception as exc:
            raise OutlookConnectionError(
                "Не удалось подключиться к Outlook. Откройте Outlook вручную, "
                "убедитесь, что выбран рабочий профиль, и повторите проверку."
            ) from exc
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _draft_key(mail: Any) -> str:
        properties = _safe_attribute(mail, "UserProperties", None)
        if properties is None:
            return ""
        try:
            property_item = properties.Find(GNS_DRAFT_KEY_PROPERTY)
        except Exception:
            return ""
        return _safe_text(_safe_attribute(property_item, "Value"))

    @classmethod
    def _find_existing_draft(cls, namespace: Any, draft_key: str) -> Any | None:
        drafts = namespace.GetDefaultFolder(OL_FOLDER_DRAFTS)
        items = _safe_attribute(drafts, "Items", None)
        item_count = int(_safe_attribute(items, "Count", 0) or 0)
        for index in range(1, item_count + 1):
            try:
                candidate = items.Item(index)
                if cls._draft_key(candidate) == draft_key:
                    return candidate
            except Exception:
                continue
        return None

    def create_draft(
        self,
        *,
        draft_key: str,
        recipient_email: str,
        subject: str,
        body: str,
        attachment_path: Path,
        attachment_name: str,
    ) -> OutlookDraftResult:
        if sys.platform != "win32":
            raise OutlookUnsupportedPlatformError(
                "Интеграция Outlook доступна только в Windows."
            )
        if not EMAIL_RE.fullmatch(recipient_email.strip().casefold()):
            raise OutlookConnectionError("Получатель Outlook указан неверно.")
        if not draft_key.strip() or len(draft_key) > 128:
            raise OutlookConnectionError("Ключ черновика Outlook указан неверно.")
        if not subject.strip() or len(subject) > 255 or "\n" in subject or "\r" in subject:
            raise OutlookConnectionError("Тема письма Outlook указана неверно.")
        source = attachment_path.resolve()
        if not source.is_file():
            raise OutlookConnectionError("Файл для вложения Outlook не найден.")
        try:
            import pythoncom  # type: ignore[import-not-found]
            from win32com import client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise OutlookComponentMissingError(
                "Не установлен локальный компонент связи с Outlook. "
                "Повторно запустите START.bat."
            ) from exc

        pythoncom.CoInitialize()
        stage = "connect"
        try:
            application = client.Dispatch("Outlook.Application")
            stage = "open_mapi"
            namespace = application.GetNamespace("MAPI")
            stage = "find_existing_draft"
            mail = self._find_existing_draft(namespace, draft_key)
            existing = mail is not None
            if mail is None:
                stage = "create_item"
                mail = application.CreateItem(OL_MAIL_ITEM)
                stage = "fill_fields"
                mail.To = recipient_email.strip().casefold()
                mail.Subject = subject.strip()
                mail.Body = body
                stage = "set_deduplication_key"
                properties = mail.UserProperties
                property_item = properties.Add(
                    GNS_DRAFT_KEY_PROPERTY,
                    OL_TEXT,
                    True,
                )
                property_item.Value = draft_key
                stage = "attach_pdf"
                mail.Attachments.Add(
                    str(source),
                    OL_BY_VALUE,
                    1,
                    attachment_name,
                )
                stage = "save_draft"
                mail.Save()
            stage = "display_draft"
            mail.Display()
            inspector = _safe_attribute(mail, "GetInspector", None)
            if inspector is not None:
                try:
                    inspector.Activate()
                except Exception:
                    pass
                focus_window_handle(
                    int(_safe_attribute(inspector, "HWND", 0) or 0)
                )
            return OutlookDraftResult(
                entry_id=_safe_text(_safe_attribute(mail, "EntryID")),
                recipient_email=_safe_text(_safe_attribute(mail, "To")),
                subject=_safe_text(_safe_attribute(mail, "Subject")),
                attachment_name=attachment_name,
                existing=existing,
            )
        except OutlookIntegrationError:
            raise
        except Exception as exc:
            record_exception(
                "outlook_worker",
                "create_draft",
                exc,
                details={"stage": stage},
            )
            raise OutlookConnectionError(
                "Outlook не смог создать или открыть черновик письма. "
                f"Этап: {stage}."
            ) from exc
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _mapi_property(item: Any, schema: str) -> str:
        accessor = _safe_attribute(item, "PropertyAccessor", None)
        if accessor is None:
            return ""
        try:
            return _safe_text(accessor.GetProperty(schema))
        except Exception:
            return ""

    @classmethod
    def _sender_smtp(cls, mail: Any) -> str:
        email_type = _safe_text(
            _safe_attribute(mail, "SenderEmailType")
        ).casefold()
        if email_type != "ex":
            return _safe_text(
                _safe_attribute(mail, "SenderEmailAddress")
            ).casefold()
        sender = _safe_attribute(mail, "Sender", None)
        if sender is not None:
            try:
                exchange_user = sender.GetExchangeUser()
                primary = _safe_text(
                    _safe_attribute(exchange_user, "PrimarySmtpAddress")
                )
                if primary:
                    return primary.casefold()
            except Exception:
                pass
            primary = cls._mapi_property(sender, PR_SMTP_ADDRESS)
            if primary:
                return primary.casefold()
        return cls._mapi_property(mail, PR_SMTP_ADDRESS).casefold()

    @classmethod
    def _source_key(cls, mail: Any) -> str:
        parent = _safe_attribute(mail, "Parent", None)
        store_id = _safe_text(_safe_attribute(parent, "StoreID"))
        internet_id = cls._mapi_property(mail, PR_INTERNET_MESSAGE_ID)
        if internet_id:
            identity = f"internet:{store_id}:{internet_id}"
        else:
            entry_id = _safe_text(_safe_attribute(mail, "EntryID"))
            if not entry_id:
                return ""
            identity = f"entry:{store_id}:{entry_id}"
        return hashlib.sha256(identity.encode("utf-8")).hexdigest()

    def scan_inbox(
        self,
        *,
        allowed_senders: frozenset[str],
        allowed_domains: frozenset[str],
        received_since: date,
        known_message_keys: frozenset[str],
        staging_dir: Path,
        max_attachment_bytes: int,
        mailbox: str = "",
    ) -> OutlookInboxScan:
        if sys.platform != "win32":
            raise OutlookUnsupportedPlatformError(
                "Интеграция Outlook доступна только в Windows."
            )
        try:
            import pythoncom  # type: ignore[import-not-found]
            from win32com import client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise OutlookComponentMissingError(
                "Не установлен локальный компонент связи с Outlook. "
                "Повторно запустите START.bat."
            ) from exc

        normalized_senders = frozenset(
            value.casefold() for value in allowed_senders
        )
        normalized_domains = frozenset(
            value.casefold().lstrip("@") for value in allowed_domains
        )
        if not staging_dir.is_dir():
            raise OutlookConnectionError(
                "Временная папка импорта Outlook не подготовлена."
            )
        pythoncom.CoInitialize()
        try:
            application = client.Dispatch("Outlook.Application")
            namespace = application.GetNamespace("MAPI")
            inbox = self._select_inbox(namespace, mailbox)
            self._ensure_outlook_window(application, inbox)
            items = inbox.Items
            inspected = 0
            eligible = 0
            known = 0
            scan_errors = 0
            messages: list[OutlookScannedMessage] = []
            item_count = int(_safe_attribute(items, "Count", 0) or 0)
            for item_index in range(1, item_count + 1):
                try:
                    mail = items.Item(item_index)
                    if int(_safe_attribute(mail, "Class", 0) or 0) != 43:
                        continue
                    inspected += 1
                    received_time = _safe_attribute(mail, "ReceivedTime", None)
                    if received_time is None or received_time.date() < received_since:
                        continue
                    sender_smtp = self._sender_smtp(mail)
                    sender_domain = sender_smtp.rpartition("@")[2]
                    if (
                        sender_smtp not in normalized_senders
                        and sender_domain not in normalized_domains
                    ):
                        continue
                    source_key = self._source_key(mail)
                    if not source_key:
                        continue
                    eligible += 1
                    if source_key in known_message_keys:
                        known += 1
                        continue
                    attachment_collection = mail.Attachments
                    attachment_count = int(
                        _safe_attribute(attachment_collection, "Count", 0) or 0
                    )
                    attachments: list[OutlookScannedAttachment] = []
                    pdf_count = 0
                    for attachment_index in range(1, attachment_count + 1):
                        attachment = attachment_collection.Item(attachment_index)
                        original_name = _safe_text(
                            _safe_attribute(attachment, "FileName")
                        ) or "document.pdf"
                        if Path(original_name).suffix.casefold() != ".pdf":
                            continue
                        pdf_count += 1
                        size_bytes = int(
                            _safe_attribute(attachment, "Size", 0) or 0
                        )
                        safe_name = sanitize_filename(original_name)
                        temporary_path = (
                            staging_dir
                            / (
                                f"{source_key[:20]}_"
                                f"{attachment_index:03d}_{safe_name}.part"
                            )
                        )
                        temporary_path_text = str(temporary_path)
                        error_message = ""
                        if size_bytes > max_attachment_bytes:
                            error_message = "PDF превышает допустимый размер"
                        else:
                            try:
                                attachment.SaveAsFile(str(temporary_path))
                            except Exception:
                                error_message = "Outlook не смог сохранить вложение"
                                temporary_path_text = ""
                        attachments.append(
                            OutlookScannedAttachment(
                                attachment_index=attachment_index,
                                original_filename=original_name,
                                temporary_path=temporary_path_text,
                                size_bytes=size_bytes,
                                error_message=error_message,
                            )
                        )
                    messages.append(
                        OutlookScannedMessage(
                            source_key=source_key,
                            sender_smtp=sender_smtp,
                            received_at=received_time.isoformat(),
                            attachment_count=attachment_count,
                            pdf_attachment_count=pdf_count,
                            attachments=tuple(attachments),
                        )
                    )
                except Exception:
                    # Одно повреждённое или служебное сообщение не скрывает
                    # остальные письма. Оно будет снова проверено при следующем
                    # обходе, потому что не попало в реестр известных ключей.
                    scan_errors += 1
                    continue
            return OutlookInboxScan(
                inspected_mail_count=inspected,
                eligible_message_count=eligible,
                known_message_count=known,
                scan_error_count=scan_errors,
                messages=tuple(messages),
            )
        except OutlookIntegrationError:
            raise
        except Exception as exc:
            raise OutlookConnectionError(
                "Не удалось проверить папку входящих Outlook."
            ) from exc
        finally:
            pythoncom.CoUninitialize()


class SubprocessOutlookGateway:
    """Изолирует COM: зависший Outlook не останавливает веб-приложение."""

    def __init__(self, timeout_seconds: int = 30):
        self.timeout_seconds = timeout_seconds

    def inspect(self, *, request_sync: bool = False) -> OutlookProbe:
        action = "sync" if request_sync else "diagnose"
        payload = self._call_bridge(action)
        probe_payload = payload.get("probe")
        if not isinstance(probe_payload, dict):
            raise OutlookConnectionError(
                "Модуль Outlook не вернул сведения о профиле."
            )
        return OutlookProbe.from_dict(probe_payload)

    def scan_inbox(
        self,
        *,
        allowed_senders: frozenset[str],
        allowed_domains: frozenset[str],
        received_since: date,
        known_message_keys: frozenset[str],
        staging_dir: Path,
        max_attachment_bytes: int,
        mailbox: str = "",
    ) -> OutlookInboxScan:
        payload = self._call_bridge(
            "scan",
            input_payload={
                "allowed_senders": sorted(allowed_senders),
                "allowed_domains": sorted(allowed_domains),
                "received_since": received_since.isoformat(),
                "known_message_keys": sorted(known_message_keys),
                "staging_dir": str(staging_dir),
                "max_attachment_bytes": max_attachment_bytes,
                "mailbox": mailbox,
            },
            timeout_seconds=max(120, self.timeout_seconds),
        )
        scan_payload = payload.get("scan")
        if not isinstance(scan_payload, dict):
            raise OutlookConnectionError(
                "Модуль Outlook не вернул результат проверки писем."
            )
        return OutlookInboxScan.from_dict(scan_payload)

    def create_draft(
        self,
        *,
        draft_key: str,
        recipient_email: str,
        subject: str,
        body: str,
        attachment_path: Path,
        attachment_name: str,
    ) -> OutlookDraftResult:
        payload = self._call_bridge(
            "draft",
            input_payload={
                "draft_key": draft_key,
                "recipient_email": recipient_email,
                "subject": subject,
                "body": body,
                "attachment_path": str(attachment_path),
                "attachment_name": attachment_name,
            },
            timeout_seconds=max(120, self.timeout_seconds),
        )
        draft_payload = payload.get("draft")
        if not isinstance(draft_payload, dict):
            raise OutlookConnectionError(
                "Модуль Outlook не вернул сведения о черновике."
            )
        return OutlookDraftResult.from_dict(draft_payload)

    def _call_bridge(
        self,
        action: str,
        *,
        input_payload: dict[str, Any] | None = None,
        timeout_seconds: int | None = None,
    ) -> dict[str, Any]:
        command = module_command(
            "gns_app.services.outlook_service",
            "--bridge-action",
            action,
        )
        timeout = timeout_seconds or self.timeout_seconds
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            completed = subprocess.run(
                command,
                input=(
                    json.dumps(input_payload, ensure_ascii=False)
                    if input_payload is not None
                    else None
                ),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                check=False,
                creationflags=creation_flags,
                cwd=Path.cwd(),
                env={**os.environ, "PYTHONIOENCODING": "utf-8"},
            )
        except subprocess.TimeoutExpired as exc:
            raise OutlookProbeTimeoutError(
                f"Outlook не ответил за {timeout} секунд. "
                "Возможно, открыто скрытое окно выбора профиля или подтверждения."
            ) from exc
        try:
            payload = json.loads(completed.stdout.strip())
        except (json.JSONDecodeError, AttributeError) as exc:
            raise OutlookConnectionError(
                "Модуль Outlook завершился без корректного результата."
            ) from exc
        if not isinstance(payload, dict):
            raise OutlookConnectionError(
                "Модуль Outlook вернул неизвестный результат."
            )
        if not payload.get("ok"):
            error_kind = str(payload.get("error_kind") or "connection")
            error_message = str(
                payload.get("message")
                or "Не удалось выполнить проверку Outlook."
            )
            error_class = {
                "unsupported_platform": OutlookUnsupportedPlatformError,
                "component_missing": OutlookComponentMissingError,
                "timeout": OutlookProbeTimeoutError,
            }.get(error_kind, OutlookConnectionError)
            raise error_class(error_message)
        return payload


class OutlookService:
    def __init__(self, gateway: OutlookGateway | None = None):
        self.gateway = gateway or SubprocessOutlookGateway()
        self._lock = RLock()
        self._last_result = OutlookDiagnosticResult(
            state="not_checked",
            checked_at="",
            message="Подключение к Outlook ещё не проверялось.",
        )

    def last_result(self) -> OutlookDiagnosticResult:
        with self._lock:
            return self._last_result

    def diagnose(self) -> OutlookDiagnosticResult:
        return self._run(request_sync=False)

    def request_send_receive(self) -> OutlookDiagnosticResult:
        return self._run(request_sync=True)

    def scan_inbox(
        self,
        *,
        allowed_senders: frozenset[str],
        allowed_domains: frozenset[str],
        received_since: date,
        known_message_keys: frozenset[str],
        staging_dir: Path,
        max_attachment_bytes: int,
        mailbox: str = "",
    ) -> OutlookInboxScan:
        scanner = getattr(self.gateway, "scan_inbox", None)
        if scanner is None:
            raise OutlookConnectionError(
                "Текущий адаптер Outlook не поддерживает импорт вложений."
            )
        return scanner(
            allowed_senders=allowed_senders,
            allowed_domains=allowed_domains,
            received_since=received_since,
            known_message_keys=known_message_keys,
            staging_dir=staging_dir,
            max_attachment_bytes=max_attachment_bytes,
            mailbox=mailbox,
        )

    def create_draft(
        self,
        *,
        draft_key: str,
        recipient_email: str,
        subject: str,
        body: str,
        attachment_path: Path,
        attachment_name: str,
    ) -> OutlookDraftResult:
        creator = getattr(self.gateway, "create_draft", None)
        if creator is None:
            raise OutlookConnectionError(
                "Текущий адаптер Outlook не поддерживает исходящие черновики."
            )
        return creator(
            draft_key=draft_key,
            recipient_email=recipient_email,
            subject=subject,
            body=body,
            attachment_path=attachment_path,
            attachment_name=attachment_name,
        )

    def _run(self, *, request_sync: bool) -> OutlookDiagnosticResult:
        started = time.monotonic()
        checked_at = datetime.now().astimezone().strftime("%d.%m.%Y %H:%M:%S")
        try:
            probe = self.gateway.inspect(request_sync=request_sync)
            if request_sync:
                state = "sync_requested"
                message = (
                    "Команда «Получить почту» передана Outlook. "
                    "Синхронизация выполняется Outlook в фоне."
                )
            elif probe.offline:
                state = "ready"
                message = (
                    "Outlook доступен, но сейчас работает автономно. "
                    "Получение новых писем может быть недоступно."
                )
            else:
                state = "ready"
                message = "Outlook доступен, папка входящих найдена."
            result = OutlookDiagnosticResult(
                state=state,
                checked_at=checked_at,
                message=message,
                probe=probe,
            )
        except OutlookUnsupportedPlatformError as exc:
            result = OutlookDiagnosticResult(
                state="unsupported_platform",
                checked_at=checked_at,
                message=str(exc),
            )
        except OutlookComponentMissingError as exc:
            result = OutlookDiagnosticResult(
                state="component_missing",
                checked_at=checked_at,
                message=str(exc),
            )
        except OutlookProbeTimeoutError as exc:
            result = OutlookDiagnosticResult(
                state="timeout",
                checked_at=checked_at,
                message=str(exc),
            )
        except OutlookIntegrationError as exc:
            result = OutlookDiagnosticResult(
                state="connection_error",
                checked_at=checked_at,
                message=str(exc),
            )
        with self._lock:
            self._last_result = result
        record_event(
            "outlook",
            "send_receive" if request_sync else "diagnose",
            result.state,
            details={
                "duration_ms": int((time.monotonic() - started) * 1000),
                "successful": result.successful,
                "offline": bool(result.probe and result.probe.offline),
                "accounts_count": len(result.probe.accounts) if result.probe else 0,
                "inbox_items": result.probe.inbox_item_count if result.probe else 0,
            },
        )
        return result


class OutlookOutgoingService:
    SUBJECT_TEMPLATE_SETTING = "outlook_outgoing_subject_template"
    DEFAULT_SUBJECT_TEMPLATE = (
        "Ответ на запрос ГНС, исх. № {outgoing_number}"
    )
    SUBJECT_FIELDS = frozenset(
        {"outgoing_number", "office_name", "recipient_name", "date"}
    )

    def __init__(
        self,
        db: Database,
        settings: Settings,
        outlook: OutlookService,
    ):
        self.db = db
        self.settings = settings
        self.outlook = outlook
        self._draft_lock = RLock()

    @classmethod
    def validate_subject_template(cls, value: str) -> str:
        template = " ".join(value.split())
        if not template or len(template) > 200:
            raise OutlookIntegrationError(
                "Тема исходящего письма должна содержать от 1 до 200 символов."
            )
        try:
            parsed = list(Formatter().parse(template))
        except ValueError as exc:
            raise OutlookIntegrationError(
                "В шаблоне темы неправильно расставлены фигурные скобки."
            ) from exc
        for _, field_name, format_spec, conversion in parsed:
            if field_name is None:
                continue
            if (
                field_name not in cls.SUBJECT_FIELDS
                or format_spec
                or conversion
            ):
                raise OutlookIntegrationError(
                    "В теме разрешены только подстановки: "
                    "{outgoing_number}, {office_name}, "
                    "{recipient_name}, {date}."
                )
        return template

    def get_subject_template(self) -> str:
        row = self.db.fetch_one(
            "SELECT value FROM settings WHERE key = ?",
            (self.SUBJECT_TEMPLATE_SETTING,),
        )
        value = str(row["value"]) if row else self.DEFAULT_SUBJECT_TEMPLATE
        try:
            return self.validate_subject_template(value)
        except OutlookIntegrationError:
            return self.DEFAULT_SUBJECT_TEMPLATE

    def update_subject_template(self, value: str) -> None:
        template = self.validate_subject_template(value)
        self.db.execute(
            """
            INSERT INTO settings(key, value, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE
            SET value = excluded.value, updated_at = excluded.updated_at
            """,
            (self.SUBJECT_TEMPLATE_SETTING, template, utc_now()),
        )
        self.db.audit(
            "settings",
            "outlook_outgoing",
            "outlook_outgoing_subject_updated",
            {
                "placeholders": sorted(
                    field_name
                    for _, field_name, _, _ in Formatter().parse(template)
                    if field_name
                )
            },
            actor="Сотрудник",
        )

    def render_subject(
        self,
        letter: dict[str, Any],
        office: dict[str, Any],
    ) -> str:
        business_date = str(letter.get("business_date") or "")
        try:
            formatted_date = date.fromisoformat(business_date).strftime(
                "%d.%m.%Y"
            )
        except ValueError:
            formatted_date = date.today().strftime("%d.%m.%Y")
        subject = self.get_subject_template().format_map(
            {
                "outgoing_number": str(letter.get("outgoing_number") or ""),
                "office_name": str(office.get("office_name") or ""),
                "recipient_name": str(
                    letter.get("recipient_display_name") or ""
                ),
                "date": formatted_date,
            }
        )
        if not subject.strip() or len(subject) > 255:
            raise OutlookIntegrationError(
                "После подстановки значений тема Outlook должна быть не длиннее "
                "255 символов."
            )
        return subject

    @staticmethod
    def _sha256_file(path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as source:
            while chunk := source.read(1024 * 1024):
                digest.update(chunk)
        return digest.hexdigest()

    def create_draft(
        self,
        workflow: OutlookWorkflow,
        letter_id: str,
    ) -> dict[str, Any]:
        started = time.monotonic()
        try:
            with self._draft_lock:
                result = self._create_draft_locked(workflow, letter_id)
        except OutlookIntegrationError as exc:
            record_exception(
                "outlook",
                "create_draft",
                exc,
                details={"duration_ms": int((time.monotonic() - started) * 1000)},
            )
            raise
        record_event(
            "outlook",
            "create_draft",
            "success",
            details={
                "duration_ms": int((time.monotonic() - started) * 1000),
                "existing_draft": bool(result.get("existing_outlook_draft")),
            },
        )
        return result

    def _create_draft_locked(
        self,
        workflow: OutlookWorkflow,
        letter_id: str,
    ) -> dict[str, Any]:
        letter = workflow.get_response_letter(letter_id)
        if not letter:
            raise OutlookIntegrationError("Готовое письмо не найдено.")
        outgoing_number = str(letter.get("outgoing_number") or "").strip()
        if not outgoing_number:
            raise OutlookIntegrationError(
                "Сначала назначьте письму исходящий номер."
            )
        scan = workflow.get_confirmed_signed_response_scan(letter_id)
        if not scan:
            raise OutlookIntegrationError(
                "Сначала загрузите и подтвердите подписанный ответ."
            )
        office = workflow.match_gns_office(
            str(letter.get("district_place") or "")
        )
        recipient_email = str(
            office.get("email_address") if office else ""
        ).strip().casefold()
        if not EMAIL_RE.fullmatch(recipient_email):
            raise OutlookIntegrationError(
                "Для подразделения ГНС не настроен однозначный email."
            )

        try:
            pdf_path = ensure_within(
                Path(str(scan["pdf_path"])),
                self.settings.runtime_dir / "signed_scans",
            )
        except (KeyError, StorageError) as exc:
            raise OutlookIntegrationError(
                "Путь подписанного ответа недоступен."
            ) from exc
        if not pdf_path.is_file():
            raise OutlookIntegrationError(
                "PDF подписанного ответа отсутствует."
            )
        subject = self.render_subject(letter, office or {})
        attachment_name = sanitize_filename(
            f"Ответ_исх_{outgoing_number}.pdf"
        )
        attachment_sha256 = self._sha256_file(pdf_path)
        message_id = uuid4().hex
        existing = self.db.fetch_one(
            "SELECT id FROM outlook_outgoing_messages WHERE signed_scan_id = ?",
            (scan["id"],),
        )
        if existing:
            message_id = str(existing["id"])
        now = utc_now()
        self.db.execute(
            """
            INSERT INTO outlook_outgoing_messages(
                id, response_letter_id, signed_scan_id, status,
                recipient_email, subject, attachment_name,
                attachment_sha256, created_at, updated_at
            ) VALUES (?, ?, ?, 'creating', ?, ?, ?, ?, ?, ?)
            ON CONFLICT(signed_scan_id) DO UPDATE SET
                status = 'creating',
                recipient_email = excluded.recipient_email,
                subject = excluded.subject,
                attachment_name = excluded.attachment_name,
                attachment_sha256 = excluded.attachment_sha256,
                error_message = NULL,
                updated_at = excluded.updated_at
            """,
            (
                message_id,
                letter_id,
                scan["id"],
                recipient_email,
                subject,
                attachment_name,
                attachment_sha256,
                now,
                now,
            ),
        )

        staging_root = ensure_within(
            self.settings.runtime_dir / "outlook_outgoing_staging",
            self.settings.runtime_dir,
        )
        run_dir = ensure_within(staging_root / message_id, staging_root)
        attachment_path = run_dir / attachment_name
        try:
            run_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(pdf_path, attachment_path)
            result = self.outlook.create_draft(
                draft_key=f"gns-scan-{scan['id']}",
                recipient_email=recipient_email,
                subject=subject,
                body="",
                attachment_path=attachment_path,
                attachment_name=attachment_name,
            )
        except OutlookIntegrationError as exc:
            self.db.execute(
                """
                UPDATE outlook_outgoing_messages
                SET status = 'technical_error', error_message = ?, updated_at = ?
                WHERE id = ?
                """,
                (str(exc)[:500], utc_now(), message_id),
            )
            self.db.audit(
                "outlook_outgoing_message",
                message_id,
                "outlook_draft_failed",
                {"signed_scan_id": scan["id"]},
                actor=workflow.get_active_employee(),
            )
            raise
        except OSError as exc:
            message = "Не удалось подготовить PDF для Outlook."
            self.db.execute(
                """
                UPDATE outlook_outgoing_messages
                SET status = 'technical_error', error_message = ?, updated_at = ?
                WHERE id = ?
                """,
                (message, utc_now(), message_id),
            )
            raise OutlookIntegrationError(message) from exc
        finally:
            if run_dir.exists():
                shutil.rmtree(run_dir, ignore_errors=True)

        self.db.execute(
            """
            UPDATE outlook_outgoing_messages
            SET status = 'draft_created', outlook_entry_id = ?,
                error_message = NULL, updated_at = ?
            WHERE id = ?
            """,
            (result.entry_id, utc_now(), message_id),
        )
        self.db.audit(
            "outlook_outgoing_message",
            message_id,
            "outlook_draft_created",
            {
                "signed_scan_id": scan["id"],
                "existing_outlook_draft": result.existing,
            },
            actor=workflow.get_active_employee(),
        )
        row = self.db.fetch_one(
            "SELECT * FROM outlook_outgoing_messages WHERE id = ?",
            (message_id,),
        ) or {}
        row["existing_outlook_draft"] = result.existing
        return row


class OutlookInboxImporter:
    ALLOWED_SENDERS_SETTING = "outlook_allowed_senders"
    IMPORT_SINCE_SETTING = "outlook_import_since"
    MAILBOX_SETTING = "outlook_mailbox"
    AUTO_ENABLED_SETTING = "outlook_auto_enabled"
    AUTO_INTERVAL_SETTING = "outlook_auto_interval_minutes"
    AUTO_STATUS_SETTING = "outlook_auto_status"
    DEFAULT_AUTO_INTERVAL_MINUTES = 5
    MIN_AUTO_INTERVAL_MINUTES = 1
    MAX_AUTO_INTERVAL_MINUTES = 1440
    DEFAULT_ALLOWED_SENDERS = frozenset({"esensabiyatov@gmail.com"})
    DEFAULT_ALLOWED_DOMAINS = frozenset({"sti.gov.kg", "salyk.kg"})

    def __init__(
        self,
        db: Database,
        settings: Settings,
        outlook: OutlookService,
    ):
        self.db = db
        self.settings = settings
        self.outlook = outlook
        self._import_lock = RLock()

    def get_allowed_senders(self) -> frozenset[str]:
        row = self.db.fetch_one(
            "SELECT value FROM settings WHERE key = ?",
            (self.ALLOWED_SENDERS_SETTING,),
        )
        if not row:
            return self.DEFAULT_ALLOWED_SENDERS
        values = {
            value.strip().casefold()
            for value in str(row["value"]).replace(";", ",").split(",")
            if value.strip()
        }
        return frozenset(values)

    def get_import_since(self) -> date:
        row = self.db.fetch_one(
            "SELECT value FROM settings WHERE key = ?",
            (self.IMPORT_SINCE_SETTING,),
        )
        if row:
            try:
                return date.fromisoformat(str(row["value"]))
            except ValueError:
                pass
        return date.today() - timedelta(days=7)

    def get_mailbox(self) -> str:
        row = self.db.fetch_one(
            "SELECT value FROM settings WHERE key = ?",
            (self.MAILBOX_SETTING,),
        )
        return str(row["value"]).strip() if row else ""

    def get_auto_enabled(self) -> bool:
        row = self.db.fetch_one(
            "SELECT value FROM settings WHERE key = ?",
            (self.AUTO_ENABLED_SETTING,),
        )
        return bool(row and str(row["value"]).strip() == "1")

    def get_auto_interval_minutes(self) -> int:
        row = self.db.fetch_one(
            "SELECT value FROM settings WHERE key = ?",
            (self.AUTO_INTERVAL_SETTING,),
        )
        try:
            value = (
                int(row["value"])
                if row
                else self.DEFAULT_AUTO_INTERVAL_MINUTES
            )
        except (TypeError, ValueError):
            value = self.DEFAULT_AUTO_INTERVAL_MINUTES
        return max(
            self.MIN_AUTO_INTERVAL_MINUTES,
            min(self.MAX_AUTO_INTERVAL_MINUTES, value),
        )

    def get_automation_status(self) -> dict[str, Any]:
        default = {
            "state": "not_started",
            "checked_at": "",
            "message": "Автоматическая проверка ещё не выполнялась.",
            "new_messages": 0,
            "saved_attachments": 0,
            "errors": 0,
        }
        row = self.db.fetch_one(
            "SELECT value FROM settings WHERE key = ?",
            (self.AUTO_STATUS_SETTING,),
        )
        if not row:
            return default
        try:
            payload = json.loads(str(row["value"]))
        except (json.JSONDecodeError, TypeError, ValueError):
            return default
        if not isinstance(payload, dict):
            return default
        return {
            "state": str(payload.get("state") or default["state"]),
            "checked_at": str(payload.get("checked_at") or ""),
            "message": str(payload.get("message") or default["message"]),
            "new_messages": int(payload.get("new_messages") or 0),
            "saved_attachments": int(payload.get("saved_attachments") or 0),
            "errors": int(payload.get("errors") or 0),
        }

    def record_automation_status(
        self,
        *,
        state: str,
        message: str,
        new_messages: int = 0,
        saved_attachments: int = 0,
        errors: int = 0,
    ) -> None:
        payload = {
            "state": state,
            "checked_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "message": message[:500],
            "new_messages": max(0, int(new_messages)),
            "saved_attachments": max(0, int(saved_attachments)),
            "errors": max(0, int(errors)),
        }
        self.db.execute(
            """
            INSERT INTO settings(key, value, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE
            SET value = excluded.value, updated_at = excluded.updated_at
            """,
            (
                self.AUTO_STATUS_SETTING,
                json.dumps(payload, ensure_ascii=False),
                utc_now(),
            ),
        )

    def get_inbox_dir(self) -> Path:
        row = self.db.fetch_one(
            "SELECT value FROM settings WHERE key = 'inbox_dir'"
        )
        value = str(row["value"]).strip() if row else ""
        return (
            Path(value).resolve()
            if value
            else self.settings.inbox_dir.resolve()
        )

    def update_settings(
        self,
        *,
        allowed_senders: str,
        import_since: str,
        mailbox: str = "",
        auto_enabled: bool = False,
        auto_interval_minutes: int = DEFAULT_AUTO_INTERVAL_MINUTES,
    ) -> None:
        values = {
            value.strip().casefold()
            for value in allowed_senders.replace(";", ",").split(",")
            if value.strip()
        }
        invalid = sorted(value for value in values if not EMAIL_RE.fullmatch(value))
        if invalid:
            raise OutlookIntegrationError(
                "Укажите корректные адреса разрешённых отправителей."
            )
        try:
            since = date.fromisoformat(import_since)
        except ValueError as exc:
            raise OutlookIntegrationError(
                "Укажите корректную начальную дату проверки Outlook."
            ) from exc
        if since > date.today():
            raise OutlookIntegrationError(
                "Начальная дата Outlook не может быть в будущем."
            )
        mailbox_name = mailbox.strip()
        if (
            len(mailbox_name) > 255
            or "\n" in mailbox_name
            or "\r" in mailbox_name
        ):
            raise OutlookIntegrationError(
                "Укажите корректное имя почтового ящика Outlook."
            )
        try:
            interval = int(auto_interval_minutes)
        except (TypeError, ValueError) as exc:
            raise OutlookIntegrationError(
                "Укажите корректный интервал автоматической проверки Outlook."
            ) from exc
        if not (
            self.MIN_AUTO_INTERVAL_MINUTES
            <= interval
            <= self.MAX_AUTO_INTERVAL_MINUTES
        ):
            raise OutlookIntegrationError(
                "Интервал Outlook должен быть от 1 минуты до 24 часов."
            )
        now = utc_now()
        self.db.executemany(
            """
            INSERT INTO settings(key, value, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE
            SET value = excluded.value, updated_at = excluded.updated_at
            """,
            [
                (self.ALLOWED_SENDERS_SETTING, ",".join(sorted(values)), now),
                (self.IMPORT_SINCE_SETTING, since.isoformat(), now),
                (self.MAILBOX_SETTING, mailbox_name, now),
                (self.AUTO_ENABLED_SETTING, "1" if auto_enabled else "0", now),
                (self.AUTO_INTERVAL_SETTING, str(interval), now),
            ],
        )
        self.db.audit(
            "settings",
            "outlook",
            "outlook_import_settings_updated",
            {
                "allowed_sender_count": len(values),
                "import_since": since.isoformat(),
                "mailbox_configured": bool(mailbox_name),
                "automatic_import_enabled": bool(auto_enabled),
                "automatic_interval_minutes": interval,
            },
            actor="Сотрудник",
        )

    def stats(self) -> dict[str, int]:
        rows = self.db.fetch_all(
            """
            SELECT status, COUNT(*) AS count
            FROM outlook_messages
            GROUP BY status
            """
        )
        result = {"messages": 0, "completed": 0, "no_pdf": 0, "errors": 0}
        for row in rows:
            count = int(row["count"])
            result["messages"] += count
            status = str(row["status"])
            if status == "completed":
                result["completed"] += count
            elif status == "no_pdf":
                result["no_pdf"] += count
            elif status == "technical_error":
                result["errors"] += count
        return result

    @staticmethod
    def _attachment_id(message_key: str, attachment_index: int) -> str:
        identity = f"{message_key}:{attachment_index}"
        return hashlib.sha256(identity.encode("utf-8")).hexdigest()

    @staticmethod
    def _hash_existing_pdf(path: Path, max_bytes: int) -> tuple[str, int]:
        digest = hashlib.sha256()
        total = 0
        with path.open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise StorageError("Сохранённый файл не имеет сигнатуру PDF")
            stream.seek(0)
            while chunk := stream.read(1024 * 1024):
                total += len(chunk)
                if total > max_bytes:
                    raise StorageError("PDF превышает допустимый размер")
                digest.update(chunk)
        return digest.hexdigest(), total

    def _upsert_message(
        self,
        message: OutlookScannedMessage,
        *,
        status: str,
        error_message: str = "",
    ) -> None:
        now = utc_now()
        self.db.execute(
            """
            INSERT INTO outlook_messages(
                source_key, sender_smtp, received_at, status,
                attachment_count, pdf_attachment_count, error_message,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(source_key) DO UPDATE SET
                sender_smtp = excluded.sender_smtp,
                received_at = excluded.received_at,
                status = excluded.status,
                attachment_count = excluded.attachment_count,
                pdf_attachment_count = excluded.pdf_attachment_count,
                error_message = excluded.error_message,
                updated_at = excluded.updated_at
            """,
            (
                message.source_key,
                message.sender_smtp,
                message.received_at,
                status,
                message.attachment_count,
                message.pdf_attachment_count,
                error_message or None,
                now,
                now,
            ),
        )

    def _upsert_attachment(
        self,
        message_key: str,
        attachment: OutlookScannedAttachment,
        *,
        status: str,
        stored_path: str = "",
        digest: str = "",
        size_bytes: int | None = None,
        upload_id: str = "",
        error_message: str = "",
    ) -> None:
        now = utc_now()
        self.db.execute(
            """
            INSERT INTO outlook_attachments(
                id, message_key, attachment_index, original_filename,
                stored_path, sha256, size_bytes, status, upload_id,
                error_message, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(message_key, attachment_index) DO UPDATE SET
                original_filename = excluded.original_filename,
                stored_path = excluded.stored_path,
                sha256 = excluded.sha256,
                size_bytes = excluded.size_bytes,
                status = excluded.status,
                upload_id = excluded.upload_id,
                error_message = excluded.error_message,
                updated_at = excluded.updated_at
            """,
            (
                self._attachment_id(message_key, attachment.attachment_index),
                message_key,
                attachment.attachment_index,
                sanitize_filename(attachment.original_filename),
                stored_path or None,
                digest or None,
                size_bytes,
                status,
                upload_id or None,
                error_message or None,
                now,
                now,
            ),
        )

    def import_new(self, workflow: OutlookWorkflow) -> dict[str, Any]:
        started = time.monotonic()
        try:
            with self._import_lock:
                summary = self._import_new_locked(workflow)
        except OutlookIntegrationError as exc:
            record_exception(
                "outlook",
                "import_inbox",
                exc,
                details={"duration_ms": int((time.monotonic() - started) * 1000)},
            )
            raise
        record_event(
            "outlook",
            "import_inbox",
            "warning"
            if summary["scan_errors"] or summary["sync_error"] or summary["errors"]
            else "success",
            details={
                "duration_ms": int((time.monotonic() - started) * 1000),
                "inspected": summary["inspected"],
                "eligible": summary["eligible"],
                "known": summary["known"],
                "new_messages": summary["new_messages"],
                "saved_attachments": summary["saved_attachments"],
                "duplicate_attachments": summary["duplicate_attachments"],
                "scan_errors": summary["scan_errors"],
                "sync_error": bool(summary["sync_error"]),
                "processing_errors": len(summary["errors"]),
            },
        )
        return summary

    def _import_new_locked(self, workflow: OutlookWorkflow) -> dict[str, Any]:
        if not workflow.get_active_employee():
            raise OutlookIntegrationError(
                "Сначала выберите или добавьте исполнителя."
            )
        known = frozenset(
            str(row["source_key"])
            for row in self.db.fetch_all(
                """
                SELECT source_key FROM outlook_messages
                WHERE status IN ('completed', 'no_pdf')
                """
            )
        )
        staging_root = (self.settings.runtime_dir / "outlook_staging").resolve()
        staging_root.mkdir(parents=True, exist_ok=True)
        run_dir = (staging_root / uuid4().hex).resolve()
        ensure_within(run_dir, staging_root)
        # Папку создаёт основной процесс. На офисном ПК политика безопасности
        # запрещала дочернему Outlook-COM процессу создавать этот каталог.
        run_dir.mkdir(parents=False, exist_ok=False)
        imported_uploads: list[str] = []
        duplicate_attachments = 0
        saved_attachments = 0
        errors: list[str] = []
        try:
            scan = self.outlook.scan_inbox(
                allowed_senders=self.get_allowed_senders(),
                allowed_domains=self.DEFAULT_ALLOWED_DOMAINS,
                received_since=self.get_import_since(),
                known_message_keys=known,
                staging_dir=run_dir,
                max_attachment_bytes=self.settings.max_upload_bytes,
                mailbox=self.get_mailbox(),
            )
            # Чтение текущей локальной папки выполняется до SendAndReceive.
            # Поэтому зависшая синхронизация не скрывает уже полученные письма,
            # а новые письма учитываются следующим полным проходом реестра.
            sync_result = self.outlook.request_send_receive()
            sync_error = "" if sync_result.successful else sync_result.message
            for message in scan.messages:
                self._upsert_message(message, status="processing")
                if message.pdf_attachment_count == 0:
                    self._upsert_message(message, status="no_pdf")
                    continue
                message_errors = 0
                try:
                    received_day = datetime.fromisoformat(
                        message.received_at
                    ).date().isoformat()
                except ValueError:
                    received_day = date.today().isoformat()
                # Учитываем путь из локальных настроек, а не только значение,
                # с которым приложение было запущено.
                inbox_dir = self.get_inbox_dir()
                message_dir = (
                    inbox_dir
                    / "Outlook"
                    / received_day
                    / message.source_key[:16]
                )
                ensure_within(message_dir, inbox_dir)
                for attachment in message.attachments:
                    if attachment.error_message or not attachment.temporary_path:
                        message_errors += 1
                        error_text = attachment.error_message or "Вложение не сохранено"
                        errors.append(error_text)
                        self._upsert_attachment(
                            message.source_key,
                            attachment,
                            status="technical_error",
                            error_message=error_text,
                        )
                        continue
                    try:
                        source = ensure_within(
                            Path(attachment.temporary_path),
                            run_dir,
                        )
                        safe_name = sanitize_filename(attachment.original_filename)
                        destination = (
                            message_dir
                            / f"{attachment.attachment_index:03d}_{safe_name}"
                        )
                        if destination.exists():
                            existing_digest, existing_size = self._hash_existing_pdf(
                                destination,
                                self.settings.max_upload_bytes,
                            )
                            source_digest, _ = self._hash_existing_pdf(
                                source,
                                self.settings.max_upload_bytes,
                            )
                            if existing_digest != source_digest:
                                raise StorageError(
                                    "Содержимое ранее сохранённого вложения изменилось"
                                )
                            digest, size_bytes = existing_digest, existing_size
                        else:
                            with source.open("rb") as stream:
                                digest, size_bytes = save_pdf_stream(
                                    stream,
                                    destination,
                                    self.settings.max_upload_bytes,
                                )
                            saved_attachments += 1
                        existing_upload = self.db.fetch_one(
                            "SELECT id FROM uploads WHERE sha256 = ?",
                            (digest,),
                        )
                        if existing_upload:
                            upload_id = str(existing_upload["id"])
                            attachment_status = "duplicate_content"
                            duplicate_attachments += 1
                        else:
                            with destination.open("rb") as stream:
                                upload_id = workflow.create_upload(
                                    safe_name,
                                    stream,
                                )
                            imported_uploads.append(upload_id)
                            attachment_status = "imported"
                        self._upsert_attachment(
                            message.source_key,
                            attachment,
                            status=attachment_status,
                            stored_path=str(destination),
                            digest=digest,
                            size_bytes=size_bytes,
                            upload_id=upload_id,
                        )
                    except Exception as exc:
                        message_errors += 1
                        error_text = str(exc)[:300] or "Ошибка сохранения PDF"
                        errors.append(error_text)
                        self._upsert_attachment(
                            message.source_key,
                            attachment,
                            status="technical_error",
                            error_message=error_text,
                        )
                self._upsert_message(
                    message,
                    status="technical_error" if message_errors else "completed",
                    error_message=(
                        f"Не обработано PDF: {message_errors}"
                        if message_errors
                        else ""
                    ),
                )
            summary = {
                "inspected": scan.inspected_mail_count,
                "eligible": scan.eligible_message_count,
                "known": scan.known_message_count,
                "new_messages": len(scan.messages),
                "saved_attachments": saved_attachments,
                "duplicate_attachments": duplicate_attachments,
                "imported": imported_uploads,
                "scan_errors": scan.scan_error_count,
                "sync_error": sync_error,
                "errors": errors,
            }
            self.db.audit(
                "outlook",
                "inbox",
                "outlook_inbox_scanned",
                {
                    "inspected": summary["inspected"],
                    "eligible": summary["eligible"],
                    "known": summary["known"],
                    "new_messages": summary["new_messages"],
                    "saved_attachments": saved_attachments,
                    "duplicates": duplicate_attachments,
                    "scan_errors": scan.scan_error_count,
                    "sync_error": bool(sync_error),
                    "processing_errors": len(errors),
                },
                actor=workflow.get_active_employee(),
            )
            return summary
        finally:
            if run_dir.exists():
                ensure_within(run_dir, staging_root)
                shutil.rmtree(run_dir, ignore_errors=True)


def _bridge_payload(
    action: str,
    input_payload: dict[str, Any] | None = None,
) -> dict[str, Any]:
    try:
        gateway = PyWin32OutlookGateway()
        if action == "draft":
            payload = input_payload or {}
            draft = gateway.create_draft(
                draft_key=str(payload.get("draft_key") or ""),
                recipient_email=str(payload.get("recipient_email") or ""),
                subject=str(payload.get("subject") or ""),
                body=str(payload.get("body") or ""),
                attachment_path=Path(
                    str(payload.get("attachment_path") or "")
                ),
                attachment_name=str(payload.get("attachment_name") or ""),
            )
            return {"ok": True, "draft": asdict(draft)}
        if action == "scan":
            payload = input_payload or {}
            scan = gateway.scan_inbox(
                allowed_senders=frozenset(
                    str(value) for value in payload.get("allowed_senders", [])
                ),
                allowed_domains=frozenset(
                    str(value) for value in payload.get("allowed_domains", [])
                ),
                received_since=date.fromisoformat(
                    str(payload.get("received_since") or "")
                ),
                known_message_keys=frozenset(
                    str(value)
                    for value in payload.get("known_message_keys", [])
                ),
                staging_dir=Path(str(payload.get("staging_dir") or "")),
                max_attachment_bytes=int(
                    payload.get("max_attachment_bytes") or 0
                ),
                mailbox=str(payload.get("mailbox") or ""),
            )
            return {"ok": True, "scan": asdict(scan)}
        probe = gateway.inspect(request_sync=action == "sync")
        return {"ok": True, "probe": asdict(probe)}
    except (TypeError, ValueError) as exc:
        return {
            "ok": False,
            "error_kind": "connection",
            "message": "Модуль Outlook получил некорректные параметры.",
        }
    except OutlookUnsupportedPlatformError as exc:
        return {
            "ok": False,
            "error_kind": "unsupported_platform",
            "message": str(exc),
        }
    except OutlookComponentMissingError as exc:
        return {
            "ok": False,
            "error_kind": "component_missing",
            "message": str(exc),
        }
    except OutlookIntegrationError as exc:
        return {
            "ok": False,
            "error_kind": "connection",
            "message": str(exc),
        }
    except Exception as exc:
        # Ошибка дочернего процесса должна вернуться основному приложению как
        # технический результат, а не выглядеть как падение всего сервера.
        record_exception(
            "outlook_worker",
            action,
            exc,
        )
        return {
            "ok": False,
            "error_kind": "connection",
            "message": "Модуль Outlook завершился технической ошибкой.",
        }


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument(
        "--bridge-action",
        choices=("diagnose", "sync", "scan", "draft"),
        required=True,
    )
    arguments = parser.parse_args()
    input_payload: dict[str, Any] | None = None
    if arguments.bridge_action in {"scan", "draft"}:
        try:
            parsed = json.loads(sys.stdin.read() or "{}")
            input_payload = parsed if isinstance(parsed, dict) else None
        except json.JSONDecodeError:
            input_payload = None
    print(
        json.dumps(
            _bridge_payload(arguments.bridge_action, input_payload),
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
