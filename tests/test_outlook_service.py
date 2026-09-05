from __future__ import annotations

import asyncio
import hashlib
import json
import shutil
import subprocess
import sys
from dataclasses import replace
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image
from pypdf import PdfWriter

from gns_app.database import utc_now

from gns_app.services.outlook_service import (
    OutlookAccountInfo,
    OutlookConnectionError,
    OutlookDraftResult,
    OutlookInboxImporter,
    OutlookInboxScan,
    OutlookProbe,
    OutlookProbeTimeoutError,
    OutlookIntegrationError,
    OutlookOutgoingService,
    OutlookService,
    OutlookScannedAttachment,
    OutlookScannedMessage,
    OutlookSentCandidate,
    OutlookSentMatch,
    OutlookSentScan,
    OutlookSendResult,
    PyWin32OutlookGateway,
    SubprocessOutlookGateway,
)


def write_test_pdf(path: Path) -> Path:
    writer = PdfWriter()
    writer.add_blank_page(width=595, height=842)
    with path.open("wb") as stream:
        writer.write(stream)
    return path


def sample_probe(*, sync_requested: bool = False) -> OutlookProbe:
    return OutlookProbe(
        outlook_version="15.0",
        profile_name="Рабочий профиль",
        current_user="Сотрудник",
        default_store="Банк",
        inbox_name="Входящие",
        inbox_path="\\\\Банк\\Входящие",
        inbox_item_count=42,
        offline=False,
        sync_group_count=1,
        accounts=(
            OutlookAccountInfo(
                display_name="Рабочая почта",
                smtp_address="employee@example.test",
                account_type="Exchange",
            ),
        ),
        sync_requested=sync_requested,
    )


class FakeGateway:
    def __init__(self):
        self.requests: list[bool] = []
        self.mailboxes: list[str] = []

    def inspect(
        self,
        *,
        request_sync: bool = False,
        mailbox: str = "",
    ) -> OutlookProbe:
        self.requests.append(request_sync)
        self.mailboxes.append(mailbox)
        return sample_probe(sync_requested=request_sync)


class FakeImportGateway(FakeGateway):
    def __init__(
        self,
        source_pdf: Path,
        *,
        sender_smtp: str = "esensabiyatov@gmail.com",
        original_sender_smtp: str = "",
        received_at: str = "2026-08-21T09:00:00+06:00",
        attachment_filename: str = "test.pdf",
    ):
        super().__init__()
        self.source_pdf = source_pdf
        self.sender_smtp = sender_smtp
        self.original_sender_smtp = original_sender_smtp
        self.received_at = received_at
        self.attachment_filename = attachment_filename
        self.scan_requests: list[frozenset[str]] = []
        self.received_since_requests: list[date] = []
        self.staging_dirs: list[Path] = []

    def scan_inbox(
        self,
        *,
        allowed_senders,
        allowed_domains,
        received_since,
        known_message_keys,
        staging_dir,
        max_attachment_bytes,
        forwarding_senders=frozenset(),
        mailbox="",
    ):
        assert mailbox in {"", "esensabiyatov@gmail.com"}
        assert staging_dir.is_dir()
        assert (staging_dir / ".gns-ready").is_file()
        self.scan_requests.append(known_message_keys)
        self.received_since_requests.append(received_since)
        self.staging_dirs.append(staging_dir)
        message_key = "a" * 64
        if date.fromisoformat(self.received_at[:10]) < received_since:
            return OutlookInboxScan(
                inspected_mail_count=1,
                eligible_message_count=0,
                known_message_count=0,
                scan_error_count=0,
                messages=(),
            )
        if message_key in known_message_keys:
            known_messages = ()
            if self.original_sender_smtp:
                known_messages = (
                    OutlookScannedMessage(
                        source_key=message_key,
                        sender_smtp=self.sender_smtp,
                        received_at=self.received_at,
                        attachment_count=0,
                        pdf_attachment_count=0,
                        attachments=(),
                        original_sender_smtp=self.original_sender_smtp,
                    ),
                )
            return OutlookInboxScan(
                inspected_mail_count=1,
                eligible_message_count=1,
                known_message_count=1,
                scan_error_count=0,
                messages=(),
                known_messages=known_messages,
            )
        message_dir = staging_dir / message_key
        message_dir.mkdir(parents=True, exist_ok=True)
        attachments = []
        for index in (1, 2):
            suffix = Path(self.attachment_filename).suffix
            temporary = message_dir / f"{index:03d}{suffix}.part"
            shutil.copyfile(self.source_pdf, temporary)
            attachments.append(
                OutlookScannedAttachment(
                    attachment_index=index,
                    original_filename=self.attachment_filename,
                    temporary_path=str(temporary),
                    size_bytes=temporary.stat().st_size,
                )
            )
        return OutlookInboxScan(
            inspected_mail_count=1,
            eligible_message_count=1,
            known_message_count=0,
            scan_error_count=0,
            messages=(
                OutlookScannedMessage(
                    source_key=message_key,
                    sender_smtp=self.sender_smtp,
                    received_at=self.received_at,
                    attachment_count=2,
                    pdf_attachment_count=2,
                    attachments=tuple(attachments),
                    original_sender_smtp=self.original_sender_smtp,
                ),
            ),
        )


def test_outlook_diagnostic_reports_profile_without_requesting_sync():
    gateway = FakeGateway()
    service = OutlookService(gateway)

    result = service.diagnose()

    assert result.state == "ready"
    assert result.successful
    assert result.probe is not None
    assert result.probe.outlook_version == "15.0"
    assert result.probe.inbox_item_count == 42
    assert gateway.requests == [False]
    assert service.last_result() == result


def test_outlook_diagnostic_uses_configured_mailbox():
    gateway = FakeGateway()

    OutlookService(gateway).diagnose(mailbox="esensabiyatov@gmail.com")

    assert gateway.mailboxes == ["esensabiyatov@gmail.com"]


def test_outlook_send_receive_is_reported_only_as_requested():
    gateway = FakeGateway()
    service = OutlookService(gateway)

    result = service.request_send_receive()

    assert result.state == "sync_requested"
    assert result.probe is not None
    assert result.probe.sync_requested
    assert "в фоне" in result.message
    assert gateway.requests == [True]


def test_outlook_connection_error_has_explicit_non_success_state():
    class FailingGateway:
        def inspect(
            self,
            *,
            request_sync: bool = False,
            mailbox: str = "",
        ) -> OutlookProbe:
            raise OutlookConnectionError("Outlook недоступен")

    result = OutlookService(FailingGateway()).diagnose()

    assert result.state == "connection_error"
    assert not result.successful
    assert result.probe is None
    assert result.message == "Outlook недоступен"


def test_outlook_timeout_has_separate_status():
    class TimeoutGateway:
        def inspect(
            self,
            *,
            request_sync: bool = False,
            mailbox: str = "",
        ) -> OutlookProbe:
            raise OutlookProbeTimeoutError("Outlook не ответил")

    result = OutlookService(TimeoutGateway()).diagnose()

    assert result.state == "timeout"
    assert not result.successful


def test_subprocess_gateway_reads_structured_probe(monkeypatch):
    payload = {
        "ok": True,
        "probe": {
            "outlook_version": "15.0",
            "profile_name": "Рабочий профиль",
            "current_user": "Сотрудник",
            "default_store": "Банк",
            "inbox_name": "Входящие",
            "inbox_path": "\\\\Банк\\Входящие",
            "inbox_item_count": 42,
            "offline": False,
            "sync_group_count": 1,
            "accounts": [
                {
                    "display_name": "Рабочая почта",
                    "smtp_address": "employee@example.test",
                    "account_type": "Exchange",
                }
            ],
            "sync_requested": True,
        },
    }

    class FakeProcess:
        def communicate(self, *, input=None, timeout=None):
            assert timeout == 30
            assert json.loads(input)["mailbox"] == "esensabiyatov@gmail.com"
            return json.dumps(payload, ensure_ascii=False), ""

    monkeypatch.setattr(subprocess, "Popen", lambda *args, **kwargs: FakeProcess())

    probe = SubprocessOutlookGateway().inspect(
        request_sync=True,
        mailbox="esensabiyatov@gmail.com",
    )

    assert probe.sync_requested
    assert probe.accounts[0].account_type == "Exchange"
    assert probe.accounts[0].smtp_address == "employee@example.test"


def test_subprocess_gateway_terminates_worker_after_timeout(monkeypatch):
    class TimedOutProcess:
        def communicate(self, *, input=None, timeout=None):
            raise subprocess.TimeoutExpired("outlook-worker", timeout)

    process = TimedOutProcess()
    gateway = SubprocessOutlookGateway()
    terminated = []
    monkeypatch.setattr(subprocess, "Popen", lambda *args, **kwargs: process)
    monkeypatch.setattr(
        gateway,
        "_terminate_bridge_process_tree",
        lambda current: terminated.append(current),
    )

    try:
        gateway._call_bridge("diagnose", timeout_seconds=1)
    except OutlookProbeTimeoutError as exc:
        assert isinstance(exc.__cause__, subprocess.TimeoutExpired)
    else:
        raise AssertionError("Ожидался тайм-аут Outlook-worker")

    assert terminated == [process]


def test_subprocess_sync_keeps_late_certificate_watcher_in_main_process(
    monkeypatch,
):
    from gns_app.services import outlook_service

    calls: list[float] = []
    monkeypatch.setattr(
        outlook_service,
        "start_outlook_certificate_dialog_watcher",
        lambda *, timeout_seconds=20.0: calls.append(timeout_seconds),
    )
    gateway = SubprocessOutlookGateway(allow_insecure_certificate=True)
    monkeypatch.setattr(
        gateway,
        "_call_bridge",
        lambda action, *, input_payload: {
            "probe": {
                "outlook_version": "15.0",
                "sync_requested": True,
            }
        },
    )

    probe = gateway.inspect(request_sync=True)

    assert probe.sync_requested
    assert calls == [20.0]


def test_subprocess_diagnose_keeps_certificate_watcher_during_com_start(
    monkeypatch,
):
    from gns_app.services import outlook_service

    calls: list[float] = []
    monkeypatch.setattr(
        outlook_service,
        "start_outlook_certificate_dialog_watcher",
        lambda *, timeout_seconds=20.0: calls.append(timeout_seconds),
    )
    gateway = SubprocessOutlookGateway(allow_insecure_certificate=True)
    monkeypatch.setattr(
        gateway,
        "_call_bridge",
        lambda action, *, input_payload: {
            "probe": {
                "outlook_version": "15.0",
                "sync_requested": False,
            }
        },
    )

    probe = gateway.inspect(request_sync=False)

    assert not probe.sync_requested
    assert calls == [20.0]


def test_com_gateway_reuses_active_outlook_before_dispatch():
    application = object()

    class Client:
        @staticmethod
        def GetActiveObject(name):
            assert name == "Outlook.Application"
            return application

        @staticmethod
        def Dispatch(_name):
            raise AssertionError("Dispatch should not start a second Outlook")

    assert (
        PyWin32OutlookGateway._connect_outlook_application(Client())
        is application
    )


def test_com_gateway_starts_outlook_when_active_object_is_unavailable():
    application = object()

    class Client:
        @staticmethod
        def GetActiveObject(_name):
            raise RuntimeError("Outlook is closed")

        @staticmethod
        def Dispatch(name):
            assert name == "Outlook.Application"
            return application

    assert (
        PyWin32OutlookGateway._connect_outlook_application(Client())
        is application
    )


def test_com_gateway_reads_metadata_and_requests_send_receive(monkeypatch):
    account = SimpleNamespace(
        DisplayName="Рабочая почта",
        SmtpAddress="employee@example.test",
        AccountType=0,
    )

    class Accounts:
        Count = 1

        @staticmethod
        def Item(index):
            assert index == 1
            return account

    class Namespace:
        CurrentProfileName = "Рабочий профиль"
        CurrentUser = SimpleNamespace(Name="Сотрудник")
        DefaultStore = SimpleNamespace(DisplayName="Банк")
        SyncObjects = SimpleNamespace(Count=1)
        Offline = False

        def __init__(self):
            self.Accounts = Accounts()
            self.sync_requested = False

        @staticmethod
        def GetDefaultFolder(folder_id):
            assert folder_id == 6
            return SimpleNamespace(
                Name="Входящие",
                FolderPath="\\\\Банк\\Входящие",
                Items=SimpleNamespace(Count=42),
            )

        def SendAndReceive(self, show_progress):
            assert show_progress is False
            self.sync_requested = True

    namespace = Namespace()
    application = SimpleNamespace(
        Version="15.0",
        GetNamespace=lambda name: namespace,
    )
    pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    client = SimpleNamespace(Dispatch=lambda name: application)
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))

    probe = PyWin32OutlookGateway().inspect(request_sync=True)

    assert namespace.sync_requested
    assert probe.profile_name == "Рабочий профиль"
    assert probe.inbox_item_count == 42
    assert probe.accounts[0].account_type == "Exchange"


def test_com_gateway_diagnostic_reads_configured_mailbox_not_default(
    monkeypatch,
):
    configured_inbox = SimpleNamespace(
        Name="Входящие",
        FolderPath="\\\\esensabiyatov@gmail.com\\Входящие",
        Items=SimpleNamespace(Count=501),
        Store=SimpleNamespace(DisplayName="esensabiyatov@gmail.com"),
    )
    configured_store = SimpleNamespace(
        DisplayName="esensabiyatov@gmail.com",
        GetDefaultFolder=lambda folder_id: configured_inbox,
    )
    namespace = SimpleNamespace(
        Accounts=SimpleNamespace(Count=0),
        CurrentProfileName="Outlook",
        CurrentUser=SimpleNamespace(Name="esensabiyatov@gmail.com"),
        DefaultStore=SimpleNamespace(DisplayName="Файл данных Outlook"),
        SyncObjects=SimpleNamespace(Count=1),
        Offline=False,
        Stores=SimpleNamespace(Count=1, Item=lambda index: configured_store),
        GetDefaultFolder=lambda _folder_id: (_ for _ in ()).throw(
            AssertionError("Нельзя подменять настроенный ящик стандартным")
        ),
    )
    application = SimpleNamespace(
        Version="16.0",
        GetNamespace=lambda name: namespace,
    )
    pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    client = SimpleNamespace(Dispatch=lambda name: application)
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))

    probe = PyWin32OutlookGateway().inspect(
        mailbox="esensabiyatov@gmail.com"
    )

    assert probe.inbox_path == "\\\\esensabiyatov@gmail.com\\Входящие"
    assert probe.inbox_item_count == 501
    assert probe.default_store == "esensabiyatov@gmail.com"


def test_com_gateway_opens_inbox_only_when_outlook_has_no_window():
    class Inbox:
        def __init__(self):
            self.display_count = 0

        def Display(self):
            self.display_count += 1

    inbox = Inbox()
    application = SimpleNamespace(Explorers=SimpleNamespace(Count=0))

    PyWin32OutlookGateway._ensure_outlook_window(application, inbox)

    assert inbox.display_count == 1

    application.Explorers.Count = 1
    PyWin32OutlookGateway._ensure_outlook_window(application, inbox)

    assert inbox.display_count == 1


def test_com_gateway_creates_and_reopens_same_outlook_draft(
    monkeypatch,
    tmp_path,
):
    attachment = tmp_path / "Ответ_исх_9544.pdf"
    attachment.write_bytes(b"%PDF-1.4\n%%EOF")

    class UserProperty:
        Value = ""

    class UserProperties:
        def __init__(self):
            self.values = {}

        def Find(self, name):
            return self.values.get(name)

        def Add(self, name, property_type, add_to_folder_fields):
            assert name == "GNS App Draft Key"
            assert property_type == 1
            assert add_to_folder_fields is False
            value = UserProperty()
            self.values[name] = value
            return value

    class Attachments:
        def __init__(self):
            self.added = []

        def Add(self, source, attachment_type, position, display_name):
            self.added.append(
                (source, attachment_type, position, display_name)
            )

    class Items:
        def __init__(self):
            self.values = []

        @property
        def Count(self):
            return len(self.values)

        def Item(self, index):
            return self.values[index - 1]

    items = Items()

    class Inspector:
        def __init__(self):
            self.WindowState = 1
            self.activate_count = 0

        def Activate(self):
            self.activate_count += 1

    class Mail:
        EntryID = "draft-entry-1"

        def __init__(self):
            self.To = ""
            self.Subject = ""
            self.Body = ""
            self.UserProperties = UserProperties()
            self.Attachments = Attachments()
            self.display_count = 0
            self.GetInspector = Inspector()

        def Save(self):
            if self not in items.values:
                items.values.append(self)

        def Display(self):
            self.display_count += 1

    sending_account = SimpleNamespace(SmtpAddress="employee@bank.kg")
    namespace = SimpleNamespace(
        GetDefaultFolder=lambda folder_id: SimpleNamespace(Items=items),
        Accounts=SimpleNamespace(
            Count=1,
            Item=lambda _index: sending_account,
        ),
    )
    application = SimpleNamespace(
        GetNamespace=lambda name: namespace,
        CreateItem=lambda item_type: Mail(),
    )
    pythoncom = SimpleNamespace(
        CoInitialize=lambda: None,
        CoUninitialize=lambda: None,
    )
    client = SimpleNamespace(Dispatch=lambda name: application)
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))

    gateway = PyWin32OutlookGateway()
    first = gateway.create_draft(
        draft_key="gns-scan-one",
        recipient_email="recipient@example.test",
        subject="Ответ № 9544",
        body="",
        attachment_path=attachment,
        attachment_name=attachment.name,
        sending_account="employee@bank.kg",
    )
    second = gateway.create_draft(
        draft_key="gns-scan-one",
        recipient_email="recipient@example.test",
        subject="Ответ № 9544",
        body="",
        attachment_path=attachment,
        attachment_name=attachment.name,
        sending_account="employee@bank.kg",
        display=False,
    )

    assert not first.existing
    assert second.existing
    assert len(items.values) == 1
    assert len(items.values[0].Attachments.added) == 1
    assert items.values[0].SendUsingAccount is sending_account
    assert items.values[0].display_count == 1
    assert items.values[0].GetInspector.WindowState == 2
    assert items.values[0].GetInspector.activate_count == 1

    items.values.clear()
    with pytest.raises(OutlookConnectionError, match="дождитесь обновления"):
        gateway.create_draft(
            draft_key="gns-scan-one",
            recipient_email="recipient@example.test",
            subject="Ответ № 9544",
            body="",
            attachment_path=attachment,
            attachment_name=attachment.name,
            create_if_missing=False,
        )
    assert items.values == []


def test_com_gateway_selects_sent_folder_from_configured_account():
    sent_folder = object()
    delivery_store = SimpleNamespace(
        GetDefaultFolder=lambda folder_id: (
            sent_folder if folder_id == 5 else None
        )
    )
    account = SimpleNamespace(
        SmtpAddress="employee@bank.kg",
        DeliveryStore=delivery_store,
    )
    namespace = SimpleNamespace(
        Accounts=SimpleNamespace(Count=1, Item=lambda _index: account),
        GetDefaultFolder=lambda _folder_id: pytest.fail(
            "default store must not be used"
        ),
    )

    selected = PyWin32OutlookGateway._select_default_folder(
        namespace,
        "employee@bank.kg",
        5,
        "Отправленные",
    )

    assert selected is sent_folder


def test_com_gateway_sends_only_existing_pdf_draft_to_exact_account(
    monkeypatch,
):
    class Property:
        Value = "gns-test-scan-one"

    class UserProperties:
        def Find(self, name):
            return Property() if name == "GNS App Draft Key" else None

    class Attachment:
        FileName = "Тест.pdf"
        Size = 1024

    class Attachments:
        Count = 1

        def Item(self, index):
            assert index == 1
            return Attachment()

    class Mail:
        EntryID = "draft-entry-1"
        Subject = "Ответ на запрос ГНС"
        Body = "Направляем ответ. Документ приложен к письму."

        def __init__(self):
            self.To = "real.person@example.com"
            self.CC = "copy@example.com"
            self.BCC = "hidden@example.com"
            self.UserProperties = UserProperties()
            self.Attachments = Attachments()
            self.SendUsingAccount = None
            self.sent = 0

        def Save(self):
            pass

        def Send(self):
            self.sent += 1

    mail = Mail()

    class Items:
        Count = 1

        def Item(self, index):
            assert index == 1
            return mail

    account = SimpleNamespace(SmtpAddress="ovo@tolubaybank.kg")

    class Accounts:
        Count = 1

        def Item(self, index):
            assert index == 1
            return account

    namespace = SimpleNamespace(
        GetDefaultFolder=lambda folder_id: SimpleNamespace(Items=Items()),
        Accounts=Accounts(),
    )
    application = SimpleNamespace(GetNamespace=lambda name: namespace)
    pythoncom = SimpleNamespace(
        CoInitialize=lambda: None,
        CoUninitialize=lambda: None,
    )
    client = SimpleNamespace(Dispatch=lambda name: application)
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))

    result = PyWin32OutlookGateway().send_draft(
        draft_key="gns-test-scan-one",
        recipient_email="esensabiyatov@gmail.com",
        sending_account="",
    )

    assert result.recipient_email == "esensabiyatov@gmail.com"
    assert mail.To == "esensabiyatov@gmail.com"
    assert mail.CC == ""
    assert mail.BCC == ""
    assert mail.SendUsingAccount is account
    assert mail.sent == 1

    try:
        PyWin32OutlookGateway().send_draft(
            draft_key="gns-test-scan-one",
            recipient_email="real.person@example.com",
            sending_account="",
        )
    except OutlookConnectionError as exc:
        assert "esensabiyatov@gmail.com" in str(exc)
    else:
        raise AssertionError("Отправка не на тестовый адрес должна быть запрещена")

    assert mail.sent == 1


def test_com_gateway_confirms_only_exact_sent_item(monkeypatch, tmp_path):
    source_pdf = write_test_pdf(tmp_path / "Ответ_исх_9544.pdf")
    expected_sha256 = hashlib.sha256(source_pdf.read_bytes()).hexdigest()

    class Property:
        Value = "gns-scan-confirmed"

    class UserProperties:
        @staticmethod
        def Find(name):
            return Property() if name == "GNS App Draft Key" else None

    class Recipient:
        Type = 1
        Address = "recipient@example.test"
        PropertyAccessor = SimpleNamespace(
            GetProperty=lambda _schema: "recipient@example.test"
        )

    class Recipients:
        Count = 1

        @staticmethod
        def Item(index):
            assert index == 1
            return Recipient()

    class Attachment:
        FileName = source_pdf.name
        Size = source_pdf.stat().st_size

        @staticmethod
        def SaveAsFile(destination):
            shutil.copyfile(source_pdf, destination)

    class Attachments:
        Count = 1

        @staticmethod
        def Item(index):
            assert index == 1
            return Attachment()

    mail = SimpleNamespace(
        Class=43,
        EntryID="sent-entry-1",
        SentOn=datetime(2026, 9, 4, 10, 30, 0),
        Subject="Ответ № 9544",
        UserProperties=UserProperties(),
        Recipients=Recipients(),
        Attachments=Attachments(),
    )

    class Items:
        Count = 1

        @staticmethod
        def Sort(field, descending):
            assert (field, descending) == ("[SentOn]", True)

        @staticmethod
        def Item(index):
            assert index in {1, 2}
            return mail

    namespace = SimpleNamespace(
        GetDefaultFolder=lambda folder_id: SimpleNamespace(Items=Items())
    )
    application = SimpleNamespace(GetNamespace=lambda _name: namespace)
    monkeypatch.setitem(
        sys.modules,
        "pythoncom",
        SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None),
    )
    monkeypatch.setitem(
        sys.modules,
        "win32com",
        SimpleNamespace(client=SimpleNamespace(Dispatch=lambda _name: application)),
    )
    from gns_app.services import outlook_service

    watcher_finishes = []
    monkeypatch.setattr(
        outlook_service,
        "start_outlook_certificate_dialog_watcher",
        lambda **_kwargs: SimpleNamespace(
            finish=lambda **kwargs: watcher_finishes.append(kwargs)
        ),
    )
    staging_dir = outlook_service._prepare_outlook_staging_dir(
        tmp_path / "sent-staging"
    )
    gateway = PyWin32OutlookGateway(allow_insecure_certificate=True)
    scan = gateway.scan_sent_items(
        candidates=(
            OutlookSentCandidate(
                draft_key="gns-scan-confirmed",
                recipient_email="recipient@example.test",
                subject="Ответ № 9544",
                attachment_name=source_pdf.name,
                attachment_sha256=expected_sha256,
            ),
        ),
        sent_since=date(2026, 9, 3),
        staging_dir=staging_dir,
        max_attachment_bytes=10 * 1024 * 1024,
    )

    assert scan.inspected_mail_count == 1
    assert scan.candidate_mail_count == 1
    assert scan.rejected_mail_count == 0
    assert scan.ambiguous_candidate_count == 0
    assert scan.matches == (
        OutlookSentMatch(
            draft_key="gns-scan-confirmed",
            entry_id="sent-entry-1",
            sent_at="2026-09-04T10:30:00",
        ),
    )

    Items.Count = 2
    ambiguous = gateway.scan_sent_items(
        candidates=(
            OutlookSentCandidate(
                draft_key="gns-scan-confirmed",
                recipient_email="recipient@example.test",
                subject="Ответ № 9544",
                attachment_name=source_pdf.name,
                attachment_sha256=expected_sha256,
            ),
        ),
        sent_since=date(2026, 9, 3),
        staging_dir=staging_dir,
        max_attachment_bytes=10 * 1024 * 1024,
    )
    assert ambiguous.matches == ()
    assert ambiguous.ambiguous_candidate_count == 1
    assert watcher_finishes == [
        {"grace_seconds": 0.0},
        {"grace_seconds": 0.0},
    ]


def test_com_gateway_rejects_sent_item_with_different_recipient(
    monkeypatch,
    tmp_path,
):
    source_pdf = write_test_pdf(tmp_path / "Ответ.pdf")
    expected_sha256 = hashlib.sha256(source_pdf.read_bytes()).hexdigest()
    property_item = SimpleNamespace(Value="gns-scan-rejected")
    recipient = SimpleNamespace(
        Type=1,
        Address="other@example.test",
        PropertyAccessor=SimpleNamespace(
            GetProperty=lambda _schema: "other@example.test"
        ),
    )
    attachment = SimpleNamespace(
        FileName=source_pdf.name,
        Size=source_pdf.stat().st_size,
        SaveAsFile=lambda destination: shutil.copyfile(source_pdf, destination),
    )
    mail = SimpleNamespace(
        Class=43,
        EntryID="sent-entry-rejected",
        SentOn=datetime(2026, 9, 4, 11, 0, 0),
        Subject="Ответ",
        UserProperties=SimpleNamespace(
            Find=lambda name: (
                property_item if name == "GNS App Draft Key" else None
            )
        ),
        Recipients=SimpleNamespace(Count=1, Item=lambda _index: recipient),
        Attachments=SimpleNamespace(Count=1, Item=lambda _index: attachment),
    )
    items = SimpleNamespace(
        Count=1,
        Sort=lambda _field, _descending: None,
        Item=lambda _index: mail,
    )
    application = SimpleNamespace(
        GetNamespace=lambda _name: SimpleNamespace(
            GetDefaultFolder=lambda _folder_id: SimpleNamespace(Items=items)
        )
    )
    monkeypatch.setitem(
        sys.modules,
        "pythoncom",
        SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None),
    )
    monkeypatch.setitem(
        sys.modules,
        "win32com",
        SimpleNamespace(client=SimpleNamespace(Dispatch=lambda _name: application)),
    )
    from gns_app.services import outlook_service

    staging_dir = outlook_service._prepare_outlook_staging_dir(
        tmp_path / "rejected-staging"
    )
    scan = PyWin32OutlookGateway().scan_sent_items(
        candidates=(
            OutlookSentCandidate(
                draft_key="gns-scan-rejected",
                recipient_email="recipient@example.test",
                subject="Ответ",
                attachment_name=source_pdf.name,
                attachment_sha256=expected_sha256,
            ),
        ),
        sent_since=date(2026, 9, 3),
        staging_dir=staging_dir,
        max_attachment_bytes=10 * 1024 * 1024,
    )

    assert scan.matches == ()
    assert scan.rejected_mail_count == 1

    recipient.Address = "recipient@example.test"
    recipient.PropertyAccessor = SimpleNamespace(
        GetProperty=lambda _schema: "recipient@example.test"
    )
    wrong_hash = PyWin32OutlookGateway().scan_sent_items(
        candidates=(
            OutlookSentCandidate(
                draft_key="gns-scan-rejected",
                recipient_email="recipient@example.test",
                subject="Ответ",
                attachment_name=source_pdf.name,
                attachment_sha256="b" * 64,
            ),
        ),
        sent_since=date(2026, 9, 3),
        staging_dir=staging_dir,
        max_attachment_bytes=10 * 1024 * 1024,
    )
    assert wrong_hash.matches == ()
    assert wrong_hash.rejected_mail_count == 1


def test_subprocess_gateway_serializes_sent_scan(monkeypatch, tmp_path):
    from gns_app.services import outlook_service

    watcher_finishes = []
    monkeypatch.setattr(
        outlook_service,
        "start_outlook_certificate_dialog_watcher",
        lambda **_kwargs: SimpleNamespace(
            finish=lambda **kwargs: watcher_finishes.append(kwargs)
        ),
    )
    gateway = SubprocessOutlookGateway(allow_insecure_certificate=True)
    captured = {}

    def fake_bridge(action, *, input_payload, timeout_seconds):
        captured.update(
            action=action,
            input_payload=input_payload,
            timeout_seconds=timeout_seconds,
        )
        return {
            "ok": True,
            "sent_scan": {
                "inspected_mail_count": 1,
                "candidate_mail_count": 0,
                "rejected_mail_count": 0,
                "ambiguous_candidate_count": 0,
                "matches": [],
            },
        }

    monkeypatch.setattr(gateway, "_call_bridge", fake_bridge)
    staging_dir = outlook_service._prepare_outlook_staging_dir(
        tmp_path / "subprocess-sent-staging"
    )
    candidate = OutlookSentCandidate(
        draft_key="gns-scan-one",
        recipient_email="recipient@example.test",
        subject="Ответ",
        attachment_name="Ответ.pdf",
        attachment_sha256="a" * 64,
    )

    result = gateway.scan_sent_items(
        candidates=(candidate,),
        sent_since=date(2026, 9, 3),
        staging_dir=staging_dir,
        max_attachment_bytes=1024,
        mailbox="mailbox@example.test",
    )

    assert result.matches == ()
    assert captured["action"] == "sent"
    assert captured["input_payload"]["candidates"] == [
        {
            "draft_key": "gns-scan-one",
            "recipient_email": "recipient@example.test",
            "subject": "Ответ",
            "attachment_name": "Ответ.pdf",
            "attachment_sha256": "a" * 64,
        }
    ]
    assert captured["input_payload"]["mailbox"] == "mailbox@example.test"
    assert captured["input_payload"]["allow_insecure_certificate"] is True
    assert watcher_finishes == [{"grace_seconds": 0.0}]


def test_sent_status_loop_runs_without_inbox_automation(monkeypatch):
    from gns_app import main

    class FakeOutgoing:
        calls = 0

        @classmethod
        def reconcile_sent_messages(cls):
            cls.calls += 1
            return {"pending": 0, "confirmed": 0}

    async def stop_after_first_check(_delay):
        raise asyncio.CancelledError

    monkeypatch.setattr(main, "outlook_outgoing", FakeOutgoing())
    monkeypatch.setattr(main.asyncio, "sleep", stop_after_first_check)

    try:
        asyncio.run(main._outlook_sent_status_loop())
    except asyncio.CancelledError:
        pass

    assert FakeOutgoing.calls == 1


def test_outlook_subject_template_accepts_only_known_placeholders(workflow):
    service = OutlookOutgoingService(
        workflow.db,
        workflow.settings,
        OutlookService(object()),
    )

    service.update_subject_template("Ответ {office_name} от {date}")

    assert service.get_subject_template() == "Ответ {office_name} от {date}"
    with pytest.raises(OutlookIntegrationError, match="только внутри Word/PDF"):
        service.update_subject_template("Ответ № {outgoing_number}")
    try:
        service.update_subject_template("Ответ {unknown}")
    except OutlookIntegrationError as exc:
        assert "разрешены только" in str(exc)
    else:
        raise AssertionError("Неизвестная подстановка должна быть отклонена")


def test_outgoing_test_send_flag_forces_only_approved_recipient(workflow):
    test_settings = replace(
        workflow.settings,
        outlook_test_email="",
        outlook_allow_test_send=True,
    )
    service = OutlookOutgoingService(
        workflow.db,
        test_settings,
        OutlookService(object()),
    )

    assert service.test_mode_enabled()
    assert service.test_send_enabled()
    assert service.get_test_recipient() == "esensabiyatov@gmail.com"


def test_single_test_send_prepares_hidden_draft_automatically(
    workflow,
    monkeypatch,
):
    test_settings = replace(
        workflow.settings,
        outlook_allow_test_send=True,
    )
    service = OutlookOutgoingService(
        workflow.db,
        test_settings,
        OutlookService(object()),
    )
    calls = []

    def prepare(active_workflow, letter_id, *, display):
        calls.append((active_workflow, letter_id, display))
        return {"status": "draft_created"}

    monkeypatch.setattr(service, "_create_draft_locked", prepare)
    monkeypatch.setattr(
        service,
        "_send_test_message_locked",
        lambda active_workflow, letter_id: {
            "status": "sent",
            "letter_id": letter_id,
        },
    )

    result = service.send_test_message(workflow, "letter-1")

    assert result == {"status": "sent", "letter_id": "letter-1"}
    assert calls == [(workflow, "letter-1", False)]


def test_test_mode_sends_once_only_to_fixed_address(workflow):
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO response_groups(
            id, business_date, group_key, abs_bucket, status,
            district_place, recipient_position, recipient_full_name,
            recipient_display_name, employee_name, taxpayer_count,
            taxpayers_per_letter, response_path, created_at, updated_at
        ) VALUES (
            'test-mail-group', '2026-08-25', 'test-mail', 'not_found',
            'created', 'Тестовое подразделение', 'Начальнику',
            'Тестовый Получатель', 'Получателю Т. П.', 'Тестовый сотрудник',
            1, 1, 'unused.docx', ?, ?
        )
        """,
        (now, now),
    )
    workflow.db.execute(
        """
        INSERT INTO response_letters(
            id, response_group_id, letter_order, taxpayer_start_order,
            taxpayer_count, outgoing_number, created_at, updated_at
        ) VALUES ('test-mail-letter', 'test-mail-group', 1, 1, 1,
                  '9544', ?, ?)
        """,
        (now, now),
    )
    scan_dir = workflow.settings.runtime_dir / "signed_scans" / "mail-test"
    scan_dir.mkdir(parents=True)
    pdf_path = scan_dir / "signed.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=595, height=842)
    with pdf_path.open("wb") as stream:
        writer.write(stream)
    workflow.db.execute(
        """
        INSERT INTO signed_response_scans(
            id, response_letter_id, status, source, original_filename,
            original_path, pdf_path, sha256, size_bytes, page_count,
            correct_letter_confirmed, signature_confirmed,
            bank_seal_confirmed, confirmed_by, confirmed_at,
            created_at, updated_at
        ) VALUES (
            'test-mail-scan', 'test-mail-letter', 'confirmed', 'upload',
            'signed.pdf', ?, ?, 'test-sha', ?, 1, 1, 1, 1,
            'Тестовый сотрудник', ?, ?, ?
        )
        """,
        (str(pdf_path), str(pdf_path), pdf_path.stat().st_size, now, now, now),
    )
    empty_path = scan_dir / "empty.pdf"
    empty_path.touch()
    workflow.db.execute(
        """
        INSERT INTO response_letters(
            id, response_group_id, letter_order, taxpayer_start_order,
            taxpayer_count, outgoing_number, created_at, updated_at
        ) VALUES ('empty-mail-letter', 'test-mail-group', 2, 2, 1,
                  '9545', ?, ?)
        """,
        (now, now),
    )
    workflow.db.execute(
        """
        INSERT INTO signed_response_scans(
            id, response_letter_id, status, source, original_filename,
            original_path, pdf_path, sha256, size_bytes, page_count,
            correct_letter_confirmed, signature_confirmed,
            bank_seal_confirmed, confirmed_by, confirmed_at,
            created_at, updated_at
        ) VALUES (
            'empty-mail-scan', 'empty-mail-letter', 'confirmed', 'upload',
            'empty.pdf', ?, ?, 'empty-sha', 0, 1, 1, 1, 1,
            'Тестовый сотрудник', ?, ?, ?
        )
        """,
        (str(empty_path), str(empty_path), now, now, now),
    )

    class Gateway:
        def __init__(self):
            self.draft_requests = []
            self.send_requests = []

        def create_draft(self, **request):
            self.draft_requests.append(request)
            return OutlookDraftResult(
                entry_id="test-draft-entry",
                recipient_email=request["recipient_email"],
                subject=request["subject"],
                attachment_name=request["attachment_name"],
            )

        def send_draft(self, **request):
            self.send_requests.append(request)
            return OutlookSendResult(
                entry_id="test-sent-entry",
                recipient_email=request["recipient_email"],
            )

    gateway = Gateway()
    test_settings = replace(
        workflow.settings,
        outlook_test_email="esensabiyatov@gmail.com",
        outlook_allow_test_send=True,
    )
    outgoing = OutlookOutgoingService(
        workflow.db,
        test_settings,
        OutlookService(gateway),
    )

    assert outgoing.ready_test_send_count() == 1
    batch = outgoing.send_all_ready_test_messages(workflow)
    repeated = outgoing.send_test_message(workflow, "test-mail-letter")

    assert batch["ready"] == 1
    assert batch["sent"] == ["test-mail-letter"]
    assert batch["already_sent"] == []
    assert batch["errors"] == []
    assert outgoing.ready_test_send_count() == 0
    with pytest.raises(OutlookIntegrationError, match="PDF.*пустой"):
        outgoing.create_draft(workflow, "empty-mail-letter")
    assert gateway.draft_requests[0]["draft_key"] == (
        "gns-test-scan-test-mail-scan"
    )
    assert gateway.draft_requests[0]["display"] is False
    assert "Направляем ответ ГНС" in gateway.draft_requests[0]["body"]
    assert "9544" not in gateway.draft_requests[0]["body"]
    assert repeated["already_sent"]
    assert len(gateway.send_requests) == 1
    assert gateway.send_requests[0]["recipient_email"] == (
        "esensabiyatov@gmail.com"
    )
    assert gateway.send_requests[0]["sending_account"] == ""


def test_com_gateway_exports_supported_documents_only(
    monkeypatch,
    tmp_path,
):
    source_pdf = write_test_pdf(tmp_path / "source.pdf")
    source_png = tmp_path / "source.png"
    Image.new("RGB", (120, 160), "white").save(source_png)

    class Accessor:
        def __init__(self, internet_id):
            self.internet_id = internet_id

        def GetProperty(self, schema):
            if schema.endswith("0x1035001E"):
                return self.internet_id
            return ""

    class Attachment:
        def __init__(self, filename, source):
            self.FileName = filename
            self.source = source
            self.Size = source.stat().st_size
            self.saved = False

        def SaveAsFile(self, destination):
            self.saved = True
            shutil.copyfile(self.source, destination)

    class Attachments:
        def __init__(self):
            self.values = [
                Attachment("letter.pdf", source_pdf),
                Attachment("photo.png", source_png),
                Attachment("logo.gif", source_png),
            ]
            self.Count = len(self.values)

        def Item(self, index):
            return self.values[index - 1]

    class Mail:
        Class = 43
        SenderEmailType = "SMTP"
        Parent = SimpleNamespace(StoreID="store")
        EntryID = "entry"

        def __init__(self, sender, internet_id, received_time):
            self.SenderEmailAddress = sender
            self.PropertyAccessor = Accessor(internet_id)
            self.Attachments = Attachments()
            self.ReceivedTime = received_time

    allowed_mail = Mail(
        "esensabiyatov@gmail.com",
        "allowed@example.test",
        datetime(2026, 8, 21, 9, 0, 0),
    )
    blocked_mail = Mail(
        "blocked@example.test",
        "blocked@example.test",
        datetime(2026, 8, 19, 9, 0, 0),
    )
    older_mail = Mail(
        "older@example.test",
        "older@example.test",
        datetime(2026, 8, 18, 9, 0, 0),
    )

    class Items:
        def __init__(self):
            self.values = [older_mail, blocked_mail, allowed_mail]
            self.sort_calls = []
            self.item_calls = 0

        @property
        def Count(self):
            return len(self.values)

        def Sort(self, property_name, descending):
            self.sort_calls.append((property_name, descending))
            self.values.sort(
                key=lambda item: item.ReceivedTime,
                reverse=descending,
            )

        def Item(self, index):
            self.item_calls += 1
            return self.values[index - 1]

    items = Items()

    class Namespace:
        def __init__(self):
            self.sync_requested = False

            configured_store = SimpleNamespace(
                DisplayName="esensabiyatov@gmail.com",
                GetDefaultFolder=lambda folder_id: SimpleNamespace(Items=items),
            )
            self.Stores = SimpleNamespace(
                Count=1,
                Item=lambda index: configured_store,
            )

        @staticmethod
        def GetDefaultFolder(folder_id):
            raise AssertionError("Должен использоваться выбранный ящик")

        def SendAndReceive(self, show_progress):
            self.sync_requested = True

    namespace = Namespace()
    application = SimpleNamespace(GetNamespace=lambda name: namespace)
    pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    client = SimpleNamespace(Dispatch=lambda name: application)
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))

    from gns_app.services import outlook_service

    staging_dir = outlook_service._prepare_outlook_staging_dir(
        tmp_path / "staging"
    )
    scan = PyWin32OutlookGateway().scan_inbox(
        allowed_senders=frozenset({"esensabiyatov@gmail.com"}),
        allowed_domains=frozenset({"sti.gov.kg", "salyk.kg"}),
        received_since=date(2026, 8, 20),
        known_message_keys=frozenset(),
        staging_dir=staging_dir,
        max_attachment_bytes=10 * 1024 * 1024,
        mailbox="esensabiyatov@gmail.com",
    )

    assert not namespace.sync_requested
    assert scan.inspected_mail_count == 2
    assert scan.eligible_message_count == 1
    assert len(scan.messages) == 1
    assert scan.messages[0].pdf_attachment_count == 2
    assert len(scan.messages[0].attachments) == 2
    assert scan.messages[0].original_sender_smtp == ""
    assert Path(scan.messages[0].attachments[0].temporary_path).is_file()
    assert (staging_dir / ".gns-ready").is_file()
    assert allowed_mail.Attachments.values[0].saved
    assert allowed_mail.Attachments.values[1].saved
    assert not allowed_mail.Attachments.values[2].saved
    assert not blocked_mail.Attachments.values[0].saved
    assert not older_mail.Attachments.values[0].saved
    assert items.sort_calls == [("[ReceivedTime]", True)]
    assert items.item_calls == 2


def test_com_gateway_accepts_direct_gns_domain_sender():
    allowed_domains = frozenset({"sti.gov.kg", "salyk.kg"})
    direct_test_senders = frozenset()
    forwarding_senders = frozenset({"reception@bank.kg"})

    for sender in ("inspector@sti.gov.kg", "executor@salyk.kg"):
        mail = SimpleNamespace(
            SenderEmailType="SMTP",
            SenderEmailAddress=sender,
            Body="",
            HTMLBody="",
        )

        assert PyWin32OutlookGateway._is_allowed_sender(
            mail,
            direct_test_senders,
            allowed_domains,
            forwarding_senders,
        )
        assert PyWin32OutlookGateway._allowed_sender_with_original(
            mail,
            direct_test_senders,
            allowed_domains,
            forwarding_senders,
        ) == (True, "")


def test_com_gateway_accepts_configured_sender_exceptions():
    exact_mail = SimpleNamespace(
        SenderEmailType="SMTP",
        SenderEmailAddress="district@gmail.com",
        Body="",
        HTMLBody="",
    )
    domain_mail = SimpleNamespace(
        SenderEmailType="SMTP",
        SenderEmailAddress="employee@custom.gov.kg",
        Body="",
        HTMLBody="",
    )

    assert PyWin32OutlookGateway._is_allowed_sender(
        exact_mail,
        frozenset({"district@gmail.com"}),
        frozenset({"sti.gov.kg", "salyk.kg"}),
        frozenset(),
    )
    assert PyWin32OutlookGateway._is_allowed_sender(
        domain_mail,
        frozenset(),
        frozenset({"sti.gov.kg", "salyk.kg", "custom.gov.kg"}),
        frozenset(),
    )


def test_com_gateway_accepts_reception_with_forwarded_gns_sender():
    allowed_domains = frozenset({"sti.gov.kg", "salyk.kg"})
    direct_test_senders = frozenset()
    forwarding_senders = frozenset({"reception@bank.kg"})
    mail = SimpleNamespace(
        SenderEmailType="SMTP",
        SenderEmailAddress="reception@bank.kg",
        Body=(
            "-----Original Message-----\n"
            "From: УГНС Первомайского района <inspector@sti.gov.kg>"
        ),
        HTMLBody="",
    )

    assert (
        PyWin32OutlookGateway._forwarded_gns_sender(mail, allowed_domains)
        == "inspector@sti.gov.kg"
    )
    assert PyWin32OutlookGateway._is_allowed_sender(
        mail,
        direct_test_senders,
        allowed_domains,
        forwarding_senders,
    )
    assert PyWin32OutlookGateway._allowed_sender_with_original(
        mail,
        direct_test_senders,
        allowed_domains,
        forwarding_senders,
    ) == (True, "inspector@sti.gov.kg")


def test_com_gateway_accepts_forwarded_exact_sender_exception():
    mail = SimpleNamespace(
        SenderEmailType="SMTP",
        SenderEmailAddress="reception@bank.kg",
        Body="From: Районное УГНС <district@gmail.com>",
        HTMLBody="",
    )

    assert PyWin32OutlookGateway._allowed_sender_with_original(
        mail,
        frozenset({"district@gmail.com"}),
        frozenset({"sti.gov.kg", "salyk.kg"}),
        frozenset({"reception@bank.kg"}),
    ) == (True, "district@gmail.com")


def test_com_gateway_accepts_reception_with_cyrillic_forwarded_sender_label():
    allowed_domains = frozenset({"sti.gov.kg", "salyk.kg"})
    direct_test_senders = frozenset()
    forwarding_senders = frozenset({"reception@bank.kg"})
    mail = SimpleNamespace(
        SenderEmailType="SMTP",
        SenderEmailAddress="reception@bank.kg",
        Body="От: Исполнитель УГНС <executor@salyk.kg>",
        HTMLBody="",
    )

    assert (
        PyWin32OutlookGateway._forwarded_gns_sender(mail, allowed_domains)
        == "executor@salyk.kg"
    )
    assert PyWin32OutlookGateway._is_allowed_sender(
        mail,
        direct_test_senders,
        allowed_domains,
        forwarding_senders,
    )


def test_com_gateway_accepts_reception_with_multiple_forwarded_gns_senders():
    allowed_domains = frozenset({"sti.gov.kg", "salyk.kg"})
    direct_test_senders = frozenset()
    forwarding_senders = frozenset({"reception@bank.kg"})
    mail = SimpleNamespace(
        SenderEmailType="SMTP",
        SenderEmailAddress="reception@bank.kg",
        Body=(
            "From: Первый отправитель <first@sti.gov.kg>\n"
            "От: Второй отправитель <second@salyk.kg>"
        ),
        HTMLBody="",
    )

    assert PyWin32OutlookGateway._forwarded_gns_sender(
        mail,
        allowed_domains,
    ) == "first@sti.gov.kg"
    assert PyWin32OutlookGateway._is_allowed_sender(
        mail,
        direct_test_senders,
        allowed_domains,
        forwarding_senders,
    )


def test_com_gateway_skips_reception_without_reliable_forwarded_sender():
    allowed_domains = frozenset({"sti.gov.kg", "salyk.kg"})
    direct_test_senders = frozenset()
    forwarding_senders = frozenset({"reception@bank.kg"})
    mail = SimpleNamespace(
        SenderEmailType="SMTP",
        SenderEmailAddress="reception@bank.kg",
        Body="Просьба проверить письмо от inspector@sti.gov.kg.",
        HTMLBody="",
    )

    assert (
        PyWin32OutlookGateway._forwarded_gns_sender(mail, allowed_domains)
        == ""
    )
    assert not PyWin32OutlookGateway._is_allowed_sender(
        mail,
        direct_test_senders,
        allowed_domains,
        forwarding_senders,
    )


def test_com_gateway_skips_reception_with_non_gns_forwarded_sender():
    allowed_domains = frozenset({"sti.gov.kg", "salyk.kg"})
    direct_test_senders = frozenset()
    forwarding_senders = frozenset({"reception@bank.kg"})
    mail = SimpleNamespace(
        SenderEmailType="SMTP",
        SenderEmailAddress="reception@bank.kg",
        Body=(
            "From: External organisation <sender@example.com>"
        ),
        HTMLBody="",
    )

    assert (
        PyWin32OutlookGateway._forwarded_gns_sender(mail, allowed_domains)
        == ""
    )
    assert not PyWin32OutlookGateway._is_allowed_sender(
        mail,
        direct_test_senders,
        allowed_domains,
        forwarding_senders,
    )


def test_subprocess_gateway_receives_existing_writable_staging_dir(
    monkeypatch,
    tmp_path,
):
    from gns_app.services import outlook_service

    gateway = SubprocessOutlookGateway()
    staging_dir = outlook_service._prepare_outlook_staging_dir(
        tmp_path / "prepared-by-main-process"
    )

    def fake_call_bridge(action, *, input_payload, timeout_seconds=None):
        assert action == "scan"
        assert timeout_seconds == 120
        assert Path(input_payload["staging_dir"]) == staging_dir
        assert staging_dir.is_dir()
        worker_probe = staging_dir / "worker-write-probe"
        worker_probe.write_text("ready", encoding="ascii")
        assert worker_probe.read_text(encoding="ascii") == "ready"
        return {
            "ok": True,
            "scan": {
                "inspected_mail_count": 0,
                "eligible_message_count": 0,
                "known_message_count": 0,
                "scan_error_count": 0,
                "messages": [],
            },
        }

    monkeypatch.setattr(gateway, "_call_bridge", fake_call_bridge)

    scan = gateway.scan_inbox(
        allowed_senders=frozenset(),
        allowed_domains=frozenset({"sti.gov.kg", "salyk.kg"}),
        received_since=date(2026, 1, 1),
        known_message_keys=frozenset(),
        staging_dir=staging_dir,
        max_attachment_bytes=1024,
    )

    assert scan.messages == ()


def test_subprocess_scan_passes_certificate_mode_to_worker(
    monkeypatch,
    tmp_path,
):
    from gns_app.services import outlook_service

    class Watcher:
        def __init__(self):
            self.finished: list[float] = []

        def finish(self, grace_seconds=5.0):
            self.finished.append(grace_seconds)
            return False

    staging_dir = outlook_service._prepare_outlook_staging_dir(
        tmp_path / "prepared-certificate-scan"
    )
    watcher = Watcher()
    calls: list[float] = []
    monkeypatch.setattr(
        outlook_service,
        "start_outlook_certificate_dialog_watcher",
        lambda *, timeout_seconds=20.0: (
            calls.append(timeout_seconds) or watcher
        ),
    )
    gateway = SubprocessOutlookGateway(allow_insecure_certificate=True)

    def fake_call_bridge(action, *, input_payload, timeout_seconds=None):
        assert action == "scan"
        assert input_payload["allow_insecure_certificate"] is True
        return {
            "ok": True,
            "scan": {
                "inspected_mail_count": 0,
                "eligible_message_count": 0,
                "known_message_count": 0,
                "scan_error_count": 0,
                "messages": [],
            },
        }

    monkeypatch.setattr(gateway, "_call_bridge", fake_call_bridge)
    gateway.scan_inbox(
        allowed_senders=frozenset(),
        allowed_domains=frozenset({"sti.gov.kg", "salyk.kg"}),
        received_since=date(2026, 1, 1),
        known_message_keys=frozenset(),
        staging_dir=staging_dir,
        max_attachment_bytes=1024,
    )

    assert calls == [20.0]
    assert watcher.finished == [0.0]


def test_outlook_staging_uses_short_ascii_path_for_cyrillic_directory(
    monkeypatch,
    tmp_path,
):
    from gns_app.services import outlook_service

    cyrillic_root = tmp_path / "временные_письма"
    # The real Windows API guarantees that this path is another spelling of
    # the same directory.  Mock the API boundary so the test does not depend
    # on whether 8.3 names are enabled on the CI volume.
    short_alias = tmp_path / "GNSO~1" / "run"
    short_alias.mkdir(parents=True)
    monkeypatch.setattr(outlook_service.sys, "platform", "win32")
    monkeypatch.setattr(
        outlook_service,
        "_get_windows_short_path",
        lambda path: short_alias,
    )

    staging = outlook_service._create_outlook_staging_directory(
        (cyrillic_root,),
        run_id="run",
    )

    assert "временные_письма" in str(staging.canonical_dir)
    assert staging.worker_dir == short_alias
    assert str(staging.worker_dir).isascii()


def test_outlook_staging_uses_next_root_when_cyrillic_path_has_no_83_alias(
    monkeypatch,
    tmp_path,
):
    from gns_app.services import outlook_service

    cyrillic_root = tmp_path / "нет_короткого_пути"
    ascii_fallback = tmp_path / "GNSO"
    monkeypatch.setattr(outlook_service.sys, "platform", "win32")
    monkeypatch.setattr(
        outlook_service,
        "_get_windows_short_path",
        lambda path: None,
    )

    staging = outlook_service._create_outlook_staging_directory(
        (cyrillic_root, ascii_fallback),
        run_id="run",
    )

    assert staging.canonical_root == ascii_fallback.resolve()
    assert staging.worker_dir == staging.canonical_dir
    assert staging.worker_dir.is_dir()
    assert str(staging.worker_dir).isascii()


def test_outlook_staging_skips_unavailable_candidate_for_writable_fallback(
    monkeypatch,
    tmp_path,
):
    from gns_app.services import outlook_service

    unavailable = tmp_path / "not-a-directory"
    unavailable.write_text("blocked", encoding="ascii")
    fallback = tmp_path / "GNSO"
    monkeypatch.setattr(outlook_service.sys, "platform", "linux")

    staging = outlook_service._create_outlook_staging_directory(
        (unavailable, fallback),
        run_id="run",
    )

    assert staging.canonical_root == fallback.resolve()
    assert staging.canonical_dir.is_dir()
    assert (
        staging.canonical_dir
        / outlook_service.OUTLOOK_STAGING_READY_MARKER
    ).read_bytes()


def test_settings_routes_use_outlook_result(monkeypatch):
    from gns_app import main

    service = OutlookService(FakeGateway())
    monkeypatch.setattr(main, "outlook", service)

    check_response = main.check_outlook_connection()
    sync_response = main.request_outlook_send_receive()

    assert check_response.status_code == 303
    assert "message=" in check_response.headers["location"]
    assert sync_response.status_code == 303
    assert "message=" in sync_response.headers["location"]


def test_saving_outlook_date_queues_import_with_new_value(monkeypatch):
    from fastapi import BackgroundTasks

    from gns_app import main

    class FakeImporter:
        saved_date = ""
        saved_sender_rules = ""

        @classmethod
        def update_settings(
            cls,
            *,
            import_since,
            direct_sender_rules="",
            **kwargs,
        ):
            cls.saved_date = import_since
            cls.saved_sender_rules = direct_sender_rules

    class FakeOutgoing:
        @staticmethod
        def validate_subject_template(value):
            assert value == "Ответ {office_name}"

        @staticmethod
        def update_subject_template(value):
            assert value == "Ответ {office_name}"

    monkeypatch.setattr(main, "outlook_importer", FakeImporter())
    monkeypatch.setattr(main, "outlook_outgoing", FakeOutgoing())
    monkeypatch.setattr(
        main,
        "_queue_outlook_import",
        lambda background_tasks: FakeImporter.saved_date == "2026-08-20",
    )
    background_tasks = BackgroundTasks()

    response = main.update_outlook_import_settings(
        background_tasks,
        outlook_allowed_senders="reception@bank.kg",
        outlook_direct_sender_rules="district@gmail.com, @custom.gov.kg",
        outlook_import_since="2026-08-20",
        outlook_mailbox="employee@bank.kg",
        outlook_auto_enabled=True,
        outlook_auto_interval=5,
        outlook_subject_template="Ответ {office_name}",
    )
    assert response.status_code == 303
    assert FakeImporter.saved_sender_rules == (
        "district@gmail.com, @custom.gov.kg"
    )
    from urllib.parse import unquote

    assert "Получение документов запущено" in unquote(
        response.headers["location"]
    )


def test_outlook_draft_route_reports_created_and_existing(monkeypatch):
    from gns_app import main

    class FakeOutgoing:
        existing = False

        @classmethod
        def create_draft(cls, workflow, letter_id):
            assert letter_id == "letter-1"
            return {"existing_outlook_draft": cls.existing}

    monkeypatch.setattr(main, "outlook_outgoing", FakeOutgoing())

    created = main.create_response_letter_outlook_draft(
        "letter-1", "group-1"
    )
    FakeOutgoing.existing = True
    reopened = main.create_response_letter_outlook_draft(
        "letter-1", "group-1"
    )

    assert created.status_code == 303
    assert "message=" in created.headers["location"]
    assert reopened.status_code == 303
    assert "message=" in reopened.headers["location"]


def test_outlook_test_send_route_reports_sent_and_idempotent(monkeypatch):
    from gns_app import main

    class FakeOutgoing:
        repeated = False

        @classmethod
        def send_test_message(cls, workflow, letter_id):
            assert letter_id == "letter-1"
            return {"already_sent": cls.repeated}

        @staticmethod
        def get_test_recipient():
            return "esensabiyatov@gmail.com"

    monkeypatch.setattr(main, "outlook_outgoing", FakeOutgoing())

    sent = main.send_response_letter_outlook_test_message(
        "letter-1", "group-1"
    )
    FakeOutgoing.repeated = True
    repeated = main.send_response_letter_outlook_test_message(
        "letter-1", "group-1"
    )

    assert sent.status_code == 303
    assert "message=" in sent.headers["location"]
    assert repeated.status_code == 303
    assert "message=" in repeated.headers["location"]


def test_outlook_test_batch_send_route_reports_partial_result(monkeypatch):
    from urllib.parse import unquote

    from gns_app import main

    class FakeOutgoing:
        @staticmethod
        def send_all_ready_test_messages(workflow):
            return {
                "ready": 2,
                "sent": ["letter-1"],
                "already_sent": [],
                "errors": [
                    {"letter_id": "letter-2", "message": "PDF пустой."}
                ],
            }

    monkeypatch.setattr(main, "outlook_outgoing", FakeOutgoing())

    response = main.send_all_ready_outlook_test_messages()

    assert response.status_code == 303
    location = unquote(response.headers["location"])
    assert "Отправлено: 1" in location
    assert "С ошибкой: 1" in location
    assert "PDF пустой" in location


def test_outlook_import_route_queues_only_new_uploads(monkeypatch):
    from fastapi import BackgroundTasks

    from gns_app import main

    class FakeImporter:
        def record_automation_status(self, **kwargs):
            self.status = kwargs

        @staticmethod
        def import_new(workflow):
            return {
                "new_messages": 1,
                "saved_attachments": 2,
                "known": 3,
                "duplicate_attachments": 1,
                "scan_errors": 0,
                "errors": [],
                "imported": ["upload-1"],
            }

    monkeypatch.setattr(main, "outlook_importer", FakeImporter())
    background_tasks = BackgroundTasks()
    main._release_outlook_import()
    try:
        response = main.import_outlook_attachments(background_tasks)

        assert response.status_code == 303
        assert "message=" in response.headers["location"]
        assert len(background_tasks.tasks) == 1
    finally:
        main._release_outlook_import()


def test_work_outlook_import_returns_immediately_as_json(monkeypatch):
    from fastapi import BackgroundTasks
    from starlette.requests import Request

    from gns_app import main

    class FakeImporter:
        @staticmethod
        def get_automation_status():
            return {
                "state": "running",
                "checked_at": "2026-08-21T09:00:00+06:00",
                "message": "Проверка почты запущена в фоне.",
                "new_messages": 0,
                "saved_attachments": 0,
                "errors": 0,
            }

    monkeypatch.setattr(main, "outlook_importer", FakeImporter())
    monkeypatch.setattr(main, "_queue_outlook_import", lambda tasks: True)
    request = Request(
        {
            "type": "http",
            "method": "POST",
            "scheme": "http",
            "path": "/work/outlook/import",
            "raw_path": b"/work/outlook/import",
            "query_string": b"",
            "headers": [(b"accept", b"application/json")],
            "client": ("127.0.0.1", 50000),
            "server": ("127.0.0.1", 8765),
        }
    )

    response = main.import_outlook_from_work(request, BackgroundTasks())

    assert response.status_code == 202
    assert json.loads(response.body) == {
        "accepted": True,
        "active": True,
        "status": FakeImporter.get_automation_status(),
    }


def test_outlook_import_saves_message_once_and_reuses_duplicate_content(
    workflow,
    tmp_path,
):
    workflow.initialize_employee_profiles()
    gateway = FakeImportGateway(write_test_pdf(tmp_path / "source.pdf"))
    service = OutlookService(gateway)
    importer = OutlookInboxImporter(workflow.db, workflow.settings, service)
    importer.update_settings(
        allowed_senders="esensabiyatov@gmail.com",
        import_since="2026-08-20",
        mailbox="esensabiyatov@gmail.com",
    )

    first = importer.import_new(workflow)
    second = importer.import_new(workflow)

    assert first["new_messages"] == 1
    assert first["saved_attachments"] == 2
    assert len(first["imported"]) == 1
    assert first["duplicate_attachments"] == 1
    assert not first["errors"]
    assert not second["imported"]
    assert second["known"] == 1
    assert gateway.requests == [True, True]
    message = workflow.db.fetch_one(
        "SELECT * FROM outlook_messages WHERE source_key = ?",
        ("a" * 64,),
    )
    attachments = workflow.db.fetch_all(
        "SELECT * FROM outlook_attachments ORDER BY attachment_index"
    )
    assert message is not None
    assert message["status"] == "completed"
    assert message["original_sender_smtp"] is None
    assert len(attachments) == 2
    assert {row["status"] for row in attachments} == {
        "imported",
        "duplicate_content",
    }
    assert all(Path(row["stored_path"]).exists() for row in attachments)

    workflow.reset_processing_data()

    preserved = workflow.db.fetch_all(
        "SELECT stored_path, upload_id FROM outlook_attachments"
    )
    assert len(preserved) == 2
    assert all(row["upload_id"] is None for row in preserved)
    assert all(Path(row["stored_path"]).exists() for row in preserved)

    restored = workflow.import_inbox()
    assert len(restored["imported"]) == 1
    assert workflow.get_upload(restored["imported"][0])["intake_source"] == (
        "inbox_folder"
    )

    restored_from_outlook = importer.import_new(workflow)
    assert len(restored_from_outlook["imported"]) == 0
    relinked = workflow.db.fetch_all(
        "SELECT upload_id FROM outlook_attachments ORDER BY attachment_index"
    )
    assert {row["upload_id"] for row in relinked} == {
        restored["imported"][0]
    }
    incoming = workflow.list_incoming_work()
    assert incoming[0]["intake_source"] == "inbox_folder"
    assert incoming[0]["sender_smtp"] is None


def test_outlook_import_accepts_png_attachment(workflow, tmp_path):
    workflow.initialize_employee_profiles()
    source = tmp_path / "letter.png"
    Image.new("RGB", (240, 320), "white").save(source)
    gateway = FakeImportGateway(source, attachment_filename="letter.png")
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(gateway),
    )
    importer.update_settings(
        allowed_senders="esensabiyatov@gmail.com",
        import_since="2026-08-20",
    )

    summary = importer.import_new(workflow)

    assert len(summary["imported"]) == 1
    assert not summary["errors"]
    upload = workflow.get_upload(summary["imported"][0])
    assert upload["original_filename"] == "letter.png"
    assert upload["page_count"] == 1


def test_manual_pdf_with_same_content_as_outlook_stays_manual(
    workflow,
    tmp_path,
):
    workflow.initialize_employee_profiles()
    source = write_test_pdf(tmp_path / "same-content.pdf")
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(FakeImportGateway(source)),
    )
    importer.update_settings(
        allowed_senders="esensabiyatov@gmail.com",
        import_since="2026-08-20",
    )
    importer.import_new(workflow)
    workflow.reset_processing_data()

    with source.open("rb") as stream:
        manual_upload_id = workflow.create_upload("manual.pdf", stream)
    importer.import_new(workflow)

    manual = workflow.get_upload(manual_upload_id)
    incoming = workflow.list_incoming_work()
    row = next(item for item in incoming if item["id"] == manual_upload_id)

    assert manual["intake_source"] == "manual_upload"
    assert row["sender_smtp"] is None
    assert row["original_sender_smtp"] is None


def test_outlook_import_uses_changed_earlier_date_for_unseen_message(
    workflow,
    tmp_path,
):
    workflow.initialize_employee_profiles()
    gateway = FakeImportGateway(
        write_test_pdf(tmp_path / "older-source.pdf"),
        received_at="2026-08-21T09:00:00+06:00",
    )
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(gateway),
    )
    importer.update_settings(
        allowed_senders="esensabiyatov@gmail.com",
        import_since="2026-08-22",
        mailbox="esensabiyatov@gmail.com",
    )

    first = importer.import_new(workflow)
    assert not first["imported"]

    importer.update_settings(
        allowed_senders="esensabiyatov@gmail.com",
        import_since="2026-08-20",
        mailbox="esensabiyatov@gmail.com",
    )
    second = importer.import_new(workflow)

    assert len(second["imported"]) == 1
    assert gateway.received_since_requests == [
        date(2026, 8, 22),
        date(2026, 8, 20),
    ]


def test_outlook_import_stores_and_shows_original_sender_from_relay(
    workflow,
    tmp_path,
    monkeypatch,
):
    from starlette.testclient import TestClient

    from gns_app import main

    workflow.initialize_employee_profiles()
    gateway = FakeImportGateway(
        write_test_pdf(tmp_path / "relay-source.pdf"),
        sender_smtp="reception@bank.kg",
        original_sender_smtp="first@sti.gov.kg",
    )
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(gateway),
    )
    importer.update_settings(
        allowed_senders="reception@bank.kg",
        import_since="2026-08-20",
    )

    result = importer.import_new(workflow)

    message = workflow.db.fetch_one(
        """
        SELECT sender_smtp, original_sender_smtp
        FROM outlook_messages
        WHERE source_key = ?
        """,
        ("a" * 64,),
    )
    assert message == {
        "sender_smtp": "reception@bank.kg",
        "original_sender_smtp": "first@sti.gov.kg",
    }
    incoming = workflow.list_incoming_work()
    assert incoming[0]["sender_smtp"] == "reception@bank.kg"
    assert incoming[0]["original_sender_smtp"] == "first@sti.gov.kg"

    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO cases(id, upload_id, status, created_at, updated_at)
        VALUES ('relay-history-case', ?, 'needs_review', ?, ?)
        """,
        (result["imported"][0], now, now),
    )
    history = workflow.list_letter_history("first@sti.gov.kg")
    assert history["total"] == 1
    assert history["items"][0]["sender_smtp"] == "reception@bank.kg"
    assert history["items"][0]["original_sender_smtp"] == "first@sti.gov.kg"

    monkeypatch.setattr(main, "workflow", workflow)
    client = TestClient(main.app)
    incoming_response = client.get("/")
    history_response = client.get("/history?query=first%40sti.gov.kg")

    assert incoming_response.status_code == 200
    assert "Транспортный: reception@bank.kg" in incoming_response.text
    assert "Исходный: first@sti.gov.kg" in incoming_response.text
    assert history_response.status_code == 200
    assert "Транспортный: reception@bank.kg" in history_response.text
    assert "Исходный: first@sti.gov.kg" in history_response.text


def test_outlook_import_backfills_known_relay_metadata_without_reimporting(
    workflow,
    tmp_path,
):
    workflow.initialize_employee_profiles()
    gateway = FakeImportGateway(
        write_test_pdf(tmp_path / "known-relay-source.pdf"),
        sender_smtp="reception@bank.kg",
    )
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(gateway),
    )
    importer.update_settings(
        allowed_senders="reception@bank.kg",
        import_since="2026-08-20",
    )

    first = importer.import_new(workflow)
    workflow.db.execute(
        """
        UPDATE outlook_messages
        SET original_sender_smtp = NULL
        WHERE source_key = ?
        """,
        ("a" * 64,),
    )
    gateway.original_sender_smtp = "first@sti.gov.kg"

    second = importer.import_new(workflow)

    message = workflow.db.fetch_one(
        """
        SELECT sender_smtp, original_sender_smtp, status
        FROM outlook_messages
        WHERE source_key = ?
        """,
        ("a" * 64,),
    )
    assert first["saved_attachments"] == 2
    assert second["saved_attachments"] == 0
    assert not second["imported"]
    assert message == {
        "sender_smtp": "reception@bank.kg",
        "original_sender_smtp": "first@sti.gov.kg",
        "status": "completed",
    }


def test_outlook_import_uses_local_staging_when_portable_runtime_is_unwritable(
    workflow,
    tmp_path,
    monkeypatch,
):
    from gns_app.services import outlook_service

    workflow.initialize_employee_profiles()
    blocked_runtime = tmp_path / "read-only-portable-runtime"
    blocked_runtime.write_text("not a directory", encoding="utf-8")
    local_app_data = tmp_path / "local-app-data"
    monkeypatch.setenv("LOCALAPPDATA", str(local_app_data))
    # This test verifies candidate order only; 8.3 hand-off representation is
    # covered separately and would otherwise obscure the selected root.
    monkeypatch.setattr(
        outlook_service,
        "_get_windows_short_path",
        lambda path: None,
    )
    gateway = FakeImportGateway(write_test_pdf(tmp_path / "source.pdf"))
    importer = OutlookInboxImporter(
        workflow.db,
        replace(workflow.settings, runtime_dir=blocked_runtime),
        OutlookService(gateway),
    )
    importer.update_settings(
        allowed_senders="esensabiyatov@gmail.com",
        import_since="2026-08-20",
        mailbox="esensabiyatov@gmail.com",
    )

    result = importer.import_new(workflow)

    expected_root = (local_app_data / "GNSO").resolve()
    assert result["saved_attachments"] == 2
    assert gateway.staging_dirs[0].parent == expected_root
    assert not gateway.staging_dirs[0].exists()


def test_outlook_import_falls_back_to_configured_inbox_staging(
    workflow,
    tmp_path,
    monkeypatch,
):
    from gns_app.services import outlook_service

    workflow.initialize_employee_profiles()
    blocked_root = tmp_path / "blocked-root"
    blocked_root.write_text("not a directory", encoding="utf-8")
    inbox_root = workflow.get_inbox_dir()
    gateway = FakeImportGateway(write_test_pdf(tmp_path / "source.pdf"))
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(gateway),
    )
    importer.update_settings(
        allowed_senders="esensabiyatov@gmail.com",
        import_since="2026-08-20",
        mailbox="esensabiyatov@gmail.com",
    )
    monkeypatch.setattr(
        outlook_service,
        "_outlook_staging_roots",
        lambda runtime_dir, inbox_dir=None: (
            blocked_root,
            inbox_root / ".gns_outlook_staging",
        ),
    )
    # Isolate candidate selection from the Windows 8.3 worker representation.
    monkeypatch.setattr(
        outlook_service,
        "_get_windows_short_path",
        lambda path: None,
    )

    result = importer.import_new(workflow)

    assert result["saved_attachments"] == 2
    assert gateway.staging_dirs[0].parent == (
        inbox_root / ".gns_outlook_staging"
    ).resolve()
    assert not gateway.staging_dirs[0].exists()


def test_outlook_import_settings_validate_sender_and_date(workflow):
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(FakeGateway()),
    )

    importer.update_settings(
        allowed_senders="first@example.test, SECOND@example.test",
        direct_sender_rules="district@gmail.com, @custom.gov.kg",
        import_since=date.today().isoformat(),
        mailbox="esensabiyatov@gmail.com",
    )

    assert importer.get_forwarding_senders() == frozenset(
        {"first@example.test", "second@example.test"}
    )
    assert importer.get_direct_sender_exceptions() == frozenset(
        {"district@gmail.com"}
    )
    assert importer.get_allowed_senders() == frozenset({"district@gmail.com"})
    assert importer.get_allowed_domains() == frozenset(
        {"sti.gov.kg", "salyk.kg", "custom.gov.kg"}
    )
    assert importer.get_sender_exception_rules() == (
        "district@gmail.com",
        "@custom.gov.kg",
    )
    assert importer.get_import_since() == date.today()
    assert importer.get_mailbox() == "esensabiyatov@gmail.com"
    assert not importer.get_auto_enabled()
    assert importer.get_auto_interval_minutes() == 5

    importer.update_settings(
        allowed_senders="first@example.test",
        import_since=date.today().isoformat(),
        mailbox="esensabiyatov@gmail.com",
        auto_enabled=True,
        auto_interval_minutes=7,
    )

    assert importer.get_auto_enabled()
    assert importer.get_auto_interval_minutes() == 7
    assert importer.get_allowed_senders() == frozenset({"district@gmail.com"})

    importer.update_settings(
        allowed_senders="first@example.test",
        direct_sender_rules="",
        import_since=date.today().isoformat(),
    )
    assert importer.get_allowed_senders() == frozenset()
    assert importer.get_allowed_domains() == frozenset(
        {"sti.gov.kg", "salyk.kg"}
    )


@pytest.mark.parametrize(
    "rules",
    ["не домен", "user@", "https://example.com"],
)
def test_outlook_sender_exceptions_reject_invalid_rules(workflow, rules):
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(FakeGateway()),
    )

    with pytest.raises(OutlookIntegrationError, match="адрес или домен"):
        importer.update_settings(
            allowed_senders="",
            direct_sender_rules=rules,
            import_since=date.today().isoformat(),
        )


def test_outlook_production_mode_defaults_to_gns_domains_only(workflow):
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(FakeGateway()),
    )

    assert importer.get_allowed_senders() == frozenset()
    assert importer.get_allowed_domains() == frozenset(
        {"sti.gov.kg", "salyk.kg"}
    )


def test_outlook_test_mode_ignores_all_other_sender_rules(workflow):
    test_settings = replace(
        workflow.settings,
        outlook_test_email="esensabiyatov@gmail.com",
    )
    importer = OutlookInboxImporter(
        workflow.db,
        test_settings,
        OutlookService(FakeGateway()),
    )
    importer.update_settings(
        allowed_senders="real.person@example.com",
        direct_sender_rules="district@gmail.com, @custom.gov.kg",
        import_since=date.today().isoformat(),
    )

    assert importer.get_allowed_senders() == frozenset(
        {"esensabiyatov@gmail.com"}
    )
    assert importer.get_allowed_domains() == frozenset()


def test_outlook_automation_status_is_persisted(workflow):
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(FakeGateway()),
    )

    assert importer.get_automation_status()["state"] == "not_started"

    importer.record_automation_status(
        state="success",
        message="Проверка завершена.",
        new_messages=2,
        saved_attachments=3,
    )

    status = importer.get_automation_status()
    assert status["state"] == "success"
    assert status["new_messages"] == 2
    assert status["saved_attachments"] == 3
    assert status["checked_at"]


def test_automatic_outlook_run_processes_new_uploads_and_records_result(
    monkeypatch,
):
    from gns_app import main

    class FakeAutomaticImporter:
        def __init__(self):
            self.statuses = []

        def record_automation_status(self, **values):
            self.statuses.append(values)

        @staticmethod
        def import_new(workflow):
            return {
                "new_messages": 1,
                "saved_attachments": 1,
                "scan_errors": 0,
                "errors": [],
                "imported": ["upload-1"],
            }

    class FakeWorkflow:
        def __init__(self):
            self.processed = []

        def process_upload(self, upload_id):
            self.processed.append(upload_id)

    importer = FakeAutomaticImporter()
    fake_workflow = FakeWorkflow()
    monkeypatch.setattr(main, "outlook_importer", importer)
    monkeypatch.setattr(main, "workflow", fake_workflow)

    asyncio.run(main._run_automated_outlook_import())

    assert fake_workflow.processed == ["upload-1"]
    assert [item["state"] for item in importer.statuses] == [
        "running",
        "success",
    ]


def test_automatic_outlook_error_is_recorded_for_next_retry(monkeypatch):
    from gns_app import main

    class FailingImporter:
        def __init__(self):
            self.statuses = []

        def record_automation_status(self, **values):
            self.statuses.append(values)

        @staticmethod
        def import_new(workflow):
            raise OutlookIntegrationError("Outlook временно недоступен")

    importer = FailingImporter()
    monkeypatch.setattr(main, "outlook_importer", importer)

    asyncio.run(main._run_automated_outlook_import())

    assert [item["state"] for item in importer.statuses] == [
        "running",
        "error",
    ]
    assert "временно недоступен" in importer.statuses[-1]["message"]


def test_settings_outlook_actions_use_saved_mailbox(workflow, monkeypatch):
    from gns_app import main

    gateway = FakeGateway()
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(gateway),
    )
    importer.update_settings(
        allowed_senders="",
        import_since=date.today().isoformat(),
        mailbox="esensabiyatov@gmail.com",
    )
    monkeypatch.setattr(main, "outlook", OutlookService(gateway))
    monkeypatch.setattr(main, "outlook_importer", importer)

    main.check_outlook_connection()
    main.request_outlook_send_receive()

    assert gateway.mailboxes == [
        "esensabiyatov@gmail.com",
        "esensabiyatov@gmail.com",
    ]


def test_settings_page_renders_outlook_diagnostic(workflow, monkeypatch):
    from starlette.requests import Request

    from gns_app import main

    service = OutlookService(FakeGateway())
    service.diagnose()
    importer = OutlookInboxImporter(workflow.db, workflow.settings, service)
    importer.update_settings(
        allowed_senders="reception@bank.kg",
        direct_sender_rules="district@gmail.com, @custom.gov.kg",
        import_since=date.today().isoformat(),
    )
    monkeypatch.setattr(main, "workflow", workflow)
    monkeypatch.setattr(main, "outlook", service)
    monkeypatch.setattr(main, "outlook_importer", importer)
    request = Request(
        {
            "type": "http",
            "app": main.app,
            "asgi": {"version": "3.0"},
            "http_version": "1.1",
            "method": "GET",
            "scheme": "http",
            "path": "/settings",
            "raw_path": b"/settings",
            "query_string": b"",
            "headers": [],
            "client": ("127.0.0.1", 50000),
            "server": ("127.0.0.1", 8765),
        }
    )

    response = main.settings_page(request)
    body = response.body.decode("utf-8")

    assert "Подключение Outlook" in body
    assert "Рабочий профиль" in body
    assert "Синхронизировать Outlook" in body
    assert 'data-settings-tab="processing"' in body
    assert 'data-settings-tab="outlook"' in body
    assert 'data-settings-tab="service"' in body
    assert 'data-settings-form' in body
    assert 'name="outlook_direct_sender_rules"' in body
    assert "district@gmail.com, @custom.gov.kg" in body
    assert "Всегда разрешены: @sti.gov.kg, @salyk.kg" in body
    for heading in (
        "Исполнитель",
        "Документы и сканер",
        "OCR",
        "АБС и реестры",
        "Интерфейс",
        "Диагностика",
    ):
        assert heading in body
