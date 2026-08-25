from __future__ import annotations

import asyncio
import json
import shutil
import subprocess
import sys
from dataclasses import replace
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace

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

    def inspect(self, *, request_sync: bool = False) -> OutlookProbe:
        self.requests.append(request_sync)
        return sample_probe(sync_requested=request_sync)


class FakeImportGateway(FakeGateway):
    def __init__(self, source_pdf: Path):
        super().__init__()
        self.source_pdf = source_pdf
        self.scan_requests: list[frozenset[str]] = []

    def scan_inbox(
        self,
        *,
        allowed_senders,
        allowed_domains,
        received_since,
        known_message_keys,
        staging_dir,
        max_attachment_bytes,
        mailbox="",
    ):
        assert mailbox in {"", "esasabiyatov@gmail.com"}
        assert staging_dir.is_dir()
        self.scan_requests.append(known_message_keys)
        message_key = "a" * 64
        if message_key in known_message_keys:
            return OutlookInboxScan(
                inspected_mail_count=1,
                eligible_message_count=1,
                known_message_count=1,
                scan_error_count=0,
                messages=(),
            )
        message_dir = staging_dir / message_key
        message_dir.mkdir(parents=True, exist_ok=True)
        attachments = []
        for index in (1, 2):
            temporary = message_dir / f"{index:03d}_test.pdf.part"
            shutil.copyfile(self.source_pdf, temporary)
            attachments.append(
                OutlookScannedAttachment(
                    attachment_index=index,
                    original_filename="test.pdf",
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
                    sender_smtp="esasabiyatov@gmail.com",
                    received_at="2026-08-21T09:00:00+06:00",
                    attachment_count=2,
                    pdf_attachment_count=2,
                    attachments=tuple(attachments),
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
        def inspect(self, *, request_sync: bool = False) -> OutlookProbe:
            raise OutlookConnectionError("Outlook недоступен")

    result = OutlookService(FailingGateway()).diagnose()

    assert result.state == "connection_error"
    assert not result.successful
    assert result.probe is None
    assert result.message == "Outlook недоступен"


def test_outlook_timeout_has_separate_status():
    class TimeoutGateway:
        def inspect(self, *, request_sync: bool = False) -> OutlookProbe:
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

    def fake_run(*args, **kwargs):
        return subprocess.CompletedProcess(
            args=args[0],
            returncode=0,
            stdout=json.dumps(payload, ensure_ascii=False),
            stderr="",
        )

    monkeypatch.setattr(subprocess, "run", fake_run)

    probe = SubprocessOutlookGateway().inspect(request_sync=True)

    assert probe.sync_requested
    assert probe.accounts[0].account_type == "Exchange"
    assert probe.accounts[0].smtp_address == "employee@example.test"


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

    class Mail:
        EntryID = "draft-entry-1"

        def __init__(self):
            self.To = ""
            self.Subject = ""
            self.Body = ""
            self.UserProperties = UserProperties()
            self.Attachments = Attachments()
            self.display_count = 0

        def Save(self):
            if self not in items.values:
                items.values.append(self)

        def Display(self):
            self.display_count += 1

    namespace = SimpleNamespace(
        GetDefaultFolder=lambda folder_id: SimpleNamespace(Items=items)
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
    )
    second = gateway.create_draft(
        draft_key="gns-scan-one",
        recipient_email="recipient@example.test",
        subject="Ответ № 9544",
        body="",
        attachment_path=attachment,
        attachment_name=attachment.name,
    )

    assert not first.existing
    assert second.existing
    assert len(items.values) == 1
    assert len(items.values[0].Attachments.added) == 1
    assert items.values[0].display_count == 2


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

    class Attachments:
        Count = 1

        def Item(self, index):
            assert index == 1
            return Attachment()

    class Mail:
        EntryID = "draft-entry-1"

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
        recipient_email="esasabiyatov@gmail.com",
        sending_account="",
    )

    assert result.recipient_email == "esasabiyatov@gmail.com"
    assert mail.To == "esasabiyatov@gmail.com"
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
        assert "esasabiyatov@gmail.com" in str(exc)
    else:
        raise AssertionError("Отправка не на тестовый адрес должна быть запрещена")

    assert mail.sent == 1


def test_outlook_subject_template_accepts_only_known_placeholders(workflow):
    service = OutlookOutgoingService(
        workflow.db,
        workflow.settings,
        OutlookService(object()),
    )

    service.update_subject_template(
        "Ответ № {outgoing_number} от {date}"
    )

    assert service.get_subject_template() == (
        "Ответ № {outgoing_number} от {date}"
    )
    try:
        service.update_subject_template("Ответ {unknown}")
    except OutlookIntegrationError as exc:
        assert "разрешены только" in str(exc)
    else:
        raise AssertionError("Неизвестная подстановка должна быть отклонена")


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
        outlook_test_email="esasabiyatov@gmail.com",
        outlook_allow_test_send=True,
    )
    outgoing = OutlookOutgoingService(
        workflow.db,
        test_settings,
        OutlookService(gateway),
    )

    draft = outgoing.create_draft(workflow, "test-mail-letter")
    sent = outgoing.send_test_message(workflow, "test-mail-letter")
    repeated = outgoing.send_test_message(workflow, "test-mail-letter")

    assert draft["recipient_email"] == "esasabiyatov@gmail.com"
    assert gateway.draft_requests[0]["draft_key"] == (
        "gns-test-scan-test-mail-scan"
    )
    assert sent["status"] == "sent"
    assert not sent["already_sent"]
    assert repeated["already_sent"]
    assert len(gateway.send_requests) == 1
    assert gateway.send_requests[0]["recipient_email"] == (
        "esasabiyatov@gmail.com"
    )
    assert gateway.send_requests[0]["sending_account"] == ""


def test_com_gateway_filters_sender_and_exports_only_pdf(
    monkeypatch,
    tmp_path,
):
    source_pdf = write_test_pdf(tmp_path / "source.pdf")

    class Accessor:
        def __init__(self, internet_id):
            self.internet_id = internet_id

        def GetProperty(self, schema):
            if schema.endswith("0x1035001E"):
                return self.internet_id
            return ""

    class Attachment:
        def __init__(self, filename):
            self.FileName = filename
            self.Size = source_pdf.stat().st_size
            self.saved = False

        def SaveAsFile(self, destination):
            self.saved = True
            shutil.copyfile(source_pdf, destination)

    class Attachments:
        def __init__(self):
            self.values = [Attachment("letter.pdf"), Attachment("logo.png")]
            self.Count = len(self.values)

        def Item(self, index):
            return self.values[index - 1]

    class Mail:
        Class = 43
        SenderEmailType = "SMTP"
        ReceivedTime = datetime(2026, 8, 21, 9, 0, 0)
        Parent = SimpleNamespace(StoreID="store")
        EntryID = "entry"

        def __init__(self, sender, internet_id):
            self.SenderEmailAddress = sender
            self.PropertyAccessor = Accessor(internet_id)
            self.Attachments = Attachments()

    allowed_mail = Mail("esasabiyatov@gmail.com", "allowed@example.test")
    blocked_mail = Mail("blocked@example.test", "blocked@example.test")

    class Items:
        values = [allowed_mail, blocked_mail]
        Count = 2

        @classmethod
        def Item(cls, index):
            return cls.values[index - 1]

    class Namespace:
        def __init__(self):
            self.sync_requested = False

            configured_store = SimpleNamespace(
                DisplayName="esasabiyatov@gmail.com",
                GetDefaultFolder=lambda folder_id: SimpleNamespace(Items=Items()),
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

    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()
    scan = PyWin32OutlookGateway().scan_inbox(
        allowed_senders=frozenset({"esasabiyatov@gmail.com"}),
        allowed_domains=frozenset({"sti.gov.kg", "salyk.kg"}),
        received_since=date(2026, 8, 20),
        known_message_keys=frozenset(),
        staging_dir=staging_dir,
        max_attachment_bytes=10 * 1024 * 1024,
        mailbox="esasabiyatov@gmail.com",
    )

    assert not namespace.sync_requested
    assert scan.inspected_mail_count == 2
    assert scan.eligible_message_count == 1
    assert len(scan.messages) == 1
    assert scan.messages[0].pdf_attachment_count == 1
    assert len(scan.messages[0].attachments) == 1
    assert Path(scan.messages[0].attachments[0].temporary_path).is_file()
    assert allowed_mail.Attachments.values[0].saved
    assert not allowed_mail.Attachments.values[1].saved
    assert not blocked_mail.Attachments.values[0].saved


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
            return "esasabiyatov@gmail.com"

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


def test_outlook_import_route_queues_only_new_uploads(monkeypatch):
    from fastapi import BackgroundTasks

    from gns_app import main

    class FakeImporter:
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

    response = main.import_outlook_attachments(background_tasks)

    assert response.status_code == 303
    assert "message=" in response.headers["location"]
    assert len(background_tasks.tasks) == 1


def test_outlook_import_saves_message_once_and_reuses_duplicate_content(
    workflow,
    tmp_path,
):
    workflow.initialize_employee_profiles()
    gateway = FakeImportGateway(write_test_pdf(tmp_path / "source.pdf"))
    service = OutlookService(gateway)
    importer = OutlookInboxImporter(workflow.db, workflow.settings, service)
    importer.update_settings(
        allowed_senders="esasabiyatov@gmail.com",
        import_since="2026-08-20",
        mailbox="esasabiyatov@gmail.com",
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


def test_outlook_import_settings_validate_sender_and_date(workflow):
    importer = OutlookInboxImporter(
        workflow.db,
        workflow.settings,
        OutlookService(FakeGateway()),
    )

    importer.update_settings(
        allowed_senders="first@example.test, SECOND@example.test",
        import_since=date.today().isoformat(),
        mailbox="esasabiyatov@gmail.com",
    )

    assert importer.get_allowed_senders() == frozenset(
        {"first@example.test", "second@example.test"}
    )
    assert importer.get_import_since() == date.today()
    assert importer.get_mailbox() == "esasabiyatov@gmail.com"
    assert not importer.get_auto_enabled()
    assert importer.get_auto_interval_minutes() == 5

    importer.update_settings(
        allowed_senders="first@example.test",
        import_since=date.today().isoformat(),
        mailbox="esasabiyatov@gmail.com",
        auto_enabled=True,
        auto_interval_minutes=7,
    )

    assert importer.get_auto_enabled()
    assert importer.get_auto_interval_minutes() == 7


def test_outlook_test_mode_ignores_all_other_sender_rules(workflow):
    test_settings = replace(
        workflow.settings,
        outlook_test_email="esasabiyatov@gmail.com",
    )
    importer = OutlookInboxImporter(
        workflow.db,
        test_settings,
        OutlookService(FakeGateway()),
    )
    importer.update_settings(
        allowed_senders="real.person@example.com",
        import_since=date.today().isoformat(),
    )

    assert importer.get_allowed_senders() == frozenset(
        {"esasabiyatov@gmail.com"}
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


def test_settings_page_renders_outlook_diagnostic(workflow, monkeypatch):
    from starlette.requests import Request

    from gns_app import main

    service = OutlookService(FakeGateway())
    service.diagnose()
    monkeypatch.setattr(main, "workflow", workflow)
    monkeypatch.setattr(main, "outlook", service)
    monkeypatch.setattr(
        main,
        "outlook_importer",
        OutlookInboxImporter(workflow.db, workflow.settings, service),
    )
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

    assert "Подключение к почте" in body
    assert "Рабочий профиль" in body
    assert "Получить почту сейчас" in body
