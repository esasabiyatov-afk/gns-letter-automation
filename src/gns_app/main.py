from __future__ import annotations

import asyncio
import os
import sys
from contextlib import asynccontextmanager, suppress
from datetime import date
from pathlib import Path
from threading import Lock
from urllib.parse import quote

import uvicorn
from fastapi import (
    BackgroundTasks,
    FastAPI,
    File,
    Form,
    HTTPException,
    Request,
    UploadFile,
)
from fastapi.responses import (
    FileResponse,
    HTMLResponse,
    JSONResponse,
    RedirectResponse,
)
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from gns_app.config import settings
from gns_app.database import Database
from gns_app.diagnostics import (
    APP_VERSION,
    create_diagnostic_bundle,
    record_event,
    record_exception,
)
from gns_app.domain import AbsStatus
from gns_app.services.storage import StorageError, ensure_within
from gns_app.services.outlook_service import (
    OutlookInboxImporter,
    OutlookIntegrationError,
    OutlookOutgoingService,
    OutlookService,
    SubprocessOutlookGateway,
)
from gns_app.services.registry_service import RegistryLookupError
from gns_app.services.scanner_service import ScannerCancelled, ScannerError
from gns_app.services.taxpayer_service import TaxpayerKind, classify_taxpayer
from gns_app.services.windows_focus import start_foreground_watcher
from gns_app.services.word_desktop import WordDesktopError, open_word_document
from gns_app.services.workflow import (
    WorkflowService,
    WorkflowValidationError,
)
from gns_app.text_cleanup import clean_taxpayer_name


PACKAGE_DIR = Path(__file__).resolve().parent
db = Database(settings.database_path)
workflow = WorkflowService(db, settings)
outlook = OutlookService(
    SubprocessOutlookGateway(
        allow_insecure_certificate=(
            settings.outlook_allow_insecure_certificate
        )
    )
)
outlook_importer = OutlookInboxImporter(db, settings, outlook)
outlook_outgoing = OutlookOutgoingService(db, settings, outlook)
_outlook_import_activity_lock = Lock()
_outlook_import_active = False
OUTLOOK_SENT_CHECK_INTERVAL_SECONDS = 30
INCOMING_PAGE_SIZE = 50


def _claim_outlook_import() -> bool:
    global _outlook_import_active
    with _outlook_import_activity_lock:
        if _outlook_import_active:
            return False
        _outlook_import_active = True
        return True


def _release_outlook_import() -> None:
    global _outlook_import_active
    with _outlook_import_activity_lock:
        _outlook_import_active = False


def _outlook_import_is_active() -> bool:
    with _outlook_import_activity_lock:
        return _outlook_import_active


def _queue_outlook_import(background_tasks: BackgroundTasks) -> bool:
    """Queue one import and return immediately to the browser."""

    if not _claim_outlook_import():
        return False
    try:
        outlook_importer.record_automation_status(
            state="running",
            message="Проверка почты запущена в фоне.",
        )
        background_tasks.add_task(
            _run_automated_outlook_import,
            claimed=True,
        )
    except Exception:
        _release_outlook_import()
        raise
    return True


async def _resume_interrupted_uploads(upload_ids: list[str]) -> None:
    for upload_id in upload_ids:
        await asyncio.to_thread(
            workflow.process_upload,
            upload_id,
            True,
        )


async def _run_automated_outlook_import(*, claimed: bool = False) -> bool:
    if not claimed and not _claim_outlook_import():
        return False
    try:
        outlook_importer.record_automation_status(
            state="running",
            message="Проверяется папка входящих Outlook.",
        )
        try:
            summary = await asyncio.to_thread(
                outlook_importer.import_new,
                workflow,
            )
        except OutlookIntegrationError as exc:
            outlook_importer.record_automation_status(
                state="error",
                message=str(exc),
                errors=1,
            )
            return True
        except Exception:
            outlook_importer.record_automation_status(
                state="error",
                message=(
                    "Автоматическая проверка завершилась технической ошибкой. "
                    "Приложение повторит попытку по расписанию."
                ),
                errors=1,
            )
            return True

        processing_errors = 0
        for upload_id in summary["imported"]:
            try:
                await asyncio.to_thread(workflow.process_upload, upload_id)
            except Exception:
                processing_errors += 1
        error_count = (
            int(summary["scan_errors"])
            + int(bool(summary.get("sync_error")))
            + len(summary["errors"])
            + processing_errors
        )
        state = "warning" if error_count else "success"
        message = (
            f"Проверка завершена. Новых писем: {summary['new_messages']}; "
            f"сохранено PDF: {summary['saved_attachments']}; ошибок: {error_count}."
        )
        outlook_importer.record_automation_status(
            state=state,
            message=message,
            new_messages=summary["new_messages"],
            saved_attachments=summary["saved_attachments"],
            errors=error_count,
        )
        return True
    finally:
        _release_outlook_import()


async def _outlook_automation_loop() -> None:
    while True:
        try:
            enabled = outlook_importer.get_auto_enabled()
            if enabled:
                await _run_automated_outlook_import()
                delay_seconds = (
                    outlook_importer.get_auto_interval_minutes() * 60
                )
            else:
                delay_seconds = 15
        except Exception:
            # Даже временная ошибка БД или адаптера не должна навсегда
            # остановить фоновую проверку.
            delay_seconds = 15
        await asyncio.sleep(delay_seconds)


async def _outlook_sent_status_loop() -> None:
    while True:
        delay_seconds = OUTLOOK_SENT_CHECK_INTERVAL_SECONDS
        try:
            await asyncio.to_thread(outlook_outgoing.reconcile_sent_messages)
        except OutlookIntegrationError as exc:
            record_exception("outlook", "scan_sent_items", exc)
            delay_seconds = 5 * 60
        except Exception as exc:
            # Ошибка Outlook не означает, что письмо не отправлено. Статус
            # черновика сохраняется, а следующая проверка повторит попытку.
            record_exception("outlook", "scan_sent_items", exc)
            delay_seconds = 5 * 60
        await asyncio.sleep(delay_seconds)


@asynccontextmanager
async def lifespan(_: FastAPI):
    record_event(
        "application",
        "startup",
        "started",
        details={
            "abs_mode": settings.abs_mode,
            "frozen": bool(getattr(sys, "frozen", False)),
        },
        runtime_dir=settings.runtime_dir,
    )
    try:
        settings.ensure_directories()
        db.initialize()
        workflow.initialize_employee_profiles()
        workflow.initialize_gns_offices()
        workflow.initialize_gns_office_emails()
        workflow.reconcile_gns_office_districts()
        workflow.reconcile_gns_office_hints()
        workflow.reconcile_structured_ocr_hints()
        workflow.reconcile_recipient_display_names()
        workflow.reprocess_incomplete_official_documents()
        workflow.reconcile_official_qr_pages()
        workflow.reconcile_case_match_reviews()
        workflow.reconcile_confident_scan_decisions()
        workflow.repair_cleaned_responses()
        ocr_health = workflow.ocr.health_check()
        record_event(
            "ocr",
            "startup_health",
            (
                "ready"
                if ocr_health.get("fast_initialized")
                else "unavailable"
            ),
            details=ocr_health,
            runtime_dir=settings.runtime_dir,
        )
    except Exception as exc:
        record_exception(
            "application",
            "startup",
            exc,
            runtime_dir=settings.runtime_dir,
        )
        raise
    record_event(
        "application",
        "startup",
        "ready",
        runtime_dir=settings.runtime_dir,
    )
    resume_ids = workflow.interrupted_upload_ids()
    if resume_ids:
        asyncio.create_task(_resume_interrupted_uploads(resume_ids))
    outlook_task = asyncio.create_task(_outlook_automation_loop())
    outlook_sent_task = asyncio.create_task(_outlook_sent_status_loop())
    try:
        yield
    finally:
        outlook_task.cancel()
        outlook_sent_task.cancel()
        with suppress(asyncio.CancelledError):
            await outlook_task
        with suppress(asyncio.CancelledError):
            await outlook_sent_task
        record_event(
            "application",
            "shutdown",
            "completed",
            runtime_dir=settings.runtime_dir,
        )


app = FastAPI(
    title="Автоматизатор писем",
    docs_url=None,
    redoc_url=None,
    lifespan=lifespan,
)
app.mount(
    "/static",
    StaticFiles(directory=PACKAGE_DIR / "static"),
    name="static",
)
templates = Jinja2Templates(directory=PACKAGE_DIR / "templates")
ASSET_VERSION = str(
    max(
        (PACKAGE_DIR / "static" / "styles.css").stat().st_mtime_ns,
        (PACKAGE_DIR / "static" / "app.js").stat().st_mtime_ns,
    )
)
templates.env.globals["asset_version"] = ASSET_VERSION


STATUS_LABELS = {
    "registered": "Зарегистрировано",
    "processing": "Обрабатывается",
    "preview_ready": "Превью готово",
    "qr_resolved": "QR обработан",
    "scan_fallback": "Обработка скана",
    "needs_review": "Требует проверки",
    "manually_confirmed": "Подтверждено",
    "ready": "Готово",
    "ready_for_abs": "Готово к АБС",
    "abs_checking": "Проверка АБС",
    "manual_period_rule": "Требуется проверка ОДБ",
    "ready_for_response": "Можно создать ответ",
    "response_created": "Ответ создан",
    "completed": "Успешно",
    "technical_error": "Техническая ошибка",
    "collecting": "Формируется",
}

PAGE_TYPE_LABELS = {
    "letter": "Письмо",
    "decision": "Решение",
    "attachment": "Приложение",
    "other": "Прочее",
    "unknown": "Не определено",
}

QR_STATUS_LABELS = {
    "not_started": "ещё не проверен",
    "found": "найден",
    "invalid_url": "посторонний адрес",
    "not_found": "не прочитан",
    "decode_error": "неоднозначный результат",
}

OCR_STATUS_LABELS = {
    "not_started": "ещё не запускался",
    "skipped_official": "не нужен: использован QR",
    "embedded_text": "текстовый слой PDF",
    "requires_engine": "модель не установлена",
    "completed": "распознан локально",
    "low_confidence": "низкая уверенность",
    "error": "ошибка OCR",
}

SOURCE_LABELS = {
    "qr_link": "QR найден, официальная версия ещё не получена",
    "qr_official": "Официальная версия по QR",
    "ocr_scan": "Скан и ручное подтверждение",
    "manual": "Введено и подтверждено сотрудником",
    "profile": "Локальный справочник",
    "system": "Системное значение",
}

ABS_STATUS_LABELS = {
    "not_checked": "не проверен",
    "checking": "проверяется",
    "found": "найден",
    "not_found": "не найден",
    "multiple": "несколько совпадений",
    "auth_error": "ошибка входа",
    "unavailable": "АБС недоступна",
    "technical_error": "техническая ошибка",
}

ODB_STATUS_LABELS = {
    "not_checked": "не проверен",
    "found": "найден в ОДБ",
    "not_found": "не найден в ОДБ",
}

REGISTRY_STATUS_LABELS = {
    "match": "название совпадает",
    "mismatch": "название отличается",
    "not_found": "ИНН не найден",
    "multiple": "несколько совпадений",
    "error": "сверка недоступна",
    "not_applicable": "не применяется к ИП/физлицу",
    "classification_uncertain": "нужно уточнить вид налогоплательщика",
}

EVENT_LABELS = {
    "upload_registered": "PDF зарегистрирован",
    "upload_reprocess_requested": "Запрошена повторная обработка",
    "page_reprocess_requested": "Запрошена повторная обработка страницы",
    "precise_ocr_requested": "Запрошен точный OCR",
    "precise_ocr_completed": "Точный OCR завершён",
    "precise_ocr_failed": "Ошибка точного OCR",
    "official_qr_auto_completed": "Страница подтверждена официальным QR",
    "official_qr_manual_confirmation_removed": (
        "Лишнее ручное подтверждение страницы отменено"
    ),
    "official_qr_source_restored": "Восстановлен официальный источник QR",
    "orphan_cases_removed": "Удалены устаревшие черновики обращений",
    "page_processed": "Страница обработана",
    "page_processing_error": "Ошибка обработки страницы",
    "scan_case_created": "Создано обращение по скану",
    "official_document_processed": "Официальная версия обработана",
    "page_manually_confirmed": "Страница подтверждена сотрудником",
    "case_review_reopened": "Обращение возвращено на ручную проверку",
    "page_reopened_for_case_review": "Лист возвращён на ручную проверку",
    "fake_abs_checked": "Выполнена тестовая проверка АБС",
    "abs_checked": "Выполнена проверка АБС Tolubay",
    "fake_abs_startup_rolled_back": "Отменена массовая автопроверка АБС",
    "fake_abs_batch_checked": "Выполнена пакетная проверка АБС",
    "abs_batch_checked": "Выполнена пакетная проверка АБС Tolubay",
    "odb_checked": "Записана ручная проверка ОДБ",
    "processing_data_reset": "Сброшены данные обработки",
    "registry_checked": "Выполнена сверка с ОсОО.KG",
    "registry_variance_accepted": "Подтверждены расхождения ОсОО.KG",
    "inbox_scanned": "Просканирована папка входящих",
    "outlook_inbox_scanned": "Проверены новые письма Outlook",
    "outlook_import_settings_updated": "Обновлены настройки Outlook",
    "outlook_outgoing_subject_updated": "Обновлена тема исходящих писем",
    "outlook_draft_created": "Создан черновик исходящего письма Outlook",
    "outlook_draft_failed": "Ошибка создания черновика Outlook",
    "outlook_test_message_sent": "Отправлено тестовое письмо Outlook",
    "outlook_test_send_failed": "Ошибка тестовой отправки Outlook",
    "outlook_sent_confirmed": "Отправка подтверждена папкой Outlook",
    "outlook_resend_draft_created": "Создан повторный черновик Outlook",
    "outlook_resend_confirmed": "Повторная отправка подтверждена Outlook",
    "gns_offices_replaced": "Обновлён справочник налоговых органов",
    "gns_office_emails_updated": "Обновлены официальные email подразделений",
    "gns_office_location_expanded": "Район дополнен областью или городом",
    "gns_office_suggested": "Предложен налоговый орган по OCR",
    "ocr_recipient_suggested": "Предложены должность и ФИО по OCR",
    "ocr_period_suggested": "Предложен период по OCR",
    "ocr_taxpayers_suggested": "Предложены налогоплательщики по OCR",
    "ocr_taxpayer_name_conflict": "OCR расходится в наименовании по одному ИНН",
    "ocr_high_resolution_render_failed": (
        "Не удалось подготовить высокое разрешение для OCR"
    ),
    "decision_without_letter_review_required": (
        "Решение остановлено: письмо не найдено"
    ),
    "confident_decision_reconciled": (
        "Уверенное решение снято с ручной проверки"
    ),
    "decision_confidence_review_required": (
        "Недостаточно признаков решения"
    ),
    "response_created": "Создан ответ Word",
    "grouped_response_created": "Создан общий ответ Word",
    "outgoing_number_assigned": "Назначен исходящий номер",
    "outgoing_number_changed": "Изменён исходящий номер",
    "abs_account_manually_checked": "Вручную проверено наличие счёта в АБС",
    "signed_response_scan_registered": "Зарегистрирован подписанный скан",
    "signed_response_scan_confirmed": "Подписанный скан проверен",
    "district_place_edge_noise_removed": (
        "Удалены лишние краевые символы в реквизите ГНС"
    ),
    "response_regenerated_after_cleanup": (
        "Ответ пересоздан после очистки реквизита ГНС"
    ),
    "employee_profile_added": "Добавлен исполнитель",
    "active_employee_selected": "Выбран активный исполнитель",
    "review_interface_settings_updated": "Обновлены настройки ручной проверки",
    "case_match_review_created": "Создано ручное сравнение обращений",
    "case_match_review_resolved": "Принято решение по сравнению обращений",
}

ENTITY_LABELS = {
    "upload": "PDF",
    "page": "Страница",
    "case": "Обращение",
    "settings": "Настройка",
    "response_group": "Общий ответ",
    "response_letter": "Исходящее письмо",
    "taxpayer": "Налогоплательщик",
    "case_match_review": "Сравнение обращений",
}


def taxpayer_word(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "лицо"
    if count % 10 in {2, 3, 4} and count % 100 not in {12, 13, 14}:
        return "лица"
    return "лиц"


def case_word(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "обращение"
    if count % 10 in {2, 3, 4} and count % 100 not in {12, 13, 14}:
        return "обращения"
    return "обращений"


def page_word(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "лист"
    if count % 10 in {2, 3, 4} and count % 100 not in {12, 13, 14}:
        return "листа"
    return "листов"


def context(request: Request, **values):
    active_employee = workflow.get_active_employee()
    return {
        "request": request,
        "status_labels": STATUS_LABELS,
        "page_type_labels": PAGE_TYPE_LABELS,
        "qr_status_labels": QR_STATUS_LABELS,
        "ocr_status_labels": OCR_STATUS_LABELS,
        "source_labels": SOURCE_LABELS,
        "abs_status_labels": ABS_STATUS_LABELS,
        "odb_status_labels": ODB_STATUS_LABELS,
        "registry_status_labels": REGISTRY_STATUS_LABELS,
        "event_labels": EVENT_LABELS,
        "entity_labels": ENTITY_LABELS,
        "taxpayer_word": taxpayer_word,
        "case_word": case_word,
        "page_word": page_word,
        "threshold": workflow.get_period_threshold().isoformat(),
        "active_employee": active_employee,
        "employee_profiles": workflow.list_employee_profiles(),
        "ui_preferences": workflow.get_ui_preferences(),
        "abs_session_active": workflow.abs_session_active(),
        "abs_login_required": False,
        "abs_session_supported": workflow.abs_session_supported(),
        "abs_is_fake": workflow.abs_is_fake(),
        "abs_tls_verification_disabled": (
            workflow.abs_tls_verification_disabled()
        ),
        "outlook_insecure_certificate_confirmation": (
            settings.outlook_allow_insecure_certificate
        ),
        "outlook_test_mode": outlook_outgoing.test_mode_enabled(),
        "outlook_test_email": outlook_outgoing.get_test_recipient(),
        "outlook_test_send_enabled": outlook_outgoing.test_send_enabled(),
        **values,
    }


def period_review_status(case: dict) -> dict[str, object]:
    route = str(case.get("period_route") or "")
    if route in {"odb", "no_odb"}:
        return {"complete": True, "requires_odb": route == "odb"}
    start_text = str(case.get("period_start") or "")
    end_text = str(case.get("period_end") or "")
    if not start_text or not end_text:
        return {"complete": False, "requires_odb": None}
    try:
        start = date.fromisoformat(start_text)
        date.fromisoformat(end_text)
    except ValueError:
        return {"complete": False, "requires_odb": None}
    return {
        "complete": True,
        "requires_odb": start < workflow.get_period_threshold(),
    }


def _incoming_workspace_values(
    incoming_sort: str,
    incoming_page: int = 1,
) -> dict[str, object]:
    safe_sort = (
        incoming_sort if incoming_sort in {"newest", "oldest"} else "newest"
    )
    total = workflow.count_incoming_work()
    page_count = max(1, (total + INCOMING_PAGE_SIZE - 1) // INCOMING_PAGE_SIZE)
    selected_page = min(max(int(incoming_page), 1), page_count)
    return {
        "uploads": workflow.list_incoming_work(
            limit=INCOMING_PAGE_SIZE,
            sort_order=safe_sort,
            offset=(selected_page - 1) * INCOMING_PAGE_SIZE,
        ),
        "incoming_sort": safe_sort,
        "incoming_page": {
            "page": selected_page,
            "page_count": page_count,
            "page_size": INCOMING_PAGE_SIZE,
            "total": total,
        },
        "inbox_dir": str(workflow.get_inbox_dir()),
        "outlook_auto_enabled": outlook_importer.get_auto_enabled(),
        "outlook_auto_status": outlook_importer.get_automation_status(),
        "outlook_import_stats": outlook_importer.stats(),
    }


@app.get("/", response_class=HTMLResponse)
def index(
    request: Request,
    tab: str = "incoming",
    incoming_sort: str = "newest",
    incoming_page: int = 1,
    response_view: str = "prepare",
    focus_letter: str = "",
    outgoing_start: str = "",
    undo_page_id: str = "",
    abs_login: bool = False,
    message: str = "",
    error: str = "",
):
    if tab not in {"incoming", "review", "responses"}:
        tab = "incoming"
    if response_view not in {"prepare", "created", "manual"}:
        response_view = "prepare"
    if incoming_sort not in {"newest", "oldest"}:
        incoming_sort = "newest"

    values: dict[str, object] = {
        "work_tab": tab,
        "work_counts": workflow.work_counts(),
        "response_view": response_view,
        "undo_page_id": undo_page_id,
        "message": message,
        "error": error,
    }
    if tab == "incoming":
        values.update(_incoming_workspace_values(incoming_sort, incoming_page))
    elif tab == "review":
        review = workflow.manual_review_overview()
        values.update(
            review=review,
            review_count=(
                len(review["pages"])
                + len(review["match_reviews"])
                + len(review["case_tasks"])
                + sum(
                    len(group["taxpayers"])
                    for group in review["account_groups"]
                )
                + sum(
                    len(case["taxpayers"])
                    for case in review["odb_cases"]
                )
            ),
        )
    else:
        outgoing_preview = None
        if outgoing_start and response_view == "created":
            try:
                outgoing_preview = workflow.preview_outgoing_numbers(
                    outgoing_start
                )
            except WorkflowValidationError as exc:
                error = str(exc)
                values["error"] = error
        if response_view == "manual":
            overview = {
                "found_groups": workflow.manual_response_groups(),
                "generated_groups": [],
                "not_found_groups": [],
                "unnumbered_letters": [],
                "ready_abs": values["work_counts"]["ready_abs"],
                "ready_response": values["work_counts"]["ready_responses"],
                "unresolved_pages": 0,
                "unresolved_matches": 0,
            }
        else:
            overview = workflow.today_overview(view=response_view)
        active_letter = None
        active_group = None
        if response_view == "created" and focus_letter:
            for group in overview["generated_groups"]:
                for letter in group.get("letters", []):
                    if letter["id"] == focus_letter:
                        active_letter = letter
                        active_group = group
                        break
                if active_letter:
                    break
        values.update(
            overview=overview,
            outgoing_preview=outgoing_preview,
            outgoing_start=outgoing_start,
            active_letter=active_letter,
            active_group=active_group,
            abs_login_required=(
                abs_login or workflow.abs_login_required()
            ),
        )
    return templates.TemplateResponse(
        request,
        "index.html",
        context(request, **values),
    )


@app.get("/work/incoming", response_class=HTMLResponse)
def incoming_workspace_fragment(
    request: Request,
    incoming_sort: str = "newest",
    incoming_page: int = 1,
):
    return templates.TemplateResponse(
        request,
        "_work_incoming.html",
        context(
            request,
            **_incoming_workspace_values(incoming_sort, incoming_page),
        ),
    )


@app.get("/today", response_class=HTMLResponse)
def today(
    request: Request,
    message: str = "",
    error: str = "",
    outgoing_start: str = "",
):
    parameters = ["tab=responses", "response_view=created"]
    if outgoing_start:
        parameters.append(f"outgoing_start={quote(outgoing_start)}")
    if message:
        parameters.append(f"message={quote(message)}")
    if error:
        parameters.append(f"error={quote(error)}")
    return RedirectResponse(
        "/?" + "&".join(parameters),
        status_code=303,
    )


@app.get("/history", response_class=HTMLResponse)
def letter_history(
    request: Request,
    query: str = "",
    page: int = 1,
    sort_order: str = "date_desc",
    message: str = "",
    error: str = "",
):
    return templates.TemplateResponse(
        request,
        "history.html",
        context(
            request,
            history=workflow.list_letter_history(
                query,
                page=page,
                sort_order=sort_order,
            ),
            message=message,
            error=error,
        ),
    )


@app.post("/outlook-messages/{message_id}/resend-draft")
def create_outlook_resend_draft(
    message_id: str,
    query: str = Form(""),
    sort_order: str = Form("date_desc"),
    page: int = Form(1),
):
    parameters = (
        f"query={quote(query)}&sort_order={quote(sort_order)}&page={max(page, 1)}"
    )
    try:
        result = outlook_outgoing.create_resend_draft(workflow, message_id)
    except OutlookIntegrationError as exc:
        return RedirectResponse(
            f"/history?{parameters}&error={quote(str(exc))}",
            status_code=303,
        )
    notice = (
        "Повторный черновик Outlook открыт."
        if result.get("existing_outlook_draft")
        else "Повторный черновик Outlook создан и открыт."
    )
    return RedirectResponse(
        f"/history?{parameters}&message={quote(notice)}",
        status_code=303,
    )


@app.get("/registry", response_class=HTMLResponse)
def registry_lookup_page(
    request: Request,
    query: str = "",
    search_mode: str = "inn",
):
    return templates.TemplateResponse(
        request,
        "registry.html",
        context(
            request,
            lookup=None,
            lookup_query=query,
            search_mode=search_mode,
            error="",
        ),
    )


@app.post("/registry", response_class=HTMLResponse)
def registry_lookup(
    request: Request,
    query: str = Form(...),
    search_mode: str = Form("inn"),
):
    try:
        result = workflow.search_registry(query, search_mode)
        error = ""
    except RegistryLookupError as exc:
        result = None
        error = str(exc)
    return templates.TemplateResponse(
        request,
        "registry.html",
        context(
            request,
            lookup=result,
            lookup_query=query,
            search_mode=search_mode,
            error=error,
        ),
    )


@app.get("/api/registry/suggestion")
def registry_suggestion(inn: str = "", name: str = ""):
    """Return a local, non-destructive ОсОО.KG hint for one entered INN."""
    digits = "".join(character for character in inn if character.isdigit())
    if len(digits) != 14:
        return {
            "status": "invalid",
            "message": "Для сверки нужно ввести ровно 14 цифр ИНН.",
        }

    taxpayer_kind = classify_taxpayer(name, digits)
    if taxpayer_kind == TaxpayerKind.INDIVIDUAL:
        return {
            "status": "not_applicable",
            "message": "",
        }
    if taxpayer_kind == TaxpayerKind.UNKNOWN:
        message = "Вид налогоплательщика нельзя определить без догадки."
        if digits.startswith("4"):
            message = (
                "ИНН на 4 неоднозначен. Укажите наименование и форму "
                "организации, если это филиал или представительство."
            )
        return {"status": "classification_uncertain", "message": message}

    try:
        result = workflow.lookup_registry(digits)
    except RegistryLookupError as exc:
        return {"status": "error", "message": str(exc)}

    provider = result.provider or "реестр"
    if result.status == "found" and result.official_name:
        return {
            "status": "found",
            "official_name": clean_taxpayer_name(result.official_name),
            "director": result.director or "",
            "provider": provider,
            "message": f"Найдена запись в {provider}.",
        }
    if result.status == "multiple":
        return {
            "status": "multiple",
            "provider": provider,
            "message": f"{provider} вернул несколько записей; нужна проверка.",
        }
    return {
        "status": "not_found",
        "provider": provider,
        "message": f"{provider} не нашёл запись по этому ИНН.",
    }


@app.get("/api/recipient-display")
def recipient_display_suggestion(full_name: str = ""):
    return {"display_name": workflow.names.recipient_display(full_name)}


@app.get("/api/recipient-suggestions")
def recipient_suggestions(query: str = ""):
    return {"items": workflow.recipient_suggestions(query)}


@app.get("/settings", response_class=HTMLResponse)
def settings_page(request: Request, message: str = "", error: str = ""):
    scanner_settings = workflow.get_scanner_settings()
    return templates.TemplateResponse(
        request,
        "settings.html",
        context(
            request,
            inbox_dir=str(workflow.get_inbox_dir()),
            registry_priority=workflow.get_registry_priority(),
            processing_workers=workflow.get_processing_workers(),
            scanner_dpi=scanner_settings["dpi"],
            scanner_color_mode=scanner_settings["color_mode"],
            outlook_status=outlook.last_result(),
            outlook_allowed_senders=", ".join(
                sorted(outlook_importer.get_forwarding_senders())
            ),
            outlook_import_since=outlook_importer.get_import_since().isoformat(),
            outlook_mailbox=outlook_importer.get_mailbox(),
            outlook_auto_enabled=outlook_importer.get_auto_enabled(),
            outlook_auto_interval=outlook_importer.get_auto_interval_minutes(),
            outlook_auto_status=outlook_importer.get_automation_status(),
            outlook_import_stats=outlook_importer.stats(),
            outlook_subject_template=outlook_outgoing.get_subject_template(),
            message=message,
            error=error,
        ),
    )


@app.get("/settings/diagnostics/download")
def download_diagnostics():
    ocr_health = workflow.ocr.last_health() or workflow.ocr.health_check()
    bundle = create_diagnostic_bundle(
        settings.runtime_dir,
        extra={
            "abs_mode": settings.abs_mode,
            "abs_tls_verification": (
                "enabled" if settings.tolubay_verify_tls else "disabled"
            ),
            "outlook_state": outlook.last_result().state,
            "outlook_test_mode": outlook_outgoing.test_mode_enabled(),
            "outlook_test_send_enabled": outlook_outgoing.test_send_enabled(),
            "outlook_insecure_certificate_confirmation": (
                settings.outlook_allow_insecure_certificate
            ),
            "database_available": settings.database_path.is_file(),
            "ocr_engine_imported": ocr_health.get("engine_imported", False),
            "ocr_fast_files_available": ocr_health.get(
                "fast_files_available", False
            ),
            "ocr_fast_initialized": ocr_health.get(
                "fast_initialized", False
            ),
            "ocr_fast_error_type": ocr_health.get("fast_error_type", ""),
            "ocr_best_files_available": ocr_health.get(
                "best_files_available", False
            ),
            "ocr_best_initialized": ocr_health.get(
                "best_initialized", False
            ),
            "ocr_best_error_type": ocr_health.get("best_error_type", ""),
        },
    )
    return FileResponse(
        bundle,
        media_type="application/zip",
        filename=bundle.name,
    )


@app.post("/settings/outlook/check")
def check_outlook_connection():
    result = outlook.diagnose(mailbox=outlook_importer.get_mailbox())
    parameter = "message" if result.successful else "error"
    return RedirectResponse(
        f"/settings?{parameter}={quote(result.message)}#outlook",
        status_code=303,
    )


@app.post("/settings/outlook/send-receive")
def request_outlook_send_receive():
    result = outlook.request_send_receive(
        mailbox=outlook_importer.get_mailbox()
    )
    parameter = "message" if result.successful else "error"
    return RedirectResponse(
        f"/settings?{parameter}={quote(result.message)}#outlook",
        status_code=303,
    )


@app.post("/settings/outlook/config")
def update_outlook_import_settings(
    background_tasks: BackgroundTasks,
    outlook_allowed_senders: str = Form(""),
    outlook_import_since: str = Form(...),
    outlook_mailbox: str = Form(""),
    outlook_auto_enabled: bool = Form(False),
    outlook_auto_interval: int = Form(5),
    outlook_subject_template: str = Form(
        OutlookOutgoingService.DEFAULT_SUBJECT_TEMPLATE
    ),
):
    try:
        outlook_outgoing.validate_subject_template(outlook_subject_template)
        outlook_importer.update_settings(
            allowed_senders=outlook_allowed_senders,
            import_since=outlook_import_since,
            mailbox=outlook_mailbox,
            auto_enabled=outlook_auto_enabled,
            auto_interval_minutes=outlook_auto_interval,
        )
        outlook_outgoing.update_subject_template(outlook_subject_template)
    except OutlookIntegrationError as exc:
        return RedirectResponse(
            f"/settings?error={quote(str(exc))}#outlook",
            status_code=303,
        )
    started = _queue_outlook_import(background_tasks)
    message = (
        "Настройки Outlook сохранены. Получение PDF запущено."
        if started
        else "Настройки Outlook сохранены. Проверка почты уже выполняется."
    )
    return RedirectResponse(
        "/settings?message=" + quote(message) + "#outlook",
        status_code=303,
    )


@app.post("/response-letters/{letter_id}/outlook-draft")
def create_response_letter_outlook_draft(
    letter_id: str,
    group_id: str = Form(""),
):
    try:
        result = outlook_outgoing.create_draft(workflow, letter_id)
    except OutlookIntegrationError as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    message = (
        "Существующий черновик Outlook открыт."
        if result.get("existing_outlook_draft")
        else "Черновик Outlook создан и открыт для проверки."
    )
    return RedirectResponse(
        f"/today?message={quote(message)}#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/response-letters/{letter_id}/outlook-test-send")
def send_response_letter_outlook_test_message(
    letter_id: str,
    group_id: str = Form(""),
):
    try:
        result = outlook_outgoing.send_test_message(workflow, letter_id)
    except OutlookIntegrationError as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    message = (
        "Тестовое письмо уже было отправлено; повторная отправка отменена."
        if result.get("already_sent")
        else f"Тестовое письмо отправлено на {outlook_outgoing.get_test_recipient()}."
    )
    return RedirectResponse(
        f"/today?message={quote(message)}#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/settings/outlook/import")
def import_outlook_attachments(background_tasks: BackgroundTasks):
    started = _queue_outlook_import(background_tasks)
    message = (
        "Проверка Outlook запущена в фоне."
        if started
        else "Проверка Outlook уже выполняется в фоне."
    )
    return RedirectResponse(
        f"/settings?message={quote(message)}#outlook",
        status_code=303,
    )


@app.post("/work/outlook/import")
def import_outlook_from_work(
    request: Request,
    background_tasks: BackgroundTasks,
):
    started = _queue_outlook_import(background_tasks)
    status = outlook_importer.get_automation_status()
    if "application/json" in request.headers.get("accept", ""):
        return JSONResponse(
            {
                "accepted": started,
                "active": True,
                "status": status,
            },
            status_code=202 if started else 200,
        )
    message = (
        "Проверка почты запущена в фоне."
        if started
        else "Проверка почты уже выполняется в фоне."
    )
    return RedirectResponse(
        f"/?tab=incoming&message={quote(message)}",
        status_code=303,
    )


@app.get("/api/outlook/import-status")
def outlook_import_status():
    return {
        "active": _outlook_import_is_active(),
        "status": outlook_importer.get_automation_status(),
    }


@app.post("/settings")
def update_settings(
    allow_multiple_taxpayers: bool = Form(False),
    require_review_checkbox: bool = Form(False),
    period_threshold: str = Form(...),
    inbox_dir: str = Form(...),
    registry_priority: str = Form("osoo"),
    processing_workers: int = Form(2),
    scanner_dpi: int = Form(150),
    scanner_color_mode: str = Form("grayscale"),
):
    try:
        workflow.update_ui_preferences(
            allow_multiple_taxpayers=allow_multiple_taxpayers,
            show_recipient_salutation=True,
            require_review_checkbox=require_review_checkbox,
        )
        workflow.update_operational_settings(
            period_threshold=period_threshold,
            inbox_dir=inbox_dir,
            registry_priority=registry_priority,
            processing_workers=processing_workers,
            scanner_dpi=scanner_dpi,
            scanner_color_mode=scanner_color_mode,
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            "/settings?error=" + quote(str(exc)) + "#processing",
            status_code=303,
        )
    return RedirectResponse(
        "/settings?message=" + quote("Настройки сохранены.") + "#processing",
        status_code=303,
    )


@app.post("/settings/reset-processing")
def reset_processing_data(
    background_tasks: BackgroundTasks,
    action: str = Form(...),
):
    if action not in {"clear", "rescan"}:
        return RedirectResponse(
            "/settings?error=" + quote("Неизвестный вариант сброса.") + "#service",
            status_code=303,
        )
    try:
        summary = workflow.reset_processing_data()
        imported: list[str] = []
        errors: list[dict[str, str]] = []
        if action == "rescan":
            inbox_summary = workflow.import_inbox(max_files=10_000)
            imported = inbox_summary["imported"]
            errors = inbox_summary["errors"]
            for upload_id in imported:
                background_tasks.add_task(workflow.process_upload, upload_id)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            "/settings?error=" + quote(str(exc)) + "#service",
            status_code=303,
        )

    if action == "rescan":
        message = (
            f"Обработка сброшена. Заново загружено PDF: {len(imported)}; "
            f"ошибок: {len(errors)}."
        )
        destination = "/"
        fragment = ""
    else:
        message = (
            "Данные обработки очищены. Исходные PDF в папке входящих сохранены."
        )
        destination = "/settings"
        fragment = "#service"
    if summary["cleanup_errors"]:
        message += " Некоторые старые служебные файлы заняты другой программой."
    return RedirectResponse(
        f"{destination}?message={quote(message)}{fragment}",
        status_code=303,
    )


@app.post("/today/abs")
def abs_check_today(
    username: str = Form(""),
    password: str = Form(""),
):
    try:
        summary = workflow.check_abs_today(username, password)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            "/?tab=responses&response_view=prepare&error="
            + quote(str(exc)),
            status_code=303,
        )
    if summary["requires_login"]:
        return RedirectResponse(
            (
                "/?tab=responses&response_view=prepare&abs_login=1&error="
                + quote(summary["error_message"])
            ),
            status_code=303,
        )
    return RedirectResponse(
        (
            "/?tab=responses&response_view=prepare&message="
            + quote(
                "Пакетная проверка завершена. Обращений: "
                f"{summary['case_count']}."
            )
        ),
        status_code=303,
    )


@app.post("/cases/{case_id}/abs-account/taxpayer")
def abs_account_check_taxpayer(
    case_id: str,
    taxpayer_inn: str = Form(...),
    account_result: str = Form(...),
    return_to: str = Form(""),
):
    destination = "/?tab=review" if return_to == "review" else "/today"
    try:
        result = workflow.confirm_abs_account_taxpayer(
            case_id,
            taxpayer_inn,
            account_result,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"{destination}&error={quote(str(exc))}"
            if "?" in destination
            else f"{destination}?error={quote(str(exc))}",
            status_code=303,
        )
    message = (
        "Наличие счёта подтверждено — оставлено для ручного ответа."
        if result["result"] == "found"
        else "Счёта нет — обращение направлено дальше по правилу периода."
    )
    return RedirectResponse(
        f"{destination}&message={quote(message)}"
        if "?" in destination
        else f"{destination}?message={quote(message)}",
        status_code=303,
    )


@app.post("/today/responses/create-all")
def create_all_grouped_responses():
    summary = workflow.generate_all_ready_daily_responses()
    created = len(summary["created"])
    errors = summary["errors"]
    if created and not errors:
        message = quote(f"Создано ответов: {created}.")
        return RedirectResponse(
            f"/?tab=responses&response_view=created&message={message}",
            status_code=303,
        )
    if created and errors:
        message = quote(
            f"Создано ответов: {created}. Не удалось создать: {len(errors)}."
        )
        return RedirectResponse(
            f"/?tab=responses&response_view=created&message={message}",
            status_code=303,
        )
    if errors:
        message = quote("Ничего не создано: " + errors[0]["message"])
        return RedirectResponse(
            f"/?tab=responses&response_view=prepare&error={message}",
            status_code=303,
        )
    return RedirectResponse(
        "/?tab=responses&response_view=prepare&message="
        + quote("Нет готовых групп для создания ответа."),
        status_code=303,
    )


@app.post("/today/responses/{group_key}")
def create_grouped_response(
    group_key: str,
    taxpayers_per_page: int = Form(0),
):
    try:
        group_id, _ = workflow.generate_daily_response(
            group_key,
            taxpayers_per_page=taxpayers_per_page or None,
        )
    except (WorkflowValidationError, ValueError) as exc:
        return RedirectResponse(
            "/?tab=responses&response_view=prepare&error="
            + quote(str(exc)),
            status_code=303,
        )
    return RedirectResponse(
        (
            "/?tab=responses&response_view=created&message="
            + quote("Общий проект ответа Word создан.")
            + f"&focus_group={group_id}#group-{group_id}"
        ),
        status_code=303,
    )


@app.post("/today/outgoing-numbers/assign")
def assign_outgoing_numbers(
    first_number: str = Form(...),
    letter_ids: list[str] = Form(...),
):
    try:
        summary = workflow.assign_outgoing_numbers(
            first_number,
            letter_ids,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            "/?tab=responses&response_view=created&error="
            + quote(str(exc)),
            status_code=303,
        )
    message = quote(
        "Назначено исходящих номеров: "
        f"{summary['letter_count']} — с {summary['first_number']} "
        f"по {summary['last_number']}."
    )
    return RedirectResponse(
        "/?tab=responses&response_view=created&message=" + message,
        status_code=303,
    )


@app.post("/response-letters/{letter_id}/outgoing-number")
def update_outgoing_number(
    letter_id: str,
    outgoing_number: str = Form(""),
    group_id: str = Form(""),
):
    try:
        workflow.set_outgoing_number(
            letter_id,
            outgoing_number,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    return RedirectResponse(
        "/today?message="
        + quote("Исходящий номер обновлён.")
        + f"#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/response-letters/{letter_id}/scan-device")
def scan_response_letter_from_device(
    letter_id: str,
    group_id: str = Form(""),
):
    try:
        workflow.acquire_signed_response_scan(
            letter_id,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except ScannerCancelled:
        return RedirectResponse(
            f"/today?message={quote('Сканирование отменено.')}"
            f"#group-{quote(group_id)}",
            status_code=303,
        )
    except (ScannerError, WorkflowValidationError) as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    return RedirectResponse(
        "/today?message="
        + quote("Скан получен. Откройте его и подтвердите подписи и печать.")
        + f"#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/response-letters/{letter_id}/scan-session/start")
def start_response_letter_scan_session(
    letter_id: str,
    group_id: str = Form(""),
):
    actor = workflow.get_active_employee() or "Сотрудник"
    try:
        session = workflow.start_signed_scan_session(letter_id, actor=actor)
        workflow.acquire_signed_scan_session_page(session["id"], actor=actor)
    except ScannerCancelled:
        return RedirectResponse(
            f"/today?message={quote('Лист не добавлен. Сессия сохранена.')}"
            f"#group-{quote(group_id)}",
            status_code=303,
        )
    except (ScannerError, WorkflowValidationError) as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    return RedirectResponse(
        "/today?message="
        + quote("Первый лист получен. Добавьте остальные или создайте PDF.")
        + f"#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/signed-scan-sessions/{session_id}/pages/acquire")
def acquire_response_scan_session_page(
    session_id: str,
    group_id: str = Form(""),
):
    try:
        workflow.acquire_signed_scan_session_page(
            session_id,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except ScannerCancelled:
        return RedirectResponse(
            f"/today?message={quote('Добавление листа отменено.')}"
            f"#group-{quote(group_id)}",
            status_code=303,
        )
    except (ScannerError, WorkflowValidationError) as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    return RedirectResponse(
        f"/today?message={quote('Лист добавлен.')}#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/signed-scan-sessions/{session_id}/pages/{page_id}/replace")
def replace_response_scan_session_page(
    session_id: str,
    page_id: str,
    group_id: str = Form(""),
):
    try:
        workflow.acquire_signed_scan_session_page(
            session_id,
            replace_page_id=page_id,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except ScannerCancelled:
        return RedirectResponse(
            f"/today?message={quote('Пересканирование отменено; старый лист сохранён.')}"
            f"#group-{quote(group_id)}",
            status_code=303,
        )
    except (ScannerError, WorkflowValidationError) as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    return RedirectResponse(
        f"/today?message={quote('Лист заменён.')}#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/signed-scan-sessions/{session_id}/pages/{page_id}/move")
def move_response_scan_session_page(
    session_id: str,
    page_id: str,
    direction: str = Form(...),
    group_id: str = Form(""),
):
    try:
        workflow.move_signed_scan_session_page(
            session_id,
            page_id,
            direction,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    return RedirectResponse(
        f"/today#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/signed-scan-sessions/{session_id}/pages/{page_id}/remove")
def remove_response_scan_session_page(
    session_id: str,
    page_id: str,
    group_id: str = Form(""),
):
    try:
        workflow.remove_signed_scan_session_page(
            session_id,
            page_id,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    return RedirectResponse(
        f"/today?message={quote('Лист удалён из будущего PDF.')}"
        f"#group-{quote(group_id)}",
        status_code=303,
    )


@app.get("/signed-scan-sessions/{session_id}/pages/{page_id}")
def preview_response_scan_session_page(session_id: str, page_id: str):
    page = workflow.db.fetch_one(
        """
        SELECT * FROM signed_scan_session_pages
        WHERE id = ? AND session_id = ?
        """,
        (page_id, session_id),
    )
    if not page:
        raise HTTPException(404, "Лист не найден")
    try:
        path = ensure_within(
            Path(page["original_path"]),
            workflow.settings.runtime_dir / "signed_scans",
        )
    except StorageError as exc:
        raise HTTPException(403, str(exc)) from exc
    if not path.is_file():
        raise HTTPException(404, "Файл листа отсутствует")
    return FileResponse(path, media_type="image/png", filename=path.name)


@app.post("/signed-scan-sessions/{session_id}/finalize")
def finalize_response_scan_session(
    session_id: str,
    group_id: str = Form(""),
):
    try:
        workflow.finalize_signed_scan_session(
            session_id,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    return RedirectResponse(
        "/today?message="
        + quote("PDF создан. Откройте его и подтвердите подписи и печать.")
        + f"#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/signed-scan-sessions/{session_id}/cancel")
def cancel_response_scan_session(
    session_id: str,
    group_id: str = Form(""),
):
    try:
        workflow.cancel_signed_scan_session(
            session_id,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    return RedirectResponse(
        f"/today?message={quote('Сессия сканирования отменена.')}"
        f"#group-{quote(group_id)}",
        status_code=303,
    )


@app.post("/response-letters/{letter_id}/scan-upload")
def upload_response_letter_scan(
    letter_id: str,
    scan_file: UploadFile = File(...),
    group_id: str = Form(""),
):
    try:
        workflow.register_signed_response_scan(
            letter_id,
            scan_file.filename or "scan.pdf",
            scan_file.file,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}#group-{quote(group_id)}",
            status_code=303,
        )
    finally:
        scan_file.file.close()
    return RedirectResponse(
        "/today?message="
        + quote("Документ загружен. Откройте его и завершите проверку.")
        + f"#group-{quote(group_id)}",
        status_code=303,
    )


@app.get("/signed-response-scans/{scan_id}")
def download_signed_response_scan(scan_id: str):
    scan = workflow.get_signed_response_scan(scan_id)
    if not scan:
        raise HTTPException(404, "Подписанный скан не найден")
    try:
        path = ensure_within(
            Path(scan["pdf_path"]),
            workflow.settings.runtime_dir / "signed_scans",
        )
    except StorageError as exc:
        raise HTTPException(403, str(exc)) from exc
    if not path.is_file():
        raise HTTPException(404, "Файл подписанного скана отсутствует")
    return FileResponse(
        path,
        media_type="application/pdf",
    )


@app.post("/signed-response-scans/{scan_id}/confirm")
def confirm_signed_response_scan(
    scan_id: str,
    group_id: str = Form(""),
):
    scan = workflow.get_signed_response_scan(scan_id)
    letter_id = str(scan.get("response_letter_id") or "") if scan else ""
    destination = "/?tab=responses&response_view=created"
    if letter_id:
        destination += "&focus_letter=" + quote(letter_id)
    try:
        workflow.confirm_signed_response_scan(
            scan_id,
            # Нажатие единственной кнопки после просмотра встроенного PDF
            # является явным подтверждением всех трёх условий.
            correct_letter=True,
            signature_present=True,
            bank_seal_present=True,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            destination + "&error=" + quote(str(exc)) + "#letter-workflow",
            status_code=303,
        )
    return RedirectResponse(
        destination
        + "&message="
        + quote("Подписанный ответ проверен и готов к отправке.")
        + "#letter-workflow",
        status_code=303,
    )


@app.get("/response-groups/{group_id}")
def download_grouped_response(group_id: str):
    group = workflow.get_response_group(group_id)
    if not group or not group.get("response_path"):
        raise HTTPException(404, "Общий ответ не найден")
    try:
        path = ensure_within(
            Path(group["response_path"]),
            workflow.settings.responses_dir,
        )
    except StorageError as exc:
        raise HTTPException(403, str(exc)) from exc
    if not path.exists():
        raise HTTPException(404, "Файл общего ответа отсутствует")
    return FileResponse(
        path,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        filename=path.name,
    )


@app.post("/response-groups/{group_id}/open-for-print")
def open_grouped_response_for_print(group_id: str):
    group = workflow.get_response_group(group_id)
    if not group or not group.get("response_path"):
        return RedirectResponse(
            "/today?error=" + quote("Общий ответ не найден"),
            status_code=303,
        )
    try:
        path = ensure_within(
            Path(group["response_path"]),
            workflow.settings.responses_dir,
        )
        if not path.exists():
            raise OSError("missing response")
        start_foreground_watcher(
            title_parts=(path.stem,),
            class_parts=("opusapp",),
            timeout_seconds=20,
            keep_foreground_seconds=2.0,
        )
        open_word_document(path)
        workflow.mark_response_group_opened_for_print(
            group_id,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except (StorageError, OSError, WordDesktopError):
        return RedirectResponse(
            "/?tab=responses&response_view=created&error="
            + quote("Не удалось открыть ответ в Word. Скачайте файл вручную."),
            status_code=303,
        )
    return RedirectResponse(
        (
            "/?tab=responses&response_view=created&message="
            + quote("Ответ открыт в Word для проверки и печати.")
            + f"#group-{quote(group_id)}"
        ),
        status_code=303,
    )


@app.post("/employees/add")
def add_employee(employee_name: str = Form(...)):
    try:
        selected = workflow.add_employee(employee_name)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/?error={quote(str(exc))}#employee-entry",
            status_code=303,
        )
    return RedirectResponse(
        f"/?message={quote(f'Исполнитель выбран: {selected}')}#employee-entry",
        status_code=303,
    )


@app.post("/employees/select")
def select_employee(employee_key: str = Form(...)):
    try:
        selected = workflow.select_employee(employee_key)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/?error={quote(str(exc))}#employee-entry",
            status_code=303,
        )
    return RedirectResponse(
        f"/?message={quote(f'Исполнитель выбран: {selected}')}#employee-entry",
        status_code=303,
    )


@app.post("/uploads")
def create_upload(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
):
    try:
        if not workflow.get_active_employee():
            raise WorkflowValidationError(
                "Сначала выберите или добавьте исполнителя"
            )
        upload_id = workflow.create_upload(file.filename or "document.pdf", file.file)
    except (StorageError, WorkflowValidationError, ValueError) as exc:
        return RedirectResponse(
            f"/?error={quote(str(exc))}",
            status_code=303,
        )
    finally:
        file.file.close()
    background_tasks.add_task(workflow.process_upload, upload_id)
    return RedirectResponse(f"/uploads/{upload_id}", status_code=303)


@app.post("/inbox/scan")
def scan_inbox(background_tasks: BackgroundTasks):
    try:
        summary = workflow.import_inbox()
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/?error={quote(str(exc))}",
            status_code=303,
        )
    for upload_id in summary["imported"]:
        background_tasks.add_task(workflow.process_upload, upload_id)
    message = (
        f"Папка проверена. Загружено PDF: {len(summary['imported'])}; "
        f"пропущено повторов: {len(summary['skipped'])}; "
        f"ошибок: {len(summary['errors'])}."
    )
    return RedirectResponse(
        f"/?message={quote(message)}",
        status_code=303,
    )


@app.get("/uploads/{upload_id}", response_class=HTMLResponse)
def upload_detail(
    request: Request,
    upload_id: str,
    page: int | None = None,
    message: str = "",
    error: str = "",
):
    upload = workflow.get_upload(upload_id)
    if not upload:
        raise HTTPException(404, "PDF не найден")
    pages = workflow.get_upload_pages(upload_id)
    selected_index = next(
        (
            index
            for index, item in enumerate(pages)
            if item["status"] in {"needs_review", "technical_error"}
        ),
        0,
    )
    if page is not None:
        selected_index = next(
            (
                index
                for index, item in enumerate(pages)
                if int(item["page_number"]) == page
            ),
            -1,
        )
        if selected_index < 0:
            raise HTTPException(404, "Страница не найдена в этом PDF")
    selected_page = pages[selected_index] if pages else None
    return templates.TemplateResponse(
        request,
        "upload_detail.html",
        context(
            request,
            upload=upload,
            pages=pages,
            selected_page=selected_page,
            previous_page=(
                pages[selected_index - 1] if selected_index > 0 else None
            ),
            next_page=(
                pages[selected_index + 1]
                if pages and selected_index + 1 < len(pages)
                else None
            ),
            message=message,
            error=error,
        ),
    )


@app.post("/uploads/{upload_id}/reprocess")
def reprocess_upload(
    background_tasks: BackgroundTasks,
    upload_id: str,
    return_page: int | None = Form(default=None),
):
    page_query = ""
    if return_page is not None and any(
        int(item["page_number"]) == return_page
        for item in workflow.get_upload_pages(upload_id)
    ):
        page_query = f"page={return_page}&"
    try:
        count = workflow.request_problem_reprocess(upload_id)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            (
                f"/uploads/{upload_id}?{page_query}error={quote(str(exc))}"
                f"{'#page-viewer' if page_query else ''}"
            ),
            status_code=303,
        )
    background_tasks.add_task(
        workflow.process_upload,
        upload_id,
        True,
    )
    return RedirectResponse(
        (
            f"/uploads/{upload_id}?{page_query}message="
            f"{quote(f'Повторно обрабатывается страниц: {count}')}"
            f"{'#page-viewer' if page_query else ''}"
        ),
        status_code=303,
    )


@app.get("/pages/{page_id}/image")
def page_image(page_id: str, variant: str = "original"):
    page = workflow.get_page(page_id)
    if not page:
        raise HTTPException(404, "Страница не найдена")
    key = "enhanced_preview_path" if variant == "enhanced" else "preview_path"
    path_text = page.get(key)
    if not path_text:
        raise HTTPException(404, "Превью ещё не готово")
    try:
        path = ensure_within(Path(path_text), settings.previews_dir)
    except StorageError as exc:
        raise HTTPException(403, str(exc)) from exc
    if not path.exists():
        raise HTTPException(404, "Файл превью отсутствует")
    return FileResponse(path, media_type="image/jpeg")


@app.get("/pages/{page_id}/pdf")
def page_pdf(page_id: str):
    page = workflow.get_page(page_id)
    if not page:
        raise HTTPException(404, "Страница не найдена")
    try:
        path = workflow.get_page_pdf_path(page_id)
    except WorkflowValidationError as exc:
        raise HTTPException(422, str(exc)) from exc
    return FileResponse(
        path,
        media_type="application/pdf",
        filename=f"page-{int(page['page_number']):04d}.pdf",
        content_disposition_type="inline",
        headers={"X-Content-Type-Options": "nosniff"},
    )


@app.post("/pages/{page_id}/reprocess")
def reprocess_page(
    background_tasks: BackgroundTasks,
    page_id: str,
):
    page_record = workflow.get_page(page_id)
    try:
        upload_id = workflow.request_page_reprocess(page_id)
    except WorkflowValidationError as exc:
        if page_record:
            return RedirectResponse(
                (
                    f"/uploads/{page_record['upload_id']}?"
                    f"page={int(page_record['page_number'])}&"
                    f"error={quote(str(exc))}#page-viewer"
                ),
                status_code=303,
            )
        return RedirectResponse(
            f"/review/{page_id}?error={quote(str(exc))}",
            status_code=303,
        )
    background_tasks.add_task(
        workflow.process_upload,
        upload_id,
        True,
    )
    page_query = (
        f"page={int(page_record['page_number'])}&"
        if page_record
        else ""
    )
    return RedirectResponse(
        (
            f"/uploads/{upload_id}?{page_query}message="
            f"{quote('Страница поставлена на повторную обработку.')}"
            "#page-viewer"
        ),
        status_code=303,
    )


@app.post("/pages/{page_id}/precise-ocr")
def precise_ocr_page(
    background_tasks: BackgroundTasks,
    page_id: str,
):
    try:
        upload_id = workflow.request_precise_ocr(page_id)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/review/{page_id}?error={quote(str(exc))}",
            status_code=303,
        )
    background_tasks.add_task(workflow.process_precise_ocr, page_id)
    return RedirectResponse(
        (
            f"/uploads/{upload_id}?message="
            f"{quote('Запущен точный OCR best. Страница останется в ручной проверке.')}"
        ),
        status_code=303,
    )


@app.get("/review", response_class=HTMLResponse)
def review_queue(request: Request, message: str = "", error: str = ""):
    parameters = ["tab=review"]
    if message:
        parameters.append(f"message={quote(message)}")
    if error:
        parameters.append(f"error={quote(error)}")
    return RedirectResponse(
        "/?" + "&".join(parameters),
        status_code=303,
    )


@app.get("/review/matches/{review_id}", response_class=HTMLResponse)
def review_case_match(
    request: Request,
    review_id: str,
    message: str = "",
    error: str = "",
):
    workflow.reconcile_case_match_reviews()
    review = workflow.get_case_match_review(review_id)
    if not review:
        raise HTTPException(404, "Проверка совпадения не найдена")
    if review["status"] != "pending":
        return RedirectResponse(
            "/?tab=review&message=" + quote("Эта проверка уже завершена."),
            status_code=303,
        )
    reviews = workflow.list_case_match_reviews()
    review_index = next(
        (
            index
            for index, item in enumerate(reviews)
            if item["id"] == review_id
        ),
        0,
    )
    return templates.TemplateResponse(
        request,
        "match_review.html",
        context(
            request,
            review=review,
            review_position=review_index + 1,
            review_total=max(len(reviews), 1),
            previous_review_id=(
                reviews[review_index - 1]["id"] if review_index > 0 else None
            ),
            next_review_id=(
                reviews[review_index + 1]["id"]
                if review_index + 1 < len(reviews)
                else None
            ),
            message=message,
            error=error,
        ),
    )


@app.post("/review/matches/{review_id}")
def confirm_case_match(
    review_id: str,
    decision: str = Form(...),
    name_choice: str = Form(""),
    inn_choice: str = Form(""),
    manual_name: str = Form(""),
    manual_inn: str = Form(""),
):
    try:
        result = workflow.resolve_case_match_review(
            review_id,
            decision=decision,
            name_choice=name_choice,
            inn_choice=inn_choice,
            manual_name=manual_name,
            manual_inn=manual_inn,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/review/matches/{review_id}?error={quote(str(exc))}",
            status_code=303,
        )
    remaining_matches = workflow.list_case_match_reviews()
    if remaining_matches:
        message = quote(
            "Решение сохранено. Осталось сравнений: "
            f"{len(remaining_matches)}."
        )
        return RedirectResponse(
            f"/review/matches/{remaining_matches[0]['id']}?message={message}",
            status_code=303,
        )
    remaining_pages = workflow.list_review_pages()
    if remaining_pages:
        return RedirectResponse(
            f"/review/{remaining_pages[0]['id']}?message="
            + quote("Решение сохранено. Перейдите к проверке страницы."),
            status_code=303,
        )
    message = (
        "Совпадение подтверждено. Обращения продолжили обработку."
        if result["decision"] == "same"
        else "Обращения подтверждены как разные и продолжили обработку."
    )
    return RedirectResponse(
        f"/?tab=review&message={quote(message)}", status_code=303
    )


@app.get("/review/{page_id}", response_class=HTMLResponse)
def review_page(request: Request, page_id: str, error: str = ""):
    page = workflow.get_page(page_id)
    if not page:
        raise HTTPException(404, "Страница не найдена")
    if page["status"] not in {"needs_review", "technical_error"}:
        if page.get("case_id"):
            return RedirectResponse(
                f"/cases/{page['case_id']}",
                status_code=303,
            )
        return RedirectResponse(
            f"/uploads/{page['upload_id']}",
            status_code=303,
        )
    case = workflow.get_case(page["case_id"]) if page.get("case_id") else None
    taxpayers = (
        workflow.get_taxpayers(case["id"]) if case else []
    )
    preferences = workflow.get_ui_preferences()
    taxpayer_conflict_candidates: list[dict[str, str]] = []
    if not preferences["allow_multiple_taxpayers"] and len(taxpayers) > 1:
        ocr_unconfirmed = all(
            not item.get("manually_confirmed")
            and item.get("name_source") == "ocr_scan"
            and item.get("inn_source") == "ocr_scan"
            for item in taxpayers
        )
        names = {
            " ".join((item.get("name") or "").split())
            for item in taxpayers
            if (item.get("name") or "").strip()
        }
        inns = list(dict.fromkeys(
            item.get("inn") or "" for item in taxpayers if item.get("inn")
        ))
        if ocr_unconfirmed and len(names) <= 1 and len(inns) > 1:
            taxpayer_conflict_candidates = [
                {"inn": inn, "name": next(iter(names), "")}
                for inn in inns
            ]
            taxpayers = [{"name": next(iter(names), ""), "inn": ""}]
    if not taxpayers:
        taxpayers = [{"name": "", "inn": ""}]
    upload = workflow.get_upload(page["upload_id"])
    review_pages = workflow.list_review_pages()
    review_index = next(
        (
            index
            for index, review_item in enumerate(review_pages)
            if review_item["id"] == page_id
        ),
        None,
    )
    previous_review_page_id = None
    next_review_page_id = None
    if review_index is not None:
        if review_index > 0:
            previous_review_page_id = review_pages[review_index - 1]["id"]
        if review_index + 1 < len(review_pages):
            next_review_page_id = review_pages[review_index + 1]["id"]
    return templates.TemplateResponse(
        request,
        "review_page.html",
        context(
            request,
            page=page,
            case=case or {},
            taxpayers=taxpayers,
            taxpayer_conflict_candidates=taxpayer_conflict_candidates,
            upload=upload or {},
            gns_offices=workflow.list_gns_offices(),
            period_status=period_review_status(case or {}),
            review_position=(review_index + 1) if review_index is not None else 1,
            review_total=max(len(review_pages), 1),
            previous_review_page_id=previous_review_page_id,
            next_review_page_id=next_review_page_id,
            error=error,
        ),
    )


@app.post("/review/{page_id}")
def confirm_review(
    page_id: str,
    page_type: str = Form(...),
    district_place: str = Form(""),
    recipient_position: str = Form(""),
    recipient_full_name: str = Form(""),
    recipient_display_name: str = Form(""),
    period_start: str = Form(""),
    period_end: str = Form(""),
    period_route: str = Form(""),
    employee_name: str = Form(""),
    critical_fields_verified: bool = Form(False),
    taxpayer_name: list[str] = Form(default=[]),
    taxpayer_inn: list[str] = Form(default=[]),
):
    preferences = workflow.get_ui_preferences()
    taxpayers = [
        {"name": name, "inn": inn}
        for name, inn in zip(taxpayer_name, taxpayer_inn, strict=False)
        if name.strip() or inn.strip()
    ]
    try:
        case_id = workflow.confirm_page(
            page_id,
            page_type=page_type,
            district_place=district_place,
            recipient_position=recipient_position,
            recipient_full_name=recipient_full_name,
            recipient_display_name=recipient_display_name,
            period_start=period_start,
            period_end=period_end,
            period_route=period_route,
            employee_name=employee_name,
            taxpayers=taxpayers,
            critical_fields_verified=(
                critical_fields_verified
                or not preferences["require_review_checkbox"]
            ),
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/review/{page_id}?error={quote(str(exc))}",
            status_code=303,
        )
    remaining = [
        page for page in workflow.list_review_pages() if page["id"] != page_id
    ]
    if remaining:
        next_page_id = remaining[0]["id"]
        message = quote(
            f"Страница подтверждена. Осталось проверить: {len(remaining)}."
        )
        return RedirectResponse(
            f"/review/{next_page_id}?message={message}",
            status_code=303,
        )
    match_reviews = workflow.list_case_match_reviews()
    if match_reviews:
        message = quote(
            "Страница подтверждена. Теперь проверьте возможное "
            "совпадение обращений."
        )
        return RedirectResponse(
            f"/review/matches/{match_reviews[0]['id']}?message={message}",
            status_code=303,
        )
    if case_id:
        return RedirectResponse(
            f"/cases/{case_id}?message=" + quote(
                "Страница подтверждена. Проблемных страниц больше нет."
            ),
            status_code=303,
        )
    return RedirectResponse(
        "/?tab=review&message=" + quote("Проблемных страниц больше нет."),
        status_code=303,
    )


@app.post("/review/{page_id}/type")
def mark_review_page_type(
    page_id: str,
    page_type: str = Form(...),
):
    try:
        removed = workflow.mark_page_type_from_queue(page_id, page_type)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            "/?tab=review&error=" + quote(str(exc)),
            status_code=303,
        )
    if removed:
        return RedirectResponse(
            "/?tab=review&message="
            + quote("Тип страницы подтверждён.")
            + "&undo_page_id="
            + quote(page_id),
            status_code=303,
        )
    return RedirectResponse(f"/review/{page_id}", status_code=303)


@app.post("/pages/{page_id}/reopen-type")
def reopen_page_type_review(page_id: str):
    try:
        workflow.reopen_page_type_review(
            page_id,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            "/?tab=review&error=" + quote(str(exc)),
            status_code=303,
        )
    return RedirectResponse(f"/review/{page_id}", status_code=303)


@app.get("/cases", response_class=HTMLResponse)
def cases_list(request: Request):
    return templates.TemplateResponse(
        request,
        "cases.html",
        context(request, cases=workflow.list_cases()),
    )


@app.post("/cases/{case_id}/reopen-review")
def reopen_incomplete_case_review(case_id: str):
    try:
        page = workflow.reopen_incomplete_case_review(
            case_id,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/cases/{case_id}?error={quote(str(exc))}",
            status_code=303,
        )
    return RedirectResponse(f"/review/{page['id']}", status_code=303)


@app.get("/cases/{case_id}", response_class=HTMLResponse)
def case_detail(
    request: Request,
    case_id: str,
    message: str = "",
    error: str = "",
    abs_login: bool = False,
):
    case = workflow.get_case(case_id)
    if not case:
        raise HTTPException(404, "Обращение не найдено")
    taxpayers = workflow.get_taxpayers(case_id)
    source_page = workflow.get_case_source_page(case_id)
    source_upload = workflow.get_upload(case["upload_id"])
    return templates.TemplateResponse(
        request,
        "case_detail.html",
        context(
            request,
            case=case,
            taxpayers=taxpayers,
            source_page=source_page,
            source_upload=source_upload,
            odb_pending_taxpayers=[
                item for item in taxpayers if not item.get("odb_result")
            ],
            odb_pending=(
                case["status"] == "manual_period_rule"
                and any(not item.get("odb_result") for item in taxpayers)
            ),
            odb_has_found=any(
                item.get("odb_result") == "found"
                or item.get("abs_result") == "found"
                for item in taxpayers
            ),
            registry_needs_confirmation=(
                case["source_kind"] != "qr_official"
                and any(
                    item.get("registry_status")
                    and item.get("registry_status")
                    not in {"match", "not_applicable"}
                    for item in taxpayers
                )
            ),
            recipient_position_display=workflow.names.position_display(
                case.get("recipient_position") or ""
            ),
            abs_login_required=(
                abs_login
                or (
                    case.get("status") == "ready_for_abs"
                    and case.get("abs_status")
                    in {
                        AbsStatus.AUTH_ERROR,
                        AbsStatus.UNAVAILABLE,
                        AbsStatus.TECHNICAL_ERROR,
                    }
                    and not workflow.abs_session_active()
                )
            ),
            message=message,
            error=error,
        ),
    )


@app.post("/cases/{case_id}/taxpayers/{taxpayer_id}/correct")
def correct_taxpayer(
    case_id: str,
    taxpayer_id: str,
    taxpayer_name: str = Form(...),
    taxpayer_inn: str = Form(...),
    return_to: str = Form(""),
):
    try:
        result = workflow.correct_taxpayer(
            case_id,
            taxpayer_id,
            name=taxpayer_name,
            inn=taxpayer_inn,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        destination = (
            "/today" if return_to == "today" else f"/cases/{case_id}"
        )
        return RedirectResponse(
            f"{destination}?error={quote(str(exc))}",
            status_code=303,
        )

    if result["requires_abs_recheck"]:
        message = quote(
            "ИНН исправлен. Результат прежней проверки сброшен — проверьте обращение в АБС повторно."
        )
        return RedirectResponse(
            f"/cases/{case_id}?message={message}",
            status_code=303,
        )
    message = quote(
        "Наименование исправлено. Группа ответов пересобрана автоматически."
    )
    destination = (
        "/today" if return_to == "today" else f"/cases/{case_id}"
    )
    return RedirectResponse(
        f"{destination}?message={message}",
        status_code=303,
    )


@app.post("/cases/{case_id}/abs")
def abs_check(
    case_id: str,
    username: str = Form(""),
    password: str = Form(""),
):
    try:
        result = workflow.check_abs(case_id, username, password)
        message = result.message
        if result.is_fake:
            message += " Используется тестовая АБС."
        if result.status in {
            AbsStatus.AUTH_ERROR,
            AbsStatus.UNAVAILABLE,
            AbsStatus.TECHNICAL_ERROR,
        }:
            return RedirectResponse(
                f"/cases/{case_id}?abs_login=1&error={quote(message)}",
                status_code=303,
            )
        return RedirectResponse(
            f"/cases/{case_id}?message={quote(message)}",
            status_code=303,
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/cases/{case_id}?error={quote(str(exc))}",
            status_code=303,
        )


@app.post("/cases/{case_id}/registry/accept")
def accept_registry_variance(case_id: str):
    try:
        workflow.accept_registry_variance(case_id)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/cases/{case_id}?error={quote(str(exc))}",
            status_code=303,
        )
    return RedirectResponse(
        (
            f"/cases/{case_id}?message="
            f"{quote('Данные письма оставлены без подмены. Можно проверить АБС.')}"
        ),
        status_code=303,
    )


@app.post("/cases/{case_id}/odb")
def odb_check(
    case_id: str,
    taxpayer_inn: list[str] = Form(default=[]),
    odb_result: list[str] = Form(default=[]),
    return_to: str = Form(default=""),
):
    results = [
        {"inn": inn, "result": result}
        for inn, result in zip(
            taxpayer_inn,
            odb_result,
            strict=False,
        )
    ]
    try:
        next_status = workflow.confirm_odb(case_id, results)
    except WorkflowValidationError as exc:
        destination = "/today" if return_to == "ready" else f"/cases/{case_id}"
        return RedirectResponse(
            f"{destination}?error={quote(str(exc))}",
            status_code=303,
        )
    message = (
        "ОДБ проверена. Можно готовить ответ."
        if next_status == "ready_for_response"
        else "ОДБ проверена. Найденные записи оставлены для ручной обработки."
    )
    destination = "/today" if return_to == "ready" else f"/cases/{case_id}"
    return RedirectResponse(
        f"{destination}?message={quote(message)}",
        status_code=303,
    )


@app.post("/cases/{case_id}/odb/taxpayer")
def odb_check_taxpayer(
    case_id: str,
    taxpayer_inn: str = Form(...),
    odb_result: str = Form(...),
):
    try:
        result = workflow.confirm_odb_taxpayer(
            case_id,
            taxpayer_inn,
            odb_result,
            actor=workflow.get_active_employee() or "Сотрудник",
        )
    except WorkflowValidationError as exc:
        return JSONResponse(
            {"ok": False, "error": str(exc)},
            status_code=422,
        )
    message = (
        "Найден в ОДБ — передан в ручной ответ."
        if result["result"] == "found"
        else "Не найден в ОДБ — проверка сохранена."
    )
    return {"ok": True, **result, "message": message}


@app.post("/cases/{case_id}/response")
def create_response(case_id: str):
    return RedirectResponse(
        (
            "/today?message="
            + quote(
                "Ответы теперь формируются общими письмами по адресату."
            )
        ),
        status_code=303,
    )


@app.get("/cases/{case_id}/response")
def download_response(case_id: str):
    case = workflow.get_case(case_id)
    if not case or not case.get("response_path"):
        raise HTTPException(404, "Ответ ещё не создан")
    try:
        path = ensure_within(Path(case["response_path"]), settings.responses_dir)
    except StorageError as exc:
        raise HTTPException(403, str(exc)) from exc
    return FileResponse(
        path,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        filename=path.name,
    )


@app.get("/journal", response_class=HTMLResponse)
def journal(request: Request):
    return templates.TemplateResponse(
        request,
        "journal.html",
        context(request, events=workflow.get_audit()),
    )


@app.get("/health")
def health():
    return {"status": "ok", "version": APP_VERSION}


def run() -> None:
    uvicorn.run(
        "gns_app.main:app",
        host="127.0.0.1",
        port=8765,
        reload=False,
        access_log=False,
        # Оконная portable-сборка не имеет stdout/stderr. Стандартная
        # конфигурация Uvicorn пытается создать консольный formatter и падает
        # ещё до запуска сервера.
        log_config=None,
    )


if __name__ == "__main__":
    run()
