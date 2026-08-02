from __future__ import annotations

from contextlib import asynccontextmanager
from pathlib import Path
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
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from gns_app.config import settings
from gns_app.database import Database
from gns_app.services.storage import StorageError, ensure_within
from gns_app.services.registry_service import RegistryLookupError
from gns_app.services.workflow import (
    WorkflowService,
    WorkflowValidationError,
)


PACKAGE_DIR = Path(__file__).resolve().parent
db = Database(settings.database_path)
workflow = WorkflowService(db, settings)


@asynccontextmanager
async def lifespan(_: FastAPI):
    settings.ensure_directories()
    db.initialize()
    workflow.initialize_employee_profiles()
    workflow.reconcile_official_qr_pages()
    workflow.reconcile_confident_scan_decisions()
    workflow.repair_cleaned_responses()
    yield


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
}

EVENT_LABELS = {
    "upload_registered": "PDF зарегистрирован",
    "upload_reprocess_requested": "Запрошена повторная обработка",
    "page_reprocess_requested": "Запрошена повторная обработка страницы",
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
    "fake_abs_checked": "Выполнена тестовая проверка АБС",
    "fake_abs_batch_checked": "Выполнена пакетная проверка АБС",
    "odb_checked": "Записана ручная проверка ОДБ",
    "registry_checked": "Выполнена сверка с ОсОО.KG",
    "registry_variance_accepted": "Подтверждены расхождения ОсОО.KG",
    "inbox_scanned": "Просканирована папка входящих",
    "gns_offices_replaced": "Обновлён справочник налоговых органов",
    "gns_office_suggested": "Предложен налоговый орган по OCR",
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
    "district_place_edge_noise_removed": (
        "Удалены лишние краевые символы в реквизите ГНС"
    ),
    "response_regenerated_after_cleanup": (
        "Ответ пересоздан после очистки реквизита ГНС"
    ),
    "employee_profile_added": "Добавлен исполнитель",
    "active_employee_selected": "Выбран активный исполнитель",
}

ENTITY_LABELS = {
    "upload": "PDF",
    "page": "Страница",
    "case": "Обращение",
    "settings": "Настройка",
    "response_group": "Общий ответ",
    "taxpayer": "Налогоплательщик",
}


def taxpayer_word(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "запись"
    if count % 10 in {2, 3, 4} and count % 100 not in {12, 13, 14}:
        return "записи"
    return "записей"


def case_word(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "обращение"
    if count % 10 in {2, 3, 4} and count % 100 not in {12, 13, 14}:
        return "обращения"
    return "обращений"


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
        "threshold": settings.period_threshold.isoformat(),
        "active_employee": active_employee,
        "employee_profiles": workflow.list_employee_profiles(),
        **values,
    }


@app.get("/", response_class=HTMLResponse)
def index(request: Request, message: str = "", error: str = ""):
    return templates.TemplateResponse(
        request,
        "index.html",
        context(
            request,
            stats=workflow.dashboard_stats(),
            uploads=workflow.list_uploads()[:12],
            cases=workflow.list_cases()[:8],
            inbox_dir=str(settings.inbox_dir),
            message=message,
            error=error,
        ),
    )


@app.get("/today", response_class=HTMLResponse)
def today(request: Request, message: str = "", error: str = ""):
    return templates.TemplateResponse(
        request,
        "today.html",
        context(
            request,
            overview=workflow.today_overview(),
            message=message,
            error=error,
        ),
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


@app.post("/today/abs")
def abs_check_today(
    username: str = Form(...),
    password: str = Form(...),
):
    try:
        summary = workflow.check_abs_today(username, password)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}",
            status_code=303,
        )
    return RedirectResponse(
        (
            "/today?message="
            + quote(
                "Пакетная проверка завершена. Обращений: "
                f"{summary['case_count']}."
            )
        ),
        status_code=303,
    )


@app.post("/today/responses/{group_key}")
def create_grouped_response(group_key: str):
    try:
        group_id, _ = workflow.generate_daily_response(group_key)
    except (WorkflowValidationError, ValueError) as exc:
        return RedirectResponse(
            f"/today?error={quote(str(exc))}",
            status_code=303,
        )
    return RedirectResponse(
        (
            "/today?message="
            + quote("Общий проект ответа Word создан.")
            + f"#group-{group_id}"
        ),
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
            settings.responses_dir,
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
    message: str = "",
    error: str = "",
):
    upload = workflow.get_upload(upload_id)
    if not upload:
        raise HTTPException(404, "PDF не найден")
    pages = workflow.get_upload_pages(upload_id)
    return templates.TemplateResponse(
        request,
        "upload_detail.html",
        context(
            request,
            upload=upload,
            pages=pages,
            message=message,
            error=error,
        ),
    )


@app.post("/uploads/{upload_id}/reprocess")
def reprocess_upload(
    background_tasks: BackgroundTasks,
    upload_id: str,
):
    try:
        count = workflow.request_problem_reprocess(upload_id)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/uploads/{upload_id}?error={quote(str(exc))}",
            status_code=303,
        )
    background_tasks.add_task(
        workflow.process_upload,
        upload_id,
        True,
    )
    return RedirectResponse(
        (
            f"/uploads/{upload_id}?message="
            f"{quote(f'Повторно обрабатывается страниц: {count}')}"
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
    try:
        upload_id = workflow.request_page_reprocess(page_id)
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/review/{page_id}?error={quote(str(exc))}",
            status_code=303,
        )
    background_tasks.add_task(
        workflow.process_upload,
        upload_id,
        True,
    )
    return RedirectResponse(
        (
            f"/uploads/{upload_id}?message="
            f"{quote('Страница поставлена на повторную обработку.')}"
        ),
        status_code=303,
    )


@app.get("/review", response_class=HTMLResponse)
def review_queue(request: Request):
    return templates.TemplateResponse(
        request,
        "review_queue.html",
        context(request, pages=workflow.list_review_pages()),
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
    if not taxpayers:
        taxpayers = [{"name": "", "inn": ""}]
    upload = workflow.get_upload(page["upload_id"])
    return templates.TemplateResponse(
        request,
        "review_page.html",
        context(
            request,
            page=page,
            case=case or {},
            taxpayers=taxpayers,
            upload=upload or {},
            gns_offices=workflow.list_gns_offices(),
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
    employee_name: str = Form(""),
    critical_fields_verified: bool = Form(False),
    taxpayer_name: list[str] = Form(default=[]),
    taxpayer_inn: list[str] = Form(default=[]),
):
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
            employee_name=employee_name,
            taxpayers=taxpayers,
            critical_fields_verified=critical_fields_verified,
        )
    except WorkflowValidationError as exc:
        return RedirectResponse(
            f"/review/{page_id}?error={quote(str(exc))}",
            status_code=303,
        )
    if case_id:
        return RedirectResponse(f"/cases/{case_id}", status_code=303)
    return RedirectResponse("/review", status_code=303)


@app.get("/cases", response_class=HTMLResponse)
def cases_list(request: Request):
    return templates.TemplateResponse(
        request,
        "cases.html",
        context(request, cases=workflow.list_cases()),
    )


@app.get("/cases/{case_id}", response_class=HTMLResponse)
def case_detail(request: Request, case_id: str, message: str = "", error: str = ""):
    case = workflow.get_case(case_id)
    if not case:
        raise HTTPException(404, "Обращение не найдено")
    taxpayers = workflow.get_taxpayers(case_id)
    return templates.TemplateResponse(
        request,
        "case_detail.html",
        context(
            request,
            case=case,
            taxpayers=taxpayers,
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
                    and item.get("registry_status") != "match"
                    for item in taxpayers
                )
            ),
            recipient_position_display=workflow.names.position_display(
                case.get("recipient_position") or ""
            ),
            message=message,
            error=error,
        ),
    )


@app.post("/cases/{case_id}/abs")
def abs_check(
    case_id: str,
    username: str = Form(...),
    password: str = Form(...),
):
    try:
        result = workflow.check_abs(case_id, username, password)
        message = result.message + " Используется тестовая АБС."
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
        return RedirectResponse(
            f"/cases/{case_id}?error={quote(str(exc))}",
            status_code=303,
        )
    message = (
        "ОДБ проверена. Можно готовить ответ."
        if next_status == "ready_for_response"
        else "ОДБ проверена. Найденные записи оставлены для ручной обработки."
    )
    return RedirectResponse(
        f"/cases/{case_id}?message={quote(message)}",
        status_code=303,
    )


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
    return {"status": "ok", "version": "0.1.0"}


def run() -> None:
    uvicorn.run(
        "gns_app.main:app",
        host="127.0.0.1",
        port=8765,
        reload=False,
        access_log=False,
    )


if __name__ == "__main__":
    run()
