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
    yield


app = FastAPI(
    title="Обработка писем ГНС",
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
    "manual_period_rule": "Ручная проверка периода",
    "ready_for_response": "Можно создать ответ",
    "response_created": "Ответ создан",
    "completed": "Завершено",
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

EVENT_LABELS = {
    "upload_registered": "PDF зарегистрирован",
    "page_processed": "Страница обработана",
    "page_processing_error": "Ошибка обработки страницы",
    "scan_case_created": "Создано обращение по скану",
    "official_document_processed": "Официальная версия обработана",
    "page_manually_confirmed": "Страница подтверждена сотрудником",
    "fake_abs_checked": "Выполнена тестовая проверка АБС",
    "response_created": "Создан ответ Word",
}

ENTITY_LABELS = {
    "upload": "PDF",
    "page": "Страница",
    "case": "Обращение",
}


def taxpayer_word(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "запись"
    if count % 10 in {2, 3, 4} and count % 100 not in {12, 13, 14}:
        return "записи"
    return "записей"


def context(request: Request, **values):
    return {
        "request": request,
        "status_labels": STATUS_LABELS,
        "page_type_labels": PAGE_TYPE_LABELS,
        "source_labels": SOURCE_LABELS,
        "abs_status_labels": ABS_STATUS_LABELS,
        "event_labels": EVENT_LABELS,
        "entity_labels": ENTITY_LABELS,
        "taxpayer_word": taxpayer_word,
        "threshold": settings.period_threshold.isoformat(),
        **values,
    }


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse(
        request,
        "index.html",
        context(
            request,
            stats=workflow.dashboard_stats(),
            uploads=workflow.list_uploads()[:12],
            cases=workflow.list_cases()[:8],
        ),
    )


@app.post("/uploads")
def create_upload(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
):
    try:
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


@app.get("/uploads/{upload_id}", response_class=HTMLResponse)
def upload_detail(request: Request, upload_id: str):
    upload = workflow.get_upload(upload_id)
    if not upload:
        raise HTTPException(404, "PDF не найден")
    pages = workflow.get_upload_pages(upload_id)
    return templates.TemplateResponse(
        request,
        "upload_detail.html",
        context(request, upload=upload, pages=pages),
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
    case = workflow.get_case(page["case_id"]) if page.get("case_id") else None
    taxpayers = (
        workflow.get_taxpayers(case["id"]) if case else []
    )
    if not taxpayers:
        taxpayers = [{"name": "", "inn": ""}]
    return templates.TemplateResponse(
        request,
        "review_page.html",
        context(
            request,
            page=page,
            case=case or {},
            taxpayers=taxpayers,
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
    return templates.TemplateResponse(
        request,
        "case_detail.html",
        context(
            request,
            case=case,
            taxpayers=workflow.get_taxpayers(case_id),
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


@app.post("/cases/{case_id}/response")
def create_response(case_id: str):
    try:
        workflow.generate_response(case_id)
        return RedirectResponse(
            f"/cases/{case_id}?message={quote('Проект ответа создан.')}",
            status_code=303,
        )
    except (WorkflowValidationError, ValueError) as exc:
        return RedirectResponse(
            f"/cases/{case_id}?error={quote(str(exc))}",
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
