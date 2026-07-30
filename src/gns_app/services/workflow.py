from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path
from typing import BinaryIO, Any
from uuid import uuid4

from docx import Document
from pypdf import PdfReader

from gns_app.config import Settings
from gns_app.database import Database, utc_now
from gns_app.domain import (
    AbsStatus,
    CaseStatus,
    PageStatus,
    PageType,
    QrStatus,
    UploadStatus,
    ValueSource,
)
from gns_app.services.abs_service import FakeAbsGateway
from gns_app.services.classifier import PageClassifier
from gns_app.services.extractor import FieldExtractor
from gns_app.services.name_service import NameService
from gns_app.services.ocr_service import OcrService
from gns_app.services.official_document import (
    OfficialDocumentClient,
    OfficialDocumentError,
)
from gns_app.services.pdf_service import PdfService
from gns_app.services.qr_service import QrService
from gns_app.services.storage import sanitize_filename, save_pdf_stream
from gns_app.services.word_service import WordTemplateService


class WorkflowValidationError(ValueError):
    pass


class WorkflowService:
    def __init__(self, db: Database, settings: Settings):
        self.db = db
        self.settings = settings
        self.pdf = PdfService()
        self.qr = QrService(
            settings.allowed_qr_hosts,
            settings.allowed_qr_paths,
        )
        self.ocr = OcrService(
            self.pdf,
            settings.ocr_fast_data_dir,
            settings.ocr_best_data_dir,
        )
        self.classifier = PageClassifier()
        self.extractor = FieldExtractor()
        self.names = NameService()
        self.official = OfficialDocumentClient(
            settings.allowed_qr_hosts,
            settings.allowed_qr_paths,
        )
        self.abs = FakeAbsGateway()
        self.word = WordTemplateService(settings.source_templates_dir)

    def create_upload(
        self,
        original_filename: str,
        stream: BinaryIO,
    ) -> str:
        upload_id = uuid4().hex
        safe_name = sanitize_filename(original_filename)
        if Path(safe_name).suffix.casefold() != ".pdf":
            raise WorkflowValidationError("Разрешены только PDF-файлы")

        upload_dir = self.settings.uploads_dir / upload_id
        stored_path = upload_dir / "source.pdf"
        digest, _ = save_pdf_stream(
            stream,
            stored_path,
            self.settings.max_upload_bytes,
        )
        try:
            page_count = self.pdf.page_count(stored_path)
        except Exception:
            stored_path.unlink(missing_ok=True)
            upload_dir.rmdir()
            raise

        created_at = utc_now()
        self.db.execute(
            """
            INSERT INTO uploads(
                id, original_filename, stored_path, sha256,
                page_count, status, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                upload_id,
                safe_name,
                str(stored_path),
                digest,
                page_count,
                UploadStatus.REGISTERED,
                created_at,
            ),
        )
        page_rows = [
            (
                uuid4().hex,
                upload_id,
                page_number,
                PageType.UNKNOWN,
                PageStatus.REGISTERED,
                created_at,
                created_at,
            )
            for page_number in range(1, page_count + 1)
        ]
        self.db.executemany(
            """
            INSERT INTO pages(
                id, upload_id, page_number, page_type,
                status, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            page_rows,
        )
        self.db.audit(
            "upload",
            upload_id,
            "upload_registered",
            {
                "filename": safe_name,
                "page_count": page_count,
                "sha256": digest,
            },
        )
        return upload_id

    def process_upload(
        self, upload_id: str, only_registered_pages: bool = False
    ) -> None:
        upload = self.get_upload(upload_id)
        if not upload:
            return
        self.db.execute(
            "UPDATE uploads SET status = ? WHERE id = ?",
            (UploadStatus.PROCESSING, upload_id),
        )
        if only_registered_pages:
            pages = self.db.fetch_all(
                """
                SELECT * FROM pages
                WHERE upload_id = ? AND status = ? AND manual_confirmed = 0
                ORDER BY page_number
                """,
                (upload_id, PageStatus.REGISTERED),
            )
        else:
            pages = self.db.fetch_all(
                "SELECT * FROM pages WHERE upload_id = ? ORDER BY page_number",
                (upload_id,),
            )
        for page in pages:
            try:
                self._process_page(upload, page)
            except Exception as exc:
                self.db.execute(
                    """
                    UPDATE pages
                    SET status = ?, issue_code = ?, issue_message = ?,
                        updated_at = ?
                    WHERE id = ?
                    """,
                    (
                        PageStatus.TECHNICAL_ERROR,
                        "page_processing_error",
                        str(exc)[:800],
                        utc_now(),
                        page["id"],
                    ),
                )
                self.db.audit(
                    "page",
                    page["id"],
                    "page_processing_error",
                    {"message": str(exc)[:800]},
                )
        self._complete_qr_companions(upload_id)
        self._remove_orphan_cases(upload_id)
        self._refresh_upload_status(upload_id)

    def _complete_qr_companions(self, upload_id: str) -> None:
        self.db.execute(
            """
            UPDATE pages
            SET status = ?, issue_code = ?, issue_message = ?, updated_at = ?
            WHERE upload_id = ? AND page_type = ? AND qr_status = ?
              AND qr_payload_hash IS NOT NULL
              AND EXISTS (
                  SELECT 1
                  FROM pages AS letter
                  WHERE letter.upload_id = pages.upload_id
                    AND letter.qr_payload_hash = pages.qr_payload_hash
                    AND letter.page_type = ?
              )
            """,
            (
                PageStatus.COMPLETED,
                "qr_group_accounted",
                (
                    "Страница учтена по тому же проверенному QR, "
                    "что и распознанное письмо."
                ),
                utc_now(),
                upload_id,
                PageType.UNKNOWN,
                QrStatus.FOUND,
                PageType.LETTER,
            ),
        )

    def _remove_orphan_cases(self, upload_id: str) -> None:
        row = self.db.fetch_one(
            """
            SELECT COUNT(*) AS count
            FROM cases
            WHERE upload_id = ? AND official_document_path IS NULL
              AND NOT EXISTS (
                  SELECT 1 FROM pages WHERE pages.case_id = cases.id
              )
              AND NOT EXISTS (
                  SELECT 1 FROM taxpayers WHERE taxpayers.case_id = cases.id
              )
            """,
            (upload_id,),
        )
        count = int(row["count"] if row else 0)
        if not count:
            return
        self.db.execute(
            """
            DELETE FROM cases
            WHERE upload_id = ? AND official_document_path IS NULL
              AND NOT EXISTS (
                  SELECT 1 FROM pages WHERE pages.case_id = cases.id
              )
              AND NOT EXISTS (
                  SELECT 1 FROM taxpayers WHERE taxpayers.case_id = cases.id
              )
            """,
            (upload_id,),
        )
        self.db.audit(
            "upload",
            upload_id,
            "orphan_cases_removed",
            {"case_count": count},
        )

    def request_problem_reprocess(self, upload_id: str) -> int:
        if not self.get_upload(upload_id):
            raise WorkflowValidationError("PDF не найден")
        row = self.db.fetch_one(
            """
            SELECT COUNT(*) AS count
            FROM pages
            WHERE upload_id = ? AND manual_confirmed = 0
              AND status IN (?, ?)
            """,
            (
                upload_id,
                PageStatus.NEEDS_REVIEW,
                PageStatus.TECHNICAL_ERROR,
            ),
        )
        count = int(row["count"] if row else 0)
        if count == 0:
            raise WorkflowValidationError(
                "Нет неподтверждённых проблемных страниц для повторной обработки"
            )
        self.db.execute(
            """
            UPDATE pages
            SET status = ?, issue_code = NULL, issue_message = NULL,
                updated_at = ?
            WHERE upload_id = ? AND manual_confirmed = 0
              AND status IN (?, ?)
            """,
            (
                PageStatus.REGISTERED,
                utc_now(),
                upload_id,
                PageStatus.NEEDS_REVIEW,
                PageStatus.TECHNICAL_ERROR,
            ),
        )
        self.db.execute(
            "UPDATE uploads SET status = ?, completed_at = NULL WHERE id = ?",
            (UploadStatus.PROCESSING, upload_id),
        )
        self.db.audit(
            "upload",
            upload_id,
            "upload_reprocess_requested",
            {"page_count": count},
        )
        return count

    def _process_page(
        self, upload: dict[str, Any], page: dict[str, Any]
    ) -> None:
        page_id = page["id"]
        page_number = page["page_number"]
        self.db.execute(
            "UPDATE pages SET status = ?, updated_at = ? WHERE id = ?",
            (PageStatus.PROCESSING, utc_now(), page_id),
        )

        preview_dir = self.settings.previews_dir / upload["id"]
        preview_path = preview_dir / f"page-{page_number:04d}.jpg"
        enhanced_path = preview_dir / f"page-{page_number:04d}-enhanced.jpg"
        rendered = self.pdf.render_page(
            Path(upload["stored_path"]),
            page_number,
            preview_path,
            enhanced_path,
        )

        qr_result = self.qr.decode(preview_path)
        ocr_result = self.ocr.recognize(
            Path(upload["stored_path"]),
            page_number,
            rendered.preview_path,
        )
        classification = self.classifier.classify(
            ocr_result.text,
            rendered.quality_score,
        )

        case_id: str | None = None
        official_complete = False
        official_issue: str | None = None
        if qr_result.status == QrStatus.FOUND:
            case_id, needs_official = self._ensure_qr_case(
                upload["id"], qr_result.payload_hash or ""
            )
            if (
                self.settings.auto_download_official
                and needs_official
                and qr_result.payload
            ):
                try:
                    official_path = self.official.download(
                        qr_result.payload,
                        self.settings.official_dir / case_id,
                    )
                    official_complete = self._apply_official_document(
                        case_id, official_path
                    )
                    if not official_complete:
                        official_issue = (
                            "Официальная версия получена по QR. "
                            "Подтвердите обязательные поля и исполнителя."
                        )
                except (OfficialDocumentError, OSError, ValueError) as exc:
                    official_issue = str(exc)
            elif self.settings.auto_download_official and not needs_official:
                existing_case = self.get_case(case_id)
                if existing_case and existing_case.get(
                    "official_document_path"
                ):
                    official_complete = bool(
                        existing_case.get("fields_confirmed")
                    )
                    if not official_complete:
                        official_issue = (
                            "Официальная версия уже получена по QR. "
                            "Подтвердите обязательные поля и исполнителя."
                        )

        extracted = None
        if classification.page_type == PageType.LETTER:
            extracted = self.extractor.extract_scan_letter(ocr_result.text)
            if case_id is None:
                case_id = self._create_scan_case(upload["id"], page_id)
            self._prefill_scan_case(case_id, extracted)

        page_status, issue_code, issue_message = self._page_outcome(
            classification.page_type,
            classification.confidence,
            qr_result.status,
            official_complete,
            official_issue,
        )
        self.db.execute(
            """
            UPDATE pages
            SET case_id = ?, preview_path = ?, enhanced_preview_path = ?,
                page_type = ?, type_confidence = ?, quality_score = ?,
                qr_status = ?, qr_payload_hash = ?, qr_safe_url = ?,
                qr_method = ?, ocr_status = ?, ocr_confidence = ?,
                extracted_text = ?, status = ?, issue_code = ?,
                issue_message = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                case_id,
                str(rendered.preview_path),
                str(rendered.enhanced_preview_path),
                classification.page_type,
                classification.confidence,
                rendered.quality_score,
                qr_result.status,
                qr_result.payload_hash,
                qr_result.safe_url,
                qr_result.method,
                ocr_result.status,
                ocr_result.confidence,
                ocr_result.text,
                page_status,
                issue_code,
                issue_message,
                utc_now(),
                page_id,
            ),
        )
        self.db.audit(
            "page",
            page_id,
            "page_processed",
            {
                "page_number": page_number,
                "page_type": classification.page_type,
                "type_confidence": classification.confidence,
                "quality_score": rendered.quality_score,
                "qr_status": qr_result.status,
                "qr_method": qr_result.method,
                "ocr_status": ocr_result.status,
                "ocr_confidence": ocr_result.confidence,
                "status": page_status,
            },
        )

    def _page_outcome(
        self,
        page_type: PageType,
        type_confidence: float,
        qr_status: QrStatus,
        official_complete: bool,
        official_issue: str | None,
    ) -> tuple[PageStatus, str | None, str | None]:
        if official_complete:
            return PageStatus.COMPLETED, None, None
        if (
            page_type == PageType.DECISION
            and type_confidence >= 0.65
            and qr_status != QrStatus.INVALID_URL
        ):
            return (
                PageStatus.COMPLETED,
                "optional_decision",
                "Решение учтено как необязательное приложение.",
            )
        if official_issue:
            return (
                PageStatus.NEEDS_REVIEW,
                "official_confirmation_required",
                official_issue,
            )
        if page_type == PageType.UNKNOWN:
            return (
                PageStatus.NEEDS_REVIEW,
                "page_type_unknown",
                "Не удалось уверенно определить тип страницы.",
            )
        if qr_status == QrStatus.FOUND:
            return (
                PageStatus.NEEDS_REVIEW,
                "official_pending",
                "QR прочитан, но обязательные данные ещё не подтверждены.",
            )
        if page_type == PageType.LETTER:
            return (
                PageStatus.NEEDS_REVIEW,
                "scan_confirmation_required",
                "Письмо распознано по скану и требует подтверждения полей.",
            )
        return (
            PageStatus.NEEDS_REVIEW,
            "manual_review_required",
            "Страница требует ручной проверки.",
        )

    def _ensure_qr_case(
        self, upload_id: str, payload_hash: str
    ) -> tuple[str, bool]:
        existing = self.db.fetch_one(
            """
            SELECT * FROM cases
            WHERE upload_id = ? AND qr_payload_hash = ?
            """,
            (upload_id, payload_hash),
        )
        if existing:
            return existing["id"], not bool(
                existing.get("official_document_path")
            )

        case_id = uuid4().hex
        now = utc_now()
        self.db.execute(
            """
            INSERT INTO cases(
                id, upload_id, status, source_kind, qr_payload_hash,
                employee_name, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                case_id,
                upload_id,
                CaseStatus.COLLECTING,
                ValueSource.QR_LINK,
                payload_hash,
                self.settings.employee_name or None,
                now,
                now,
            ),
        )
        return case_id, True

    def _create_scan_case(self, upload_id: str, page_id: str) -> str:
        case_id = uuid4().hex
        now = utc_now()
        self.db.execute(
            """
            INSERT INTO cases(
                id, upload_id, status, source_kind,
                employee_name, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                case_id,
                upload_id,
                CaseStatus.NEEDS_REVIEW,
                ValueSource.OCR_SCAN,
                self.settings.employee_name or None,
                now,
                now,
            ),
        )
        self.db.audit(
            "case",
            case_id,
            "scan_case_created",
            {"source_page_id": page_id},
        )
        return case_id

    def _prefill_scan_case(self, case_id: str, fields) -> None:
        current = self.get_case(case_id)
        if not current:
            return
        self.db.execute(
            """
            UPDATE cases
            SET period_start = COALESCE(period_start, ?),
                period_end = COALESCE(period_end, ?),
                status = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                fields.period_start,
                fields.period_end,
                CaseStatus.NEEDS_REVIEW,
                utc_now(),
                case_id,
            ),
        )
        if fields.taxpayers and not self.get_taxpayers(case_id):
            rows = []
            now = utc_now()
            for index, taxpayer in enumerate(fields.taxpayers, 1):
                if not taxpayer.name or not taxpayer.inn:
                    continue
                rows.append(
                    (
                        uuid4().hex,
                        case_id,
                        index,
                        taxpayer.name,
                        taxpayer.inn,
                        ValueSource.OCR_SCAN,
                        ValueSource.OCR_SCAN,
                        0,
                        now,
                        now,
                    )
                )
            if rows:
                self.db.executemany(
                    """
                    INSERT INTO taxpayers(
                        id, case_id, display_order, name, inn,
                        name_source, inn_source, manually_confirmed,
                        created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    rows,
                )

    def _apply_official_document(
        self, case_id: str, official_path: Path
    ) -> bool:
        text = self._official_text(official_path)
        fields = self.extractor.extract_official_letter(text)
        recipient_display = (
            self.names.recipient_display(fields.recipient_full_name or "")
            if fields.recipient_full_name
            else None
        )
        complete = bool(
            fields.confidence >= 0.95
            and fields.district_place
            and fields.recipient_position
            and fields.recipient_full_name
            and recipient_display
            and fields.period_start
            and fields.period_end
            and fields.taxpayers
        )
        employee = self.settings.employee_name or None
        ready = complete and bool(employee)
        self.db.execute(
            """
            UPDATE cases
            SET official_document_path = ?, source_kind = ?,
                district_place = ?,
                recipient_position = ?, recipient_full_name = ?,
                recipient_display_name = ?, period_start = ?,
                period_end = ?, employee_name = ?,
                fields_confirmed = ?, status = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                str(official_path),
                ValueSource.QR_OFFICIAL,
                fields.district_place,
                fields.recipient_position,
                fields.recipient_full_name,
                recipient_display,
                fields.period_start,
                fields.period_end,
                employee,
                1 if ready else 0,
                CaseStatus.READY_FOR_ABS if ready else CaseStatus.NEEDS_REVIEW,
                utc_now(),
                case_id,
            ),
        )
        self.db.execute("DELETE FROM taxpayers WHERE case_id = ?", (case_id,))
        rows = []
        now = utc_now()
        for index, taxpayer in enumerate(fields.taxpayers, 1):
            if taxpayer.name and taxpayer.inn:
                rows.append(
                    (
                        uuid4().hex,
                        case_id,
                        index,
                        taxpayer.name,
                        taxpayer.inn,
                        ValueSource.QR_OFFICIAL,
                        ValueSource.QR_OFFICIAL,
                        1,
                        now,
                        now,
                    )
                )
        if rows:
            self.db.executemany(
                """
                INSERT INTO taxpayers(
                    id, case_id, display_order, name, inn,
                    name_source, inn_source, manually_confirmed,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                rows,
            )
        self.db.audit(
            "case",
            case_id,
            "official_document_processed",
            {
                "extension": official_path.suffix.lower(),
                "fields_complete": complete,
                "employee_profile_present": bool(employee),
                "issues": fields.issues,
            },
        )
        return ready

    @staticmethod
    def _official_text(path: Path) -> str:
        if path.suffix.casefold() == ".pdf":
            reader = PdfReader(str(path))
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        if path.suffix.casefold() == ".docx":
            document = Document(str(path))
            parts = [paragraph.text for paragraph in document.paragraphs]
            for table in document.tables:
                for row in table.rows:
                    parts.extend(cell.text for cell in row.cells)
            return "\n".join(parts)
        raise WorkflowValidationError(
            "Неподдерживаемый формат официального документа"
        )

    def confirm_page(
        self,
        page_id: str,
        *,
        page_type: str,
        district_place: str = "",
        recipient_position: str = "",
        recipient_full_name: str = "",
        recipient_display_name: str = "",
        period_start: str = "",
        period_end: str = "",
        employee_name: str = "",
        taxpayers: list[dict[str, str]] | None = None,
        actor: str = "Сотрудник",
    ) -> str | None:
        page = self.get_page(page_id)
        if not page:
            raise WorkflowValidationError("Страница не найдена")
        try:
            selected_type = PageType(page_type)
        except ValueError as exc:
            raise WorkflowValidationError("Неизвестный тип страницы") from exc
        if selected_type == PageType.UNKNOWN:
            raise WorkflowValidationError("Нужно обозначить тип страницы")

        case_id = page.get("case_id")
        if selected_type == PageType.LETTER:
            clean_taxpayers = self._validate_manual_fields(
                district_place,
                recipient_position,
                recipient_full_name,
                period_start,
                period_end,
                employee_name,
                taxpayers or [],
            )
            if not case_id:
                case_id = self._create_scan_case(
                    page["upload_id"], page_id
                )
            display_name = (
                recipient_display_name.strip()
                or self.names.recipient_display(recipient_full_name)
            )
            if not display_name:
                raise WorkflowValidationError(
                    "Не удалось сформировать обращение к адресату"
                )

            self.db.execute(
                """
                UPDATE cases
                SET district_place = ?, recipient_position = ?,
                    recipient_full_name = ?, recipient_display_name = ?,
                    period_start = ?, period_end = ?, employee_name = ?,
                    source_kind = ?, fields_confirmed = 1,
                    status = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    district_place.strip(),
                    recipient_position.strip(),
                    recipient_full_name.strip(),
                    display_name,
                    period_start,
                    period_end,
                    employee_name.strip(),
                    ValueSource.MANUAL,
                    CaseStatus.READY_FOR_ABS,
                    utc_now(),
                    case_id,
                ),
            )
            self._replace_taxpayers(case_id, clean_taxpayers)

        self.db.execute(
            """
            UPDATE pages
            SET case_id = ?, page_type = ?, type_confidence = 1,
                manual_confirmed = 1, status = ?,
                issue_code = NULL, issue_message = NULL, updated_at = ?
            WHERE id = ?
            """,
            (
                case_id,
                selected_type,
                PageStatus.MANUALLY_CONFIRMED,
                utc_now(),
                page_id,
            ),
        )
        self.db.audit(
            "page",
            page_id,
            "page_manually_confirmed",
            {
                "page_type": selected_type,
                "case_id": case_id,
                "taxpayer_count": len(taxpayers or []),
            },
            actor=actor,
        )
        self._refresh_upload_status(page["upload_id"])
        return case_id

    @staticmethod
    def _validate_manual_fields(
        district_place: str,
        recipient_position: str,
        recipient_full_name: str,
        period_start: str,
        period_end: str,
        employee_name: str,
        taxpayers: list[dict[str, str]],
    ) -> list[dict[str, str]]:
        fields = {
            "Район и место": district_place,
            "Должность адресата": recipient_position,
            "ФИО адресата": recipient_full_name,
            "Начало периода": period_start,
            "Конец периода": period_end,
            "Исполнитель банка": employee_name,
        }
        missing = [label for label, value in fields.items() if not value.strip()]
        if missing:
            raise WorkflowValidationError(
                "Не заполнены поля: " + ", ".join(missing)
            )
        try:
            start = date.fromisoformat(period_start)
            end = date.fromisoformat(period_end)
        except ValueError as exc:
            raise WorkflowValidationError("Период содержит неверную дату") from exc
        if start > end:
            raise WorkflowValidationError(
                "Начало периода не может быть позже окончания"
            )
        if not taxpayers:
            raise WorkflowValidationError(
                "Нужно указать хотя бы одного налогоплательщика"
            )

        clean: list[dict[str, str]] = []
        for index, taxpayer in enumerate(taxpayers, 1):
            name = taxpayer.get("name", "").strip()
            inn = re.sub(r"\D", "", taxpayer.get("inn", ""))
            if not name or len(inn) != 14:
                raise WorkflowValidationError(
                    f"Налогоплательщик {index}: нужно наименование и 14 цифр ИНН"
                )
            clean.append({"name": name, "inn": inn})
        return clean

    def _replace_taxpayers(
        self, case_id: str, taxpayers: list[dict[str, str]]
    ) -> None:
        self.db.execute("DELETE FROM taxpayers WHERE case_id = ?", (case_id,))
        now = utc_now()
        self.db.executemany(
            """
            INSERT INTO taxpayers(
                id, case_id, display_order, name, inn,
                name_source, inn_source, manually_confirmed,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    uuid4().hex,
                    case_id,
                    index,
                    taxpayer["name"],
                    taxpayer["inn"],
                    ValueSource.MANUAL,
                    ValueSource.MANUAL,
                    1,
                    now,
                    now,
                )
                for index, taxpayer in enumerate(taxpayers, 1)
            ],
        )

    def check_abs(
        self,
        case_id: str,
        username: str,
        password: str,
    ):
        case = self.get_case(case_id)
        if not case:
            raise WorkflowValidationError("Обращение не найдено")
        if not case["fields_confirmed"]:
            raise WorkflowValidationError(
                "Перед АБС нужно подтвердить все поля письма"
            )
        taxpayers = self.get_taxpayers(case_id)
        if not taxpayers:
            raise WorkflowValidationError("Нет налогоплательщиков")

        self.db.execute(
            "UPDATE cases SET status = ?, abs_status = ?, updated_at = ? WHERE id = ?",
            (
                CaseStatus.ABS_CHECKING,
                AbsStatus.CHECKING,
                utc_now(),
                case_id,
            ),
        )
        result = self.abs.check(username, password, taxpayers)
        # Пароль и логин не сохраняются и не передаются в журнал.
        username = ""
        password = ""

        for taxpayer_result in result.taxpayers:
            self.db.execute(
                """
                UPDATE taxpayers SET abs_result = ?, updated_at = ?
                WHERE case_id = ? AND inn = ?
                """,
                (
                    taxpayer_result["result"],
                    utc_now(),
                    case_id,
                    taxpayer_result["inn"],
                ),
            )

        next_status = CaseStatus.NEEDS_REVIEW
        if result.status == AbsStatus.NOT_FOUND:
            start = date.fromisoformat(case["period_start"])
            next_status = (
                CaseStatus.MANUAL_PERIOD_RULE
                if start < self.settings.period_threshold
                else CaseStatus.READY_FOR_RESPONSE
            )

        self.db.execute(
            """
            UPDATE cases
            SET status = ?, abs_status = ?, updated_at = ?
            WHERE id = ?
            """,
            (next_status, result.status, utc_now(), case_id),
        )
        self.db.audit(
            "case",
            case_id,
            "fake_abs_checked",
            {
                "status": result.status,
                "is_fake": True,
                "taxpayers": result.taxpayers,
                "next_status": next_status,
            },
        )
        return result

    def generate_response(self, case_id: str) -> Path:
        case = self.get_case(case_id)
        if not case:
            raise WorkflowValidationError("Обращение не найдено")
        if case["status"] != CaseStatus.READY_FOR_RESPONSE:
            raise WorkflowValidationError(
                "Обращение ещё не готово к созданию ответа"
            )
        taxpayers = self.get_taxpayers(case_id)
        output = self.settings.responses_dir / f"response-{case_id}.docx"
        self.word.render(output, case, taxpayers)
        self.db.execute(
            """
            UPDATE cases
            SET status = ?, response_status = ?, response_path = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                CaseStatus.RESPONSE_CREATED,
                "created",
                str(output),
                utc_now(),
                case_id,
            ),
        )
        self.db.audit(
            "case",
            case_id,
            "response_created",
            {
                "filename": output.name,
                "taxpayer_count": len(taxpayers),
            },
        )
        return output

    def _refresh_upload_status(self, upload_id: str) -> None:
        page_counts = self.db.fetch_one(
            """
            SELECT
                COUNT(*) AS total,
                SUM(CASE WHEN status IN ('needs_review', 'technical_error')
                    THEN 1 ELSE 0 END) AS problem,
                SUM(CASE WHEN status IN ('registered', 'preview_ready', 'processing')
                    THEN 1 ELSE 0 END) AS pending
            FROM pages WHERE upload_id = ?
            """,
            (upload_id,),
        ) or {"total": 0, "problem": 0, "pending": 0}
        incomplete_cases = self.db.fetch_one(
            """
            SELECT COUNT(*) AS count
            FROM cases
            WHERE upload_id = ? AND status IN ('collecting', 'needs_review')
            """,
            (upload_id,),
        ) or {"count": 0}

        if page_counts["pending"]:
            status = UploadStatus.PROCESSING
        elif page_counts["problem"] or incomplete_cases["count"]:
            status = UploadStatus.NEEDS_REVIEW
        else:
            status = UploadStatus.READY
        self.db.execute(
            "UPDATE uploads SET status = ? WHERE id = ?",
            (status, upload_id),
        )

    def get_upload(self, upload_id: str) -> dict[str, Any] | None:
        return self.db.fetch_one("SELECT * FROM uploads WHERE id = ?", (upload_id,))

    def list_uploads(self) -> list[dict[str, Any]]:
        return self.db.fetch_all(
            "SELECT * FROM uploads ORDER BY created_at DESC"
        )

    def get_upload_pages(self, upload_id: str) -> list[dict[str, Any]]:
        return self.db.fetch_all(
            "SELECT * FROM pages WHERE upload_id = ? ORDER BY page_number",
            (upload_id,),
        )

    def get_page(self, page_id: str) -> dict[str, Any] | None:
        return self.db.fetch_one("SELECT * FROM pages WHERE id = ?", (page_id,))

    def list_review_pages(self) -> list[dict[str, Any]]:
        return self.db.fetch_all(
            """
            SELECT pages.*, uploads.original_filename
            FROM pages
            JOIN uploads ON uploads.id = pages.upload_id
            WHERE pages.status IN ('needs_review', 'technical_error')
            ORDER BY uploads.created_at, pages.page_number
            """
        )

    def get_case(self, case_id: str) -> dict[str, Any] | None:
        return self.db.fetch_one("SELECT * FROM cases WHERE id = ?", (case_id,))

    def get_taxpayers(self, case_id: str) -> list[dict[str, Any]]:
        return self.db.fetch_all(
            """
            SELECT * FROM taxpayers
            WHERE case_id = ? ORDER BY display_order
            """,
            (case_id,),
        )

    def list_cases(self) -> list[dict[str, Any]]:
        return self.db.fetch_all(
            """
            SELECT cases.*, uploads.original_filename,
                (SELECT COUNT(*) FROM taxpayers
                 WHERE taxpayers.case_id = cases.id) AS taxpayer_count
            FROM cases
            JOIN uploads ON uploads.id = cases.upload_id
            ORDER BY cases.created_at DESC
            """
        )

    def get_audit(self, limit: int = 200) -> list[dict[str, Any]]:
        rows = self.db.fetch_all(
            """
            SELECT * FROM audit_events
            ORDER BY id DESC LIMIT ?
            """,
            (limit,),
        )
        for row in rows:
            try:
                row["payload"] = json.loads(row["payload_json"])
            except json.JSONDecodeError:
                row["payload"] = {}
        return rows

    def dashboard_stats(self) -> dict[str, int]:
        return {
            "uploads": (
                self.db.fetch_one("SELECT COUNT(*) AS n FROM uploads") or {"n": 0}
            )["n"],
            "pages": (
                self.db.fetch_one("SELECT COUNT(*) AS n FROM pages") or {"n": 0}
            )["n"],
            "review": (
                self.db.fetch_one(
                    """
                    SELECT COUNT(*) AS n FROM pages
                    WHERE status IN ('needs_review', 'technical_error')
                    """
                )
                or {"n": 0}
            )["n"],
            "ready_abs": (
                self.db.fetch_one(
                    "SELECT COUNT(*) AS n FROM cases WHERE status = 'ready_for_abs'"
                )
                or {"n": 0}
            )["n"],
            "ready_response": (
                self.db.fetch_one(
                    """
                    SELECT COUNT(*) AS n FROM cases
                    WHERE status = 'ready_for_response'
                    """
                )
                or {"n": 0}
            )["n"],
        }
