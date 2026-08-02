from __future__ import annotations

import hashlib
import json
import re
from datetime import UTC, date, datetime, time, timedelta, timezone
from pathlib import Path
from typing import Any, BinaryIO
from uuid import uuid4

from docx import Document
from pypdf import PdfReader

from gns_app.config import Settings
from gns_app.database import Database, utc_now
from gns_app.domain import (
    AbsStatus,
    CaseStatus,
    OdbStatus,
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
from gns_app.services.pdf_service import PdfProcessingError, PdfService
from gns_app.services.qr_service import QrService
from gns_app.services.registry_service import (
    OsooRegistryClient,
    RegistryLookupError,
    RegistryLookupResult,
)
from gns_app.services.storage import (
    ensure_within,
    sanitize_filename,
    save_pdf_stream,
)
from gns_app.services.word_service import WordTemplateService
from gns_app.text_cleanup import clean_location


class WorkflowValidationError(ValueError):
    pass


class WorkflowService:
    ACTIVE_EMPLOYEE_SETTING = "active_employee_key"
    BUSINESS_TIMEZONE = timezone(timedelta(hours=6), "Asia/Bishkek")

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
        self.registry = OsooRegistryClient()
        self.word = WordTemplateService(
            settings.source_templates_dir,
            self.names,
        )

    def lookup_registry(self, inn: str) -> RegistryLookupResult:
        return self.registry.lookup_by_inn(inn)

    def search_registry(
        self,
        query: str,
        search_mode: str,
    ) -> RegistryLookupResult:
        if search_mode == "name":
            return self.registry.search_by_name(query)
        return self.registry.lookup_by_inn(query)

    def initialize_employee_profiles(self) -> None:
        configured = self.settings.employee_name.strip()
        if not configured:
            return
        name = self._normalize_employee_name(configured)
        name_key = name.casefold()
        now = utc_now()
        self.db.execute(
            """
            INSERT INTO employee_profiles(
                name, name_key, created_at, updated_at
            ) VALUES (?, ?, ?, ?)
            ON CONFLICT(name_key) DO NOTHING
            """,
            (name, name_key, now, now),
        )
        if not self._active_employee_from_database():
            self._store_active_employee(name_key)
            self._apply_employee_to_open_cases(name)

    @staticmethod
    def _normalize_office_text(value: str) -> str:
        normalized = value.casefold().replace("ё", "е")
        normalized = re.sub(r"[^\w\s-]", " ", normalized)
        return " ".join(normalized.split())

    def replace_gns_offices(
        self,
        records: list[dict[str, Any]],
        actor: str = "Сотрудник",
    ) -> int:
        prepared: list[tuple[str, str, str, str, str, str, str]] = []
        seen: set[str] = set()
        now = utc_now()
        for record in records:
            office_name = " ".join(
                str(record.get("office_name") or "").split()
            )
            district_place = clean_location(
                str(record.get("district_place") or "")
            )
            postal_address = " ".join(
                str(record.get("postal_address") or "").split()
            )
            raw_aliases = record.get("aliases") or []
            if isinstance(raw_aliases, str):
                raw_aliases = [raw_aliases]
            aliases = [
                " ".join(str(alias).split())
                for alias in raw_aliases
                if " ".join(str(alias).split())
            ]
            if not office_name or not district_place:
                raise WorkflowValidationError(
                    "У каждого налогового органа нужны название и район/место"
                )
            office_key = self._normalize_office_text(
                f"{office_name} {district_place}"
            )
            if office_key in seen:
                raise WorkflowValidationError(
                    "В справочнике повторяется один налоговый орган"
                )
            seen.add(office_key)
            prepared.append(
                (
                    office_key,
                    office_name,
                    district_place,
                    postal_address,
                    json.dumps(aliases, ensure_ascii=False),
                    now,
                    now,
                )
            )
        with self.db.connect() as connection:
            connection.execute("DELETE FROM gns_offices")
            connection.executemany(
                """
                INSERT INTO gns_offices(
                    office_key, office_name, district_place,
                    postal_address, aliases_json, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                prepared,
            )
        self.db.audit(
            "settings",
            "gns_offices",
            "gns_offices_replaced",
            {"office_count": len(prepared)},
            actor=actor,
        )
        return len(prepared)

    def list_gns_offices(self) -> list[dict[str, Any]]:
        rows = self.db.fetch_all(
            """
            SELECT * FROM gns_offices
            WHERE active = 1
            ORDER BY district_place, office_name
            """
        )
        for row in rows:
            try:
                row["aliases"] = json.loads(
                    row.get("aliases_json") or "[]"
                )
            except json.JSONDecodeError:
                row["aliases"] = []
        return rows

    def match_gns_office(self, text: str) -> dict[str, Any] | None:
        normalized_text = self._normalize_office_text(text)
        if not normalized_text:
            return None
        matches: list[dict[str, Any]] = []
        for office in self.list_gns_offices():
            raw_terms = [
                office.get("office_name") or "",
                office.get("district_place") or "",
                office.get("postal_address") or "",
                *(office.get("aliases") or []),
            ]
            normalized_terms = {
                normalized
                for term in raw_terms
                if len(normalized := self._normalize_office_text(term)) >= 8
            }
            if any(term in normalized_text for term in normalized_terms):
                matches.append(office)
        return matches[0] if len(matches) == 1 else None

    def list_employee_profiles(self) -> list[dict[str, Any]]:
        return self.db.fetch_all(
            """
            SELECT id, name, name_key, created_at, updated_at
            FROM employee_profiles
            ORDER BY name COLLATE NOCASE
            """
        )

    def get_active_employee(self) -> str | None:
        active = self._active_employee_from_database()
        if active:
            return active
        configured = self.settings.employee_name.strip()
        return configured or None

    def add_employee(self, employee_name: str) -> str:
        name = self._normalize_employee_name(employee_name)
        name_key = name.casefold()
        existing = self.db.fetch_one(
            "SELECT name FROM employee_profiles WHERE name_key = ?",
            (name_key,),
        )
        added = existing is None
        if added:
            now = utc_now()
            self.db.execute(
                """
                INSERT INTO employee_profiles(
                    name, name_key, created_at, updated_at
                ) VALUES (?, ?, ?, ?)
                """,
                (name, name_key, now, now),
            )
        else:
            name = existing["name"]

        self._store_active_employee(name_key)
        self._apply_employee_to_open_cases(name)
        if added:
            self.db.audit(
                "settings",
                "employee_profiles",
                "employee_profile_added",
                {"employee_name": name},
            )
        self.db.audit(
            "settings",
            "active_employee",
            "active_employee_selected",
            {"employee_name": name},
        )
        return name

    def select_employee(self, employee_key: str) -> str:
        row = self.db.fetch_one(
            """
            SELECT name, name_key
            FROM employee_profiles
            WHERE name_key = ?
            """,
            (employee_key.strip().casefold(),),
        )
        if not row:
            raise WorkflowValidationError(
                "Выбранный исполнитель не найден в локальном справочнике"
            )
        self._store_active_employee(row["name_key"])
        self._apply_employee_to_open_cases(row["name"])
        self.db.audit(
            "settings",
            "active_employee",
            "active_employee_selected",
            {"employee_name": row["name"]},
        )
        return row["name"]

    def _active_employee_from_database(self) -> str | None:
        row = self.db.fetch_one(
            """
            SELECT employee_profiles.name
            FROM settings
            JOIN employee_profiles
              ON employee_profiles.name_key = settings.value
            WHERE settings.key = ?
            """,
            (self.ACTIVE_EMPLOYEE_SETTING,),
        )
        return row["name"] if row else None

    def _store_active_employee(self, name_key: str) -> None:
        self.db.execute(
            """
            INSERT INTO settings(key, value, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE
            SET value = excluded.value, updated_at = excluded.updated_at
            """,
            (self.ACTIVE_EMPLOYEE_SETTING, name_key, utc_now()),
        )

    def _apply_employee_to_open_cases(self, employee_name: str) -> None:
        self.db.execute(
            """
            UPDATE cases
            SET employee_name = ?, updated_at = ?
            WHERE response_path IS NULL
              AND status NOT IN (?, ?)
            """,
            (
                employee_name,
                utc_now(),
                CaseStatus.RESPONSE_CREATED,
                CaseStatus.COMPLETED,
            ),
        )
        self.reconcile_official_qr_pages()

    @staticmethod
    def _normalize_employee_name(employee_name: str) -> str:
        name = " ".join(employee_name.split())
        if len(name) < 2:
            raise WorkflowValidationError(
                "Введите имя или ФИО исполнителя"
            )
        if len(name) > 120:
            raise WorkflowValidationError(
                "Имя исполнителя не должно быть длиннее 120 символов"
            )
        if not any(character.isalpha() for character in name):
            raise WorkflowValidationError(
                "Имя исполнителя должно содержать буквы"
            )
        if any(ord(character) < 32 for character in name):
            raise WorkflowValidationError(
                "Имя исполнителя содержит недопустимые символы"
            )
        return name

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

    def import_inbox(self, max_files: int = 200) -> dict[str, Any]:
        if not self.get_active_employee():
            raise WorkflowValidationError(
                "Сначала выберите или добавьте исполнителя"
            )
        inbox = self.settings.inbox_dir.resolve()
        inbox.mkdir(parents=True, exist_ok=True)
        candidates = sorted(
            (
                path
                for path in inbox.rglob("*")
                if path.is_file() and path.suffix.casefold() == ".pdf"
            ),
            key=lambda path: str(path).casefold(),
        )[:max_files]
        imported: list[str] = []
        skipped: list[str] = []
        errors: list[dict[str, str]] = []

        for source in candidates:
            try:
                safe_source = ensure_within(source, inbox)
                digest = hashlib.sha256()
                total = 0
                with safe_source.open("rb") as stream:
                    while chunk := stream.read(1024 * 1024):
                        total += len(chunk)
                        if total > self.settings.max_upload_bytes:
                            raise WorkflowValidationError(
                                "PDF превышает допустимый размер"
                            )
                        digest.update(chunk)
                existing = self.db.fetch_one(
                    "SELECT id FROM uploads WHERE sha256 = ?",
                    (digest.hexdigest(),),
                )
                if existing:
                    skipped.append(safe_source.name)
                    continue
                with safe_source.open("rb") as stream:
                    imported.append(
                        self.create_upload(safe_source.name, stream)
                    )
            except Exception as exc:
                errors.append(
                    {
                        "filename": source.name,
                        "message": str(exc)[:300],
                    }
                )

        self.db.audit(
            "settings",
            "inbox",
            "inbox_scanned",
            {
                "folder": str(inbox),
                "found": len(candidates),
                "imported": len(imported),
                "skipped_duplicates": len(skipped),
                "errors": errors,
            },
        )
        return {
            "folder": str(inbox),
            "found": len(candidates),
            "imported": imported,
            "skipped": skipped,
            "errors": errors,
        }

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
        self.reconcile_official_qr_pages(upload_id)
        self._complete_qr_companions(upload_id)
        self._require_letter_for_scan_only_decisions(upload_id)
        self._remove_orphan_cases(upload_id)
        self._refresh_upload_status(upload_id)

    def _require_letter_for_scan_only_decisions(self, upload_id: str) -> int:
        letter = self.db.fetch_one(
            """
            SELECT id FROM pages
            WHERE upload_id = ? AND page_type = ?
            LIMIT 1
            """,
            (upload_id, PageType.LETTER),
        )
        if letter:
            return 0
        decisions = self.db.fetch_all(
            """
            SELECT id, page_number FROM pages
            WHERE upload_id = ?
              AND page_type = ?
              AND status = ?
              AND qr_status != ?
              AND manual_confirmed = 0
            ORDER BY page_number
            """,
            (
                upload_id,
                PageType.DECISION,
                PageStatus.COMPLETED,
                QrStatus.FOUND,
            ),
        )
        if not decisions:
            return 0
        now = utc_now()
        for decision in decisions:
            self.db.execute(
                """
                UPDATE pages
                SET status = ?, issue_code = ?, issue_message = ?,
                    updated_at = ?
                WHERE id = ?
                """,
                (
                    PageStatus.NEEDS_REVIEW,
                    "confirmed_letter_missing",
                    (
                        "Страница уверенно похожа на решение, но в PDF не "
                        "найдено ни одного письма. Проверьте тип, чтобы письмо "
                        "не было пропущено."
                    ),
                    now,
                    decision["id"],
                ),
            )
            self.db.audit(
                "page",
                decision["id"],
                "decision_without_letter_review_required",
                {"page_number": decision["page_number"]},
            )
        return len(decisions)

    def reconcile_confident_scan_decisions(
        self,
        upload_id: str | None = None,
    ) -> int:
        """Safely re-evaluate legacy decision pages after rule updates."""
        params: list[Any] = [
            PageType.DECISION,
            QrStatus.FOUND,
            PageStatus.NEEDS_REVIEW,
            PageStatus.COMPLETED,
        ]
        upload_filter = ""
        if upload_id:
            upload_filter = " AND upload_id = ?"
            params.append(upload_id)
        pages = self.db.fetch_all(
            f"""
            SELECT * FROM pages
            WHERE page_type = ?
              AND qr_status != ?
              AND status IN (?, ?)
              AND manual_confirmed = 0
              {upload_filter}
            ORDER BY upload_id, page_number
            """,
            tuple(params),
        )
        changed = 0
        affected_uploads: set[str] = set()
        letter_cache: dict[str, bool] = {}
        for page in pages:
            classification = self.classifier.classify(
                page.get("extracted_text") or "",
                float(page.get("quality_score") or 0.0),
            )
            current_upload_id = page["upload_id"]
            if current_upload_id not in letter_cache:
                letter_cache[current_upload_id] = bool(
                    self.db.fetch_one(
                        """
                        SELECT id FROM pages
                        WHERE upload_id = ? AND page_type = ?
                        LIMIT 1
                        """,
                        (current_upload_id, PageType.LETTER),
                    )
                )

            if classification.automatic_terminal and letter_cache[
                current_upload_id
            ]:
                status = PageStatus.COMPLETED
                issue_code = None
                issue_message = None
                event_type = "confident_decision_reconciled"
            elif not letter_cache[current_upload_id]:
                status = PageStatus.NEEDS_REVIEW
                issue_code = "confirmed_letter_missing"
                issue_message = (
                    "Это решение распознано уверенно, но во всём PDF не "
                    "найдено письмо. Проверьте, чтобы основной документ не "
                    "был пропущен."
                )
                event_type = "decision_without_letter_review_required"
            else:
                status = PageStatus.NEEDS_REVIEW
                issue_code = "decision_type_not_confident"
                issue_message = (
                    "Страница похожа на решение, но для автоматического "
                    "завершения не хватило независимых признаков формы. "
                    "Проверьте тип страницы."
                )
                event_type = "decision_confidence_review_required"

            if (
                page.get("status") == status
                and page.get("issue_code") == issue_code
                and page.get("issue_message") == issue_message
            ):
                continue
            self.db.execute(
                """
                UPDATE pages
                SET status = ?, issue_code = ?, issue_message = ?,
                    type_confidence = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    status,
                    issue_code,
                    issue_message,
                    classification.confidence,
                    utc_now(),
                    page["id"],
                ),
            )
            self.db.audit(
                "page",
                page["id"],
                event_type,
                {
                    "page_number": page["page_number"],
                    "automatic_terminal": (
                        classification.automatic_terminal
                    ),
                    "reasons": classification.reasons,
                },
            )
            changed += 1
            affected_uploads.add(current_upload_id)

        for affected_upload_id in affected_uploads:
            self._refresh_upload_status(affected_upload_id)
        return changed

    def reconcile_official_qr_pages(
        self, upload_id: str | None = None
    ) -> int:
        parameters: list[Any] = [
            PageStatus.NEEDS_REVIEW,
            QrStatus.FOUND,
            ValueSource.QR_OFFICIAL,
        ]
        upload_filter = ""
        if upload_id:
            upload_filter = "AND pages.upload_id = ?"
            parameters.append(upload_id)
        pages = self.db.fetch_all(
            f"""
            SELECT pages.id, pages.upload_id, pages.page_number
            FROM pages
            JOIN cases ON cases.id = pages.case_id
            WHERE pages.status = ?
              AND pages.manual_confirmed = 0
              AND pages.qr_status = ?
              AND cases.source_kind = ?
              AND cases.fields_confirmed = 1
              AND cases.official_document_path IS NOT NULL
              {upload_filter}
            ORDER BY pages.upload_id, pages.page_number
            """,
            parameters,
        )
        if not pages:
            return 0

        now = utc_now()
        for page in pages:
            self.db.execute(
                """
                UPDATE pages
                SET status = ?, issue_code = NULL, issue_message = NULL,
                    updated_at = ?
                WHERE id = ?
                """,
                (PageStatus.COMPLETED, now, page["id"]),
            )
            self.db.audit(
                "page",
                page["id"],
                "official_qr_auto_completed",
                {
                    "page_number": page["page_number"],
                    "reason": "official_fields_already_confirmed",
                },
            )

        for affected_upload_id in {
            page["upload_id"] for page in pages
        }:
            self._refresh_upload_status(affected_upload_id)
        return len(pages)

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

    def request_page_reprocess(self, page_id: str) -> str:
        page = self.get_page(page_id)
        if not page:
            raise WorkflowValidationError("Страница не найдена")
        if page["manual_confirmed"]:
            raise WorkflowValidationError(
                "Подтверждённую сотрудником страницу нельзя перезапускать"
            )
        if page["status"] not in {
            PageStatus.NEEDS_REVIEW,
            PageStatus.TECHNICAL_ERROR,
        }:
            raise WorkflowValidationError(
                "Эта страница не требует повторной обработки"
            )
        self.db.execute(
            """
            UPDATE pages
            SET status = ?, issue_code = NULL, issue_message = NULL,
                updated_at = ?
            WHERE id = ?
            """,
            (PageStatus.REGISTERED, utc_now(), page_id),
        )
        self.db.execute(
            """
            UPDATE uploads
            SET status = ?, completed_at = NULL
            WHERE id = ?
            """,
            (UploadStatus.PROCESSING, page["upload_id"]),
        )
        self.db.audit(
            "page",
            page_id,
            "page_reprocess_requested",
            {"page_number": page["page_number"]},
        )
        return page["upload_id"]

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
        # Превью является самостоятельным результатом этапа. Сохраняем пути
        # до QR и сетевой загрузки, чтобы таймаут официального сервера не
        # оставлял сотрудника без изображения проблемной страницы.
        self.db.execute(
            """
            UPDATE pages
            SET preview_path = ?, enhanced_preview_path = ?,
                quality_score = ?, status = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                str(rendered.preview_path),
                str(rendered.enhanced_preview_path),
                rendered.quality_score,
                PageStatus.PREVIEW_READY,
                utc_now(),
                page_id,
            ),
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
                            "Не удалось надёжно извлечь все обязательные поля."
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
                        official_complete = self._apply_official_document(
                            case_id,
                            Path(existing_case["official_document_path"]),
                        )
                    if not official_complete:
                        official_issue = (
                            "Официальная версия уже получена по QR. "
                            "Не удалось надёжно извлечь все обязательные поля."
                        )
            # Решение и письмо часто имеют один QR. Перед итоговым статусом
            # повторно читаем карточку из БД: если соседняя страница уже
            # получила и полностью разобрала официальный документ, никакое
            # ручное подтверждение этой страницы больше не требуется.
            refreshed_case = self.get_case(case_id)
            if (
                refreshed_case
                and refreshed_case.get("source_kind")
                == ValueSource.QR_OFFICIAL
                and refreshed_case.get("fields_confirmed")
                and refreshed_case.get("official_document_path")
            ):
                official_complete = True
                official_issue = None

        extracted = None
        if classification.page_type == PageType.LETTER:
            extracted = self.extractor.extract_scan_letter(ocr_result.text)
            office_suggestion = self.match_gns_office(ocr_result.text)
            if case_id is None:
                case_id = self._create_scan_case(upload["id"], page_id)
            self._prefill_scan_case(
                case_id,
                extracted,
                ocr_result.critical_fields_agree,
                office_suggestion,
            )

        page_status, issue_code, issue_message = self._page_outcome(
            classification.page_type,
            classification.confidence,
            qr_result.status,
            official_complete,
            official_issue,
            classification.automatic_terminal,
        )
        if (
            classification.page_type == PageType.LETTER
            and not official_complete
            and not ocr_result.critical_fields_agree
        ):
            issue_code = "ocr_critical_fields_disagree"
            issue_message = (
                "OCR-модели не подтвердили одинаково наименование, ИНН "
                "и период. Критические поля нужно ввести вручную, сверяя "
                "с изображением."
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
                "automatic_terminal": classification.automatic_terminal,
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
        automatic_terminal: bool = False,
    ) -> tuple[PageStatus, str | None, str | None]:
        if official_complete:
            return PageStatus.COMPLETED, None, None
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
        if page_type == PageType.DECISION and automatic_terminal:
            return PageStatus.COMPLETED, None, None
        if page_type == PageType.DECISION:
            return (
                PageStatus.NEEDS_REVIEW,
                "decision_type_not_confident",
                (
                    "Страница похожа на решение, но для автоматического "
                    "завершения не хватило независимых признаков формы. "
                    "Проверьте тип страницы."
                ),
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
                self.get_active_employee(),
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
                self.get_active_employee(),
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

    def _prefill_scan_case(
        self,
        case_id: str,
        fields,
        critical_fields_agree: bool = False,
        office_suggestion: dict[str, Any] | None = None,
    ) -> None:
        current = self.get_case(case_id)
        if not current:
            return
        if (
            current.get("source_kind") == ValueSource.QR_OFFICIAL
            and current.get("fields_confirmed")
        ):
            return
        if office_suggestion and not current.get("district_place"):
            self.db.execute(
                """
                UPDATE cases
                SET district_place = COALESCE(district_place, ?),
                    updated_at = ?
                WHERE id = ?
                """,
                (
                    office_suggestion["district_place"],
                    utc_now(),
                    case_id,
                ),
            )
            self.db.audit(
                "case",
                case_id,
                "gns_office_suggested",
                {
                    "office_id": office_suggestion["id"],
                    "district_place": office_suggestion["district_place"],
                    "source": "ocr_exact_directory_match",
                },
            )
        if not critical_fields_agree:
            self.db.execute(
                """
                UPDATE cases
                SET status = ?, updated_at = ?
                WHERE id = ?
                """,
                (CaseStatus.NEEDS_REVIEW, utc_now(), case_id),
            )
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
                self.check_registry_case(case_id, automatic=True)

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
        employee = self.get_active_employee()
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
                1 if complete else 0,
                (
                    CaseStatus.READY_FOR_ABS
                    if complete
                    else CaseStatus.NEEDS_REVIEW
                ),
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
        return complete

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
        critical_fields_verified: bool = False,
        actor: str = "Сотрудник",
    ) -> str | None:
        page = self.get_page(page_id)
        if not page:
            raise WorkflowValidationError("Страница не найдена")
        if page["status"] not in {
            PageStatus.NEEDS_REVIEW,
            PageStatus.TECHNICAL_ERROR,
        }:
            raise WorkflowValidationError(
                "Страница уже обработана и не требует подтверждения"
            )
        try:
            selected_type = PageType(page_type)
        except ValueError as exc:
            raise WorkflowValidationError("Неизвестный тип страницы") from exc
        if selected_type == PageType.UNKNOWN:
            raise WorkflowValidationError("Нужно обозначить тип страницы")

        case_id = page.get("case_id")
        if selected_type == PageType.LETTER:
            if not critical_fields_verified:
                raise WorkflowValidationError(
                    "Подтвердите, что ИНН, наименования и период сверены "
                    "с изображением страницы"
                )
            employee_name = (
                employee_name.strip() or self.get_active_employee() or ""
            )
            district_place = clean_location(district_place)
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
                    district_place,
                    recipient_position.strip(),
                    recipient_full_name.strip(),
                    display_name,
                    period_start,
                    period_end,
                    employee_name,
                    ValueSource.MANUAL,
                    CaseStatus.READY_FOR_ABS,
                    utc_now(),
                    case_id,
                ),
            )
            self._replace_taxpayers(case_id, clean_taxpayers)
            if page.get("qr_status") != QrStatus.FOUND:
                self.check_registry_case(case_id, automatic=True)

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
    def _registry_name_key(value: str) -> str:
        normalized = value.casefold().replace("ё", "е")
        normalized = re.sub(
            r"\bобщество\s+с\s+ограниченной\s+ответственностью\b",
            "осоо",
            normalized,
        )
        normalized = re.sub(
            r"\bиндивидуальный\s+предприниматель\b",
            "ип",
            normalized,
        )
        return "".join(character for character in normalized if character.isalnum())

    def check_registry_case(
        self,
        case_id: str,
        *,
        automatic: bool = False,
    ) -> dict[str, int]:
        if automatic and not self.settings.auto_registry_check:
            return {"skipped": 1}
        taxpayers = self.get_taxpayers(case_id)
        summary: dict[str, int] = {}
        for taxpayer in taxpayers:
            status = "error"
            official_name = None
            director = None
            try:
                result = self.registry.lookup_by_inn(taxpayer["inn"])
                official_name = result.official_name
                director = result.director
                if result.status == "found" and official_name:
                    status = (
                        "match"
                        if self._registry_name_key(taxpayer["name"])
                        == self._registry_name_key(official_name)
                        else "mismatch"
                    )
                else:
                    status = result.status
            except RegistryLookupError:
                status = "error"

            checked_at = utc_now()
            self.db.execute(
                """
                UPDATE taxpayers
                SET registry_status = ?, registry_name = ?,
                    registry_director = ?, registry_checked_at = ?,
                    updated_at = ?
                WHERE id = ?
                """,
                (
                    status,
                    official_name,
                    director,
                    checked_at,
                    checked_at,
                    taxpayer["id"],
                ),
            )
            summary[status] = summary.get(status, 0) + 1
            self.db.audit(
                "taxpayer",
                taxpayer["id"],
                "registry_checked",
                {
                    "automatic": automatic,
                    "status": status,
                    "provider": "ОсОО.KG",
                },
            )
        case = self.get_case(case_id)
        needs_confirmation = bool(
            automatic
            and case
            and case.get("source_kind") != ValueSource.QR_OFFICIAL
            and any(status != "match" for status in summary)
        )
        if needs_confirmation:
            self.db.execute(
                "UPDATE cases SET status = ?, updated_at = ? WHERE id = ?",
                (CaseStatus.NEEDS_REVIEW, utc_now(), case_id),
            )
        return summary

    def accept_registry_variance(
        self,
        case_id: str,
        actor: str = "Сотрудник",
    ) -> None:
        case = self.get_case(case_id)
        if not case:
            raise WorkflowValidationError("Обращение не найдено")
        if case.get("source_kind") == ValueSource.QR_OFFICIAL:
            raise WorkflowValidationError(
                "Официальные данные QR не требуют подтверждения по реестру"
            )
        taxpayers = self.get_taxpayers(case_id)
        if not case.get("fields_confirmed") or not taxpayers:
            raise WorkflowValidationError(
                "Сначала подтвердите данные письма и налогоплательщиков"
            )
        if not any(
            taxpayer.get("registry_status")
            and taxpayer.get("registry_status") != "match"
            for taxpayer in taxpayers
        ):
            raise WorkflowValidationError(
                "В карточке нет расхождений ОсОО.KG для подтверждения"
            )
        self.db.execute(
            "UPDATE cases SET status = ?, updated_at = ? WHERE id = ?",
            (CaseStatus.READY_FOR_ABS, utc_now(), case_id),
        )
        self.db.audit(
            "case",
            case_id,
            "registry_variance_accepted",
            {
                "results": [
                    {
                        "inn": taxpayer["inn"],
                        "status": taxpayer.get("registry_status"),
                    }
                    for taxpayer in taxpayers
                ]
            },
            actor=actor,
        )

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
                UPDATE taxpayers
                SET abs_result = ?, odb_result = NULL, updated_at = ?
                WHERE case_id = ? AND inn = ?
                """,
                (
                    taxpayer_result["result"],
                    utc_now(),
                    case_id,
                    taxpayer_result["inn"],
                ),
            )

        failed_statuses = {
            AbsStatus.AUTH_ERROR,
            AbsStatus.UNAVAILABLE,
            AbsStatus.TECHNICAL_ERROR,
        }
        start = date.fromisoformat(case["period_start"])
        if result.status in failed_statuses:
            next_status = CaseStatus.READY_FOR_ABS
        elif start < self.settings.period_threshold:
            next_status = CaseStatus.MANUAL_PERIOD_RULE
        elif result.status == AbsStatus.NOT_FOUND:
            next_status = CaseStatus.READY_FOR_RESPONSE
        else:
            next_status = CaseStatus.NEEDS_REVIEW

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

    def confirm_odb(
        self,
        case_id: str,
        results: list[dict[str, str]],
        actor: str = "Сотрудник",
    ) -> str:
        case = self.get_case(case_id)
        if not case:
            raise WorkflowValidationError("Обращение не найдено")
        if case["status"] != CaseStatus.MANUAL_PERIOD_RULE:
            raise WorkflowValidationError(
                "Проверка ОДБ для этого обращения не требуется"
            )
        if date.fromisoformat(case["period_start"]) >= (
            self.settings.period_threshold
        ):
            raise WorkflowValidationError(
                "Период обращения не относится к старой АБС"
            )

        taxpayers = self.get_taxpayers(case_id)
        expected = {taxpayer["inn"] for taxpayer in taxpayers}
        supplied = {item.get("inn", "") for item in results}
        if expected != supplied:
            raise WorkflowValidationError(
                "Нужно отметить результат ОДБ для каждого налогоплательщика"
            )
        allowed = {OdbStatus.FOUND, OdbStatus.NOT_FOUND}
        normalized: dict[str, OdbStatus] = {}
        for item in results:
            try:
                normalized[item["inn"]] = OdbStatus(item["result"])
            except (KeyError, ValueError) as exc:
                raise WorkflowValidationError(
                    "Неизвестный результат проверки ОДБ"
                ) from exc
        if set(normalized.values()) - allowed:
            raise WorkflowValidationError(
                "Неизвестный результат проверки ОДБ"
            )

        now = utc_now()
        for inn, odb_result in normalized.items():
            self.db.execute(
                """
                UPDATE taxpayers
                SET odb_result = ?, updated_at = ?
                WHERE case_id = ? AND inn = ?
                """,
                (odb_result, now, case_id, inn),
            )

        refreshed = self.get_taxpayers(case_id)
        all_absent = all(
            taxpayer.get("abs_result") == AbsStatus.NOT_FOUND
            and taxpayer.get("odb_result") == OdbStatus.NOT_FOUND
            for taxpayer in refreshed
        )
        next_status = (
            CaseStatus.READY_FOR_RESPONSE
            if all_absent
            else CaseStatus.MANUAL_PERIOD_RULE
        )
        self.db.execute(
            "UPDATE cases SET status = ?, updated_at = ? WHERE id = ?",
            (next_status, now, case_id),
        )
        self.db.audit(
            "case",
            case_id,
            "odb_checked",
            {
                "results": [
                    {"inn": inn, "result": result}
                    for inn, result in normalized.items()
                ],
                "next_status": next_status,
            },
            actor=actor,
        )
        return next_status

    def check_abs_today(
        self,
        username: str,
        password: str,
        business_date: date | None = None,
    ) -> dict[str, Any]:
        day = business_date or datetime.now(self.BUSINESS_TIMEZONE).date()
        start_utc, end_utc = self._business_day_bounds(day)
        cases = self.db.fetch_all(
            """
            SELECT cases.id
            FROM cases
            JOIN uploads ON uploads.id = cases.upload_id
            WHERE uploads.created_at >= ? AND uploads.created_at < ?
              AND cases.status = ?
              AND cases.fields_confirmed = 1
            ORDER BY cases.created_at, cases.id
            """,
            (start_utc, end_utc, CaseStatus.READY_FOR_ABS),
        )
        if not cases:
            raise WorkflowValidationError(
                "За сегодня нет подтверждённых обращений, ожидающих АБС"
            )

        counts: dict[str, int] = {}
        for case in cases:
            result = self.check_abs(case["id"], username, password)
            key = str(result.status)
            counts[key] = counts.get(key, 0) + 1

        # Явно разрываем ссылки после завершения пакетного запроса.
        username = ""
        password = ""
        self.db.audit(
            "settings",
            f"abs-batch-{day.isoformat()}",
            "fake_abs_batch_checked",
            {
                "business_date": day.isoformat(),
                "case_count": len(cases),
                "results": counts,
                "is_fake": True,
            },
        )
        return {
            "business_date": day.isoformat(),
            "case_count": len(cases),
            "results": counts,
        }

    def today_overview(
        self,
        business_date: date | None = None,
    ) -> dict[str, Any]:
        day = business_date or datetime.now(self.BUSINESS_TIMEZONE).date()
        start_utc, end_utc = self._business_day_bounds(day)
        counts = self.db.fetch_one(
            """
            SELECT
                SUM(CASE WHEN cases.status = 'ready_for_abs'
                    THEN 1 ELSE 0 END) AS ready_abs,
                SUM(CASE WHEN cases.status = 'ready_for_abs'
                          AND cases.source_kind = 'qr_official'
                    THEN 1 ELSE 0 END) AS ready_qr,
                SUM(CASE WHEN cases.status = 'ready_for_response'
                    THEN 1 ELSE 0 END) AS ready_response,
                SUM(CASE
                    WHEN cases.status = 'manual_period_rule'
                     AND EXISTS (
                        SELECT 1
                        FROM taxpayers AS odb_taxpayers
                        WHERE odb_taxpayers.case_id = cases.id
                          AND odb_taxpayers.odb_result IS NULL
                     )
                    THEN 1 ELSE 0
                END) AS odb_pending
            FROM cases
            JOIN uploads ON uploads.id = cases.upload_id
            WHERE uploads.created_at >= ? AND uploads.created_at < ?
            """,
            (start_utc, end_utc),
        ) or {}
        unresolved = self.db.fetch_one(
            """
            SELECT COUNT(*) AS count
            FROM pages
            JOIN uploads ON uploads.id = pages.upload_id
            WHERE uploads.created_at >= ? AND uploads.created_at < ?
              AND pages.status IN ('needs_review', 'technical_error')
            """,
            (start_utc, end_utc),
        ) or {"count": 0}
        generated = self.db.fetch_all(
            """
            SELECT *
            FROM response_groups
            WHERE business_date = ?
            ORDER BY created_at DESC
            """,
            (day.isoformat(),),
        )
        return {
            "business_date": day.isoformat(),
            "ready_abs": int(counts.get("ready_abs") or 0),
            "ready_qr": int(counts.get("ready_qr") or 0),
            "ready_response": int(counts.get("ready_response") or 0),
            "odb_pending": int(counts.get("odb_pending") or 0),
            "unresolved_pages": int(unresolved["count"]),
            "odb_cases": self.db.fetch_all(
                """
                SELECT cases.id, cases.recipient_display_name,
                       cases.district_place, cases.period_start,
                       COUNT(taxpayers.id) AS taxpayer_count,
                       SUM(CASE WHEN taxpayers.odb_result IS NULL
                           THEN 1 ELSE 0 END) AS pending_count
                FROM cases
                JOIN uploads ON uploads.id = cases.upload_id
                JOIN taxpayers ON taxpayers.case_id = cases.id
                WHERE uploads.created_at >= ? AND uploads.created_at < ?
                  AND cases.status = 'manual_period_rule'
                GROUP BY cases.id
                HAVING SUM(CASE WHEN taxpayers.odb_result IS NULL
                    THEN 1 ELSE 0 END) > 0
                ORDER BY cases.created_at
                """,
                (start_utc, end_utc),
            ),
            "not_found_groups": self._daily_response_groups(
                day, AbsStatus.NOT_FOUND
            ),
            "found_groups": self._daily_response_groups(
                day, AbsStatus.FOUND
            ),
            "generated_groups": generated,
        }

    def _daily_response_groups(
        self,
        business_date: date,
        abs_bucket: AbsStatus,
    ) -> list[dict[str, Any]]:
        start_utc, end_utc = self._business_day_bounds(business_date)
        if abs_bucket == AbsStatus.NOT_FOUND:
            extra_where = (
                "cases.status = 'ready_for_response' "
                "AND taxpayers.abs_result = 'not_found'"
            )
        else:
            extra_where = (
                "(taxpayers.abs_result = 'found' "
                "OR taxpayers.odb_result = 'found') "
                "AND ("
                "cases.status != 'manual_period_rule' "
                "OR NOT EXISTS ("
                "SELECT 1 FROM taxpayers AS pending_odb "
                "WHERE pending_odb.case_id = cases.id "
                "AND pending_odb.odb_result IS NULL"
                ")"
                ")"
            )
        rows = self.db.fetch_all(
            f"""
            SELECT
                cases.id AS case_id,
                cases.district_place,
                cases.recipient_position,
                cases.recipient_full_name,
                cases.recipient_display_name,
                cases.employee_name,
                cases.created_at AS case_created_at,
                taxpayers.name AS taxpayer_name,
                taxpayers.inn AS taxpayer_inn,
                taxpayers.display_order
            FROM cases
            JOIN uploads ON uploads.id = cases.upload_id
            JOIN taxpayers ON taxpayers.case_id = cases.id
            WHERE uploads.created_at >= ? AND uploads.created_at < ?
              AND {extra_where}
            ORDER BY cases.created_at, cases.id, taxpayers.display_order
            """,
            (start_utc, end_utc),
        )

        grouped: dict[str, dict[str, Any]] = {}
        for row in rows:
            district = clean_location(row.get("district_place"))
            recipient = " ".join(
                (row.get("recipient_full_name") or "").split()
            )
            group_key = self._response_group_key(
                business_date,
                abs_bucket,
                district,
                recipient,
            )
            group = grouped.setdefault(
                group_key,
                {
                    "group_key": group_key,
                    "business_date": business_date.isoformat(),
                    "abs_bucket": str(abs_bucket),
                    "district_place": district,
                    "recipient_position": row.get("recipient_position") or "",
                    "recipient_full_name": recipient,
                    "recipient_display_name": (
                        row.get("recipient_display_name") or ""
                    ),
                    "employee_name": row.get("employee_name") or "",
                    "case_ids": [],
                    "taxpayers": [],
                    "issues": [],
                    "_positions": set(),
                    "_display_names": set(),
                    "_employees": set(),
                    "_taxpayers_by_inn": {},
                },
            )
            if row["case_id"] not in group["case_ids"]:
                group["case_ids"].append(row["case_id"])
            group["_positions"].add(
                self._normalize_group_value(row.get("recipient_position"))
            )
            group["_display_names"].add(
                self._normalize_group_value(
                    row.get("recipient_display_name")
                )
            )
            group["_employees"].add(
                self._normalize_group_value(row.get("employee_name"))
            )

            inn = re.sub(r"\D", "", row.get("taxpayer_inn") or "")
            name = " ".join((row.get("taxpayer_name") or "").split())
            existing = group["_taxpayers_by_inn"].get(inn)
            if existing is None:
                taxpayer = {
                    "source_case_id": row["case_id"],
                    "name": name,
                    "inn": inn,
                }
                group["_taxpayers_by_inn"][inn] = taxpayer
                group["taxpayers"].append(taxpayer)
            elif self._normalize_group_value(existing["name"]) != (
                self._normalize_group_value(name)
            ):
                issue = (
                    f"ИНН {inn} встречается с разными наименованиями. "
                    "Нужно проверить вручную."
                )
                if issue not in group["issues"]:
                    group["issues"].append(issue)

        result: list[dict[str, Any]] = []
        for group in grouped.values():
            if len(group["_positions"]) != 1:
                group["issues"].append(
                    "У одного адресата отличаются должности в письмах."
                )
            if len(group["_display_names"]) != 1:
                group["issues"].append(
                    "У одного адресата отличаются формы обращения."
                )
            if len(group["_employees"]) != 1 or not group["employee_name"]:
                group["issues"].append(
                    "Для группы не определён единый исполнитель банка."
                )
            if abs_bucket == AbsStatus.FOUND:
                group["issues"].append(
                    "Налогоплательщики найдены в АБС. Для них нужен "
                    "отдельный утверждённый шаблон ответа."
                )
            group["case_count"] = len(group["case_ids"])
            group["taxpayer_count"] = len(group["taxpayers"])
            group["can_generate"] = bool(
                abs_bucket == AbsStatus.NOT_FOUND
                and group["taxpayers"]
                and not group["issues"]
            )
            for private_key in (
                "_positions",
                "_display_names",
                "_employees",
                "_taxpayers_by_inn",
            ):
                group.pop(private_key, None)
            result.append(group)

        return sorted(
            result,
            key=lambda item: (
                item["recipient_full_name"].casefold(),
                item["district_place"].casefold(),
            ),
        )

    def generate_daily_response(
        self,
        group_key: str,
        business_date: date | None = None,
    ) -> tuple[str, Path]:
        day = business_date or datetime.now(self.BUSINESS_TIMEZONE).date()
        group = next(
            (
                item
                for item in self._daily_response_groups(
                    day, AbsStatus.NOT_FOUND
                )
                if item["group_key"] == group_key
            ),
            None,
        )
        if not group:
            raise WorkflowValidationError(
                "Группа уже обработана или больше не готова к ответу"
            )
        if not group["can_generate"]:
            raise WorkflowValidationError(
                "Нельзя создать ответ: " + " ".join(group["issues"])
            )

        group_id = uuid4().hex
        output = (
            self.settings.responses_dir
            / f"response-group-{day.isoformat()}-{group_id}.docx"
        )
        case_snapshot = {
            "district_place": group["district_place"],
            "recipient_position": group["recipient_position"],
            "recipient_display_name": group["recipient_display_name"],
            "employee_name": group["employee_name"],
        }
        self.word.render(output, case_snapshot, group["taxpayers"])

        now = utc_now()
        with self.db.connect() as connection:
            connection.execute(
                """
                INSERT INTO response_groups(
                    id, business_date, group_key, abs_bucket, status,
                    district_place, recipient_position, recipient_full_name,
                    recipient_display_name, employee_name, taxpayer_count,
                    response_path, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    group_id,
                    day.isoformat(),
                    group_key,
                    str(AbsStatus.NOT_FOUND),
                    "created",
                    group["district_place"],
                    group["recipient_position"],
                    group["recipient_full_name"],
                    group["recipient_display_name"],
                    group["employee_name"],
                    group["taxpayer_count"],
                    str(output),
                    now,
                    now,
                ),
            )
            connection.executemany(
                """
                INSERT INTO response_group_cases(response_group_id, case_id)
                VALUES (?, ?)
                """,
                [(group_id, case_id) for case_id in group["case_ids"]],
            )
            connection.executemany(
                """
                INSERT INTO response_group_taxpayers(
                    response_group_id, display_order, source_case_id,
                    name, inn
                ) VALUES (?, ?, ?, ?, ?)
                """,
                [
                    (
                        group_id,
                        index,
                        taxpayer["source_case_id"],
                        taxpayer["name"],
                        taxpayer["inn"],
                    )
                    for index, taxpayer in enumerate(
                        group["taxpayers"], 1
                    )
                ],
            )
            connection.executemany(
                """
                UPDATE cases
                SET status = ?, response_status = ?,
                    response_path = ?, updated_at = ?
                WHERE id = ? AND status = ?
                """,
                [
                    (
                        CaseStatus.RESPONSE_CREATED,
                        "grouped",
                        str(output),
                        now,
                        case_id,
                        CaseStatus.READY_FOR_RESPONSE,
                    )
                    for case_id in group["case_ids"]
                ],
            )

        self.db.audit(
            "response_group",
            group_id,
            "grouped_response_created",
            {
                "business_date": day.isoformat(),
                "case_count": group["case_count"],
                "taxpayer_count": group["taxpayer_count"],
                "filename": output.name,
            },
        )
        return group_id, output

    def get_response_group(
        self, group_id: str
    ) -> dict[str, Any] | None:
        return self.db.fetch_one(
            "SELECT * FROM response_groups WHERE id = ?",
            (group_id,),
        )

    def repair_cleaned_responses(self) -> None:
        cases = self.db.fetch_all(
            """
            SELECT cases.*
            FROM cases
            WHERE cases.response_path IS NOT NULL
              AND EXISTS (
                  SELECT 1
                  FROM audit_events AS cleanup
                  WHERE cleanup.entity_type = 'case'
                    AND cleanup.entity_id = cases.id
                    AND cleanup.event_type =
                        'district_place_edge_noise_removed'
              )
              AND NOT EXISTS (
                  SELECT 1
                  FROM audit_events AS repaired
                  WHERE repaired.entity_type = 'case'
                    AND repaired.entity_id = cases.id
                    AND repaired.event_type =
                        'response_regenerated_after_cleanup'
              )
            """
        )
        for case in cases:
            path = Path(case["response_path"])
            try:
                self.word.render(
                    path,
                    case,
                    self.get_taxpayers(case["id"]),
                )
            except (OSError, ValueError):
                continue
            self.db.audit(
                "case",
                case["id"],
                "response_regenerated_after_cleanup",
                {"filename": path.name},
            )

    @classmethod
    def _business_day_bounds(
        cls, business_date: date
    ) -> tuple[str, str]:
        local_start = datetime.combine(
            business_date,
            time.min,
            tzinfo=cls.BUSINESS_TIMEZONE,
        )
        local_end = datetime.combine(
            business_date + timedelta(days=1),
            time.min,
            tzinfo=cls.BUSINESS_TIMEZONE,
        )
        return (
            local_start.astimezone(UTC).isoformat(timespec="seconds"),
            local_end.astimezone(UTC).isoformat(timespec="seconds"),
        )

    @staticmethod
    def _normalize_group_value(value: str | None) -> str:
        return " ".join((value or "").split()).casefold()

    @classmethod
    def _response_group_key(
        cls,
        business_date: date,
        abs_bucket: AbsStatus,
        district_place: str,
        recipient_full_name: str,
    ) -> str:
        identity = "\x1f".join(
            (
                business_date.isoformat(),
                str(abs_bucket),
                cls._normalize_group_value(district_place),
                cls._normalize_group_value(recipient_full_name),
            )
        )
        return hashlib.sha256(identity.encode("utf-8")).hexdigest()[:24]

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

    def get_page_pdf_path(self, page_id: str) -> Path:
        page = self.get_page(page_id)
        if not page:
            raise WorkflowValidationError("Страница не найдена")
        upload = self.get_upload(page["upload_id"])
        if not upload:
            raise WorkflowValidationError("Исходный PDF не найден")
        try:
            source = ensure_within(
                Path(upload["stored_path"]),
                self.settings.uploads_dir,
            )
            output = ensure_within(
                self.settings.previews_dir
                / upload["id"]
                / f"page-{int(page['page_number']):04d}.pdf",
                self.settings.previews_dir,
            )
            if not output.exists() or (
                source.stat().st_mtime_ns > output.stat().st_mtime_ns
            ):
                self.pdf.extract_page_pdf(
                    source,
                    int(page["page_number"]),
                    output,
                )
            return output
        except (OSError, PdfProcessingError, ValueError) as exc:
            raise WorkflowValidationError(str(exc)) from exc

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
