from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any


class UploadStatus(StrEnum):
    REGISTERED = "registered"
    PROCESSING = "processing"
    NEEDS_REVIEW = "needs_review"
    READY = "ready"
    COMPLETED = "completed"
    TECHNICAL_ERROR = "technical_error"


class PageStatus(StrEnum):
    REGISTERED = "registered"
    PREVIEW_READY = "preview_ready"
    PROCESSING = "processing"
    QR_RESOLVED = "qr_resolved"
    SCAN_FALLBACK = "scan_fallback"
    NEEDS_REVIEW = "needs_review"
    MANUALLY_CONFIRMED = "manually_confirmed"
    COMPLETED = "completed"
    TECHNICAL_ERROR = "technical_error"


class PageType(StrEnum):
    LETTER = "letter"
    DECISION = "decision"
    ATTACHMENT = "attachment"
    OTHER = "other"
    UNKNOWN = "unknown"


class QrStatus(StrEnum):
    NOT_STARTED = "not_started"
    FOUND = "found"
    INVALID_URL = "invalid_url"
    NOT_FOUND = "not_found"
    DECODE_ERROR = "decode_error"


class OcrStatus(StrEnum):
    NOT_STARTED = "not_started"
    SKIPPED_OFFICIAL = "skipped_official"
    EMBEDDED_TEXT = "embedded_text"
    REQUIRES_ENGINE = "requires_engine"
    COMPLETED = "completed"
    LOW_CONFIDENCE = "low_confidence"
    ERROR = "error"


class CaseStatus(StrEnum):
    COLLECTING = "collecting"
    NEEDS_REVIEW = "needs_review"
    READY_FOR_ABS = "ready_for_abs"
    ABS_CHECKING = "abs_checking"
    MANUAL_PERIOD_RULE = "manual_period_rule"
    READY_FOR_RESPONSE = "ready_for_response"
    RESPONSE_CREATED = "response_created"
    COMPLETED = "completed"
    TECHNICAL_ERROR = "technical_error"


class AbsStatus(StrEnum):
    NOT_CHECKED = "not_checked"
    CHECKING = "checking"
    FOUND = "found"
    NOT_FOUND = "not_found"
    MULTIPLE = "multiple"
    AUTH_ERROR = "auth_error"
    UNAVAILABLE = "unavailable"
    TECHNICAL_ERROR = "technical_error"
    INVALID_INN = "invalid_inn"


class OdbStatus(StrEnum):
    NOT_CHECKED = "not_checked"
    FOUND = "found"
    NOT_FOUND = "not_found"


class ValueSource(StrEnum):
    QR_LINK = "qr_link"
    QR_OFFICIAL = "qr_official"
    OCR_SCAN = "ocr_scan"
    MANUAL = "manual"
    PROFILE = "profile"
    SYSTEM = "system"


@dataclass(slots=True)
class QrDecodeResult:
    status: QrStatus
    payload: str | None = None
    payload_hash: str | None = None
    safe_url: str | None = None
    method: str | None = None
    issue: str | None = None


@dataclass(slots=True)
class OcrResult:
    status: OcrStatus
    text: str
    confidence: float
    language: str
    issue: str | None = None
    critical_fields_agree: bool = False
    taxpayer_fields_agree: bool = False
    period_fields_agree: bool = False
    recipient_fields_agree: bool = False


@dataclass(slots=True)
class ClassificationResult:
    page_type: PageType
    confidence: float
    reasons: list[str] = field(default_factory=list)
    automatic_terminal: bool = False


@dataclass(frozen=True, slots=True)
class VisualPageEvidence:
    decision_layout: bool
    confidence: float
    horizontal_line_groups: int = 0
    vertical_line_groups: int = 0
    ink_density: float = 0.0
    reasons: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ExtractedTaxpayer:
    name: str | None
    inn: str | None
    confidence: float
    issues: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ExtractedFields:
    district_place: str | None = None
    recipient_position: str | None = None
    recipient_full_name: str | None = None
    period_start: str | None = None
    period_end: str | None = None
    taxpayers: list[ExtractedTaxpayer] = field(default_factory=list)
    confidence: float = 0.0
    issues: list[str] = field(default_factory=list)


@dataclass(slots=True)
class AbsCheckResult:
    status: AbsStatus
    taxpayers: list[dict[str, Any]]
    message: str
    is_fake: bool = True
