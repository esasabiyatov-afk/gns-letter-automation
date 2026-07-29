from __future__ import annotations

import re
import unicodedata
from pathlib import Path

from gns_app.domain import OcrResult, OcrStatus
from gns_app.services.pdf_service import PdfService


class OcrService:
    """Безопасный fallback до подключения проверенной модели rus+kir.

    Существующий текстовый слой PDF используется только как подсказка
    для сотрудника. Он никогда не получает достаточную уверенность для
    автоматического ответа.
    """

    def __init__(self, pdf_service: PdfService):
        self.pdf_service = pdf_service

    def recognize(self, pdf_path: Path, page_number: int) -> OcrResult:
        text = unicodedata.normalize(
            "NFC",
            self.pdf_service.extract_embedded_text(pdf_path, page_number),
        )
        cleaned = "\n".join(line.rstrip() for line in text.splitlines()).strip()
        if not cleaned:
            return OcrResult(
                status=OcrStatus.REQUIRES_ENGINE,
                text="",
                confidence=0.0,
                language="rus+kir",
                issue=(
                    "В PDF нет текстового слоя. Нужна локальная OCR-модель "
                    "с поддержкой русского и кыргызского."
                ),
            )

        replacement_count = cleaned.count("\ufffd")
        non_space = max(1, len(re.sub(r"\s", "", cleaned)))
        replacement_ratio = replacement_count / non_space
        length_factor = min(1.0, len(cleaned) / 800)
        confidence = max(0.15, min(0.58, 0.35 + 0.25 * length_factor))
        confidence *= max(0.25, 1.0 - replacement_ratio * 8)

        return OcrResult(
            status=OcrStatus.EMBEDDED_TEXT,
            text=cleaned,
            confidence=round(confidence, 3),
            language="unknown (Unicode preserved)",
            issue=(
                "Использован существующий текстовый слой PDF. "
                "Результат не подтверждается автоматически."
            ),
        )

