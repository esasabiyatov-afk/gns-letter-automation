from __future__ import annotations

import re
import unicodedata
from pathlib import Path

from PIL import Image, ImageFilter, ImageOps

from gns_app.domain import OcrResult, OcrStatus
from gns_app.services.pdf_service import PdfService

try:
    import tesserocr
except ImportError:  # pragma: no cover - зависит от локальной Windows-среды
    tesserocr = None


class OcrService:
    """Локальный OCR rus+kir без генеративного восстановления символов."""

    def __init__(
        self,
        pdf_service: PdfService,
        fast_data_dir: Path | None = None,
        best_data_dir: Path | None = None,
    ):
        self.pdf_service = pdf_service
        self.fast_data_dir = fast_data_dir
        self.best_data_dir = best_data_dir

    def recognize(
        self,
        pdf_path: Path,
        page_number: int,
        image_path: Path | None = None,
    ) -> OcrResult:
        text = unicodedata.normalize(
            "NFC",
            self.pdf_service.extract_embedded_text(pdf_path, page_number),
        )
        cleaned = "\n".join(line.rstrip() for line in text.splitlines()).strip()
        embedded_text_rejected = bool(
            cleaned and not self._embedded_text_is_reliable(cleaned)
        )
        if cleaned and not embedded_text_rejected:
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
                language="Unicode из PDF",
                issue=(
                    "Использован существующий текстовый слой PDF. "
                    "Результат не подтверждается автоматически."
                ),
            )

        if image_path and self._model_available(self.fast_data_dir):
            try:
                fast = self._recognize_image(
                    image_path, self.fast_data_dir, "fast"
                )
                if self._model_available(self.best_data_dir):
                    best = self._recognize_image(
                        image_path, self.best_data_dir, "best"
                    )
                    result = (
                        best
                        if self._quality_key(best) > self._quality_key(fast)
                        else fast
                    )
                    result.critical_fields_agree = (
                        self._critical_field_signature(fast.text)
                        == self._critical_field_signature(best.text)
                        != ()
                    )
                else:
                    result = fast
                issue_parts: list[str] = []
                if embedded_text_rejected:
                    issue_parts.append(
                        "Повреждённый текстовый слой PDF отклонён. "
                        "Показан результат локального OCR rus+kir без "
                        "догадок и автоматической подмены символов."
                    )
                if result.critical_fields_agree:
                    issue_parts.append(
                        "Два OCR-прохода одинаково прочитали наименование, "
                        "ИНН и период. Их всё равно нужно сверить с "
                        "изображением."
                    )
                else:
                    issue_parts.append(
                        "Два независимых OCR-прохода не подтвердили "
                        "одинаково наименование, ИНН и период. Критические "
                        "поля не предзаполняются."
                    )
                result.issue = " ".join(issue_parts)
                return result
            except Exception as exc:
                return OcrResult(
                    status=OcrStatus.ERROR,
                    text="",
                    confidence=0.0,
                    language="rus+kir",
                    issue=f"Локальный OCR завершился ошибкой: {str(exc)[:300]}",
                )

        return OcrResult(
            status=OcrStatus.REQUIRES_ENGINE,
            text="",
            confidence=0.0,
            language="rus+kir",
            issue=(
                (
                    "Встроенный текстовый слой PDF выглядит повреждённым. "
                    if embedded_text_rejected
                    else "В PDF нет текстового слоя. "
                )
                + "Локальная модель rus+kir не установлена или недоступна."
            ),
        )

    @staticmethod
    def _embedded_text_is_reliable(text: str) -> bool:
        """Отбрасывает ложный OCR-слой, записанный латиницей вместо кириллицы."""
        letters = [character for character in text if character.isalpha()]
        if len(letters) < 40:
            return True

        cyrillic = sum(
            "\u0400" <= character <= "\u052f" for character in letters
        )
        latin = sum(
            "a" <= character.casefold() <= "z" for character in letters
        )
        replacement_ratio = text.count("\ufffd") / max(1, len(text))
        cyrillic_ratio = cyrillic / len(letters)
        latin_ratio = latin / len(letters)

        if replacement_ratio > 0.005:
            return False
        return not (cyrillic_ratio < 0.45 and latin_ratio > 0.35)

    @staticmethod
    def _model_available(path: Path | None) -> bool:
        return bool(
            tesserocr
            and path
            and (path / "rus.traineddata").is_file()
            and (path / "kir.traineddata").is_file()
        )

    @staticmethod
    def _quality_key(result: OcrResult) -> tuple[float, int]:
        return result.confidence, len(result.text)

    @staticmethod
    def _critical_field_signature(text: str) -> tuple[str, ...]:
        normalized = unicodedata.normalize("NFC", text)
        names = [
            re.sub(r"\s+", " ", match.group(1)).strip().casefold()
            for match in re.finditer(
                r"наименование\s*[:;]\s*(.+?)(?=\s+инн\s*[:;])",
                normalized,
                re.IGNORECASE | re.DOTALL,
            )
        ]
        inns = []
        for match in re.finditer(
            r"инн\s*[:;]\s*([\d\s]{14,28})",
            normalized,
            re.IGNORECASE,
        ):
            digits = re.sub(r"\D", "", match.group(1))
            if len(digits) == 14:
                inns.append(digits)
        periods = [
            (
                match.group(1).replace("/", ".").replace("-", ".")
                + "|"
                + match.group(2).replace("/", ".").replace("-", ".")
            )
            for match in re.finditer(
                r"период\s*[:;]?\s*с\s*"
                r"(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4})\s*по\s*"
                r"(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4})",
                normalized,
                re.IGNORECASE,
            )
        ]
        if not names or not inns or not periods:
            return ()
        return (
            *(f"name:{value}" for value in names),
            *(f"inn:{value}" for value in inns),
            *(f"period:{value}" for value in periods),
        )

    @staticmethod
    def _recognize_image(
        image_path: Path,
        data_dir: Path,
        model_name: str,
    ) -> OcrResult:
        image = Image.open(image_path).convert("L")
        image = ImageOps.autocontrast(image, cutoff=1).filter(
            ImageFilter.UnsharpMask(radius=1.1, percent=115, threshold=4)
        )

        with tesserocr.PyTessBaseAPI(
            path=str(data_dir),
            lang="rus+kir",
            psm=tesserocr.PSM.AUTO,
        ) as api:
            api.SetVariable("preserve_interword_spaces", "1")
            api.SetImage(image)
            text = unicodedata.normalize("NFC", api.GetUTF8Text() or "")
            mean_confidence = max(0, min(100, api.MeanTextConf()))

        cleaned = "\n".join(
            line.rstrip() for line in text.splitlines()
        ).strip()
        confidence = round(min(0.84, mean_confidence / 100), 3)
        status = (
            OcrStatus.COMPLETED
            if confidence >= 0.65 and len(cleaned) >= 80
            else OcrStatus.LOW_CONFIDENCE
        )
        return OcrResult(
            status=status,
            text=cleaned,
            confidence=confidence,
            language=f"rus+kir (Tesseract {model_name})",
            issue=(
                "Текст получен локальным OCR rus+kir. Поля из скана "
                "требуют подтверждения сотрудника."
            ),
        )
