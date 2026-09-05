from __future__ import annotations

import re
import time
import unicodedata
from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageFilter, ImageOps

from gns_app.domain import OcrResult, OcrStatus
from gns_app.diagnostics import record_event, record_exception
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
        self._last_health: dict[str, bool | str] = {}

    def health_check(self) -> dict[str, bool | str]:
        """Check that packaged OCR data can initialize the native engine."""
        result: dict[str, bool | str] = {
            "engine_imported": bool(tesserocr),
        }
        for model_name, data_dir in (
            ("fast", self.fast_data_dir),
            ("best", self.best_data_dir),
        ):
            files_available = self._model_available(data_dir)
            result[f"{model_name}_files_available"] = files_available
            initialized = False
            error_type = ""
            if files_available and data_dir is not None:
                try:
                    with tesserocr.PyTessBaseAPI(
                        path=str(data_dir),
                        lang="rus+kir",
                    ):
                        initialized = True
                except Exception as exc:  # pragma: no cover - native runtime
                    error_type = type(exc).__name__
            result[f"{model_name}_initialized"] = initialized
            result[f"{model_name}_error_type"] = error_type
        self._last_health = result
        return dict(result)

    def last_health(self) -> dict[str, bool | str]:
        return dict(self._last_health)

    def recognize(
        self,
        pdf_path: Path,
        page_number: int,
        image_path: Path | None = None,
        model_name: str = "fast",
    ) -> OcrResult:
        started = time.monotonic()
        if model_name not in {"fast", "best"}:
            raise ValueError("Неизвестная локальная OCR-модель")
        data_dir = (
            self.fast_data_dir if model_name == "fast" else self.best_data_dir
        )
        text = unicodedata.normalize(
            "NFC",
            self.pdf_service.extract_embedded_text(pdf_path, page_number),
        )
        cleaned = "\n".join(line.rstrip() for line in text.splitlines()).strip()
        embedded_text_rejected = bool(
            cleaned and not self._embedded_text_is_reliable(cleaned)
        )
        record_event(
            "ocr",
            "recognize",
            "started",
            details={
                "model": model_name,
                "image_available": bool(image_path and image_path.is_file()),
                "model_available": self._model_available(data_dir),
                "embedded_text_length": len(cleaned),
                "embedded_text_rejected": embedded_text_rejected,
            },
        )
        if image_path and self._model_available(data_dir):
            try:
                result = self._recognize_image(
                    image_path,
                    data_dir,
                    model_name,
                    supplement_regions=bool(
                        cleaned and not embedded_text_rejected
                    ),
                )
                issue_parts: list[str] = []
                if embedded_text_rejected:
                    issue_parts.append(
                        "Повреждённый текстовый слой PDF отклонён. "
                        "Показан результат локального OCR rus+kir без "
                        "догадок и автоматической подмены символов."
                    )
                elif cleaned:
                    result.text = self._combine_text_sources(
                        cleaned, result.text
                    )
                    result.language = "Unicode PDF + rus+kir (вся страница)"
                    issue_parts.append(
                        "Текстовый слой PDF может быть неполным, поэтому "
                        "дополнительно распознано изображение всей страницы. "
                        "Оба варианта показаны для ручной сверки."
                    )
                issue_parts.append(
                    "Локальный OCR показан только как неподтверждённая "
                    "подсказка. Наименование, ИНН и период необходимо "
                    "посимвольно сверить с изображением."
                )
                result.issue = " ".join(issue_parts)
                self._record_result(result, model_name, started)
                return result
            except Exception as exc:
                record_exception(
                    "ocr",
                    "recognize_image",
                    exc,
                    details={
                        "model": model_name,
                        "duration_ms": int((time.monotonic() - started) * 1000),
                        "embedded_fallback_available": bool(
                            cleaned and not embedded_text_rejected
                        ),
                    },
                )
                if cleaned and not embedded_text_rejected:
                    result = self._embedded_result(
                        cleaned,
                        "Дополнительный OCR изображения завершился "
                        "ошибкой; показан только текстовый слой PDF.",
                    )
                    self._record_result(result, model_name, started)
                    return result
                result = OcrResult(
                    status=OcrStatus.ERROR,
                    text="",
                    confidence=0.0,
                    language=f"rus+kir ({model_name})",
                    issue=f"Локальный OCR завершился ошибкой: {str(exc)[:300]}",
                )
                self._record_result(result, model_name, started)
                return result

        if cleaned and not embedded_text_rejected:
            result = self._embedded_result(cleaned)
            self._record_result(result, model_name, started)
            return result

        result = OcrResult(
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
                + f"Локальная модель rus+kir {model_name} не установлена "
                "или недоступна."
            ),
        )
        self._record_result(result, model_name, started)
        return result

    @staticmethod
    def _record_result(
        result: OcrResult,
        model_name: str,
        started: float,
    ) -> None:
        record_event(
            "ocr",
            "recognize",
            str(result.status),
            details={
                "model": model_name,
                "duration_ms": int((time.monotonic() - started) * 1000),
                "text_length": len(result.text),
                "confidence": result.confidence,
            },
        )

    @staticmethod
    def _embedded_result(text: str, extra_issue: str = "") -> OcrResult:
        replacement_count = text.count("\ufffd")
        non_space = max(1, len(re.sub(r"\s", "", text)))
        replacement_ratio = replacement_count / non_space
        length_factor = min(1.0, len(text) / 800)
        confidence = max(0.15, min(0.58, 0.35 + 0.25 * length_factor))
        confidence *= max(0.25, 1.0 - replacement_ratio * 8)
        issue = (
            "Использован существующий текстовый слой PDF. "
            "Результат не подтверждается автоматически."
        )
        if extra_issue:
            issue = f"{issue} {extra_issue}"
        return OcrResult(
            status=OcrStatus.EMBEDDED_TEXT,
            text=text,
            confidence=round(confidence, 3),
            language="Unicode из PDF",
            issue=issue,
        )

    @staticmethod
    def _combine_text_sources(embedded_text: str, image_text: str) -> str:
        return (
            "=== Текстовый слой PDF ===\n"
            f"{embedded_text.strip()}\n\n"
            "=== OCR изображения всей страницы ===\n"
            f"{image_text.strip()}"
        ).strip()

    def recognize_type_markers(self, image_path: Path) -> str:
        """Read only the document header for classification, trying rotations."""
        if not self._model_available(self.fast_data_dir):
            return ""
        try:
            with Image.open(image_path) as source:
                base = ImageOps.exif_transpose(source).convert("L")
        except OSError:
            return ""

        best_text = ""
        best_key = (0, 0, 0)
        marker_weights = {
            "решение": 5,
            "audit sti-010": 5,
            "sti-010": 3,
            "проводимых на счетах организаций": 3,
            "принято решение": 3,
            "раздел i": 2,
        }
        for degrees in (0, 180, 90, 270):
            oriented = base.rotate(degrees, expand=True)
            width, height = oriented.size
            header = oriented.crop((0, 0, width, int(height * 0.32)))
            header = ImageOps.autocontrast(header, cutoff=1)
            header = header.resize((header.width * 2, header.height * 2))
            try:
                with tesserocr.PyTessBaseAPI(
                    path=str(self.fast_data_dir),
                    lang="rus+kir",
                    psm=tesserocr.PSM.SINGLE_BLOCK,
                ) as api:
                    api.SetVariable("preserve_interword_spaces", "1")
                    api.SetImage(header)
                    text = unicodedata.normalize(
                        "NFC", api.GetUTF8Text() or ""
                    ).strip()
                    confidence = max(0, min(100, api.MeanTextConf()))
            except Exception:
                continue
            normalized = re.sub(r"\s+", " ", text.casefold())
            marker_score = sum(
                weight
                for marker, weight in marker_weights.items()
                if marker in normalized
            )
            key = (marker_score, confidence, len(text))
            if key > best_key:
                best_key = key
                best_text = text

        return best_text if best_key[0] >= 3 else ""

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
        return tuple(
            sorted(
                {
                    *(f"name:{value}" for value in names),
                    *(f"inn:{value}" for value in inns),
                    *(f"period:{value}" for value in periods),
                }
            )
        )

    @staticmethod
    def _recipient_field_signature(text: str) -> tuple[str, ...]:
        name_word = r"[А-ЯЁҢӨҮ][А-Яа-яЁёҢңӨөҮү\-]+"
        matches = re.finditer(
            r"(?P<position>(?:зам\.?[ \t]+)?начальник[ау]?[ \t]+"
            r"управления)(?:[ \t]+|(?:\r?\n[ \t]*){1,3})(?P<name>"
            + name_word
            + r"(?:[ \t]+"
            + name_word
            + r"){1,3})[ \t]*(?=\r?$)",
            unicodedata.normalize("NFC", text),
            re.IGNORECASE | re.MULTILINE,
        )
        signatures = []
        for match in matches:
            position = re.sub(
                r"\s+", " ", match.group("position")
            ).strip().casefold()
            name = re.sub(r"\s+", " ", match.group("name")).strip().casefold()
            signatures.append(f"recipient:{position}|{name}")
        return tuple(sorted(set(signatures)))

    @staticmethod
    def _recognize_image(
        image_path: Path,
        data_dir: Path,
        model_name: str,
        supplement_regions: bool = False,
    ) -> OcrResult:
        # Цвет сохраняется: он помогает Tesseract отделить чёрный печатный
        # текст от синих печатей и подписей. Для чёрно-белых сканов поведение
        # остаётся тем же.
        with Image.open(image_path) as source:
            source_image = ImageOps.exif_transpose(source).convert("RGB")
        image = ImageOps.autocontrast(source_image, cutoff=1).filter(
            ImageFilter.UnsharpMask(radius=1.1, percent=115, threshold=4)
        )

        width, height = image.size
        # Полосы идут до общего OCR: нижняя подпись и другие локальные области
        # не наследуют ошибочную адаптацию шрифта от всей сложной страницы.
        bands = (
            (0.78, 0.90),
            (0.78, 1.00),
            (0.60, 0.84),
            (0.40, 0.64),
            (0.20, 0.44),
            (0.00, 0.24),
        )
        # На письмах синяя печать или подпись часто проходит поверх чёрного
        # печатного ФИО. Для дополнительных областей удаляем только явно
        # цветные пиксели. Это не дорисовывает текст: основной OCR по
        # оригиналу остаётся в результате, а маска лишь убирает помеху.
        region_source = (
            OcrService._suppress_colored_artifacts(source_image)
            if width >= 1500
            else image
        )
        pass_texts: list[str] = []
        confidences: list[float] = []
        for top_ratio, bottom_ratio in bands:
            band = region_source.crop(
                (
                    0,
                    int(height * top_ratio),
                    width,
                    int(height * bottom_ratio),
                )
            )
            if width < 1500:
                band = band.resize(
                    (
                        max(1, int(band.width * 1.35)),
                        max(1, int(band.height * 1.35)),
                    )
                )
            with tesserocr.PyTessBaseAPI(
                path=str(data_dir),
                lang="rus+kir",
                psm=tesserocr.PSM.SPARSE_TEXT,
            ) as region_api:
                region_api.SetVariable("preserve_interword_spaces", "1")
                region_api.SetImage(band)
                pass_texts.append(
                    unicodedata.normalize(
                        "NFC", region_api.GetUTF8Text() or ""
                    )
                )
                confidences.append(
                    max(0, min(100, region_api.MeanTextConf()))
                )

        if width >= 1500:
            adaptive_source = OcrService._prepare_high_resolution_regions(
                source_image
            )
            adaptive_band = adaptive_source.crop(
                (0, int(height * 0.48), width, int(height * 0.82))
            )
            with tesserocr.PyTessBaseAPI(
                path=str(data_dir),
                lang="rus+kir",
                psm=tesserocr.PSM.SPARSE_TEXT,
            ) as adaptive_api:
                adaptive_api.SetVariable("preserve_interword_spaces", "1")
                adaptive_api.SetImage(adaptive_band)
                pass_texts.append(
                    unicodedata.normalize(
                        "NFC", adaptive_api.GetUTF8Text() or ""
                    )
                )
                confidences.append(
                    max(0, min(100, adaptive_api.MeanTextConf()))
                )

        with tesserocr.PyTessBaseAPI(
            path=str(data_dir),
            lang="rus+kir",
            psm=tesserocr.PSM.AUTO,
        ) as api:
            api.SetVariable("preserve_interword_spaces", "1")
            api.SetImage(image)
            full_text = unicodedata.normalize(
                "NFC", api.GetUTF8Text() or ""
            )
            confidences.append(max(0, min(100, api.MeanTextConf())))

        # Для чтения и извлечения сохраняем логический полный проход первым,
        # затем независимые области сверху вниз. Между областями нет
        # перемешивания строк.
        text = OcrService._merge_ocr_passes(
            [full_text, *reversed(pass_texts)]
        )
        mean_confidence = sum(confidences) / len(confidences)

        cleaned = OcrService._clean_ocr_text(text)
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

    @staticmethod
    def _suppress_colored_artifacts(image: Image.Image) -> Image.Image:
        """Whiten clearly chromatic ink while preserving black print."""
        pixels = np.asarray(image.convert("RGB"), dtype=np.uint8).copy()
        red = pixels[:, :, 0].astype(np.int16)
        green = pixels[:, :, 1].astype(np.int16)
        blue = pixels[:, :, 2].astype(np.int16)
        blue_ink = (blue - red > 20) & (blue - green > 5)
        red_ink = (red - blue > 24) & (red - green > 8)
        pixels[blue_ink | red_ink] = 255
        return Image.fromarray(pixels, mode="RGB")

    @staticmethod
    def _prepare_high_resolution_regions(image: Image.Image) -> Image.Image:
        """Remove colored stamps and normalize uneven scan illumination."""
        cleaned = OcrService._suppress_colored_artifacts(image).convert("L")
        grayscale = np.asarray(
            ImageOps.autocontrast(cleaned, cutoff=1), dtype=np.uint8
        )
        binary = cv2.adaptiveThreshold(
            grayscale,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            41,
            13,
        )
        return Image.fromarray(binary, mode="L")

    @staticmethod
    def _clean_ocr_text(text: str) -> str:
        return "\n".join(
            line.rstrip() for line in text.splitlines()
        ).strip()

    @staticmethod
    def _merge_ocr_passes(pass_texts: list[str]) -> str:
        blocks = [OcrService._clean_ocr_text(text) for text in pass_texts]
        blocks = [block for block in blocks if block]
        if not blocks:
            return ""
        parts = [blocks[0]]
        for index, block in enumerate(blocks[1:], 1):
            parts.append(
                f"=== Дополнительная OCR-область {index} ===\n{block}"
            )
        return "\n\n".join(parts)
