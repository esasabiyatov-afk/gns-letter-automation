from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageFilter, ImageOps, ImageStat
from pypdf import PdfReader


class PdfProcessingError(ValueError):
    pass


@dataclass(frozen=True, slots=True)
class RenderedPage:
    preview_path: Path
    enhanced_preview_path: Path
    quality_score: float


class PdfService:
    def page_count(self, path: Path) -> int:
        try:
            reader = PdfReader(str(path))
            if reader.is_encrypted:
                raise PdfProcessingError("Зашифрованные PDF пока не поддерживаются")
            count = len(reader.pages)
        except PdfProcessingError:
            raise
        except Exception as exc:
            raise PdfProcessingError(f"Не удалось прочитать PDF: {exc}") from exc

        if count < 1:
            raise PdfProcessingError("PDF не содержит страниц")
        return count

    def render_page(
        self,
        pdf_path: Path,
        page_number: int,
        preview_path: Path,
        enhanced_path: Path,
        scale: float = 1.6,
    ) -> RenderedPage:
        preview_path.parent.mkdir(parents=True, exist_ok=True)
        enhanced_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            pdf = pdfium.PdfDocument(str(pdf_path))
            try:
                page = pdf[page_number - 1]
                image = page.render(scale=scale).to_pil().convert("L")
            finally:
                pdf.close()
        except Exception as exc:
            raise PdfProcessingError(
                f"Не удалось отрисовать страницу {page_number}: {exc}"
            ) from exc

        image.save(preview_path, format="JPEG", quality=90, subsampling=0)
        enhanced = self._enhance_for_review(image)
        enhanced.save(enhanced_path, format="JPEG", quality=92, subsampling=0)

        return RenderedPage(
            preview_path=preview_path,
            enhanced_preview_path=enhanced_path,
            quality_score=self._quality_score(image),
        )

    def extract_embedded_text(self, pdf_path: Path, page_number: int) -> str:
        try:
            reader = PdfReader(str(pdf_path))
            return reader.pages[page_number - 1].extract_text() or ""
        except Exception:
            return ""

    @staticmethod
    def _enhance_for_review(image: Image.Image) -> Image.Image:
        # Только классические операции. Никакого генеративного восстановления.
        enhanced = ImageOps.autocontrast(image, cutoff=1)
        return enhanced.filter(
            ImageFilter.UnsharpMask(radius=1.4, percent=135, threshold=4)
        )

    @staticmethod
    def _quality_score(image: Image.Image) -> float:
        sample = image.copy()
        sample.thumbnail((900, 1200))
        edges = sample.filter(ImageFilter.FIND_EDGES)
        edge_std = ImageStat.Stat(edges).stddev[0]
        contrast_std = ImageStat.Stat(sample).stddev[0]
        raw = (edge_std / 38.0) * 0.65 + (contrast_std / 75.0) * 0.35
        if math.isnan(raw):
            return 0.0
        return round(max(0.0, min(1.0, raw)), 3)

