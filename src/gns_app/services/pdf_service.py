from __future__ import annotations

import argparse
import json
import math
import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from PIL import Image, ImageFilter, ImageOps, ImageStat
from pypdf import PdfReader, PdfWriter

from gns_app.runtime_commands import module_command


class PdfProcessingError(ValueError):
    pass


@dataclass(frozen=True, slots=True)
class RenderedPage:
    preview_path: Path
    enhanced_preview_path: Path
    quality_score: float


class PdfService:
    RENDER_TIMEOUT_SECONDS = 180

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

    def image_to_pdf(self, image_path: Path, output_path: Path) -> Path:
        """Create a one-page working PDF while preserving the source image."""
        output_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = output_path.with_name(f".{output_path.name}.{uuid4().hex}.part")
        try:
            with Image.open(image_path) as source:
                image = ImageOps.exif_transpose(source)
                if image.mode in {"RGBA", "LA"} or (
                    image.mode == "P" and "transparency" in image.info
                ):
                    rgba = image.convert("RGBA")
                    flattened = Image.new("RGB", rgba.size, "white")
                    flattened.paste(rgba, mask=rgba.getchannel("A"))
                    image = flattened
                else:
                    image = image.convert("RGB")
                image.save(
                    temporary,
                    format="PDF",
                    resolution=200.0,
                    quality=95,
                    subsampling=0,
                )
            self.page_count(temporary)
            os.replace(temporary, output_path)
            return output_path
        except PdfProcessingError:
            temporary.unlink(missing_ok=True)
            raise
        except Exception as exc:
            temporary.unlink(missing_ok=True)
            raise PdfProcessingError(
                f"Не удалось подготовить изображение для обработки: {exc}"
            ) from exc

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
        payload = self._run_isolated_renderer(
            action="preview",
            pdf_path=pdf_path,
            page_number=page_number,
            output_path=preview_path,
            secondary_output_path=enhanced_path,
            scale=scale,
        )

        return RenderedPage(
            preview_path=preview_path,
            enhanced_preview_path=enhanced_path,
            quality_score=float(payload.get("quality_score", 0.0)),
        )

    def render_page_for_qr(
        self,
        pdf_path: Path,
        page_number: int,
        output_path: Path,
        scale: float = 3.2,
    ) -> Path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        self._run_isolated_renderer(
            action="qr",
            pdf_path=pdf_path,
            page_number=page_number,
            output_path=output_path,
            scale=scale,
        )
        return output_path

    def render_page_high_resolution(
        self,
        pdf_path: Path,
        page_number: int,
        output_path: Path,
        scale: float = 3.2,
    ) -> Path:
        """Render one color source shared by high-resolution QR and OCR."""
        output_path.parent.mkdir(parents=True, exist_ok=True)
        self._run_isolated_renderer(
            action="high_resolution",
            pdf_path=pdf_path,
            page_number=page_number,
            output_path=output_path,
            scale=scale,
        )
        return output_path

    def _run_isolated_renderer(
        self,
        *,
        action: str,
        pdf_path: Path,
        page_number: int,
        output_path: Path,
        scale: float,
        secondary_output_path: Path | None = None,
    ) -> dict[str, object]:
        command = module_command(
            "gns_app.services.pdf_service",
            "--render-action",
            action,
            "--pdf-path",
            str(pdf_path),
            "--page-number",
            str(page_number),
            "--output-path",
            str(output_path),
            "--scale",
            str(scale),
        )
        if secondary_output_path is not None:
            command.extend(
                ["--secondary-output-path", str(secondary_output_path)]
            )
        outputs = [output_path]
        if secondary_output_path is not None:
            outputs.append(secondary_output_path)
        for path in outputs:
            path.unlink(missing_ok=True)
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=self.RENDER_TIMEOUT_SECONDS,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                cwd=Path.cwd(),
                env={**os.environ, "PYTHONIOENCODING": "utf-8"},
            )
        except subprocess.TimeoutExpired as exc:
            for path in outputs:
                path.unlink(missing_ok=True)
            raise PdfProcessingError(
                f"Отрисовка страницы {page_number} превысила лимит времени"
            ) from exc
        if completed.returncode != 0:
            for path in outputs:
                path.unlink(missing_ok=True)
            raise PdfProcessingError(
                f"Изолированный модуль PDF аварийно завершился на странице "
                f"{page_number}; сервер продолжает работу"
            )
        try:
            payload = json.loads(completed.stdout)
        except json.JSONDecodeError as exc:
            for path in outputs:
                path.unlink(missing_ok=True)
            raise PdfProcessingError(
                f"Модуль PDF не вернул результат для страницы {page_number}"
            ) from exc
        if not isinstance(payload, dict) or not payload.get("ok"):
            for path in outputs:
                path.unlink(missing_ok=True)
            raise PdfProcessingError(
                f"Не удалось отрисовать страницу {page_number}"
            )
        if not all(path.is_file() for path in outputs):
            for path in outputs:
                path.unlink(missing_ok=True)
            raise PdfProcessingError(
                f"Модуль PDF не создал изображение страницы {page_number}"
            )
        return payload

    def render_page_for_ocr(
        self,
        pdf_path: Path,
        page_number: int,
        output_path: Path,
        scale: float = 3.0,
    ) -> Path:
        """Compatibility wrapper for callers that only need OCR rendering."""
        return self.render_page_high_resolution(
            pdf_path,
            page_number,
            output_path,
            scale=scale,
        )

    def extract_embedded_text(self, pdf_path: Path, page_number: int) -> str:
        try:
            reader = PdfReader(str(pdf_path))
            return reader.pages[page_number - 1].extract_text() or ""
        except Exception:
            return ""

    def extract_page_pdf(
        self,
        pdf_path: Path,
        page_number: int,
        output_path: Path,
    ) -> Path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = output_path.with_name(
            f".{output_path.name}.{uuid4().hex}.part"
        )
        try:
            reader = PdfReader(str(pdf_path))
            if reader.is_encrypted:
                raise PdfProcessingError(
                    "Зашифрованные PDF пока не поддерживаются"
                )
            if page_number < 1 or page_number > len(reader.pages):
                raise PdfProcessingError(
                    f"Страница {page_number} отсутствует в PDF"
                )
            writer = PdfWriter()
            writer.add_page(reader.pages[page_number - 1])
            with temporary.open("wb") as stream:
                writer.write(stream)
            temporary.replace(output_path)
            return output_path
        except PdfProcessingError:
            temporary.unlink(missing_ok=True)
            raise
        except Exception as exc:
            temporary.unlink(missing_ok=True)
            raise PdfProcessingError(
                f"Не удалось подготовить PDF страницы {page_number}: {exc}"
            ) from exc

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


def _render_pdfium_image(
    pdf_path: Path,
    page_number: int,
    scale: float,
    mode: str,
) -> Image.Image:
    # PDFium загружается только в дочернем процессе. Его нативный сбой не
    # может завершить локальный веб-сервер.
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        page = pdf[page_number - 1]
        try:
            bitmap = page.render(scale=scale)
            try:
                return bitmap.to_pil().convert(mode).copy()
            finally:
                bitmap.close()
        finally:
            page.close()
    finally:
        pdf.close()


def _render_bridge(args: argparse.Namespace) -> dict[str, object]:
    try:
        pdf_path = Path(args.pdf_path)
        output_path = Path(args.output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if args.render_action == "preview":
            if not args.secondary_output_path:
                raise PdfProcessingError("Не указан путь улучшенного превью")
            secondary = Path(args.secondary_output_path)
            secondary.parent.mkdir(parents=True, exist_ok=True)
            image = _render_pdfium_image(
                pdf_path,
                args.page_number,
                args.scale,
                "L",
            )
            image.save(output_path, format="JPEG", quality=90, subsampling=0)
            enhanced = PdfService._enhance_for_review(image)
            enhanced.save(
                secondary,
                format="JPEG",
                quality=92,
                subsampling=0,
            )
            return {
                "ok": True,
                "quality_score": PdfService._quality_score(image),
            }
        mode = "L" if args.render_action == "qr" else "RGB"
        image = _render_pdfium_image(
            pdf_path,
            args.page_number,
            args.scale,
            mode,
        )
        image.save(output_path, format="PNG", optimize=True)
        return {"ok": True}
    except Exception:
        return {"ok": False}


def _run_render_bridge() -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument(
        "--render-action",
        choices=("preview", "qr", "high_resolution"),
        required=True,
    )
    parser.add_argument("--pdf-path", required=True)
    parser.add_argument("--page-number", type=int, required=True)
    parser.add_argument("--output-path", required=True)
    parser.add_argument("--secondary-output-path", default="")
    parser.add_argument("--scale", type=float, required=True)
    args = parser.parse_args()
    sys.stdout.write(json.dumps(_render_bridge(args), ensure_ascii=False))
    return 0


def main() -> int:
    return _run_render_bridge()


if __name__ == "__main__":
    raise SystemExit(main())
