from __future__ import annotations

import hashlib
from pathlib import Path
from urllib.parse import urlparse

import zxingcpp
from PIL import Image, ImageEnhance, ImageOps

from gns_app.domain import QrDecodeResult, QrStatus


class QrService:
    def __init__(
        self,
        allowed_hosts: frozenset[str],
        allowed_paths: frozenset[str],
    ):
        self.allowed_hosts = allowed_hosts
        self.allowed_paths = allowed_paths

    def decode(self, image_path: Path) -> QrDecodeResult:
        image = Image.open(image_path).convert("L")

        attempts: list[tuple[str, Image.Image, zxingcpp.Binarizer]] = [
            ("original", image, zxingcpp.Binarizer.LocalAverage),
            (
                "autocontrast",
                ImageOps.autocontrast(image, cutoff=1),
                zxingcpp.Binarizer.LocalAverage,
            ),
        ]

        for method, candidate, binarizer in attempts:
            result = self._try_decode(candidate, method, binarizer)
            if result:
                return self._validate(result[0], result[1])

        width, height = image.size
        crop = image.crop(
            (
                int(width * 0.60),
                int(height * 0.52),
                width,
                int(height * 0.98),
            )
        )
        crop_variants: list[tuple[str, Image.Image]] = [
            ("crop", crop),
            ("crop-autocontrast", ImageOps.autocontrast(crop, cutoff=1)),
            ("crop-contrast", ImageEnhance.Contrast(crop).enhance(1.8)),
        ]

        for name, variant in crop_variants:
            for scale in (1, 2, 3):
                scaled = (
                    variant
                    if scale == 1
                    else variant.resize(
                        (variant.width * scale, variant.height * scale),
                        Image.Resampling.LANCZOS,
                    )
                )
                for binarizer in (
                    zxingcpp.Binarizer.LocalAverage,
                    zxingcpp.Binarizer.GlobalHistogram,
                    zxingcpp.Binarizer.FixedThreshold,
                ):
                    method = f"{name}-x{scale}-{binarizer.name}"
                    result = self._try_decode(scaled, method, binarizer)
                    if result:
                        return self._validate(result[0], result[1])

            for threshold in (110, 140, 170, 200, 220):
                binary = variant.point(
                    lambda pixel, limit=threshold: 255 if pixel > limit else 0
                ).resize(
                    (variant.width * 2, variant.height * 2),
                    Image.Resampling.NEAREST,
                )
                result = self._try_decode(
                    binary,
                    f"{name}-threshold-{threshold}",
                    zxingcpp.Binarizer.BoolCast,
                )
                if result:
                    return self._validate(result[0], result[1])

            for angle in (-3, -2, -1, 1, 2, 3):
                rotated = variant.rotate(
                    angle,
                    Image.Resampling.BICUBIC,
                    expand=True,
                    fillcolor=255,
                )
                result = self._try_decode(
                    rotated,
                    f"{name}-rotate-{angle}",
                    zxingcpp.Binarizer.GlobalHistogram,
                )
                if result:
                    return self._validate(result[0], result[1])

        return QrDecodeResult(
            status=QrStatus.NOT_FOUND,
            issue="QR не найден или повреждён. Требуется OCR/ручная проверка.",
        )

    @staticmethod
    def _try_decode(
        image: Image.Image,
        method: str,
        binarizer: zxingcpp.Binarizer,
    ) -> tuple[str, str] | None:
        results = zxingcpp.read_barcodes(
            image,
            formats=zxingcpp.BarcodeFormat.QRCode,
            try_rotate=True,
            try_downscale=True,
            try_invert=True,
            binarizer=binarizer,
            return_errors=False,
        )
        payloads = {item.text for item in results if item.text}
        if len(payloads) == 1:
            return payloads.pop(), method
        if len(payloads) > 1:
            # Несколько разных QR нельзя выбирать по догадке.
            return "", "multiple-values"
        return None

    def _validate(self, payload: str, method: str) -> QrDecodeResult:
        if not payload:
            return QrDecodeResult(
                status=QrStatus.DECODE_ERROR,
                method=method,
                issue="На странице обнаружены разные QR. Нужна ручная проверка.",
            )

        parsed = urlparse(payload)
        host = (parsed.hostname or "").lower()
        if (
            parsed.scheme != "https"
            or host not in self.allowed_hosts
            or parsed.path not in self.allowed_paths
        ):
            return QrDecodeResult(
                status=QrStatus.INVALID_URL,
                method=method,
                issue="QR ведёт на адрес вне разрешённого списка.",
            )

        payload_hash = hashlib.sha256(payload.encode("utf-8")).hexdigest()
        safe_url = f"{parsed.scheme}://{host}{parsed.path}"
        return QrDecodeResult(
            status=QrStatus.FOUND,
            payload=payload,
            payload_hash=payload_hash,
            safe_url=safe_url,
            method=method,
        )

