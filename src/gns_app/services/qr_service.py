from __future__ import annotations

import hashlib
from pathlib import Path
from urllib.parse import urlparse

import zxingcpp
from PIL import Image, ImageEnhance, ImageOps
from pyzbar.pyzbar import ZBarSymbol, decode as zbar_decode

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
        deferred_result: QrDecodeResult | None = None
        zxing_candidate: tuple[str, str] | None = None

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
                zxing_candidate = result
                break

        # ZBar uses a different grid sampler and is materially more tolerant of
        # low-resolution scan jitter.  It receives the untouched grayscale
        # image: no reconstructed or generated pixels are introduced.
        zbar_candidate = self._try_decode_zbar(image, "zbar-original")
        validated = self._validate_decoder_candidates(
            zxing_candidate,
            zbar_candidate,
        )
        if validated:
            if validated.status in {QrStatus.FOUND, QrStatus.DECODE_ERROR}:
                return validated
            deferred_result = validated

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
                        validated = self._validate(result[0], result[1])
                        if validated.status == QrStatus.FOUND:
                            return validated
                        deferred_result = deferred_result or validated

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
                    validated = self._validate(result[0], result[1])
                    if validated.status == QrStatus.FOUND:
                        return validated
                    deferred_result = deferred_result or validated

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
                    validated = self._validate(result[0], result[1])
                    if validated.status == QrStatus.FOUND:
                        return validated
                    deferred_result = deferred_result or validated

        recovery_profiles = (
            (
                "bottom-right",
                (0.62, 0.62, 0.99, 0.96),
                (
                    (5.95, zxingcpp.Binarizer.LocalAverage),
                    (5.00, zxingcpp.Binarizer.FixedThreshold),
                    (3.95, zxingcpp.Binarizer.FixedThreshold),
                ),
            ),
            (
                "middle-right",
                (0.76, 0.58, 0.98, 0.84),
                (
                    (3.05, zxingcpp.Binarizer.FixedThreshold),
                    (6.05, zxingcpp.Binarizer.GlobalHistogram),
                ),
            ),
        )
        for profile_name, bounds, variants in recovery_profiles:
            x1, y1, x2, y2 = bounds
            focused = image.crop(
                (
                    int(width * x1),
                    int(height * y1),
                    int(width * x2),
                    int(height * y2),
                )
            )
            for factor, binarizer in variants:
                scaled = focused.resize(
                    (
                        int(focused.width * factor),
                        int(focused.height * factor),
                    ),
                    Image.Resampling.LANCZOS,
                )
                method = (
                    f"recovery-{profile_name}-x{factor:g}-{binarizer.name}"
                )
                result = self._try_decode(scaled, method, binarizer)
                if result:
                    validated = self._validate(result[0], result[1])
                    if validated.status == QrStatus.FOUND:
                        return validated
                    deferred_result = deferred_result or validated

        return deferred_result or QrDecodeResult(
            status=QrStatus.NOT_FOUND,
            issue="QR не найден или повреждён. Требуется OCR/ручная проверка.",
        )

    def decode_high_resolution(self, image_path: Path) -> QrDecodeResult:
        """A bounded retry for a genuine high-resolution PDF render."""
        image = Image.open(image_path).convert("L")
        deferred_result: QrDecodeResult | None = None
        zxing_candidate = self._try_decode(
            image,
            "highres-original",
            zxingcpp.Binarizer.LocalAverage,
        )
        zbar_candidate = self._try_decode_zbar(
            image,
            "zbar-highres-original",
        )
        validated = self._validate_decoder_candidates(
            zxing_candidate,
            zbar_candidate,
        )
        if validated:
            if validated.status in {QrStatus.FOUND, QrStatus.DECODE_ERROR}:
                return validated
            deferred_result = validated
        width, height = image.size
        crop = image.crop(
            (
                int(width * 0.60),
                int(height * 0.52),
                width,
                int(height * 0.98),
            )
        )
        variants = (
            ("highres-crop", crop),
            (
                "highres-crop-autocontrast",
                ImageOps.autocontrast(crop, cutoff=1),
            ),
            (
                "highres-crop-contrast",
                ImageEnhance.Contrast(crop).enhance(1.8),
            ),
        )
        for name, variant in variants:
            for binarizer in (
                zxingcpp.Binarizer.LocalAverage,
                zxingcpp.Binarizer.GlobalHistogram,
                zxingcpp.Binarizer.FixedThreshold,
            ):
                result = self._try_decode(variant, name, binarizer)
                if result:
                    validated = self._validate(result[0], result[1])
                    if validated.status == QrStatus.FOUND:
                        return validated
                    deferred_result = deferred_result or validated
            for threshold in (110, 140, 170, 200, 220):
                binary = variant.point(
                    lambda pixel, limit=threshold: (
                        255 if pixel > limit else 0
                    )
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
                    validated = self._validate(result[0], result[1])
                    if validated.status == QrStatus.FOUND:
                        return validated
                    deferred_result = deferred_result or validated
        return deferred_result or QrDecodeResult(
            status=QrStatus.NOT_FOUND,
            issue="QR не найден и на повышенном разрешении.",
        )

    def _try_decode(
        self,
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
            return_errors=True,
        )
        payloads = {
            item.text.strip()
            for item in results
            if item.text and item.valid and item.error is None
        }
        return self._select_payload(payloads, method, "multiple-values")

    def _validate_decoder_candidates(
        self,
        zxing_candidate: tuple[str, str] | None,
        zbar_candidate: tuple[str, str] | None,
    ) -> QrDecodeResult | None:
        if zxing_candidate and zbar_candidate:
            if not zxing_candidate[0] or not zbar_candidate[0]:
                return QrDecodeResult(
                    status=QrStatus.DECODE_ERROR,
                    method="zxing-zbar-conflict",
                    issue=(
                        "QR-декодеры получили разные значения. "
                        "Требуется ручная проверка."
                    ),
                )
            zxing_result = self._validate(*zxing_candidate)
            zbar_result = self._validate(*zbar_candidate)
            if zxing_result.status == QrStatus.FOUND:
                if zbar_result.status != QrStatus.FOUND:
                    return zxing_result
                if zxing_candidate[0] == zbar_candidate[0]:
                    return self._validate(
                        zxing_candidate[0],
                        f"{zxing_candidate[1]}+{zbar_candidate[1]}",
                    )
            elif zbar_result.status == QrStatus.FOUND:
                return zbar_result
            elif zxing_candidate[0] == zbar_candidate[0]:
                return zxing_result
            return QrDecodeResult(
                status=QrStatus.DECODE_ERROR,
                method="zxing-zbar-conflict",
                issue=(
                    "QR-декодеры получили разные значения. "
                    "Требуется ручная проверка."
                ),
            )
        candidate = zxing_candidate or zbar_candidate
        if candidate:
            return self._validate(candidate[0], candidate[1])
        return None

    def _try_decode_zbar(
        self,
        image: Image.Image,
        method: str,
    ) -> tuple[str, str] | None:
        payloads: set[str] = set()
        for symbol in zbar_decode(image, symbols=[ZBarSymbol.QRCODE]):
            try:
                payload = symbol.data.decode("utf-8", errors="strict").strip()
            except UnicodeDecodeError:
                continue
            if payload:
                payloads.add(payload)
        return self._select_payload(
            payloads,
            method,
            "multiple-values-zbar",
        )

    def _select_payload(
        self,
        payloads: set[str],
        method: str,
        multiple_method: str,
    ) -> tuple[str, str] | None:
        allowed = {payload for payload in payloads if self._is_allowed(payload)}
        if len(allowed) == 1:
            return allowed.pop(), method
        if len(allowed) > 1:
            # Несколько разных разрешённых QR нельзя выбирать по догадке.
            return "", multiple_method
        if len(payloads) == 1:
            return payloads.pop(), method
        if len(payloads) > 1:
            return "", multiple_method
        return None

    def _is_allowed(self, payload: str) -> bool:
        parsed = urlparse(payload)
        return (
            parsed.scheme == "https"
            and (parsed.hostname or "").lower() in self.allowed_hosts
            and parsed.path in self.allowed_paths
        )

    def _validate(self, payload: str, method: str) -> QrDecodeResult:
        if not payload:
            return QrDecodeResult(
                status=QrStatus.DECODE_ERROR,
                method=method,
                issue="На странице обнаружены разные QR. Нужна ручная проверка.",
            )

        parsed = urlparse(payload)
        host = (parsed.hostname or "").lower()
        if not self._is_allowed(payload):
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
