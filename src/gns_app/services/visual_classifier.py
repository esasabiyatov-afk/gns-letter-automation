from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image, ImageOps

from gns_app.domain import VisualPageEvidence


class PageVisualAnalyzer:
    """Detect the dense ruled layout of STI-010 without reading or inventing text."""

    @staticmethod
    def _group_count(mask: np.ndarray, max_gap: int = 2) -> int:
        indexes = np.flatnonzero(mask)
        if not len(indexes):
            return 0
        return 1 + int(np.sum(np.diff(indexes) > max_gap))

    @staticmethod
    def _adaptive_dark_map(
        pixels: np.ndarray,
        window: int = 41,
        delta: float = 12.0,
    ) -> np.ndarray:
        padding = window // 2
        padded = np.pad(
            pixels.astype(np.float32), padding, mode="reflect"
        )
        integral = np.pad(
            padded, ((1, 0), (1, 0)), constant_values=0
        ).cumsum(axis=0).cumsum(axis=1)
        local_mean = (
            integral[window:, window:]
            - integral[:-window, window:]
            - integral[window:, :-window]
            + integral[:-window, :-window]
        ) / (window * window)
        return pixels < (local_mean - delta)

    def analyze(self, image_path: Path) -> VisualPageEvidence:
        try:
            with Image.open(image_path) as source:
                image = ImageOps.autocontrast(source.convert("L"), cutoff=1)
                image.thumbnail((800, 1200))
                pixels = np.asarray(image)
        except (OSError, ValueError):
            return VisualPageEvidence(
                decision_layout=False,
                confidence=0.0,
                reasons=["Не удалось проанализировать структуру изображения"],
            )

        dark = self._adaptive_dark_map(pixels)
        height, width = dark.shape
        region = dark[
            int(height * 0.15) : int(height * 0.96),
            int(width * 0.03) : int(width * 0.97),
        ]
        if min(region.shape) < 50:
            return VisualPageEvidence(
                decision_layout=False,
                confidence=0.0,
                reasons=["Изображение слишком мало для анализа структуры"],
            )

        horizontal_window = max(20, int(region.shape[1] * 0.42))
        horizontal_sum = np.pad(
            np.cumsum(region, axis=1), ((0, 0), (1, 0))
        )
        horizontal_density = (
            horizontal_sum[:, horizontal_window:]
            - horizontal_sum[:, :-horizontal_window]
        ).max(axis=1) / horizontal_window

        vertical_window = max(20, int(region.shape[0] * 0.10))
        vertical_sum = np.pad(
            np.cumsum(region, axis=0), ((1, 0), (0, 0))
        )
        vertical_density = (
            vertical_sum[vertical_window:]
            - vertical_sum[:-vertical_window]
        ).max(axis=0) / vertical_window

        horizontal_groups = self._group_count(horizontal_density > 0.80)
        vertical_groups = self._group_count(vertical_density > 0.80)
        ink_density = float(region.mean())
        decision_layout = bool(
            horizontal_groups >= 8
            and vertical_groups >= 2
            and 0.06 <= ink_density <= 0.32
        )
        if decision_layout:
            confidence = min(
                0.98,
                0.80
                + min(horizontal_groups, 20) / 250
                + min(vertical_groups, 15) / 200,
            )
            reasons = [
                "Найдена плотная табличная структура формы STI-010",
                (
                    f"Группы линий: {horizontal_groups} по горизонтали, "
                    f"{vertical_groups} по вертикали"
                ),
            ]
        else:
            confidence = 0.0
            reasons = ["Табличная структура STI-010 не подтверждена"]

        return VisualPageEvidence(
            decision_layout=decision_layout,
            confidence=round(confidence, 3),
            horizontal_line_groups=horizontal_groups,
            vertical_line_groups=vertical_groups,
            ink_density=round(ink_density, 4),
            reasons=reasons,
        )
