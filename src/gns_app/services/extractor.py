from __future__ import annotations

import re

from gns_app.domain import ExtractedFields, ExtractedTaxpayer


class FieldExtractor:
    PERIOD_RE = re.compile(
        r"период\s*[:;]?\s*с\s*"
        r"(?P<start>\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4})\s*"
        r"по\s*(?P<end>\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4})",
        re.IGNORECASE,
    )
    NAME_RE = re.compile(
        r"наименование\s*[:;]\s*(?P<name>.+?)(?=\s+инн\s*[:;])",
        re.IGNORECASE | re.DOTALL,
    )
    INN_RE = re.compile(r"инн\s*[:;]\s*(?P<inn>[\d\s]{14,28})", re.IGNORECASE)
    DISTRICT_RE = re.compile(
        r"управление\s+государственной\s+налоговой\s+службы\s+"
        r"(?P<district>по\s+.{5,180}?)\s+"
        r"(?:в\s+соответствии|на\s+основании)",
        re.IGNORECASE | re.DOTALL,
    )
    RECIPIENT_RE = re.compile(
        r"(?P<position>(?:зам\.?\s+)?начальник[ау]?\s+управления)\s+"
        r"(?P<name>[А-ЯЁҢӨҮ][А-Яа-яЁёҢңӨөҮү\-]+\s+"
        r"[А-ЯЁҢӨҮ][А-Яа-яЁёҢңӨөҮү\-]+\s+"
        r"[А-ЯЁҢӨҮ][А-Яа-яЁёҢңӨөҮү\-]+)",
        re.IGNORECASE,
    )

    def extract_scan_letter(self, text: str) -> ExtractedFields:
        result = ExtractedFields(
            confidence=0.0,
            issues=[
                "Поля извлечены из низкодоверенного текстового слоя скана.",
                "Перед проверкой АБС требуется подтверждение сотрудника.",
            ],
        )

        period_matches = list(self.PERIOD_RE.finditer(text))
        if period_matches:
            period = period_matches[-1]
            result.period_start = self._date_iso(period.group("start"))
            result.period_end = self._date_iso(period.group("end"))
        else:
            result.issues.append("Период не найден.")

        names = [self._clean(match.group("name")) for match in self.NAME_RE.finditer(text)]
        inns: list[str] = []
        for match in self.INN_RE.finditer(text):
            digits = re.sub(r"\D", "", match.group("inn"))
            if len(digits) == 14:
                inns.append(digits)

        # Заголовок ГНС также содержит ИНН. Берём только последнее совпадение
        # после маркированного наименования, но всё равно требуем подтверждение.
        if names and inns:
            result.taxpayers.append(
                ExtractedTaxpayer(
                    name=names[-1],
                    inn=inns[-1],
                    confidence=0.5,
                    issues=["Не подтверждено сотрудником."],
                )
            )
        else:
            result.issues.append("ИНН или наименование не найдены однозначно.")

        found_required = sum(
            (
                bool(result.period_start),
                bool(result.period_end),
                bool(result.taxpayers),
            )
        )
        result.confidence = round(found_required / 3 * 0.55, 3)
        return result

    def extract_official_letter(self, text: str) -> ExtractedFields:
        result = self.extract_scan_letter(text)
        result.issues = []

        district_matches = list(self.DISTRICT_RE.finditer(text))
        if district_matches:
            result.district_place = self._clean(
                district_matches[-1].group("district")
            )
        else:
            result.issues.append("Район и место не найдены.")

        recipient_matches = list(self.RECIPIENT_RE.finditer(text))
        if recipient_matches:
            recipient = recipient_matches[-1]
            result.recipient_position = self._clean(
                recipient.group("position")
            )
            result.recipient_full_name = self._clean(recipient.group("name"))
        else:
            result.issues.append("Должность или ФИО адресата не найдены.")

        if not result.period_start or not result.period_end:
            result.issues.append("Период не найден.")
        if not result.taxpayers:
            result.issues.append("Налогоплательщик не найден.")

        if not result.issues:
            result.confidence = 0.96
            for taxpayer in result.taxpayers:
                taxpayer.confidence = 0.96
                taxpayer.issues.clear()
        else:
            result.confidence = 0.0
        return result

    @staticmethod
    def _clean(value: str) -> str:
        return re.sub(r"\s+", " ", value).strip()

    @staticmethod
    def _date_iso(value: str) -> str | None:
        parts = re.split(r"[.\-/]", value)
        if len(parts) != 3:
            return None
        day, month, year = parts
        try:
            day_i, month_i, year_i = int(day), int(month), int(year)
            from datetime import date

            return date(year_i, month_i, day_i).isoformat()
        except ValueError:
            return None
