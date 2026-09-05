from __future__ import annotations

import re

from gns_app.domain import ExtractedFields, ExtractedTaxpayer
from gns_app.text_cleanup import clean_location


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
    # Длинные наименования (особенно ОсОО) часто переносятся визуально на
    # вторую строку без отдельной метки — просто продолжение текста без
    # двоеточия. Старая версия этой регулярки останавливалась на первом же
    # переводе строки и отрезала хвост названия (например, "Строй Компани"
    # вместо "Строй Компани Каракол"). Новая версия захватывает текст до
    # ближайшей следующей строки-метки (короткая строка с двоеточием —
    # "ИНН:", "Период:", повреждённые варианты вроде "ПИ:"/"АГИ:" и т.п.),
    # до следующего "Наименование:" (несколько налогоплательщиков подряд),
    # до пустой строки или до конца текста — в зависимости от того, что
    # встретится раньше.
    SCAN_NAME_RE = re.compile(
        r"наименование\s*[:;]\s*(?P<name>.+?)"
        r"(?=\r?\n[ \t]*[^\r\n]{0,40}[:;]|\r?\n\s*\r?\n|\Z)",
        re.IGNORECASE | re.DOTALL,
    )
    INN_RE = re.compile(r"инн\s*[:;]\s*(?P<inn>[\d\s]{14,28})", re.IGNORECASE)
    BARE_INN_RE = re.compile(r"(?<!\d)(?P<inn>(?:\d[\s-]*){14})(?!\d)")
    DISTRICT_RE = re.compile(
        r"управление\s+государственной\s+налоговой\s+службы\s+"
        r"(?P<district>по\s+.{5,180}?)\s+"
        r"(?:в\s+соответствии|на\s+основании)",
        re.IGNORECASE | re.DOTALL,
    )
    # ФИО должно начинаться с прописной буквы даже при re.IGNORECASE у
    # всего шаблона: иначе «отдела» ошибочно становится первой частью ФИО.
    NAME_WORD = r"(?-i:[А-ЯЁҢӨҮ][А-Яа-яЁёҢңӨөҮү\-]+)"
    RECIPIENT_NAME_PART = rf"(?:{NAME_WORD}|(?-i:уулу|кызы))"
    # В официальных QR-документах должность адресата не ограничивается
    # справочником: «начальник отдела», «заместитель директора» и другие
    # формулировки — такие же данные первоисточника. Берём строку перед ФИО,
    # но не поля с двоеточием (ИНН, Наименование и т. п.). Для OCR-скана
    # ниже остаётся более строгий шаблон, чтобы не принять шум за должность.
    OFFICIAL_RECIPIENT_RE = re.compile(
        r"^[ \t]*(?P<position>[^\r\n:]{3,100}?)"
        r"(?:[ \t]+|(?:\r?\n[ \t]*){1,3})(?P<name>"
        + RECIPIENT_NAME_PART
        + r"(?:[ \t]+"
        + RECIPIENT_NAME_PART
        + r"){1,3})[ \t]*(?=\r?$)",
        re.IGNORECASE | re.MULTILINE,
    )
    RECIPIENT_RE = re.compile(
        r"(?P<position>(?:зам\.?[ \t]+)?начальник[ау]?[ \t]+управления)"
        r"(?:[ \t]+|(?:\r?\n[ \t]*){1,3})(?P<name>"
        + RECIPIENT_NAME_PART
        + r"(?:[ \t]+"
        + RECIPIENT_NAME_PART
        + r"){1,3})[ \t]*(?=\r?$)",
        re.IGNORECASE | re.MULTILINE,
    )
    RECIPIENT_POSITION_RE = re.compile(
        r"(?:зам\.?[ \t]+)?начальник[ау]?[ \t]+управления",
        re.IGNORECASE,
    )
    RECIPIENT_NAME_LINE_RE = re.compile(
        r"(?m)^[ \t]*(?P<name>"
        + NAME_WORD
        + r"(?:[ \t]+"
        + NAME_WORD
        + r"){1,3})[ \t]*[.,]?[ \t]*$"
    )

    def extract_scan_letter(self, text: str) -> ExtractedFields:
        result = ExtractedFields(
            confidence=0.0,
            issues=[
                "Поля извлечены из низкодоверенного текстового слоя скана.",
                "Перед проверкой АБС требуется подтверждение сотрудника.",
            ],
        )

        district_matches = list(self.DISTRICT_RE.finditer(text))
        if district_matches:
            result.district_place = clean_location(
                district_matches[-1].group("district")
            )

        position_matches = list(self.RECIPIENT_POSITION_RE.finditer(text))
        if position_matches:
            result.recipient_position = self._clean(
                position_matches[-1].group(0)
            )

        recipient_matches = [
            match
            for match in self.RECIPIENT_RE.finditer(text)
            if self._plausible_recipient_name(match.group("name"))
        ]
        if recipient_matches:
            recipient = recipient_matches[-1]
            result.recipient_position = self._clean(
                recipient.group("position")
            )
            result.recipient_full_name = self._clean(recipient.group("name"))
        elif result.recipient_position:
            # Печать может вставить несколько коротких шумовых строк между
            # должностью и подписью. Берём только ближайшую полноценную строку
            # ФИО после явной должности в пределах той же OCR-секции.
            for section in re.split(r"(?m)^\s*=== .+ ===\s*$", text):
                for position_match in self.RECIPIENT_POSITION_RE.finditer(section):
                    tail = section[position_match.end() : position_match.end() + 240]
                    candidate = next(
                        (
                            match.group("name")
                            for match in self.RECIPIENT_NAME_LINE_RE.finditer(tail)
                            if self._plausible_recipient_name(match.group("name"))
                        ),
                        None,
                    )
                    if candidate:
                        result.recipient_full_name = self._clean(candidate)
                        break
                if result.recipient_full_name:
                    break

        period_matches = list(self.PERIOD_RE.finditer(text))
        if period_matches:
            period = period_matches[-1]
            result.period_start = self._date_iso(period.group("start"))
            result.period_end = self._date_iso(period.group("end"))
        else:
            result.issues.append("Период не найден.")

        taxpayer_candidates: list[tuple[str, str, bool]] = []
        # OCR-блоки могут быть с отступом (например, в тестах или после
        # объединения результатов). Не даём периоду одного прохода
        # подтверждать ИНН из соседнего прохода.
        sections = re.split(r"(?m)^\s*=== .+ ===\s*$", text)
        bare_inn_section_counts: dict[str, int] = {}
        for section in sections:
            visible_inns = {
                re.sub(r"\D", "", match.group("inn"))
                for match in self.BARE_INN_RE.finditer(section)
            }
            for visible_inn in visible_inns:
                if len(visible_inn) == 14:
                    bare_inn_section_counts[visible_inn] = (
                        bare_inn_section_counts.get(visible_inn, 0) + 1
                    )
        for section in sections:
            name_matches = list(self.SCAN_NAME_RE.finditer(section))
            for index, name_match in enumerate(name_matches):
                section_start = (
                    name_matches[index - 1].end()
                    if index > 0
                    else 0
                )
                section_end = (
                    name_matches[index + 1].start()
                    if index + 1 < len(name_matches)
                    else len(section)
                )
                window_start = max(section_start, name_match.start() - 600)
                window_end = min(section_end, name_match.end() + 600)
                context = section[window_start:window_end]
                name_center = (
                    name_match.start() + name_match.end()
                ) / 2 - window_start
                period_match = self.PERIOD_RE.search(context)
                valid_inns: list[tuple[float, str]] = []
                for inn_match in self.INN_RE.finditer(context):
                    digits = re.sub(r"\D", "", inn_match.group("inn"))
                    if len(digits) == 14:
                        center = (inn_match.start() + inn_match.end()) / 2
                        valid_inns.append((abs(center - name_center), digits))
                # На плохом скане сама метка «ИНН» часто превращается в
                # «ПИ», «НИ» или «АГИ». Если рядом уже есть наименование и
                # тот же блок заканчивается корректным периодом, можно взять
                # ровно 14 видимых цифр без догадки о повреждённой метке.
                if not valid_inns and (period_match or period_matches):
                    for inn_match in self.BARE_INN_RE.finditer(context):
                        digits = re.sub(r"\D", "", inn_match.group("inn"))
                        independently_repeated = (
                            bare_inn_section_counts.get(digits, 0) >= 2
                        )
                        if len(digits) == 14 and (
                            period_match or independently_repeated
                        ):
                            center = (inn_match.start() + inn_match.end()) / 2
                            valid_inns.append(
                                (abs(center - name_center), digits)
                            )
                if valid_inns:
                    valid_inns.sort(key=lambda item: item[0])
                    taxpayer_candidates.append(
                        (
                            self._clean(name_match.group("name")),
                            valid_inns[0][1],
                            period_match is not None,
                        )
                    )

        strong_names = {
            name.casefold()
            for name, _inn, has_period in taxpayer_candidates
            if has_period
        }
        taxpayer_pairs = [
            (name, inn)
            for name, inn, has_period in taxpayer_candidates
            if has_period or name.casefold() not in strong_names
        ]
        for name, inn in dict.fromkeys(taxpayer_pairs):
            result.taxpayers.append(
                ExtractedTaxpayer(
                    name=name,
                    inn=inn,
                    confidence=0.5,
                    issues=["Не подтверждено сотрудником."],
                )
            )
        if not result.taxpayers:
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
            result.district_place = clean_location(
                district_matches[-1].group("district")
            )
        else:
            result.issues.append("Район и место не найдены.")

        recipient_matches = [
            match
            for match in self.OFFICIAL_RECIPIENT_RE.finditer(text)
            if self._plausible_recipient_name(match.group("name"))
        ]
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
    def _plausible_recipient_name(value: str) -> bool:
        service_words = {
            "банк",
            "входящего",
            "документ",
            "исполнению",
            "количество",
            "контроль",
            "листов",
            "основной",
            "приложение",
            "роспись",
            "сведению",
        }
        raw_words = [
            word.strip(".,:;()[]{}") for word in value.split() if word.strip()
        ]
        words = {word.casefold() for word in raw_words}
        if words & service_words:
            return False
        return all(
            word.casefold() in {"уулу", "кызы"} or word[:1].isupper()
            for word in raw_words
        )

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
