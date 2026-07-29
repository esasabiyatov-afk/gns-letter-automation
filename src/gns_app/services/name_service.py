from __future__ import annotations

import re


class NameService:
    KYRGYZ_CONNECTORS = {"уулу", "кызы"}

    def recipient_display(self, full_name: str) -> str:
        words = [item for item in re.split(r"\s+", full_name.strip()) if item]
        if not words:
            return ""

        lower = [item.casefold() for item in words]
        for connector in self.KYRGYZ_CONNECTORS:
            if connector in lower:
                index = lower.index(connector)
                prefix = words[: index + 1]
                suffix = words[index + 1 :]
                if suffix:
                    return " ".join(prefix + [self._initial(suffix[0])])
                return " ".join(prefix)

        surname = words[0]
        given = words[1] if len(words) > 1 else ""
        patronymic = words[2] if len(words) > 2 else ""
        gender = self._gender(patronymic)
        declined = self._decline_surname(surname, gender)
        initials = " ".join(
            self._initial(item) for item in (given, patronymic) if item
        )
        return f"{declined} {initials}".strip()

    @staticmethod
    def _initial(value: str) -> str:
        return f"{value[0].upper()}." if value else ""

    @staticmethod
    def _gender(patronymic: str) -> str | None:
        normalized = patronymic.casefold()
        if normalized.endswith(("ович", "евич", "ич")):
            return "male"
        if normalized.endswith(("овна", "евна", "ична")):
            return "female"
        return None

    @staticmethod
    def _decline_surname(surname: str, gender: str | None) -> str:
        lower = surname.casefold()
        if gender == "male":
            if lower.endswith(("ов", "ев", "ёв", "ин", "ын")):
                return surname + "у"
            if lower.endswith("ский"):
                return surname[:-4] + "скому"
            if lower.endswith("цкий"):
                return surname[:-4] + "цкому"
        if gender == "female":
            for ending in ("ова", "ева", "ёва", "ина", "ына"):
                if lower.endswith(ending):
                    return surname[:-1] + "ой"
            if lower.endswith("ская"):
                return surname[:-4] + "ской"
            if lower.endswith("цкая"):
                return surname[:-4] + "цкой"
        return surname

