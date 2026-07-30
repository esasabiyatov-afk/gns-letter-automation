from __future__ import annotations

import re

import pymorphy3


class NameService:
    KYRGYZ_CONNECTORS = {"уулу", "кызы"}
    RUSSIAN_WORD = re.compile(r"[А-Яа-яЁё]+(?:-[А-Яа-яЁё]+)*")
    KYRGYZ_LETTERS = frozenset("ҢңӨөҮү")

    def __init__(self):
        self.morph = pymorphy3.MorphAnalyzer()

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

    def position_display(self, position: str) -> str:
        """Возвращает консервативную форму должности в дательном падеже."""
        normalized = " ".join(position.split())
        if not normalized or any(
            letter in normalized for letter in self.KYRGYZ_LETTERS
        ):
            return normalized

        expanded = self._expand_position_prefix(normalized)
        if expanded is not None:
            return expanded

        matches = list(self.RUSSIAN_WORD.finditer(normalized))
        if not matches:
            return normalized

        replacements: dict[tuple[int, int], str] = {}
        head_found = False
        for index, match in enumerate(matches):
            word = match.group(0)
            parses = self.morph.parse(word)
            top = parses[0]
            if "datv" in top.tag and top.tag.POS in {
                "NOUN",
                "ADJF",
                "PRTF",
            }:
                return normalized

            nominative_head_ahead = self._has_nominative_head_ahead(
                matches,
                index,
            )

            parse = top
            if nominative_head_ahead:
                modifier = next(
                    (
                        candidate
                        for candidate in parses
                        if candidate.tag.POS in {"ADJF", "PRTF"}
                        and "nomn" in candidate.tag
                    ),
                    None,
                )
                if modifier:
                    parse = modifier

            if parse.tag.POS in {"ADJF", "PRTF"} and "nomn" in parse.tag:
                inflected = parse.inflect({"datv"})
                if not inflected:
                    return normalized
                replacements[match.span()] = self._restore_case(
                    word,
                    inflected.word,
                )
                if not nominative_head_ahead:
                    head_found = True
                    break
                continue

            if parse.tag.POS == "NOUN" and "nomn" in parse.tag:
                inflected = parse.inflect({"datv"})
                if not inflected:
                    return normalized
                replacements[match.span()] = self._restore_case(
                    word,
                    inflected.word,
                )
                head_found = True
                break

        if not head_found:
            return normalized

        parts: list[str] = []
        cursor = 0
        for match in matches:
            if match.span() not in replacements:
                continue
            parts.append(normalized[cursor : match.start()])
            parts.append(replacements[match.span()])
            cursor = match.end()
        parts.append(normalized[cursor:])
        return "".join(parts)

    def _has_nominative_head_ahead(
        self,
        matches: list[re.Match[str]],
        current_index: int,
    ) -> bool:
        for match in matches[current_index + 1 :]:
            parses = self.morph.parse(match.group(0))
            top = parses[0]
            if top.tag.POS == "NOUN" and "nomn" in top.tag:
                return True
            modifier = any(
                candidate.tag.POS in {"ADJF", "PRTF"}
                and "nomn" in candidate.tag
                for candidate in parses
            )
            if modifier:
                continue
            return False
        return False

    @staticmethod
    def _expand_position_prefix(position: str) -> str | None:
        patterns = (
            (
                re.compile(r"^зам\.\s*", re.IGNORECASE),
                "Заместителю",
            ),
            (
                re.compile(
                    r"^(?:вр\.\s*)?и\.\s*о\.\s*",
                    re.IGNORECASE,
                ),
                "Исполняющему обязанности",
            ),
            (
                re.compile(r"^врио\s+", re.IGNORECASE),
                "Временно исполняющему обязанности ",
            ),
        )
        for pattern, replacement in patterns:
            match = pattern.match(position)
            if not match:
                continue
            if position[0].islower():
                replacement = replacement[0].lower() + replacement[1:]
            rest = position[match.end() :]
            separator = "" if replacement.endswith(" ") else " "
            return f"{replacement}{separator}{rest}".strip()
        return None

    @staticmethod
    def _restore_case(source: str, value: str) -> str:
        if source.isupper():
            return value.upper()
        if source[:1].isupper():
            return value[:1].upper() + value[1:]
        return value

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
