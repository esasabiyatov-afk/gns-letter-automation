from __future__ import annotations

import os
import re
import zipfile
from copy import deepcopy
from datetime import date
from pathlib import Path
from uuid import uuid4

from docx import Document
from docx.text.paragraph import Paragraph
from lxml import etree

from gns_app.services.name_service import NameService


class WordTemplateError(ValueError):
    pass


MONTHS_RU = (
    "",
    "января",
    "февраля",
    "марта",
    "апреля",
    "мая",
    "июня",
    "июля",
    "августа",
    "сентября",
    "октября",
    "ноября",
    "декабря",
)


class WordTemplateService:
    TOKENS = {
        "today": "[Дата.Сегодня]",
        "district": "[Район.Место]",
        "position": "[Должность.Отправитель]",
        "recipient": "[ФИО. Отправитель]",
        "subjects": "[Перечисления.Субьект]",
        "inns": "[ИНН.Субьект]",
        "employee": "[ФИО.Исп]",
    }

    def __init__(
        self,
        templates_dir: Path,
        name_service: NameService | None = None,
    ):
        self.templates_dir = templates_dir
        self.names = name_service or NameService()

    def render(
        self,
        output_path: Path,
        case: dict[str, str],
        taxpayers: list[dict[str, str]],
    ) -> Path:
        self._validate(case, taxpayers)
        template_name = (
            "шаблон ответа одиночный.docx"
            if len(taxpayers) == 1
            else "шаблон ответа много.docx"
        )
        template_path = self.templates_dir / template_name
        if not template_path.exists():
            raise WordTemplateError(f"Не найден шаблон: {template_name}")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        working_path = output_path.with_suffix(".working.docx")
        document = Document(str(template_path))

        replacements = {
            self.TOKENS["today"]: self._format_date(date.today()),
            self.TOKENS["district"]: case["district_place"].strip(),
            self.TOKENS["position"]: self.names.position_display(
                case["recipient_position"]
            ),
            self.TOKENS["recipient"]: case["recipient_display_name"].strip(),
            self.TOKENS["employee"]: case["employee_name"].strip(),
        }

        if len(taxpayers) == 1:
            taxpayer = taxpayers[0]
            replacements[self.TOKENS["subjects"]] = (
                f" {taxpayer['name'].strip()}"
            )
            replacements[self.TOKENS["inns"]] = (
                f"ИНН: {taxpayer['inn'].strip()} "
            )
        else:
            self._replace_multi_taxpayer_paragraph(document, taxpayers)
            replacements[self.TOKENS["subjects"]] = self._plain_taxpayer_list(
                taxpayers
            )
            replacements[self.TOKENS["inns"]] = ""

        for paragraph in self._iter_visible_paragraphs(document):
            for token, value in replacements.items():
                self._replace_across_runs(paragraph, token, value)

        document.save(working_path)
        self._patch_all_xml(working_path, replacements)
        self._assert_no_tokens(working_path)
        os.replace(working_path, output_path)
        return output_path

    @staticmethod
    def _validate(
        case: dict[str, str], taxpayers: list[dict[str, str]]
    ) -> None:
        required = (
            "district_place",
            "recipient_position",
            "recipient_display_name",
            "employee_name",
        )
        missing = [field for field in required if not case.get(field)]
        if missing:
            raise WordTemplateError(
                "Не заполнены обязательные поля: " + ", ".join(missing)
            )
        if not taxpayers:
            raise WordTemplateError("Нет подтверждённых налогоплательщиков")
        for taxpayer in taxpayers:
            inn = re.sub(r"\D", "", taxpayer.get("inn", ""))
            if len(inn) != 14 or not taxpayer.get("name", "").strip():
                raise WordTemplateError(
                    "У каждого налогоплательщика нужны наименование и 14-значный ИНН"
                )

    @staticmethod
    def _format_date(value: date) -> str:
        return f"«{value.day:02d}» {MONTHS_RU[value.month]} {value.year} г."

    def _replace_multi_taxpayer_paragraph(
        self,
        document: Document,
        taxpayers: list[dict[str, str]],
    ) -> None:
        target = next(
            (
                paragraph
                for paragraph in document.paragraphs
                if self.TOKENS["subjects"] in paragraph.text
                and self.TOKENS["inns"] in paragraph.text
            ),
            None,
        )
        if target is None:
            raise WordTemplateError(
                "В множественном шаблоне не найдена строка налогоплательщиков"
            )

        reference_properties = None
        if target.runs and target.runs[0]._r.rPr is not None:
            reference_properties = deepcopy(target.runs[0]._r.rPr)

        target.clear()
        for index, taxpayer in enumerate(taxpayers, 1):
            if index > 1:
                target.add_run().add_break()
            run = target.add_run(
                f"{index}. {taxpayer['name'].strip()} "
                f"ИНН: {taxpayer['inn'].strip()};"
            )
            if reference_properties is not None:
                run._r.insert(0, deepcopy(reference_properties))

    @staticmethod
    def _plain_taxpayer_list(taxpayers: list[dict[str, str]]) -> str:
        return "; ".join(
            f"{index}. {item['name'].strip()} ИНН: {item['inn'].strip()}"
            for index, item in enumerate(taxpayers, 1)
        )

    @staticmethod
    def _iter_visible_paragraphs(document: Document):
        yield from document.paragraphs
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from cell.paragraphs
        for section in document.sections:
            for part in (section.header, section.footer):
                yield from part.paragraphs
                for table in part.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            yield from cell.paragraphs

    @staticmethod
    def _replace_across_runs(
        paragraph: Paragraph,
        token: str,
        replacement: str,
    ) -> None:
        while token in "".join(run.text for run in paragraph.runs):
            full_text = "".join(run.text for run in paragraph.runs)
            start = full_text.index(token)
            end = start + len(token)
            cursor = 0
            inserted = False
            for run in paragraph.runs:
                run_start = cursor
                run_end = cursor + len(run.text)
                cursor = run_end
                if run_end <= start or run_start >= end:
                    continue

                local_start = max(0, start - run_start)
                local_end = min(len(run.text), end - run_start)
                prefix = run.text[:local_start]
                suffix = run.text[local_end:]
                if not inserted:
                    run.text = prefix + replacement + suffix
                    inserted = True
                else:
                    run.text = prefix + suffix

    def _patch_all_xml(
        self, docx_path: Path, replacements: dict[str, str]
    ) -> None:
        temporary = docx_path.with_name(
            f"{docx_path.stem}.{uuid4().hex}.zip"
        )
        namespace = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        }

        with zipfile.ZipFile(docx_path, "r") as source, zipfile.ZipFile(
            temporary, "w", zipfile.ZIP_DEFLATED
        ) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(
                    ".xml"
                ):
                    try:
                        root = etree.fromstring(data)
                        for paragraph in root.xpath(
                            ".//w:p", namespaces=namespace
                        ):
                            text_nodes = paragraph.xpath(
                                ".//w:t", namespaces=namespace
                            )
                            for token, value in replacements.items():
                                self._replace_xml_nodes(
                                    text_nodes, token, value
                                )
                        data = etree.tostring(
                            root,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )
                    except etree.XMLSyntaxError:
                        pass
                target.writestr(item, data)

        os.replace(temporary, docx_path)

    @staticmethod
    def _replace_xml_nodes(
        nodes: list[etree._Element],
        token: str,
        replacement: str,
    ) -> None:
        while token in "".join(node.text or "" for node in nodes):
            full = "".join(node.text or "" for node in nodes)
            start = full.index(token)
            end = start + len(token)
            cursor = 0
            inserted = False
            for node in nodes:
                value = node.text or ""
                node_start = cursor
                node_end = cursor + len(value)
                cursor = node_end
                if node_end <= start or node_start >= end:
                    continue
                local_start = max(0, start - node_start)
                local_end = min(len(value), end - node_start)
                prefix = value[:local_start]
                suffix = value[local_end:]
                if not inserted:
                    node.text = prefix + replacement + suffix
                    inserted = True
                else:
                    node.text = prefix + suffix

    def _assert_no_tokens(self, docx_path: Path) -> None:
        with zipfile.ZipFile(docx_path) as archive:
            remaining: set[str] = set()
            for name in archive.namelist():
                if not (name.startswith("word/") and name.endswith(".xml")):
                    continue
                text = archive.read(name).decode("utf-8", errors="ignore")
                for token in self.TOKENS.values():
                    if token in text:
                        remaining.add(token)
        if remaining:
            raise WordTemplateError(
                "В ответе остались незаполненные коды: "
                + ", ".join(sorted(remaining))
            )
