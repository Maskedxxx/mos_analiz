# START_MODULE_CONTRACT
# PURPOSE: Типы-контракты формат-парсеров. Единая форма возвращаемого словаря для docx/pptx/... .
# INPUTS: —
# OUTPUTS: `ParsedDocument` (TypedDict) — словарь с полями `filename`, `path`, `raw_text`.
# KEYWORDS: typeddict, contract, parsed-document.
# LINKS: src/format_parsers/docx.py, src/format_parsers/pptx.py, src/format_parsers/__init__.py.
# RATIONALE: Явный типизированный контракт фиксирует форму результата парсеров без рантайм-оверхеда: сохраняет dict-API, добавляет IDE-автодополнение и возможность статического type-check.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
from typing import TypedDict
# END_IMPORTS


# START_PARSED_DOCUMENT
# PURPOSE: Общий контракт возвращаемого значения формат-парсеров.
# INPUTS: —
# OUTPUTS: TypedDict `ParsedDocument`.
# KEYWORDS: parsed-document, filename, path, raw-text.
class ParsedDocument(TypedDict):
    """
    Назначение:
        Контракт возвращаемого значения формат-парсеров (`parse_docx`, `parse_pptx`, ...).

    Поля:
        filename: Базовое имя исходного файла (без директории).
        path: Абсолютный путь к исходному файлу.
        raw_text: Полный текст документа как единая строка, без разбивки по страницам
            или чанкам.

    Логика:
        Формат-парсер читает файл своего типа и отдаёт строго эту форму.
        Смысловая нарезка документа происходит в верхних слоях конвейера.
    """

    filename: str
    path: str
    raw_text: str
# END_PARSED_DOCUMENT
