# START_MODULE_CONTRACT
# PURPOSE: Пакет низкоуровневых формат-парсеров. Каждый формат живёт в отдельном модуле, а здесь — только публичный API.
# INPUTS: Импорт-время: подмодули docx и pptx.
# OUTPUTS: Публичные функции `parse_docx`, `parse_pptx` и служебные имена, которые ещё ждёт текущий bridge в main.py.
# KEYWORDS: package, format-parser, docx, pptx, public-api, weak-coupling.
# LINKS: src/format_parsers/docx.py, src/format_parsers/pptx.py, main.py.
# RATIONALE: Разнесение по файлам делает пакет самоописательным: один формат — один модуль. `__init__.py` хранит только публичную поверхность, чтобы потребителям не нужно было знать внутреннюю структуру пакета.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_REEXPORTS
# Публичная поверхность пакета для текущих потребителей (main.py bridge, тесты).
from src.format_parsers._types import ParsedDocument
from src.format_parsers.docx import (
    docx_parser__table_to_html,
    docx_parser_logger,
    parse_docx,
    _extract_elements_in_order,
    _ocr_header_images,
)
from src.format_parsers.pptx import (
    parse_pptx,
    pptx_parser__table_to_html,
    pptx_parser_logger,
    _extract_slide_text,
    _textframe_to_text,
)

__all__ = [
    "ParsedDocument",
    "parse_docx",
    "parse_pptx",
    "docx_parser_logger",
    "pptx_parser_logger",
    "docx_parser__table_to_html",
    "pptx_parser__table_to_html",
    "_extract_elements_in_order",
    "_ocr_header_images",
    "_extract_slide_text",
    "_textframe_to_text",
]
# END_REEXPORTS
