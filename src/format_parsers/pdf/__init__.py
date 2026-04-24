# START_MODULE_CONTRACT
# PURPOSE: Публичный API пакета PDF-парсера. Внутренности (клиенты, парсинг, оркестратор) — package-private. Конфиг-классы импортируются из `config.parsers` (единый источник правды).
# INPUTS: Импорт-время — подмодули пакета + config.parsers.
# OUTPUTS: `parse_pdf`, клиенты, чистая логика, типы-контракты конфига.
# KEYWORDS: package, pdf-parser, public-api, paddle-internal.
# LINKS: config/parsers.py, src/format_parsers/pdf/_clients.py, src/format_parsers/pdf/_parsing.py, src/format_parsers/pdf/parse.py.
# RATIONALE: Снаружи пакет воспринимается как «PDF-парсер» независимо от внутренней реализации (сейчас Paddle).
# END_MODULE_CONTRACT

from __future__ import annotations

# START_REEXPORTS
# Типы конфига — из единого config.parsers (там же дефолтные значения).
from config.parsers import (
    PdfExtractorConfig,
    PdfLayoutConfig,
    PdfParserConfig,
    PdfParsingConfig,
    PdfVlmConfig,
)
from src.format_parsers.pdf._clients import (
    LayoutDetector,
    RemoteLayoutDetector,
    VLMClient,
)
from src.format_parsers.pdf._parsing import (
    Cell,
    build_page_markdown,
    is_table_format,
    merge_text_and_tables,
    normalize_paddle_text,
    parse_paddle_table_html,
    parse_paddle_table_markdown,
    sort_by_reading_order,
)
from src.format_parsers.pdf.parse import PaddleExtractor, parse_pdf

__all__ = [
    # Публичный entry
    "parse_pdf",
    # Конфиг
    "PdfExtractorConfig",
    "PdfLayoutConfig",
    "PdfParserConfig",
    "PdfParsingConfig",
    "PdfVlmConfig",
    # Клиенты (IO)
    "LayoutDetector",
    "RemoteLayoutDetector",
    "VLMClient",
    # Чистая логика
    "Cell",
    "build_page_markdown",
    "is_table_format",
    "merge_text_and_tables",
    "normalize_paddle_text",
    "parse_paddle_table_html",
    "parse_paddle_table_markdown",
    "sort_by_reading_order",
    # Оркестратор (внутренний, но export для тестов/отладки)
    "PaddleExtractor",
]
# END_REEXPORTS
